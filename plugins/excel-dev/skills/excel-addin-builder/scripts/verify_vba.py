"""
verify_vba.py - prove VBA compiles and runs, in a THROWAWAY Excel, before you
hand anything over.

What it does, in order:
  1. Copies the workbook / add-in to a temp folder under a NEW name. Excel
     refuses two open workbooks with the same file name, and an add-in that is
     ticked in File > Options > Add-ins can load into the new Excel too, so a
     copy with the original name may fail to open.
  2. Starts a SEPARATE, invisible Excel (DispatchEx) and records its process
     id from its window handle. Only that process is ever closed or killed;
     any Excel you (or anything else) have open is never touched.
  3. Imports the .bas / .cls files you name over the copy's modules (removing
     the old module first - VBA's Import adds "Module1"-style copies instead
     of replacing), and pastes form / document-module code from .txt files.
  4. Forces a FULL compile with the VBA editor's Debug > Compile command.
     Running a macro does NOT do this: with "Compile On Demand" (the default)
     a module with a compile error is only compiled when something in it is
     called, so a probe can pass while another module is broken.
     If compile fails, the error text and the offending module and line are
     reported.
  5. Runs a probe function, then any macros you name (--run) and ribbon
     callbacks you name (--ribbon, through a generated shim, because
     Application.Run cannot supply an IRibbonControl).
  6. Checks every imageMso in the ribbon XML exists (an unknown name is not
     an error in Excel - the icon is just blank).
  7. Closes the copy without saving, quits that Excel, and deletes the copy.

Any dialog box that appears while a step runs (a MsgBox, a run-time error) is
read, reported as a FAIL, and dismissed - in an invisible Excel nobody could
click it and the run would hang forever. Stub confirmations in test code
anyway; this is the safety net, not the plan.

A --run macro that returns a string starting with "FAIL" is reported as a
failure, so a test function can return "PASS ..." / "FAIL ..." text.

Usage:
  python verify_vba.py BOOK.xlsm
  python verify_vba.py BOOK.xlam --import build\\modRibbon.bas build\\AddInStorage.bas
  python verify_vba.py BOOK.xlsm --form frmSettings=frmSettings.txt \\
         --document ThisWorkbook=ThisWorkbook.txt --run modTests.RunAll --ribbon Ribbon_ShowAbout

Exit code: 0 = all passed, 1 = something failed, 2 = could not run.
Needs: Windows, Excel, pywin32, and Excel's File > Options > Trust Center >
Trust Center Settings > Macro Settings > "Trust access to the VBA project
object model" ticked.
"""
import argparse
import os
import re
import shutil
import sys
import tempfile
import threading
import time
import zipfile

try:
    import pythoncom
    import pywintypes
    import win32api
    import win32con
    import win32event
    import win32gui
    import win32process
    import win32com.client
except ImportError:                                   # pragma: no cover
    print("verify_vba.py needs pywin32:  python -m pip install pywin32")
    sys.exit(2)

# The shared gen_py cache (early-binding wrappers pywin32 generates for Excel)
# can be left corrupt by an interrupted run, and then every Dispatch fails with
# "has no attribute 'CLSIDToClassMap'". Deleting a shared cache affects other
# tools, so bypass it instead: everything below is late-bound.
win32com.client.gencache.GetClassForCLSID = lambda clsid: None

VBEXT_STD, VBEXT_CLASS, VBEXT_FORM, VBEXT_DOCUMENT = 1, 2, 3, 100
COMPILE_CONTROL_ID = 578                               # Debug > Compile VBAProject
PROBE_MODULE = "zzVerifyProbe"
RPC_BUSY = (-2147418111, -2147417846)                  # "call was rejected" while Excel is busy
E_UNEXPECTED = -2147418113
E_INVALIDARG = -2147024809

ROWS = []


def row(status, check, detail=""):
    ROWS.append((status, check, detail))
    print(f"  [{status}] {check}" + (f"  {detail}" if detail else ""), flush=True)


def retry(fn, tries=120, wait=0.25):
    """Excel rejects calls while it is busy; that is not an error, so retry."""
    for _ in range(tries):
        try:
            return fn()
        except pywintypes.com_error as e:
            if e.args[0] not in RPC_BUSY:
                raise
            time.sleep(wait)
    raise RuntimeError("Excel stayed busy")


def com_message(e):
    """The human part of a pywintypes.com_error."""
    try:
        exc = e.args[2]
        if exc and exc[2]:
            return str(exc[2]).strip()
    except Exception:
        pass
    return str(e)


# ----------------------------------------------------------------- dialogs
def dialogs_of(pid):
    """Visible standard dialog windows (#32770) owned by our Excel only."""
    found = []

    def cb(h, _):
        try:
            if (win32gui.IsWindowVisible(h) and win32gui.GetClassName(h) == "#32770"
                    and win32process.GetWindowThreadProcessId(h)[1] == pid):
                found.append(h)
        except win32gui.error:
            pass
        return True
    win32gui.EnumWindows(cb, None)
    return found


def dialog_text(h):
    parts = []

    def cb(ch, _):
        if win32gui.GetClassName(ch) == "Static":
            t = win32gui.GetWindowText(ch).strip()
            if t:
                parts.append(t)
        return True
    try:
        win32gui.EnumChildWindows(h, cb, None)
    except win32gui.error:
        pass
    title = win32gui.GetWindowText(h)
    return (title + ": " if title else "") + " ".join(" ".join(parts).split())


def dismiss(h):
    """Press End (run-time error), else Cancel / No / OK, else close."""
    buttons = {}

    def cb(ch, _):
        if win32gui.GetClassName(ch) == "Button":
            buttons[win32gui.GetWindowText(ch).replace("&", "").strip().lower()] = ch
        return True
    try:
        win32gui.EnumChildWindows(h, cb, None)
        for want in ("end", "cancel", "no", "ok"):
            if want in buttons:
                btn = buttons[want]
                win32gui.PostMessage(h, win32con.WM_COMMAND, win32gui.GetDlgCtrlID(btn), btn)
                return
        win32gui.PostMessage(h, win32con.WM_CLOSE, 0, 0)
    except win32gui.error:
        pass


class Watch:
    """While a step runs: read and dismiss any dialog our Excel raises, and
    kill our Excel (only ours) if the step overruns its time limit."""

    def __init__(self, session, label, timeout):
        self.s, self.label, self.timeout = session, label, timeout
        self.seen, self.killed, self._stop = [], False, threading.Event()

    def __enter__(self):
        self.t = threading.Thread(target=self._loop, daemon=True)
        self.t.start()
        return self

    def _loop(self):
        t0 = time.time()
        while not self._stop.wait(0.4):
            for h in dialogs_of(self.s.pid):
                txt = dialog_text(h)
                if txt and txt not in self.seen:
                    self.seen.append(txt)
                dismiss(h)
            if time.time() - t0 > self.timeout:
                self.killed = True
                self.s.kill()
                return

    def __exit__(self, *exc):
        self._stop.set()
        self.t.join(2)
        return False


# ----------------------------------------------------------------- session
class Session:
    def __init__(self, workbook, timeout):
        self.src, self.timeout = os.path.abspath(workbook), timeout
        self.xl = self.wb = self.vbp = None
        self.pid, self.handle = None, None

    def start(self):
        stem, ext = os.path.splitext(os.path.basename(self.src))
        self.tmp = tempfile.mkdtemp(prefix="verify_vba_")
        self.name = f"{stem} VERIFY{ext}"
        self.path = os.path.join(self.tmp, self.name)
        shutil.copy2(self.src, self.path)

        pythoncom.CoInitialize()
        self.xl = win32com.client.DispatchEx("Excel.Application")
        # The PID from Excel's own window handle - exact, even if another
        # process starts an Excel at the same moment.
        self.pid = win32process.GetWindowThreadProcessId(self.xl.Hwnd)[1]
        self.handle = win32api.OpenProcess(win32con.SYNCHRONIZE | win32con.PROCESS_TERMINATE, False, self.pid)
        self.xl.Visible = False
        self.xl.DisplayAlerts = False
        self.xl.AutomationSecurity = 1              # msoAutomationSecurityLow: macros enabled
        print(f"Throwaway Excel started (PID {self.pid}); working on a copy: {self.name}")

    def open(self):
        with Watch(self, "open", self.timeout) as w:
            self.wb = retry(lambda: self.xl.Workbooks.Open(self.path))
        for t in w.seen:
            row("WARN", "a dialog appeared while opening", t)
        self.vbp = self.wb.VBProject                # raises if VBA project access is not trusted
        _ = self.vbp.VBComponents.Count

    def run(self, macro, *args):
        target = "'" + self.name.replace("'", "''") + "'!" + macro
        return retry(lambda: self.xl.Run(target, *args))

    def kill(self):
        if self.handle:
            try:
                win32api.TerminateProcess(self.handle, 1)
            except pywintypes.error:
                pass

    def close(self):
        try:
            if self.xl is not None:
                for b in list(self.xl.Workbooks):
                    b.Close(False)
                self.xl.Quit()
        except Exception:
            pass
        self.wb = self.vbp = self.xl = None
        if self.handle:
            # Give Excel 20 seconds to exit cleanly, then end it. Only ours.
            if win32event.WaitForSingleObject(self.handle, 20000) != win32event.WAIT_OBJECT_0:
                for h in dialogs_of(self.pid):
                    dismiss(h)
                if win32event.WaitForSingleObject(self.handle, 5000) != win32event.WAIT_OBJECT_0:
                    self.kill()
                    win32event.WaitForSingleObject(self.handle, 5000)
            still = win32event.WaitForSingleObject(self.handle, 0) != win32event.WAIT_OBJECT_0
            print(f"Throwaway Excel (PID {self.pid}) " + ("STILL RUNNING - end it in Task Manager"
                                                          if still else "has exited."))
            win32api.CloseHandle(self.handle)
        pythoncom.CoUninitialize()
        shutil.rmtree(getattr(self, "tmp", ""), ignore_errors=True)


# ----------------------------------------------------------------- steps
def module_name_of(path):
    head = open(path, "rb").read(600).decode("cp1252", errors="replace")
    m = re.search(r'Attribute VB_Name = "([^"]+)"', head)
    return m.group(1) if m else None


def component(vbp, name):
    for c in vbp.VBComponents:
        if c.Name.lower() == name.lower():
            return c
    return None


def import_module(s, path):
    name = module_name_of(path)
    label = f"import {os.path.basename(path)}"
    if not name:
        row("FAIL", label, 'no  Attribute VB_Name = "..."  line - it would import as Module1')
        return
    old = component(s.vbp, name)
    if old is not None:
        if old.Type not in (VBEXT_STD, VBEXT_CLASS):
            row("FAIL", label, f"{name} is a form or document module here; use --form / --document")
            return
        s.vbp.VBComponents.Remove(old)
    s.vbp.VBComponents.Import(os.path.abspath(path))
    if component(s.vbp, name) is None:
        stray = [c.Name for c in s.vbp.VBComponents if c.Name.lower().startswith(name.lower())]
        row("FAIL", label, f"imported as {', '.join(stray) or '?'} instead of {name}")
    else:
        row("PASS", label, f"-> {name}" + (" (replaced)" if old is not None else " (new)"))


def paste_code(s, name, path, kind):
    """--form / --document: replace a code module's text with a .txt file."""
    label = f"paste {os.path.basename(path)} into {name}"
    text = open(path, "rb").read().decode("cp1252", errors="replace").replace("\r\n", "\n").replace("\n", "\r\n")
    comp = component(s.vbp, name)
    if comp is None:
        if kind != "form":
            row("FAIL", label, f"no document module named {name} (use the code name, e.g. ThisWorkbook or Sheet1)")
            return
        comp = s.vbp.VBComponents.Add(VBEXT_FORM)
        comp.Name = name
    elif kind == "form" and comp.Type != VBEXT_FORM:
        row("FAIL", label, f"{name} exists but is not a UserForm")
        return
    cm = comp.CodeModule
    if cm.CountOfLines:
        cm.DeleteLines(1, cm.CountOfLines)       # also removes an auto-inserted Option Explicit
    cm.AddFromString(text)
    row("PASS", label, f"{cm.CountOfLines} lines")


def shim_name(cb):
    return "zzShim_" + re.sub(r"\W", "_", cb)


def add_probe(s, shims):
    lines = ["Option Explicit", "Public Function zzVerifyProbe() As String",
             '    zzVerifyProbe = "probe-ok"', "End Function"]
    for cb in shims:
        lines += [f"Public Sub {shim_name(cb)}()",
                  "    Dim c As IRibbonControl          ' Nothing - fine unless the callback reads it",
                  f"    {cb} c", "End Sub"]
    comp = s.vbp.VBComponents.Add(VBEXT_STD)
    comp.Name = PROBE_MODULE
    cm = comp.CodeModule
    if cm.CountOfLines:
        cm.DeleteLines(1, cm.CountOfLines)
    cm.AddFromString("\r\n".join(lines) + "\r\n")


def compile_project(s):
    """Debug > Compile. A clean compile greys the command out; a failed one
    leaves it enabled and raises the editor's "Compile error" dialog."""
    s.xl.VBE.ActiveVBProject = s.vbp                # other loaded add-ins may be active
    with Watch(s, "compile", s.timeout) as w:
        ctl = s.xl.VBE.CommandBars.FindControl(Type=1, Id=COMPILE_CONTROL_ID)
        ctl.Execute()
        time.sleep(1.0)                             # let the error dialog appear, if any
        compiled = not s.xl.VBE.CommandBars.FindControl(Type=1, Id=COMPILE_CONTROL_ID).Enabled
        where = ""
        if not compiled:
            try:
                pane = s.xl.VBE.ActiveCodePane
                ln = pane.GetSelection()[0]
                mod = pane.CodeModule.Parent.Name
                try:
                    proc = pane.CodeModule.ProcOfLine(ln, 0)
                    if isinstance(proc, tuple):         # late binding returns (name, kind)
                        proc = proc[0]
                except Exception:
                    proc = ""
                where = f" in {mod}" + (f".{proc}" if proc else "") + f" (near line {ln} of the module)"
            except Exception:
                pass
        time.sleep(0.6)                             # one more poll so the dialog text is caught
    msg = "; ".join(t for t in w.seen) or "(no message captured)"
    row("PASS" if compiled else "FAIL", "project compiles (Debug > Compile)",
        "" if compiled else msg + where)
    return compiled


def run_step(s, label, macro, *args):
    with Watch(s, label, s.timeout) as w:
        try:
            result, err = s.run(macro, *args), None
        except pywintypes.com_error as e:
            result, err = None, com_message(e)
        except Exception as e:                      # Excel killed by the watchdog, etc.
            result, err = None, str(e)
    if w.killed:
        row("FAIL", label, f"still running after {s.timeout}s - stopped that Excel")
        return False
    problems = []
    if w.seen:
        problems.append("dialog: " + " | ".join(w.seen))
    if err:
        problems.append(err)
    if isinstance(result, str) and result.upper().startswith("FAIL"):
        problems.append(result)
    shown = "" if result is None else f"returned {result!r}"
    row("FAIL" if problems else "PASS", label, "; ".join(problems) or shown)
    return not problems


def ribbon_xml(path):
    try:
        z = zipfile.ZipFile(path)
    except zipfile.BadZipFile:
        return ""
    for n in z.namelist():
        if re.fullmatch(r"customUI/customUI(14)?\.xml", n, re.I):
            return z.read(n).decode("utf-8", errors="replace")
    return ""


def check_icons(s, xml_text):
    names = sorted(set(re.findall(r'\bimageMso="([^"]+)"', xml_text)))
    if not names:
        return
    # GetImageMso cannot hand a picture to ANOTHER process, so for a real icon
    # it fails with E_UNEXPECTED ("Catastrophic failure"). For a name that does
    # not exist - including a wrong upper/lower case, names are case-sensitive -
    # it fails with E_INVALIDARG instead. That difference is the test.
    unknown, unsure = [], []
    for n in names:
        try:
            s.xl.CommandBars.GetImageMso(n, 16, 16)
        except pywintypes.com_error as e:
            scode = e.args[2][5] if e.args[2] and len(e.args[2]) > 5 else None
            if scode == E_INVALIDARG:
                unknown.append(n)
            elif e.args[0] != E_UNEXPECTED:
                unsure.append(n)
    row("FAIL" if unknown else "PASS", "every imageMso icon exists in this Office",
        f"unknown (would show blank; names are case-sensitive): {', '.join(unknown)}" if unknown
        else f"{len(names)} icons")
    if unsure:
        row("WARN", "imageMso icons that could not be checked", ", ".join(unsure))


# ----------------------------------------------------------------- main
def pair(text):
    if "=" not in text:
        raise argparse.ArgumentTypeError("use NAME=FILE")
    k, v = text.split("=", 1)
    return k.strip(), v.strip()


def main():
    ap = argparse.ArgumentParser(description="Compile and smoke-test VBA in a throwaway Excel.")
    ap.add_argument("workbook", help=".xlsm or .xlam (it is copied; the original is never opened)")
    ap.add_argument("--import", dest="imports", nargs="+", default=[], metavar="FILE",
                    help=".bas / .cls files to import over the copy's modules")
    ap.add_argument("--form", type=pair, action="append", default=[], metavar="NAME=FILE",
                    help="paste a .txt into a UserForm's code module (created if missing)")
    ap.add_argument("--document", type=pair, action="append", default=[], metavar="NAME=FILE",
                    help="paste a .txt into ThisWorkbook or a sheet's module (by code name)")
    ap.add_argument("--run", nargs="+", default=[], metavar="MACRO",
                    help="macros to run after compiling (no arguments); 'Module.Proc' is fine")
    ap.add_argument("--ribbon", nargs="+", default=[], metavar="CALLBACK",
                    help="button callbacks (control As IRibbonControl) to run through a shim")
    ap.add_argument("--xml", action="append", default=[], metavar="FILE",
                    help="also check the imageMso icons in this ribbon XML file or snippet")
    ap.add_argument("--no-icons", action="store_true", help="skip the imageMso check")
    ap.add_argument("--timeout", type=int, default=90, help="seconds allowed per step (default 90)")
    a = ap.parse_args()

    for f in [a.workbook] + a.imports + a.xml + [p for _, p in a.form + a.document]:
        if not os.path.isfile(f):
            print(f"Not found: {f}")
            return 2

    s = Session(a.workbook, a.timeout)
    try:
        s.start()
        try:
            s.open()
        except pywintypes.com_error as e:
            msg = com_message(e)
            if "not trusted" in msg.lower() or "programmatic access" in msg.lower():
                print("\nExcel will not let a program read the VBA project. Tick: File > Options > "
                      "Trust Center > Trust Center Settings > Macro Settings > "
                      "'Trust access to the VBA project object model'.")
            else:
                print(f"\nCould not open the copy: {msg}")
            return 2

        print("\nModules")
        for f in a.imports:
            import_module(s, f)
        for name, f in a.form:
            paste_code(s, name, f, "form")
        for name, f in a.document:
            paste_code(s, name, f, "document")
        add_probe(s, a.ribbon)

        print("\nCompile")
        if compile_project(s):
            print("\nRun")
            run_step(s, "probe function runs", f"{PROBE_MODULE}.zzVerifyProbe")
            for m in a.run:
                run_step(s, f"run {m}", m)
            for cb in a.ribbon:
                run_step(s, f"ribbon callback {cb} (via shim)", f"{PROBE_MODULE}.{shim_name(cb)}")
        else:
            row("INFO", "macros not run", "fix the compile error first")

        if not a.no_icons:
            xml_text = ribbon_xml(a.workbook)
            for x in a.xml:
                xml_text += open(x, encoding="utf-8-sig").read()
            if xml_text:
                print("\nRibbon icons")
                check_icons(s, xml_text)
    except Exception as e:
        row("FAIL", "verification run", f"{type(e).__name__}: {e}")
    finally:
        print()
        s.close()

    fails = sum(1 for st, _, _ in ROWS if st == "FAIL")
    print(f"\n{'=' * 60}\nRESULT: {'FAIL' if fails else 'PASS'}  "
          f"({sum(1 for st, _, _ in ROWS if st == 'PASS')} passed, {fails} failed)")
    return 1 if fails else 0


if __name__ == "__main__":
    sys.exit(main())
