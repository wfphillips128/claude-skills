# Verifying VBA in an isolated Excel

Reading code and believing it works is not verification. Excel's behaviour
often cannot be deduced from the source. For example, `Range.HasFormula`
returns `Null` (not True or False) on a mixed range, so
`If rng.HasFormula Then` silently does nothing. Only running the code shows
that. Before any handover, **compile and run the new code in a throwaway
Excel**, driven from Python. `scripts/verify_vba.py` does this. This page
explains how it works and how to write tests for it.

## 1. Quick use

```
python scripts/verify_vba.py MyAddIn.xlsm
python scripts/verify_vba.py MyAddIn.xlsm --import build/modReports.bas build/modTests.bas `
       --form "frmSettings=build/frmSettings (paste into form module).txt" `
       --document "ThisWorkbook=build/ThisWorkbook (paste).txt" `
       --run modTests.TestReports modTests.TestSettings `
       --ribbon Ribbon_ShowAbout `
       --xml "build/ribbon - reports group.xml"
```

| Option | Does |
|---|---|
| `--import F.bas ...` | Removes the module named in each file's `VB_Name`, then imports the file. FAILs if it lands as `Name1`. |
| `--form NAME=F.txt` | Pastes code into a UserForm, creating a blank one if needed. For forms that build their own controls. |
| `--document NAME=F.txt` | Pastes code into `ThisWorkbook` or a sheet module, by code name. |
| `--run MACRO ...` | Runs each with `Application.Run`. A String result starting with `FAIL` counts as a failure. |
| `--ribbon CALLBACK ...` | Runs a `(control As IRibbonControl)` callback through a generated shim. |
| `--xml FILE` | Also checks the `imageMso` names in a snippet you have not pasted yet. |
| `--timeout N` | Seconds per step before that Excel is stopped (default 90). |

Requirements: Windows, desktop Excel, `pywin32`, and *File > Options > Trust
Center > Trust Center Settings > Macro Settings > **Trust access to the VBA
project object model***. That setting lets a program read and change VBA
code. Explain to the user what it allows before asking them to tick it, and
that they can untick it afterwards.

## 2. What the script does, and the trap behind each step

**A separate Excel, closed by its own process id.**
`win32com.client.DispatchEx("Excel.Application")` always starts a **new**
Excel process, unlike `Dispatch`, which may attach to the one the user is
working in. The script reads that process's id from its window handle
(`win32process.GetWindowThreadProcessId(xl.Hwnd)`), which is exact even if
another program starts an Excel at the same moment. Comparing "Excel
processes before and after" is not exact. The script only ever ends that
one process. **Never kill Excel by name.** The user's unsaved work, or
another tool's session, may be in a different Excel process.

**Release every COM reference before Quit.** Excel stays alive while Python
holds any object from it (a workbook, a range, a module). `Quit()` then
appears to do nothing. The script drops its references, waits up to 20
seconds for that process to exit, and only then terminates it by handle.

**A copy with a new name.** The workbook is copied to a temp folder as
`<name> VERIFY.xlsm`, so the original is never opened or locked. The new name
matters too. Excel refuses to open two workbooks with the same file name, and
an add-in ticked in *File > Options > Add-ins* may load into the new instance
under its own name. A same-named copy then fails to open. Address macros
through the copy's name: `xl.Run("'MyAddIn VERIFY.xlsm'!modTests.RunAll")`.

**Late binding, to survive a damaged gen_py cache.** pywin32 generates
wrapper code for Excel into a shared `gen_py` cache. If a run is interrupted,
that cache can be left corrupt, and every `Dispatch` then fails with
`module ... has no attribute 'CLSIDToClassMap'`. Deleting a shared cache
affects other tools, so the script bypasses it instead:

```python
import win32com.client
win32com.client.gencache.GetClassForCLSID = lambda clsid: None   # always late-bound
```

**A full compile, which a probe alone does not give you.** VBA's default
*Compile On Demand* compiles a module only when something in it is first
called. Tested: with a compile error in one module, a probe function in
another module ran fine. So the script runs *Debug > Compile VBAProject*
through the VBA editor's command bar (control id 578). A clean compile greys
that command out (`Enabled = False`). A failed compile leaves it enabled and
raises the editor's "Compile error" dialog. The script reads the dialog's
text and the module, procedure and line under the cursor, then dismisses it.
Example output:

```
[FAIL] project compiles (Debug > Compile)  Compile error: Variable not defined in modBroken.NeverCalled (near line 4 of the module)
```

The probe function still runs afterwards, to prove the project is executable.

**Dialogs are read and dismissed, never left to block.** In an invisible
Excel, any `MsgBox`, `InputBox` or run-time error dialog waits for a click
that will never come, and the run hangs forever. While each step runs, a
watcher thread looks for dialog windows owned by **that Excel's process**,
records their text, and presses *End* (run-time error) or *Cancel*, *No* or
*OK*. The step is then reported as a FAIL with the dialog text:

```
[FAIL] ribbon callback Ribbon_ReloadSettings (via shim)  dialog: My Add-in: Settings reloaded from the add-in's storage sheet.
```

This is a safety net. In tests, call the plain procedure that does the work,
not the UI wrapper that shows a message. Where a confirmation cannot be
avoided, add a module-level `Public TestMode As Boolean` that makes the
confirm function return True without asking. If a step still overruns
`--timeout`, that Excel is terminated and the step FAILs.

**Ribbon callbacks through shims.** `Application.Run` cannot create an
`IRibbonControl`, so a callback cannot be run directly. The script adds a
temporary module (`zzVerifyProbe`) with one shim per `--ribbon` name:

```vba
Public Sub zzShim_Ribbon_ShowAbout()
    Dim c As IRibbonControl          ' Nothing
    Ribbon_ShowAbout c
End Sub
```

This works for any callback that does not use `control` (thin callbacks don't).
A callback that reads `control.ID` gets error 91 here. Test the procedure it
hands off to instead.

**Icons checked against the real Office.** `imageMso` names are
case-sensitive, and an unknown one is not an error, just a blank icon. The
script asks Excel for every name in the ribbon XML (and any `--xml` file)
with `CommandBars.GetImageMso`. Across processes, a real icon cannot be
returned as a picture, so Excel answers `E_UNEXPECTED` ("Catastrophic
failure"). A name that does not exist gets `E_INVALIDARG`. That difference is
the test.

## 3. What it cannot verify

- **The ribbon appearing.** A hidden Excel does not load ribbon UI. In
  testing, `onLoad` never fired there, so `get*` callbacks and invalidation
  cannot be exercised. `review_addin.py` covers the XML statically. The
  handover must ask the user to open the workbook and click the buttons.
- **Forms on screen.** Showing a modal form in a hidden Excel blocks the run.
  Compiling proves the form's code is valid. Test the form's logic by keeping
  it in procedures the form calls, and leave layout to the user.
- **Behaviour that needs a visible window**: `ActiveWindow`, `Selection` in a
  specific window, screen positions.
- **`Workbook.Activate` is unreliable in an invisible instance.** It does not
  always move `ActiveWorkbook`. In tests, create the target workbook
  (`Workbooks.Add`, which becomes active) or pass ranges explicitly. Assert on
  `ActiveWorkbook.Name` rather than assuming a switch happened.

## 4. Writing a test module

Put tests in their own module (`modTests.bas`), import it with
`--import`, and do **not** ship it in the add-in.

```vba
Attribute VB_Name = "modTests"
Option Explicit

' Each test returns "PASS <what was proven>" or "FAIL <what went wrong>".
' No MsgBox, no InputBox, no form .Show.

Public Function TestRowHeightSetting() As String
    If Not AddInStorage.SetSetting("DEFAULT_ROW_HEIGHT", 20) Then
        TestRowHeightSetting = "FAIL write: " & AddInStorage.LastError: Exit Function
    End If
    If DEFAULT_ROW_HEIGHT <> 20 Then TestRowHeightSetting = "FAIL read-back " & DEFAULT_ROW_HEIGHT: Exit Function
    TestRowHeightSetting = "PASS setting round-trips"
End Function

Public Function TestInsertIntoNewBook() As String
    Dim wb As Workbook
    Set wb = Workbooks.Add                        ' a target workbook; never the user's
    modRibbon.InsertNextItem wb.Worksheets(1).Range("A1")
    TestInsertIntoNewBook = IIf(Len(wb.Worksheets(1).Range("A1").Value) > 0, "PASS", "FAIL A1 empty")
    wb.Close False
End Function
```

Good tests:
- **Assert the claim, not just "no error".** Read the value back and compare.
- **Include the case that should be refused** (a bad value, a table in the
  way) and assert it was refused and nothing changed.
- **Test against a disposable target** (`Workbooks.Add`), never a file the user
  owns.
- **When behaviour is uncertain, write the test as an experiment first.** Five
  one-line variations in a throwaway harness settle in one run what an hour
  of reasoning will not. Example: does `ListRows.Add` fail with a table below?
  Only when that table spans different columns. The test proved it, and the
  first guess was wrong.

Tests that call `AddInStorage` must run inside the add-in's own project, which
`verify_vba.py` does by importing into the copy. One VBA project cannot see
another's modules without a reference.

## 5. Reading the result

```
RESULT: PASS  (17 passed, 0 failed)
Throwaway Excel (PID 1012) has exited.
```

Paste the table into the handover. If the last line ever says the throwaway
Excel is **still running**, end that PID (and only that one) in Task Manager.

Checking from PowerShell that nothing was left behind, before and after a
run:

```powershell
Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object Id, StartTime
```

Any process listed after the run that was not there before is a leak, unless
another program started it. Compare start times, and never end a process you
did not start.
