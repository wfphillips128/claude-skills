"""LAMBDA test harness template -- evaluate a LAMBDA in real Excel and prove it.

A Python shadow model proves the MATHS. This proves the FORMULA: it installs the
LAMBDA into a throwaway workbook in a separate, hidden Excel, evaluates it on many
inputs, and compares every cell against an independent Python reference.

What it checks, in order:

  1. Static lint      -- traps Excel reports only as a generic "problem with this
                         formula" or as a silently wrong number (see pitfalls.md)
  2. Install          -- the flattened, comment-free form goes in via Names.Add
  3. Round trip       -- the definition Excel holds equals the text we wrote, both
                         as COM reads it back and as it is STORED in the saved file,
                         and the stored form (prefixes included) is under 8,192
  4. Help             -- calling with no arguments spills the help table
  5. Values           -- fixed and seeded-random cases match the reference
                         within tolerance; invalid inputs return an Excel error
  6. Survives save    -- the saved file reopens and recalculates to the same values

Exit code: 0 all passed, 1 any check failed, 2 the harness itself could not run.

HOW TO ADAPT IT
  Copy this file, then edit the blocks marked  ### EDIT ###:
    EDIT 1  the LAMBDA source, AFE form (or pass --source FILE instead)
    EDIT 2  FUNC -- the name under test
    EDIT 3  reference() -- an independent Python implementation. Write it
            differently from the LAMBDA (a plain loop, not a port of SCAN), so
            one misunderstanding cannot appear identically in both
    EDIT 4  cases() -- fixed hand cases, edge cases, then random ones
  Everything below the EDIT blocks is generic.

SAFETY
  - Starts its OWN Excel with DispatchEx and records that process id. It never
    attaches to, closes or kills any other Excel. On exit it quits its own
    instance, and only if that one process is still alive does it end it.
  - Never opens or writes a user workbook. All files go to a temp folder that is
    deleted afterwards (keep it with --keep-files).
  - Run it with no other heavy Excel automation going at the same time: a busy
    machine turns COM calls into odd transient exceptions (retry() absorbs most).

Requirements: Windows, desktop Excel 365 (LAMBDA + dynamic arrays), pywin32.

Usage:
  python lambda_test_harness_template.py                  # embedded example
  python lambda_test_harness_template.py --source my.txt  # your AFE module file
  python lambda_test_harness_template.py --cases 200 --seed 7 --tol 1e-9
  python lambda_test_harness_template.py --visible        # watch it (debugging)

This file is pure ASCII on purpose: the lambda character and the help-text
delimiters are built with chr(), so no editor or shell can mangle them.
"""

from __future__ import annotations

import argparse
import ctypes
import html
import math
import os
import random
import re
import shutil
import sys
import tempfile
import time
import zipfile

L = chr(0x03BB)      # the lambda suffix on 5G function names
COL = chr(0x2192)    # help-table column delimiter (right arrow)
ROW = chr(0x00B6)    # help-table row delimiter (pilcrow)
# Python's \w already matches the lambda character, so the regexes below need
# no special case for it.

# ================================================================================
# ### EDIT 1 ### The LAMBDA under test, in AFE / Excel Labs form.
#
# Comments and indentation are fine here: the harness strips // and /* */ and
# collapses whitespace outside string literals before installing, exactly as an
# importer must. Several definitions may be given (NAME = LAMBDA(...); each);
# all are installed, so helpers the function calls can live in the same module.
#
# The worked example: a straight-line depreciation schedule, 3 rows x Periods
# columns -- Opening, -Depreciation, Closing. Depreciation stops after Life.
# ================================================================================
DEMO_SOURCE = (
    "/** Straight-line depreciation schedule, one column per period. */\n"
    "Demo.DepreciateSL" + L + " = LAMBDA([Cost], [Salvage], [Life], [Periods],\n"
    "  LET(\n"
    "    // Inline help, shown when the function is called with no arguments\n"
    "    Help, TRIM(TEXTSPLIT(\n"
    "      \"FUNCTION:     " + COL + "Demo.DepreciateSL" + L + "(Cost, Salvage, Life, [Periods])" + ROW + "\" &\n"
    "      \"DESCRIPTION:  " + COL + "Straight-line depreciation, one column per period." + ROW + "\" &\n"
    "      \"VERSION:      " + COL + "Example template" + ROW + "\" &\n"
    "      \"PARAMETERS:   " + COL + ROW + "\" &\n"
    "      \"Cost          " + COL + "(Required) Cost of the asset." + ROW + "\" &\n"
    "      \"Salvage       " + COL + "(Required) Value at the end of its life." + ROW + "\" &\n"
    "      \"Life          " + COL + "(Required) Useful life in whole periods, at least 1." + ROW + "\" &\n"
    "      \"Periods       " + COL + "(Optional) Columns to return. Defaults to Life." + ROW + "\" &\n"
    "      \"EXAMPLE:      " + COL + "=Demo.DepreciateSL" + L + "(1000, 100, 5, 8)\",\n"
    "      \"" + COL + "\", \"" + ROW + "\")),\n"
    "\n"
    "    Cols,    IF(ISOMITTED(Periods), Life, Periods),\n"
    "    Bad,     OR(Life < 1, Life <> INT(Life), Cols < 1, Cols <> INT(Cols),\n"
    "               Salvage > Cost),\n"
    "    Charge,  (Cost - Salvage) / Life,\n"
    "    Per,     SEQUENCE(1, Cols),\n"
    "    Dep,     IF(Per <= Life, Charge, 0),           // nothing after Life\n"
    "    Closing, Cost - SCAN(0, Dep, LAMBDA(Acc, D, Acc + D)),\n"
    "    // DROP(Closing, , -1) alone is EMPTY when Cols = 1 and returns #CALC!;\n"
    "    // stacking first and dropping after keeps at least one column.\n"
    "    Opening, DROP(HSTACK(Cost, Closing), , -1),\n"
    "    Result,  VSTACK(Opening, -Dep, Closing),\n"
    "\n"
    "    Output,  IF(ISOMITTED(Cost), Help, IF(Bad, #NUM!, Result)),\n"
    "    Output\n"
    "  )\n"
    ");\n"
)

# ================================================================================
# ### EDIT 2 ### The function the value tests call.
# ================================================================================
FUNC = "Demo.DepreciateSL" + L


# ================================================================================
# ### EDIT 3 ### Independent Python reference.
#
# Return the expected 2-D result as a list of rows, or raise ValueError for
# inputs the LAMBDA must reject with an Excel error. Deliberately written as a
# plain loop, not as a translation of SCAN/HSTACK -- an independent route.
# ================================================================================
def reference(cost, salvage, life, periods=None):
    cols = life if periods is None else periods
    if life < 1 or life != int(life) or cols < 1 or cols != int(cols) or salvage > cost:
        raise ValueError("invalid input -- the LAMBDA must return an Excel error")
    opening, dep, closing = [], [], []
    balance = float(cost)
    for period in range(1, int(cols) + 1):
        charge = (cost - salvage) / life if period <= life else 0.0
        opening.append(balance)
        dep.append(-charge)
        balance -= charge
        closing.append(balance)
    return [opening, dep, closing]


# ================================================================================
# ### EDIT 4 ### Test cases.
#
# Each case: (label, args). args is the positional argument list; None means
# "omit this argument" (trailing Nones are dropped, inner ones leave an empty
# slot). Lists become array constants {a,b,c}; nested lists become 2-D.
# Whether an Excel error is expected comes from reference() raising ValueError.
#
# Cover, at minimum: a hand-checkable case, both timeline boundaries, the
# optional argument omitted, invalid inputs, and random inputs whose SHAPE
# varies (here: Life and Periods), not just their values.
# ================================================================================
def cases(n_random, seed):
    fixed = [
        ("hand case 1000/100/5 over 8",   [1000, 100, 5, 8]),
        ("Periods omitted -> Life",       [1000, 100, 5]),
        ("Periods = Life (boundary)",     [500, 0, 4, 4]),
        ("Periods < Life (truncated)",    [500, 0, 4, 2]),
        ("one period",                    [90, 10, 1, 1]),
        ("zero salvage, long tail",       [1200, 0, 3, 12]),
        ("salvage = cost (no charge)",    [700, 700, 6, 6]),
        ("ERR: Life = 0",                 [1000, 100, 0, 5]),
        ("ERR: fractional Life",          [1000, 100, 2.5, 5]),
        ("ERR: salvage > cost",           [100, 200, 5, 5]),
        ("ERR: Periods = 0",              [1000, 100, 5, 0]),
    ]
    rng = random.Random(seed)
    rand = []
    for k in range(n_random):
        cost = round(rng.uniform(1, 1_000_000), 2)
        salvage = round(rng.uniform(0, cost), 2)
        life = rng.randint(1, 30)
        periods = None if rng.random() < 0.2 else rng.randint(1, 40)
        rand.append((f"random {k + 1:03d}", [cost, salvage, life, periods]))
    return fixed + rand


# ================================================================================
# Everything below is generic.
# ================================================================================

CAP = 8192                      # Excel's limit on a stored formula, in characters
PREFIXES = ("_xlfn._xlws.", "_xlfn.", "_xlws.", "_xlpm.", "_xlop.")


# ---------------------------------------------------------------- source handling
def _scan(src):
    """Yield (index, char, in_string) for src, honouring "" escapes in strings."""
    instr, i, n = False, 0, len(src)
    while i < n:
        ch = src[i]
        if instr and ch == '"' and i + 1 < n and src[i + 1] == '"':
            yield i, ch, True
            yield i + 1, '"', True
            i += 2
            continue
        if ch == '"':
            instr = not instr
            yield i, ch, True
        else:
            yield i, ch, instr
        i += 1


def strip_comments(src):
    """Remove // and /* */ comments outside string literals. A // left in is
    illegal in an Excel formula -- an importer that forgets this fails."""
    out, i, n, instr = [], 0, len(src), False
    while i < n:
        ch = src[i]
        if instr:
            out.append(ch)
            if ch == '"':
                if i + 1 < n and src[i + 1] == '"':
                    out.append('"')
                    i += 2
                    continue
                instr = False
            i += 1
            continue
        if ch == '"':
            instr = True
            out.append(ch)
            i += 1
        elif src.startswith("//", i):
            while i < n and src[i] not in "\r\n":
                i += 1
        elif src.startswith("/*", i):
            j = src.find("*/", i + 2)
            i = n if j < 0 else j + 2
        else:
            out.append(ch)
            i += 1
    return "".join(out)


def flatten(src):
    """Collapse whitespace runs OUTSIDE string literals to one space. Indentation
    counts against the 8,192 limit if it is written in verbatim; whitespace
    inside a string is help text and must survive untouched."""
    out = []
    pending_space = False
    for _, ch, instr in _scan(src):
        if not instr and ch in " \t\r\n":
            pending_space = True
            continue
        if pending_space and out:
            out.append(" ")
        pending_space = False
        out.append(ch)
    return "".join(out).strip()


def parse_module(src):
    """Split an AFE module into [(name, '=' + flattened formula)]. Definitions are
    NAME = <formula>; with the ';' at paren depth 0 outside strings."""
    code = strip_comments(src)
    entries, depth, start = [], 0, 0
    for i, ch, instr in _scan(code):
        if instr:
            continue
        if ch == "(":
            depth += 1
        elif ch == ")":
            depth -= 1
        elif ch == ";" and depth == 0:
            entries.append(code[start:i])
            start = i + 1
    if code[start:].strip():
        entries.append(code[start:])
    out = []
    for chunk in entries:
        chunk = chunk.strip()
        if not chunk:
            continue
        m = re.match(r"([A-Za-z_\\][\w.]*)\s*=\s*(.*)\Z", chunk, re.S)
        if not m:
            raise SystemExit(f"cannot parse definition starting: {chunk[:60]!r}")
        out.append((m.group(1), "=" + flatten(m.group(2))))
    return out


def outside_strings(text):
    return "".join(ch if not instr else " " for _, ch, instr in _scan(text))


def strip_prefixes(text):
    """Turn Excel's STORED form back into user syntax (outside strings):
    _xlop.Name -> [Name] (an optional LAMBDA parameter, as declared), and the
    other prefixes are simply removed."""
    out, i, n, instr = [], 0, len(text), False
    while i < n:
        ch = text[i]
        if ch == '"':
            instr = not instr
        if not instr:
            if text.startswith("_xlop.", i):
                m = re.match(r"_xlop\.([\w.]+)", text[i:])
                out.append("[" + m.group(1) + "]")
                i += m.end()
                continue
            hit = next((p for p in PREFIXES if text.startswith(p, i)), None)
            if hit:
                i += len(hit)
                continue
        out.append(ch)
        i += 1
    return "".join(out)


def canonical(text):
    """Drop whitespace OUTSIDE string literals. Excel rewrites it on the way in
    (verified: it removes the spaces in a LAMBDA's parameter list, and tightens
    {array constants}), so a byte-exact compare would fail on noise. Every
    token and every string literal must still match exactly.

    Caveat: a space between two range references is Excel's intersection
    operator. A self-contained 5G function has no range references, so this
    is safe there; it is NOT a general-purpose formula comparer."""
    return "".join(ch for _, ch, instr in _scan(text) if instr or ch not in " \t\r\n")


def first_difference(a, b):
    for k, (x, y) in enumerate(zip(a, b)):
        if x != y:
            return k
    return min(len(a), len(b)) if len(a) != len(b) else -1


# ------------------------------------------------------------------- static lint
# Built-in function names. A LET-bound LAMBDA given one of these names is NOT
# called when you write name(...): Excel calls the built-in instead. Verified in
# Excel 365 by experiment: na / step / mid / row are refused outright, while
# t(2) silently returns "" (the T function), sec(2) returns -2.40 (SEC) and
# var(2) returns #DIV/0! (VAR). As plain VALUE names (no call) they are accepted.
# Not exhaustive -- extend it with any function your library shadows.
BUILTIN_FUNCTIONS = set("""
ABS ADDRESS AND AREAS AVERAGE AVERAGEIF AVERAGEIFS BYCOL BYROW CEILING CELL CHAR
CHOOSE CHOOSECOLS CHOOSEROWS CLEAN CODE COLUMN COLUMNS CONCAT COS COSH COT COUNT
COUNTA COUNTIF COUNTIFS CSC DATE DAY DAYS DB DDB DELTA DROP EDATE EOMONTH EVEN EXACT
EXP EXPAND FACT FALSE FILTER FIND FIXED FLOOR FV GCD HLOOKUP HOUR HSTACK IF IFERROR
IFNA IFS INDEX INDIRECT INFO INT IPMT IRR ISBLANK ISERR ISERROR ISNA ISNUMBER ISODD
ISOMITTED ISTEXT LAMBDA LARGE LCM LEFT LEN LET LINEST LN LOG LOOKUP LOWER MAKEARRAY
MAP MATCH MAX MEDIAN MID MIN MINUTE MOD MODE MONTH MROUND N NA NOT NOW NPER NPV ODD
OFFSET OR PI PMT POWER PPMT PRICE PRODUCT PV RAND RANDARRAY RANK RATE REDUCE REPLACE
REPT RIGHT ROUND ROW ROWS SCAN SEARCH SEC SECOND SEQUENCE SIGN SIN SINH SLN SMALL SORT
SORTBY SQRT STDEV STEP SUBSTITUTE SUM SUMIF SUMIFS SUMPRODUCT SWITCH SYD T TAKE TAN
TEXT TEXTJOIN TEXTSPLIT TIME TOCOL TODAY TOROW TRANSPOSE TREND TRIM TRUE TRUNC TYPE
UNIQUE UPPER VALUE VAR VLOOKUP VSTACK WEEKDAY WRAPCOLS WRAPROWS XLOOKUP XMATCH YEAR
""".split())


def _col_number(letters):
    v = 0
    for ch in letters.upper():
        v = v * 26 + (ord(ch) - 64)
    return v


def looks_like_cell(name):
    m = re.fullmatch(r"([A-Za-z]{1,3})([0-9]+)", name)
    if m and _col_number(m.group(1)) <= 16384 and 1 <= int(m.group(2)) <= 1048576:
        return True
    return re.fullmatch(r"[RrCc][0-9]+|[Rr][0-9]+[Cc][0-9]+", name) is not None


def lint(name, formula):
    """Return a list of problems found without running Excel."""
    probs = []
    code = outside_strings(formula)
    tokens = list(re.finditer(r"(?<![\w.])([A-Za-z_]\w*)(?!\w)", code))
    declared = set()
    for t in tokens:
        after = code[t.end():t.end() + 1]
        if after not in ("(", "."):
            declared.add(t.group(1))
    for nm in sorted(declared):
        if nm.upper() in ("TRUE", "FALSE"):
            continue
        if looks_like_cell(nm):
            probs.append(f"name '{nm}' parses as a cell reference -- Excel refuses the "
                         f"whole definition (use a non-numeric suffix)")
    # LET-bound LAMBDAs: "Name, LAMBDA(" -- calling one named like a built-in
    # calls the built-in.
    for m in re.finditer(r"(?<![\w.])([A-Za-z_]\w*)\s*,\s*LAMBDA\s*\(", code):
        nm = m.group(1)
        if nm.upper() in BUILTIN_FUNCTIONS and re.search(
                r"(?<![\w.])" + re.escape(nm) + r"\s*\(", code[m.end():]):
            probs.append(f"LET-bound LAMBDA '{nm}' is called, but {nm.upper()}() is a "
                         f"built-in -- Excel calls the built-in or refuses the formula")
    lower = {}
    for nm in declared:
        lower.setdefault(nm.lower(), set()).add(nm)
    for group in lower.values():
        if len(group) > 1:
            probs.append(f"names {sorted(group)} differ only by case -- Excel names are "
                         f"case-insensitive: refused if both are bound in one LET, "
                         f"silent shadowing otherwise")
    for nm in sorted(declared):
        # e.g. a parameter Beta turns BETA.INV(...) into a field access on the
        # parameter, and the function returns #FIELD!
        if re.search(r"(?<![\w.])" + re.escape(nm) + r"\.[A-Za-z][\w.]*\s*\(", code, re.I):
            probs.append(f"'{nm}' is a name AND the head of a dotted function "
                         f"'{nm}.<...>(' -- Excel returns #FIELD!")
    if re.search(r"[(,=]\s*-\s*[A-Za-z_][\w.]*\s*\^", code):
        probs.append("unary minus on a power (-x^2): Excel computes (-x)^2; write -(x^2)")
    if "[[" in code:
        probs.append("'[[Arg]]' -- an optional parameter takes ONE bracket pair")
    if "//" in code or "/*" in code:
        probs.append("comment text survived into the formula")
    return probs


# ------------------------------------------------------------------ Excel driving
def retry(fn, tries=120, wait=0.25):
    """Run a COM call, absorbing the transient failures a busy Excel produces.

    Late-bound pywin32 reports "Excel is busy" in several disguises: com_error
    'call rejected' / 'server busy' / 'member not found'; AttributeError
    '<unknown>.Name'; and TypeError "'bool' object is not callable", which means
    a no-argument method was ALREADY run during the name lookup.
    """
    import pywintypes
    for _ in range(tries):
        try:
            return fn()
        except pywintypes.com_error as e:
            if e.args[0] not in (-2147418111, -2147417846, -2147352573):
                raise
        except AttributeError as e:
            if not (re.fullmatch(r"[\w<>.]+\.\w+", str(e))
                    or re.fullmatch(r"Property '[\w<>.]+' can not be set\.", str(e))):
                raise
        except TypeError as e:
            if "'bool' object is not callable" in str(e):
                return True
            if "'NoneType' object is not callable" in str(e):
                return None
            raise
        time.sleep(wait)
    raise RuntimeError("Excel stayed busy")


class OwnExcel:
    """A private, hidden Excel instance. Only this process is ever closed."""

    def __init__(self, visible=False):
        import pythoncom
        import win32com.client
        import win32com.client.gencache
        import win32process
        pythoncom.CoInitialize()
        # A corrupt makepy (gen_py) cache breaks early binding with "no attribute
        # 'CLSIDToClassMap'". Force late binding instead of deleting a shared cache.
        win32com.client.gencache.GetClassForCLSID = lambda clsid: None
        self.xl = win32com.client.DispatchEx("Excel.Application")   # NEW process
        self.pid = win32process.GetWindowThreadProcessId(self.xl.Hwnd)[1]
        retry(lambda: setattr(self.xl, "Visible", bool(visible)))
        retry(lambda: setattr(self.xl, "DisplayAlerts", False))

    def close(self):
        import pythoncom
        try:
            for wb in retry(lambda: list(self.xl.Workbooks)):
                retry(lambda: setattr(wb, "Saved", True))
        except Exception:
            pass
        try:
            retry(lambda: self.xl.Quit())
        except Exception:
            pass
        self.xl = None
        pythoncom.CoUninitialize()
        # Wait for OUR process to exit; end it only if it is still there.
        SYNCHRONIZE, TERMINATE = 0x00100000, 0x0001
        k32 = ctypes.windll.kernel32
        h = k32.OpenProcess(SYNCHRONIZE | TERMINATE, False, self.pid)
        if h:
            if k32.WaitForSingleObject(h, 15000) != 0:          # 0 = exited
                k32.TerminateProcess(h, 1)
                print(f"  note: own Excel (pid {self.pid}) did not quit; ended it")
            k32.CloseHandle(h)


def is_error_value(v):
    """Excel error cells reach COM as large negative ints (#N/A = -2146826246)."""
    return isinstance(v, int) and v <= -2146826200


def read_result(cell, tries=120):
    """The spilled block (or single value) as a list of rows, or an error string."""
    for _ in range(tries):
        txt = str(retry(lambda: cell.Text))
        if txt.startswith("#") and txt != "#":
            return txt
        if retry(lambda: cell.HasSpill):
            v = retry(lambda: cell.SpillingToRange.Value)
            return [list(r) for r in v]
        v = retry(lambda: cell.Value)
        if v is not None:
            return [[v]]
        time.sleep(0.25)
    return "#TIMEOUT"


def put(ws, addr, formula):
    """Formula2, never Formula: Range.Formula inserts the implicit-intersection
    operator (an at-sign) into a dynamic-array formula and quietly truncates the
    spill to one cell."""
    cell = retry(lambda: ws.Range(addr))
    retry(lambda: setattr(cell, "Formula2", formula))
    return cell


def excel_literal(v):
    if v is None:
        return ""
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, int):
        return str(v)
    if isinstance(v, float):
        return str(int(v)) if v.is_integer() else repr(v)   # repr round-trips exactly
    if isinstance(v, str):
        return '"' + v.replace('"', '""') + '"'
    if isinstance(v, (list, tuple)):
        if v and isinstance(v[0], (list, tuple)):
            return "{" + ";".join(",".join(excel_literal(x) for x in r) for r in v) + "}"
        return "{" + ",".join(excel_literal(x) for x in v) + "}"
    raise TypeError(f"no Excel literal for {type(v).__name__}")


def call_formula(func, args):
    args = list(args)
    while args and args[-1] is None:
        args.pop()
    return "=" + func + "(" + ",".join(excel_literal(a) for a in args) + ")"


def compare(expected, actual, tol):
    """Return (ok, max_abs_diff, note). Tolerance is relative above 1, absolute below."""
    if isinstance(actual, str):
        return False, math.inf, f"Excel returned {actual}"
    if len(expected) != len(actual) or any(len(e) != len(a) for e, a in zip(expected, actual)):
        return False, math.inf, (f"shape: expected {len(expected)}x{len(expected[0])}, "
                                 f"got {len(actual)}x{len(actual[0]) if actual else 0}")
    worst, ok = 0.0, True
    for er, ar in zip(expected, actual):
        for e, a in zip(er, ar):
            if is_error_value(a) or not isinstance(a, (int, float)):
                return False, math.inf, f"non-numeric cell {a!r}"
            d = abs(float(e) - float(a))
            worst = max(worst, d)
            if d > tol * max(1.0, abs(float(e))):
                ok = False
    return ok, worst, ""


def stored_definitions(xlsx_path):
    """Read <definedName> text straight out of the saved file -- the form Excel
    STORES, prefixes and all, which is what the 8,192 limit is measured on."""
    with zipfile.ZipFile(xlsx_path) as z:
        xml = z.read("xl/workbook.xml").decode("utf-8")
    out = {}
    for m in re.finditer(r'<definedName\b([^>]*)>(.*?)</definedName>', xml, re.S):
        nm = re.search(r'name="([^"]*)"', m.group(1))
        if nm:
            out[html.unescape(nm.group(1))] = html.unescape(m.group(2))
    return out


# --------------------------------------------------------------------- the run
class Results:
    def __init__(self):
        self.rows = []

    def add(self, section, label, ok, detail=""):
        self.rows.append((section, label, ok, detail))
        print(f"  [{'PASS' if ok else 'FAIL'}] {label}" + (f"  -- {detail}" if detail else ""))

    def failed(self):
        return [r for r in self.rows if not r[2]]


def main():
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    ap.add_argument("--source", help="AFE-form module file (UTF-8). Default: embedded example")
    ap.add_argument("--cases", type=int, default=50, help="random cases (default 50)")
    ap.add_argument("--seed", type=int, default=20261015, help="random seed (reproducible)")
    ap.add_argument("--tol", type=float, default=1e-9,
                    help="tolerance, relative above |1|, absolute below (default 1e-9)")
    ap.add_argument("--visible", action="store_true", help="show the Excel window")
    ap.add_argument("--keep-files", action="store_true", help="keep the temp workbook")
    ap.add_argument("--quiet-cases", action="store_true",
                    help="print only failing value cases")
    ap.add_argument("--lint-only", action="store_true",
                    help="run the static checks only; do not start Excel")
    opt = ap.parse_args()

    src = DEMO_SOURCE
    if opt.source:
        with open(opt.source, encoding="utf-8-sig") as fh:
            src = fh.read()
    entries = parse_module(src)
    res = Results()

    print(f"LAMBDA test harness   function {FUNC}   tolerance {opt.tol:g}   "
          f"seed {opt.seed}")
    print(f"  {len(entries)} definition(s) in source: " + ", ".join(n for n, _ in entries))

    print("\n1. static lint")
    for name, formula in entries:
        probs = lint(name, formula)
        res.add("lint", f"{name}: lint clean", not probs, "; ".join(probs))
        res.add("lint", f"{name}: flattened length {len(formula):,} <= {CAP:,}",
                len(formula) <= CAP)
    if opt.lint_only:
        return 1 if res.failed() else 0
    if FUNC not in [n for n, _ in entries]:
        print(f"  {FUNC} is not defined in the source -- edit FUNC (EDIT 2)")
        return 2

    tmpdir = tempfile.mkdtemp(prefix="lambda_harness_")
    path = os.path.join(tmpdir, "harness.xlsx")
    excel = None
    try:
        excel = OwnExcel(visible=opt.visible)
        xl = excel.xl
        print(f"\n   (started a private Excel, pid {excel.pid})")
        wb = retry(lambda: xl.Workbooks.Add())
        ws = retry(lambda: wb.Worksheets(1))

        print("\n2. install")
        installed, reasons = set(), {}
        for _ in range(len(entries)):              # helpers may be defined after callers
            progress = False
            for name, formula in entries:
                if name in installed:
                    continue
                try:
                    retry(lambda: wb.Names.Add(Name=name, RefersTo=formula))
                    installed.add(name)
                    progress = True
                except Exception as ex:
                    reasons[name] = str(ex).splitlines()[0][:160]
            if not progress or len(installed) == len(entries):
                break
        for name, _ in entries:
            res.add("install", f"{name}: installed via Names.Add", name in installed,
                    "" if name in installed else reasons.get(name, ""))
        if FUNC not in installed:
            raise RuntimeError("function under test did not install")

        print("\n3. round trip")
        for name, formula in entries:
            if name not in installed:
                continue
            back = str(retry(lambda: wb.Names(name).RefersTo))
            a, b = canonical(strip_prefixes(back)), canonical(formula)
            k = first_difference(a, b)
            exact = "byte-exact" if back == formula else \
                f"Excel rewrote whitespace: {len(formula) - len(back):+d} chars"
            res.add("roundtrip", f"{name}: COM read-back equals written text", k < 0,
                    exact if k < 0 else f"first difference at char {k}: wrote "
                    f"{b[max(0, k - 30):k + 30]!r} got {a[max(0, k - 30):k + 30]!r}")
        retry(lambda: wb.SaveAs(path, FileFormat=51))
        stored = stored_definitions(path)
        for name, formula in entries:
            if name not in installed:
                continue
            s = stored.get(name)
            if s is None:
                res.add("roundtrip", f"{name}: present in saved file", False)
                continue
            a, b = canonical(strip_prefixes(s)), canonical(formula[1:])
            k = first_difference(a, b)
            res.add("roundtrip", f"{name}: stored text (prefixes stripped) equals written",
                    k < 0, "" if k < 0 else f"first difference at char {k}: wrote "
                    f"{b[max(0, k - 30):k + 30]!r} stored {a[max(0, k - 30):k + 30]!r}")
            res.add("roundtrip", f"{name}: stored length {len(s):,} < {CAP:,} "
                    f"(written {len(formula) - 1:,}; storage adds {len(s) - len(formula) + 1:+,})",
                    len(s) < CAP)

        print("\n4. inline help")
        help_out = read_result(put(ws, "A1", "=" + FUNC + "()"))
        ok = isinstance(help_out, list) and str(help_out[0][0]).startswith("FUNCTION")
        res.add("help", f"{FUNC}() spills a help table starting FUNCTION:", ok,
                "" if ok else f"got {help_out if isinstance(help_out, str) else help_out[0]}")
        retry(lambda: ws.Cells.Clear())

        print(f"\n5. values against the Python reference ({opt.cases} random + fixed)")
        print(f"   {'case':<34} {'shape':>7} {'max |diff|':>12}  result")
        snapshot = None
        for label, args in cases(opt.cases, opt.seed):
            try:
                expected, want_error = reference(*args), False
            except ValueError:
                expected, want_error = None, True
            formula = call_formula(FUNC, args)
            cell = put(ws, "A1", formula)
            got = read_result(cell)
            if want_error:
                ok = isinstance(got, str) and got.startswith("#")
                shape, diff = "-", "-"
                note = "" if ok else "expected an Excel error, got a value"
            else:
                ok, worst, note = compare(expected, got, opt.tol)
                shape = f"{len(expected)}x{len(expected[0])}"
                diff = "-" if worst == math.inf else f"{worst:.2e}"
                if snapshot is None and ok:
                    snapshot = (formula, expected)
            res.rows.append(("values", label, ok, note))
            if not ok or not opt.quiet_cases:
                print(f"   {label:<34} {shape:>7} {diff:>12}  "
                      f"{'PASS' if ok else 'FAIL'}" + (f"  {note}" if note else ""))
                if not ok:
                    print(f"      formula: {formula}")
            retry(lambda: ws.Cells.Clear())     # never write into a live spill range
        n_val = [r for r in res.rows if r[0] == "values"]
        print(f"   values: {sum(r[2] for r in n_val)} of {len(n_val)} passed")

        print("\n6. survives save and reopen")
        if snapshot:
            put(ws, "A1", snapshot[0])
            retry(lambda: wb.Save())
            retry(lambda: wb.Close(SaveChanges=False))
            wb = retry(lambda: xl.Workbooks.Open(path))
            retry(lambda: xl.CalculateFull())
            got = read_result(retry(lambda: wb.Worksheets(1).Range("A1")))
            ok, worst, note = compare(snapshot[1], got, opt.tol)
            res.add("reopen", "reopened workbook recalculates to the reference", ok, note)
        else:
            res.add("reopen", "no passing value case to reopen with", False)
    except Exception as ex:
        print(f"\nHARNESS ERROR: {type(ex).__name__}: {ex}")
        if excel:
            excel.close()
            excel = None
        if not opt.keep_files:
            shutil.rmtree(tmpdir, ignore_errors=True)
        return 2
    finally:
        if excel:
            excel.close()
            print(f"\n   (closed the private Excel, pid {excel.pid})")
    if opt.keep_files:
        print(f"   workbook kept: {path}")
    else:
        shutil.rmtree(tmpdir, ignore_errors=True)

    # ------------------------------------------------------------------ summary
    print("\n" + "=" * 72)
    print(f"{'section':<10} {'passed':>7} {'failed':>7}")
    for sec in ("lint", "install", "roundtrip", "help", "values", "reopen"):
        rows = [r for r in res.rows if r[0] == sec]
        if rows:
            print(f"{sec:<10} {sum(r[2] for r in rows):>7} {sum(not r[2] for r in rows):>7}")
    total, bad = len(res.rows), len(res.failed())
    print("-" * 72)
    if bad:
        print(f"FAIL -- {bad} of {total} checks failed:")
        for sec, label, _, detail in res.failed()[:25]:
            print(f"  [{sec}] {label}" + (f": {detail}" if detail else ""))
        return 1
    print(f"PASS -- all {total} checks passed (tolerance {opt.tol:g})")
    return 0


if __name__ == "__main__":
    sys.exit(main())
