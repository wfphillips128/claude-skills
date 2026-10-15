"""Inventory the LAMBDA functions defined in an Excel workbook.

Reads the defined-name collection, strips Excel's internal prefixes, re-indents the
source, and reports each function's signature plus which other functions it calls.

Read-only: the workbook is never modified.

Usage:
    python dump_lambdas.py <workbook.xlsx> [function-name-fragment]

    python dump_lambdas.py model.xlsx              # all LAMBDAs
    python dump_lambdas.py model.xlsx Interest     # only names containing "Interest"
    python dump_lambdas.py model.xlsx --signatures # signature table only
"""

import re
import sys

try:
    import openpyxl
except ImportError:
    sys.exit("openpyxl is required:  python -m pip install openpyxl")

# Excel's internal prefixes on stored formulas.
#   _xlfn.  post-2019 function          _xlpm.  LET name / required LAMBDA parameter
#   _xlop.  OPTIONAL LAMBDA parameter   _xlws.  worksheet-scoped function
PREFIXES = ("_xlfn.", "_xlpm.", "_xlop.", "_xlws.")


def unmangle(formula):
    """Turn stored formula text into what the user would type."""
    out = formula
    # ANCHORARRAY('Sheet'!$F$46)  ->  'Sheet'!$F$46#
    out = re.sub(r"_xlfn\.ANCHORARRAY\(([^()]*)\)", r"\1#", out)
    # SINGLE(x) is the implicit-intersection operator @
    out = re.sub(r"_xlfn\.SINGLE\(", "@(", out)
    for p in PREFIXES:
        out = out.replace(p, "")
    return out


def optional_params(formula):
    """Names carrying the _xlop. prefix are the optional LAMBDA parameters.

    5G declares every argument optional so the inline help can display, so in a
    correctly built 5G function this should be the full parameter list.
    """
    return set(re.findall(r"_xlop\.([A-Za-z0-9_.]+)", formula))


def signature(clean):
    """Extract the parameter list from an un-mangled LAMBDA definition."""
    m = re.match(r"\s*LAMBDA\s*\(", clean)
    if not m:
        return []
    depth, params, current = 0, [], ""
    for ch in clean[m.end() - 1:]:
        if ch == "(":
            depth += 1
            if depth == 1:
                continue
        elif ch == ")":
            depth -= 1
            if depth == 0:
                break
        if depth == 1 and ch == ",":
            params.append(current.strip())
            current = ""
        else:
            current += ch
    # The final "parameter" is the function body (a LET, usually) - drop it.
    return [p for p in params if p and re.fullmatch(r"\[?[A-Za-z0-9_.]+\]?", p)]


def strip_strings(formula):
    """Blank out string literals.

    The help text contains lines like "=DebtSculptIssuanceλ(...)" which would
    otherwise be counted as function calls.
    """
    return re.sub(r'"[^"]*"', '""', formula)


def called_functions(clean, own_name):
    """Which other lambda-suffixed functions this one calls, self-call flagged."""
    found = re.findall(r"([A-Za-z0-9_.]*λ)\s*\(", strip_strings(clean))
    calls = {}
    for name in found:
        calls[name] = calls.get(name, 0) + 1
    return [(n, c, n == own_name) for n, c in sorted(calls.items())]


def split_top_level(text, sep=","):
    """Split on `sep`, ignoring anything inside (), {} or a string literal."""
    parts, buf, depth, i = [], "", 0, 0
    while i < len(text):
        ch = text[i]
        if ch == '"':                       # copy the whole string literal verbatim
            j = text.find('"', i + 1)
            j = len(text) - 1 if j == -1 else j
            buf += text[i:j + 1]
            i = j + 1
            continue
        if ch in "({":
            depth += 1
        elif ch in ")}":
            depth -= 1
        if depth == 0 and text.startswith(sep, i):
            parts.append(buf)
            buf = ""
            i += len(sep)
            continue
        buf += ch
        i += 1
    parts.append(buf)
    return [p.strip() for p in parts]


def body_of(text, func):
    """Return the argument text inside `func(...)`, or None if it doesn't match."""
    m = re.match(r"\s*" + func + r"\s*\(", text)
    if not m:
        return None
    return text[m.end():text.rfind(")")]


def wrap_help(value, indent):
    """Put each concatenated help-text line on its own row.

    Splits only between two string literals ("..." & "..."), which is the help-text
    concatenation pattern and nothing else.
    """
    pad = " " * (indent + 4)
    return re.sub(r'"\s*&\s*"', '" &\n' + pad + '"', value)


def pretty(clean):
    """Re-indent a stored formula so each LET step is on its own line.

    Deliberately narrow: it splits LAMBDA parameters and LET steps and leaves each
    step's expression alone. Trying to format every nested IF() produces noise.
    """
    args = body_of(clean, "LAMBDA")
    if args is None:
        return clean

    parts = split_top_level(args)
    params, body = parts[:-1], parts[-1]

    out = ["LAMBDA("]
    if params:
        out.append("    " + ", ".join(params) + ",")

    let_args = body_of(body, "LET")
    if let_args is None:
        out.append("    " + body)
        out.append(")")
        return "\n".join(out)

    steps = split_top_level(let_args)
    out.append("    LET(")
    # LET is name, value, name, value, ..., final-expression
    pairs, tail = steps[:-1], steps[-1]
    for i in range(0, len(pairs) - 1, 2):
        name, value = pairs[i], pairs[i + 1]
        if "TEXTSPLIT" in value and "&" in value:
            value = wrap_help(value, 8)
        out.append(f"        {name}, {value},")
    if len(pairs) % 2:                       # odd count: something didn't pair up
        out.append(f"        {pairs[-1]},")
    out.append(f"        {tail}")
    out.append("    )")
    out.append(")")
    return "\n".join(out)


def main():
    if len(sys.argv) < 2:
        sys.exit(__doc__)

    path = sys.argv[1]
    args = sys.argv[2:]
    sigs_only = "--signatures" in args
    needle = next((a for a in args if not a.startswith("--")), "").lower()

    wb = openpyxl.load_workbook(path, read_only=False)
    lambdas = []
    for name in wb.defined_names.keys():
        raw = str(wb.defined_names[name].value)
        if "LAMBDA" not in raw:
            continue
        if needle and needle not in name.lower():
            continue
        lambdas.append((name, raw))

    if not lambdas:
        print("No LAMBDA definitions found"
              + (f" matching {needle!r}" if needle else "") + ".")
        others = [k for k in wb.defined_names.keys()]
        if others:
            print(f"({len(others)} other defined names in this workbook.)")
        return

    print(f"{len(lambdas)} LAMBDA function(s) in {path}\n")

    # ---- signature table ------------------------------------------------------
    print("SIGNATURES  ([x] = optional, per Excel's _xlop. prefix)")
    print("-" * 78)
    for name, raw in lambdas:
        clean = unmangle(raw)
        opt = optional_params(raw)
        params = signature(clean)
        shown = [f"[{p}]" if p in opt else p for p in params]
        print(f"{name}(")
        print("    " + ", ".join(shown) if shown else "    <none parsed>")
        print(")")
        calls = called_functions(clean, name)
        for callee, count, is_self in calls:
            tag = "  <-- SELF-CALL" if is_self else ""
            print(f"    calls {callee} x{count}{tag}")
        # A recursive function that never calls itself is a red flag: the self-call
        # probably still names the function this one was copied from.
        if calls and not any(s for _, _, s in calls):
            print("    !! calls other lambdas but never itself - check the recursion "
                  "target was renamed")
        print()

    if sigs_only:
        return

    # ---- full source ----------------------------------------------------------
    for name, raw in lambdas:
        print("=" * 78)
        print(name)
        print("=" * 78)
        print(pretty(unmangle(raw)))
        print()


if __name__ == "__main__":
    main()
