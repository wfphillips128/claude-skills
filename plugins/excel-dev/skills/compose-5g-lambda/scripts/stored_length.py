"""Measure what Excel will STORE for a 5G LAMBDA, against its 8192-character cap.

Excel caps a single formula at 8192 characters and it measures the form it
stores, not the text you paste. Storage adds a prefix to every name:

    _xlfn.   post-2007 functions (LET, LAMBDA, TEXTSPLIT, VSTACK, ...)
    _xlws.   worksheet-only functions (FILTER, SORT, ...)
    _xlpm.   LET step names and REQUIRED lambda parameters
    _xlop.   OPTIONAL lambda parameters -- the [Bracketed] ones

On a function with a dozen arguments those prefixes run to ~2000 characters, so
a 6200-character paste is stored at over 8100. Over the limit, .xlsx saves warn
and Name Manager refuses the paste outright.

This is a MEASUREMENT, not a converter.

CALIBRATION, against three independent points:

    large function, nodoc build, on disk     actual 4566, est 4538  (-0.6%)
    large function, measured inside Excel    actual 7328, est 7351  (+0.3%)
    small example (1,017 chars written)      actual 1359, est 1390  (+2.3%)

Within about half a percent on large functions, worse on small ones (the
estimate does not model Excel dropping the spaces in the parameter list), and
the sign is not guaranteed. Treat
anything inside 300 characters of the cap as needing confirmation in Excel
rather than trusting this number.

Usage
    python stored_length.py <source.txt> [--show]

    <source.txt>  AFE-form source: comments, indentation and all. Comments are
                  stripped first, exactly as AFE does on import, so they cost
                  nothing.
    --show        print the reconstructed stored form as well as the count

Read-only. Never writes to the workbook or the source.
"""

from __future__ import annotations

import re
import sys

CAP = 8192

# Functions that post-date Excel 2007 carry _xlfn. in storage.
XLFN = {
    "LAMBDA", "LET", "TEXTSPLIT", "TEXTJOIN", "TEXTBEFORE", "TEXTAFTER",
    "VSTACK", "HSTACK", "TOROW", "TOCOL", "WRAPROWS", "WRAPCOLS",
    "TAKE", "DROP", "EXPAND", "CHOOSEROWS", "CHOOSECOLS",
    "SCAN", "REDUCE", "MAP", "BYROW", "BYCOL", "MAKEARRAY", "ISOMITTED",
    "SWITCH", "IFS", "MAXIFS", "MINIFS", "CONCAT", "XLOOKUP", "XMATCH",
    "ANCHORARRAY", "SINGLE", "ARRAYTOTEXT", "VALUETOTEXT", "GROUPBY",
    "PIVOTBY", "PERCENTOF", "TRIMRANGE", "REGEXTEST", "REGEXEXTRACT",
    "REGEXREPLACE", "SEQUENCE", "RANDARRAY",
}

# Dynamic-array functions that only work on a worksheet carry BOTH prefixes,
# stored as _xlfn._xlws.NAME. Verified: SEQUENCE and RANDARRAY are NOT in this
# set -- they store as plain _xlfn.
XLWS = {"FILTER", "SORT", "SORTBY", "UNIQUE"}

# Everything else that looks like a function but needs no prefix.
NATIVE = {
    "IF", "AND", "OR", "NOT", "SUM", "MAX", "MIN", "ABS", "AVERAGE", "COUNT",
    "COLUMNS", "ROWS", "INDEX", "MATCH", "FLOOR", "CEILING", "ROUND", "TRIM",
    "ISERROR", "ISNUMBER", "ISBLANK", "ISTEXT", "IFERROR", "TEXT", "VALUE",
    "LEFT", "RIGHT", "MID", "LEN", "SUBSTITUTE", "DATE", "YEAR", "MONTH",
    "DAY", "EDATE", "EOMONTH", "PMT", "NPV", "IRR", "SUMPRODUCT", "TRANSPOSE",
    "CHOOSE", "OFFSET", "INDIRECT", "SIGN", "MOD", "INT", "POWER", "SQRT",
    "COUNTA", "LOOKUP", "VLOOKUP", "HLOOKUP", "SUMIF", "SUMIFS", "COUNTIF",
    "COUNTIFS", "AVERAGEIF", "AVERAGEIFS", "NA", "TRUE", "FALSE",
}


def strip_comments(src: str) -> str:
    """Remove /* */ and // comments. AFE discards these on import."""
    out, i, n = [], 0, len(src)
    in_string = False
    while i < n:
        ch = src[i]
        if in_string:
            out.append(ch)
            if ch == '"':
                # "" inside a string is an escaped quote
                if i + 1 < n and src[i + 1] == '"':
                    out.append(src[i + 1])
                    i += 2
                    continue
                in_string = False
            i += 1
            continue
        if ch == '"':
            in_string = True
            out.append(ch)
            i += 1
            continue
        if src.startswith("/*", i):
            j = src.find("*/", i + 2)
            i = n if j < 0 else j + 2
            continue
        if src.startswith("//", i):
            j = src.find("\n", i)
            i = n if j < 0 else j
            continue
        out.append(ch)
        i += 1
    return "".join(out)


def collapse(src: str) -> str:
    """Flatten to one line. Excel PRESERVES the spaces you paste, so runs of
    whitespace collapse to a single space rather than disappearing.

    Verified against a known paste/stored pair: 373 spaces in, 377 out.
    """
    out, i, n = [], 0, len(src)
    in_string = False
    while i < n:
        ch = src[i]
        if in_string:
            out.append(ch)
            if ch == '"':
                if i + 1 < n and src[i + 1] == '"':
                    out.append(src[i + 1])
                    i += 2
                    continue
                in_string = False
            i += 1
            continue
        if ch == '"':
            in_string = True
            out.append(ch)
        elif ch.isspace():
            if out and out[-1] != " ":
                out.append(" ")
        else:
            out.append(ch)
        i += 1
    return "".join(out).strip()


def optional_params(src: str) -> set:
    """Names declared [Bracketed] in a LAMBDA signature -> _xlop."""
    return set(re.findall(r"\[\s*([A-Za-z_][A-Za-z0-9_.?]*)\s*\]", src))


def prefix_for(name: str) -> str:
    upper = name.upper()
    if upper in XLFN:
        return "_xlfn."
    if upper in XLWS:
        return "_xlfn._xlws."      # worksheet-only DA functions take both
    if upper in NATIVE:
        return ""
    # Everything else -- LET step names, parameter references, required
    # parameters -- stores as _xlpm. Note that a bare reference to an OPTIONAL
    # parameter is _xlpm. too: _xlop. marks the declaration in the signature
    # and nothing else. Verified against a stored pair: 12 _xlop. (one per
    # parameter) against 250 _xlpm.
    return "_xlpm."


def stored_form(src: str) -> str:
    body = collapse(strip_comments(src))

    # Drop the "Name = " assignment head if the AFE source carries one.
    m = re.match(r"^[A-Za-z_][A-Za-z0-9_.]*λ?=", body)
    if m:
        body = body[m.end():]
    body = body.rstrip(";")

    # A call to a WORKBOOK-LEVEL 5g function keeps its bare name. A lambda
    # held in a LET step (Creditλ, Partsλ, Solveλ) does NOT -- it is a LET name
    # and stores as _xlpm. The two are told apart by the module qualifier: a
    # workbook-level name is dotted (fx.DebtSizeIssueλ), a LET step is not.
    calls = set(m for m in re.findall(r"([A-Za-z_][A-Za-z0-9_.]*λ)\s*\(", body)
                if "." in m)

    out, i, n = [], 0, len(body)
    while i < n:
        ch = body[i]
        if ch == '"':
            j = i + 1
            while j < n:
                if body[j] == '"':
                    if j + 1 < n and body[j + 1] == '"':
                        j += 2
                        continue
                    break
                j += 1
            out.append(body[i:j + 1])
            i = j + 1
            continue
        if ch == "[":                      # [Arg] -> _xlop.Arg, signature only
            j = body.find("]", i)
            if j > 0:
                inner = body[i + 1:j].strip()
                if re.fullmatch(r"[A-Za-z_][A-Za-z0-9_.?]*", inner or ""):
                    out.append("_xlop." + inner)
                    i = j + 1
                    continue
        m = re.match(r"[A-Za-z_λ][A-Za-z0-9_.λ?]*", body[i:])
        if m:
            name = m.group(0)
            if name in calls:
                out.append(name)           # sibling 5g function, no prefix
            else:
                out.append(prefix_for(name) + name)
            i += len(name)
            continue
        out.append(ch)
        i += 1
    return "".join(out)


def main() -> int:
    if len(sys.argv) < 2:
        print(__doc__)
        return 2
    path = sys.argv[1]
    show = "--show" in sys.argv

    with open(path, encoding="utf-8") as fh:
        src = fh.read()

    stored = stored_form(src)
    length = len(stored)
    headroom = CAP - length

    print("Source file      %s" % path)
    print("Pasted (no comments, collapsed)  %6d chars"
          % len(collapse(strip_comments(src))))
    print("STORED (estimated)               %6d chars" % length)
    print("Cap                              %6d" % CAP)
    print("Headroom                         %6d %s"
          % (headroom, "OK" if headroom >= 0 else "*** OVER THE LIMIT ***"))
    if 0 <= headroom < 300:
        print("\nWithin 300 characters of the cap. This estimate is an upper")
        print("bound but not exact -- confirm by pasting into Excel before")
        print("relying on it.")
    if show:
        print("\n--- stored form ---")
        print(stored)
    return 0 if headroom >= 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())
