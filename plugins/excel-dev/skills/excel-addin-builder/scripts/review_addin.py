"""
review_addin.py - static review of an Excel VBA add-in (.xlsm / .xlam).

Reads the file only. It NEVER opens Excel, so it is safe to run on a workbook
that is open, and safe to run before anything else.

What it checks, and why each one matters:

  RIBBON (customUI/customUI14.xml or customUI.xml inside the zip)
    part + relationship  The XML part exists AND _rels/.rels points at it with
                        the right relationship type. A part with no
                        relationship is ignored by Excel: no tab, no error.
    well-formed + namespace  A malformed file or wrong namespace hides the WHOLE
                        tab, silently.
    unique ids          A duplicate id hides the whole tab, silently.
    exclusive pairs     label + getLabel (and similar pairs) on one control is
                        a schema error: the whole tab disappears, silently.
    callbacks           Every onAction / getLabel / getEnabled / getImage /
                        onLoad / ... names a Public procedure in a STANDARD
                        module, defined once, with the parameter count the
                        ribbon passes. Otherwise the button silently does
                        nothing, or says "cannot run the macro" with no clue.
    custom images       Every image="x" is embedded (customUI14.xml.rels +
                        the image file), unless loadImage supplies images.

  VBA (extracted with oletools; nothing is executed)
    duplicate imports   modFoo1 next to modFoo: File > Import does not replace,
                        it adds a numbered copy, and every procedure becomes
                        ambiguous.
    duplicate Public names  The same Public name in two standard modules.
    Option Explicit     Missing in a module (warning).
    non-ASCII code      Characters that a UTF-8 editor would destroy
                        (warning; see reference/bas-modules.md).

  .bas FILES ON DISK (optional, --bas-dir)
    encoding            UTF-8 multi-byte sequences or U+FFFD replacement bytes
                        in a file the VBA editor will read as Windows-1252.
    VB_Name             Missing header = imports as "Module1".
    drift               Each file matches the module of the same name in the
                        workbook. With --pending the files are treated as
                        not-yet-imported: callbacks may resolve to them and
                        drift is information only.

Usage:
  python review_addin.py BOOK.xlsm
  python review_addin.py BOOK.xlam --bas-dir build
  python review_addin.py BOOK.xlsm --bas-dir build --pending --xml "ribbon - new group.xml"
  python review_addin.py BOOK.xlsm --strict        (warnings also fail)

Exit code: 0 = no FAIL, 1 = at least one FAIL, 2 = could not run.
Needs: Python 3.8+, oletools (pip install oletools).
"""
import argparse
import difflib
import os
import re
import sys
import zipfile
import xml.etree.ElementTree as ET

try:
    from oletools.olevba import VBA_Parser
except ImportError:                                   # pragma: no cover
    print("review_addin.py needs oletools:  python -m pip install oletools")
    sys.exit(2)

NS14 = "http://schemas.microsoft.com/office/2009/07/customui"
NS12 = "http://schemas.microsoft.com/office/2006/01/customui"
REL_UI14 = "http://schemas.microsoft.com/office/2007/relationships/ui/extensibility"
REL_UI12 = "http://schemas.microsoft.com/office/2006/relationships/ui/extensibility"
REL_IMAGE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"

# Attribute -> parameter count the ribbon passes. onAction depends on the
# element, so it is handled separately.
CALLBACK_ARGS = {
    "onLoad": 1, "loadImage": 2, "onChange": 2,
    "getLabel": 2, "getEnabled": 2, "getVisible": 2, "getImage": 2,
    "getScreentip": 2, "getSupertip": 2, "getKeytip": 2, "getPressed": 2,
    "getShowLabel": 2, "getShowImage": 2, "getSize": 2, "getDescription": 2,
    "getText": 2, "getItemCount": 2, "getSelectedItemIndex": 2,
    "getSelectedItemID": 2, "getContent": 2, "getTitle": 2, "getHelperText": 2,
    "getItemLabel": 3, "getItemID": 3, "getItemImage": 3,
    "getItemScreentip": 3, "getItemSupertip": 3, "getItemWidth": 2, "getItemHeight": 2,
}
ONACTION_ARGS = {"button": 1, "toggleButton": 2, "checkBox": 2,
                 "dropDown": 3, "gallery": 3, "command": (2, 3)}

# A property and its callback are mutually exclusive on one control.
EXCLUSIVE = [("label", "getLabel"), ("enabled", "getEnabled"), ("visible", "getVisible"),
             ("screentip", "getScreentip"), ("supertip", "getSupertip"),
             ("keytip", "getKeytip"), ("description", "getDescription"),
             ("showLabel", "getShowLabel"), ("showImage", "getShowImage"),
             ("size", "getSize"), ("title", "getTitle")]
IMAGE_ATTRS = ("image", "imageMso", "getImage")

ROWS = []                                             # (status, check, detail)


def row(status, check, detail=""):
    ROWS.append((status, check, detail))


# --------------------------------------------------------------------------- VBA
PROC_RE = re.compile(
    r"^[ \t]*(?:(Public|Private|Friend)[ \t]+)?(?:Static[ \t]+)?"
    r"(Sub|Function|Property[ \t]+(?:Get|Let|Set))[ \t]+([A-Za-z_]\w*)[ \t]*\(",
    re.I | re.M)


def join_continuations(code):
    """VBA ' _' line continuations -> one logical line."""
    return re.sub(r"[ \t]_[ \t]*\r?\n", " ", code)


def read_params(code, start):
    """Text between the '(' just before start and its matching ')'."""
    depth, i = 1, start
    while i < len(code) and depth:
        if code[i] == "(":
            depth += 1
        elif code[i] == ")":
            depth -= 1
        i += 1
    return code[start:i - 1]


def split_params(text):
    out, depth, cur = [], 0, ""
    for ch in text:
        if ch == "(":
            depth += 1
        elif ch == ")":
            depth -= 1
        if ch == "," and depth == 0:
            out.append(cur)
            cur = ""
        else:
            cur += ch
    out.append(cur)
    return [p.strip() for p in out if p.strip()]


def param_type(p):
    m = re.search(r"\bAs\s+([\w.]+)", p, re.I)
    return m.group(1) if m else "Variant"


def parse_procs(module, kind, code):
    """[(name, scope, kind_of_proc, [param types], module, module kind)]"""
    code = join_continuations(code)
    procs = []
    for m in PROC_RE.finditer(code):
        scope = (m.group(1) or "Public").capitalize()
        params = split_params(read_params(code, m.end()))
        required = sum(1 for p in params if not re.match(r"(Optional|ParamArray)\s", p, re.I))
        most = 99 if any(re.match(r"ParamArray\s", p, re.I) for p in params) else len(params)
        procs.append({"name": m.group(3), "scope": scope, "type": m.group(2).split()[0].capitalize(),
                      "ptypes": [param_type(p) for p in params], "required": required, "most": most,
                      "module": module, "mkind": kind})
    return procs


def module_kind(filename, code):
    ext = os.path.splitext(filename)[1].lower()
    if ext == ".bas":
        return "standard"
    if ext == ".frm":
        return "form"
    if re.search(r'Attribute VB_Base = "0\{0002082', code):
        return "document"                            # ThisWorkbook / Sheet modules
    return "class"


def workbook_modules(path):
    mods = {}
    vp = VBA_Parser(path)
    try:
        if not vp.detect_vba_macros():
            return mods
        for (_, _, vba_filename, code) in vp.extract_macros():
            name = os.path.splitext(vba_filename)[0]
            mods[name] = {"code": code, "kind": module_kind(vba_filename, code), "source": "workbook"}
    finally:
        vp.close()
    return mods


def norm_code(code):
    """Compare code the way VBA does: ignore case (the VBE re-cases
    identifiers), line endings, blank lines, trailing spaces, Attribute lines."""
    return [ln.rstrip().lower() for ln in code.replace("\r\n", "\n").replace("\r", "\n").split("\n")
            if ln.strip() and not ln.startswith("Attribute VB_")]


def read_bas_file(path):
    """Returns (text, problems). VBE reads .bas files as Windows-1252."""
    raw = open(path, "rb").read()
    problems = []
    if b"\xef\xbf\xbd" in raw:
        problems.append("contains U+FFFD replacement bytes: a character was destroyed by a UTF-8 "
                        "round trip. Recover the original from the workbook (olevba --code).")
    if raw.startswith(b"\xef\xbb\xbf"):
        problems.append("starts with a UTF-8 BOM, which the VBA editor imports as junk characters.")
    try:
        raw.decode("ascii")
    except UnicodeDecodeError:
        try:
            decoded = raw.decode("utf-8")
            if any(ord(c) > 127 for c in decoded) and b"\xef\xbf\xbd" not in raw:
                problems.append("is saved as UTF-8 with non-ASCII characters; the VBA editor will "
                                "import them as mojibake. Save as Windows-1252 or keep it pure ASCII.")
        except UnicodeDecodeError:
            pass                                       # genuine Windows-1252: fine
    return raw.decode("cp1252", errors="replace"), problems


# ------------------------------------------------------------------------ ribbon
def local(tag):
    return tag.split("}", 1)[1] if "}" in tag else tag


def ribbon_parts(z):
    """[(part name, relationship type)] for every ribbon part _rels/.rels declares."""
    names = set(z.namelist())
    parts = []
    if "_rels/.rels" in names:
        rels = ET.fromstring(z.read("_rels/.rels"))
        for r in rels:
            t = r.get("Type", "")
            if t in (REL_UI14, REL_UI12):
                parts.append((r.get("Target", "").lstrip("/"), t))
    return parts


def check_relationships(z):
    names = set(z.namelist())
    declared = ribbon_parts(z)
    on_disk = [n for n in names if re.fullmatch(r"customUI/customUI(14)?\.xml", n, re.I)]
    for target, typ in declared:
        expect = REL_UI14 if target.lower().endswith("customui14.xml") else REL_UI12
        if target not in names:
            row("FAIL", f"relationship target {target} exists", "declared in _rels/.rels but missing from the file")
        elif typ != expect:
            row("FAIL", f"relationship type for {target}",
                f"is {typ.rsplit('/', 3)[-3]}; {os.path.basename(target)} needs {expect}")
        else:
            row("PASS", f"{target} is registered in _rels/.rels", "")
    for n in on_disk:
        if n not in [t for t, _ in declared]:
            row("FAIL", f"{n} is registered in _rels/.rels",
                "the part is in the file but no relationship points at it, so Excel ignores it")
    return [t for t, _ in declared if t in names]


def parse_xml(text, label, fragment=False):
    """ElementTree root, or None after recording a FAIL."""
    try:
        if fragment and not re.search(r"<customUI\b", text):
            text = re.sub(r"^\s*<\?xml[^>]*\?>", "", text)
            text = f'<customUI xmlns="{NS14}"><snippet>{text}</snippet></customUI>'
        return ET.fromstring(text.encode("utf-8") if isinstance(text, str) else text)
    except ET.ParseError as e:
        row("FAIL", f"{label} is well-formed XML",
            f"{e}. Excel would hide the WHOLE tab. (Two hyphens inside a comment are illegal too.)")
        return None


def check_ribbon_tree(root, label, part, procs, image_rels, z_names, all_ids):
    ns = root.tag.split("}")[0].strip("{") if "}" in root.tag else ""
    expect = NS14 if part.lower().endswith("customui14.xml") or part == "snippet" else NS12
    if local(root.tag) != "customUI":
        row("FAIL", f"{label} root element", f"is <{local(root.tag)}>, must be <customUI>")
    elif ns != expect:
        row("FAIL", f"{label} namespace", f"is {ns or '(none)'}; must be {expect}")
    else:
        row("PASS", f"{label} is well-formed with the right namespace", "")

    # unique ids (a snippet is also checked against the ribbon it will be
    # pasted into - the caller passes those ids in all_ids)
    ids = [el.get("id") for el in root.iter() if el.get("id")]
    dups = sorted({i for i in ids if ids.count(i) > 1} | {i for i in ids if i in all_ids})
    row("FAIL" if dups else "PASS", f"{label}: every id is unique",
        f"duplicated: {', '.join(dups)}" if dups else f"{len(ids)} ids")
    all_ids.update(ids)

    # exclusive pairs
    clashes = []
    for el in root.iter():
        for a, b in EXCLUSIVE:
            if el.get(a) is not None and el.get(b) is not None:
                clashes.append(f"{el.get('id') or local(el.tag)}: {a} + {b}")
        used = [a for a in IMAGE_ATTRS if el.get(a) is not None]
        if len(used) > 1:
            clashes.append(f"{el.get('id') or local(el.tag)}: {' + '.join(used)}")
    row("FAIL" if clashes else "PASS", f"{label}: no control sets a property AND its callback",
        "; ".join(clashes))

    # callbacks
    has_loadimage = root.get("loadImage") is not None
    n_cb, bad = 0, []
    for el in root.iter():
        tag = local(el.tag)
        for attr, val in el.attrib.items():
            if attr == "onAction":
                want = ONACTION_ARGS.get(tag, 1)
            elif attr in CALLBACK_ARGS:
                want = CALLBACK_ARGS[attr]
            else:
                continue
            n_cb += 1
            problem = resolve_callback(val, want, attr, procs)
            if problem:
                bad.append(f"{el.get('id') or tag} {attr}=\"{val}\": {problem}")
    row("FAIL" if bad else "PASS", f"{label}: every callback has a matching Public procedure",
        "\n".join(bad) if bad else f"{n_cb} callbacks")

    # custom images
    used = sorted({el.get("image") for el in root.iter() if el.get("image")})
    if used:
        if has_loadimage:
            row("INFO", f"{label}: custom images", f"{len(used)} supplied at run time by loadImage=\"{root.get('loadImage')}\"")
        else:
            missing = [u for u in used if u not in image_rels]
            broken = [u for u in used if u in image_rels and image_rels[u] not in z_names]
            row("FAIL" if (missing or broken) else "PASS", f"{label}: every custom image is embedded",
                (f"not registered: {', '.join(missing)}. " if missing else "")
                + (f"registered but file missing: {', '.join(broken)}" if broken else "")
                or f"{len(used)} images")


def resolve_callback(val, want, attr, procs):
    """'' if fine, else the reason the ribbon would not reach it."""
    if "." in val:
        mod, name = val.rsplit(".", 1)
        cands = [p for p in procs if p["name"].lower() == name.lower() and p["module"].lower() == mod.lower()]
    else:
        name = val
        cands = [p for p in procs if p["name"].lower() == name.lower()]
    cands = [p for p in cands if p["type"] in ("Sub", "Function")]
    if not cands:
        known = {p["name"].lower(): p["name"] for p in procs if p["type"] in ("Sub", "Function")}
        close = [known[c] for c in difflib.get_close_matches(name.lower(), list(known), n=3, cutoff=0.75)]
        return ("no Sub or Function by this name - the button will silently do nothing"
                + (f" (did you mean {' / '.join(close)}?)" if close else ""))
    usable = [p for p in cands if p["scope"] == "Public" and p["mkind"] == "standard"]
    if not usable:
        where = ", ".join(f"{p['scope']} in {p['module']} ({p['mkind']} module)" for p in cands)
        return f"found only as {where}; it must be Public in a standard module"
    if len(usable) > 1 and "." not in val:
        return "defined in " + " and ".join(p["module"] for p in usable) + ": ambiguous, the ribbon cannot pick"
    p = usable[0]
    counts = want if isinstance(want, tuple) else (want,)
    # Trailing Optional parameters are fine: the ribbon passes fewer, VBA fills the rest.
    if not any(p["required"] <= c <= p["most"] for c in counts):
        sig = {1: "(control As IRibbonControl)", 2: "(control As IRibbonControl, ByRef returnedVal)",
               3: "(control As IRibbonControl, index/id..., ByRef returnedVal)"}
        hint = "(ribbon As IRibbonUI)" if attr == "onLoad" else sig.get(counts[0], "")
        return (f"{p['module']}.{p['name']} takes {len(p['ptypes'])} parameter(s); "
                f"{attr} passes {'/'.join(map(str, counts))}. Expected {hint}")
    first = p["ptypes"][0].split(".")[-1].lower() if p["ptypes"] else ""
    ok_first = {"iribbonui", "variant", "object"} if attr == "onLoad" else \
               {"string", "variant"} if attr == "loadImage" else {"iribboncontrol", "variant", "object"}
    if first and first not in ok_first:
        return f"{p['module']}.{p['name']}: first parameter is As {p['ptypes'][0]}"
    return ""


# -------------------------------------------------------------------------- main
def main():
    ap = argparse.ArgumentParser(description="Static review of an Excel VBA add-in. Never opens Excel.")
    ap.add_argument("workbook", help=".xlsm or .xlam")
    ap.add_argument("--bas-dir", help="folder of .bas/.cls files to check and compare with the workbook")
    ap.add_argument("--pending", action="store_true",
                    help="the --bas-dir files are NOT imported yet: resolve callbacks against them, drift is info")
    ap.add_argument("--xml", action="append", default=[],
                    help="ribbon XML file or snippet to check as well (e.g. before pasting it in)")
    ap.add_argument("--strict", action="store_true", help="treat WARN as FAIL")
    a = ap.parse_args()

    if not os.path.isfile(a.workbook):
        print(f"Not found: {a.workbook}")
        return 2
    try:
        z = zipfile.ZipFile(a.workbook)
    except zipfile.BadZipFile:
        print(f"{a.workbook} is not an Office Open XML file (.xls / .xla are not supported).")
        return 2
    names = set(z.namelist())
    print(f"Reviewing {os.path.basename(a.workbook)}")

    # ---------------------------------------------------------------- VBA
    mods = workbook_modules(a.workbook)
    if not mods:
        row("FAIL", "workbook contains a VBA project", "none found: is it saved as .xlsm / .xlam?")
    else:
        kinds = {}
        for m in mods.values():
            kinds[m["kind"]] = kinds.get(m["kind"], 0) + 1
        row("INFO", "VBA modules", f"{len(mods)}: " + ", ".join(f"{v} {k}" for k, v in sorted(kinds.items())))

    # .bas files on disk
    disk = {}
    if a.bas_dir:
        files = sorted(f for f in os.listdir(a.bas_dir) if f.lower().endswith((".bas", ".cls")))
        if not files:
            row("WARN", f"--bas-dir {a.bas_dir}", "no .bas or .cls files found")
        for f in files:
            path = os.path.join(a.bas_dir, f)
            text, problems = read_bas_file(path)
            m = re.search(r'^Attribute VB_Name = "([^"]+)"', text, re.M)
            if not m:
                problems.append('has no  Attribute VB_Name = "..."  line, so it would import as Module1')
            name = m.group(1) if m else os.path.splitext(f)[0]
            row("FAIL" if problems else "PASS", f"{f}: encoding and header", " ".join(problems))
            disk[name] = {"code": text, "kind": "standard" if f.lower().endswith(".bas") else "class",
                          "source": f}
            if name in mods:
                same = norm_code(mods[name]["code"]) == norm_code(text)
                if same:
                    row("PASS", f"{f} matches module {name} in the workbook", "")
                else:
                    row("INFO" if a.pending else "FAIL", f"{f} matches module {name} in the workbook",
                        "differs - " + ("will replace it on import (remove the old one FIRST)" if a.pending
                                        else "which one is current? Re-export, or import the file"))
            else:
                row("INFO" if a.pending else "FAIL", f"{f}: module {name} is in the workbook",
                    "not yet - new module to import" if a.pending else "missing - import it")

    # the code callbacks resolve against: workbook, overlaid by pending files
    effective = dict(mods)
    if a.pending:
        effective.update(disk)
    procs = []
    for name, m in effective.items():
        procs += parse_procs(name, m["kind"], m["code"])

    # duplicate imports (modFoo1 beside modFoo)
    strays = sorted(n for n in mods if re.search(r"\d$", n) and re.sub(r"\d+$", "", n) in mods)
    row("FAIL" if strays else "PASS", "no duplicate-import modules (modFoo1 beside modFoo)",
        f"{', '.join(strays)}: VBA Import adds a numbered copy instead of replacing" if strays else "")

    # Public names defined twice in standard modules
    seen = {}
    for p in procs:
        if p["scope"] == "Public" and p["mkind"] == "standard" and p["type"] in ("Sub", "Function"):
            seen.setdefault(p["name"].lower(), set()).add(p["module"])
    dups = {k: sorted(v) for k, v in seen.items() if len(v) > 1}
    row("WARN" if dups else "PASS", "no Public Sub/Function name in two standard modules",
        "; ".join(f"{k} in {' and '.join(v)}" for k, v in sorted(dups.items())))

    # Option Explicit, non-ASCII
    no_oe, non_ascii = [], []
    for name, m in effective.items():
        body = [ln for ln in m["code"].splitlines() if not ln.startswith("Attribute ")]
        if any(ln.strip() for ln in body) and not re.search(r"^\s*Option Explicit\b", m["code"], re.M | re.I):
            no_oe.append(name)
        for i, ln in enumerate(m["code"].splitlines(), 1):
            if any(ord(c) > 127 for c in ln):
                non_ascii.append(f"{name}:{i}")
    row("WARN" if no_oe else "PASS", "Option Explicit in every module", ", ".join(sorted(no_oe)))
    row("WARN" if non_ascii else "PASS", "VBA code is pure ASCII",
        (f"{len(non_ascii)} line(s), e.g. {', '.join(non_ascii[:6])}. Edit these modules only with a "
         "cp1252-aware script - a UTF-8 editor destroys the characters.") if non_ascii else "")

    # ---------------------------------------------------------------- ribbon
    image_rels = {}
    all_ids = set()
    parts = check_relationships(z)
    for part in parts:
        rels_name = os.path.dirname(part) + "/_rels/" + os.path.basename(part) + ".rels"
        if rels_name in names:
            for r in ET.fromstring(z.read(rels_name)):
                if r.get("Type") == REL_IMAGE:
                    image_rels[r.get("Id")] = os.path.dirname(part) + "/" + r.get("Target", "").lstrip("/")
    for part in parts:
        root = parse_xml(z.read(part), part)
        if root is not None:
            check_ribbon_tree(root, part, part, procs, image_rels, names, all_ids)
    if not parts and not a.xml:
        row("WARN", "ribbon", "no customUI part registered: this add-in has no ribbon of its own")
    for x in a.xml:
        text = open(x, encoding="utf-8-sig").read()
        whole = re.search(r"<customUI\b", text) is not None
        root = parse_xml(text, os.path.basename(x), fragment=True)
        if root is not None:
            # A whole <customUI> file REPLACES the embedded ribbon, so its ids
            # stand alone. A snippet is pasted INTO it, so its ids must not
            # clash with what is already there.
            check_ribbon_tree(root, os.path.basename(x), "snippet", procs, image_rels, names,
                              set() if whole else set(all_ids))

    # ---------------------------------------------------------------- report
    if a.strict:
        ROWS[:] = [("FAIL" if s == "WARN" else s, c, d) for s, c, d in ROWS]
    width = max(len(c) for _, c, _ in ROWS) if ROWS else 10
    width = min(width, 70)
    print()
    for status, check, detail in ROWS:
        lines = (detail or "").split("\n")
        print(f"  [{status}] {check:<{width}}  {lines[0]}")
        for extra in lines[1:]:
            print(f"  {'':6} {'':<{width}}  {extra}")
    fails = sum(1 for s, _, _ in ROWS if s == "FAIL")
    warns = sum(1 for s, _, _ in ROWS if s == "WARN")
    print(f"\n{'=' * 60}\nRESULT: {'FAIL' if fails else 'PASS'}  ({fails} failed, {warns} warnings)")
    return 1 if fails else 0


if __name__ == "__main__":
    sys.exit(main())
