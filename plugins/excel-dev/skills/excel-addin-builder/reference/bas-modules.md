# .bas modules: encoding, line endings, import, extraction

A `.bas` file is a text export of one standard module. The VBA editor (VBE)
writes one with *File > Export File* and reads one with *File > Import File*.
Simple as that sounds, the format has traps that silently damage code. This
page covers them.

## 1. The header

```vba
Attribute VB_Name = "modReports"
Option Explicit
```

- `Attribute VB_Name` is how Import knows the module's name. Without it the
  file imports as `Module1`. The VBE hides this line once the module is
  imported, and you never type it inside the editor.
- `Option Explicit` makes every variable need a `Dim`. Without it a typo
  (`totl = totl + x`) silently creates a new variable, and the macro returns a
  wrong number instead of failing to compile. Put it in every module, and tick
  *Tools > Options > Require Variable Declaration* so new modules get it
  automatically.
- Class modules (`.cls`) carry more `Attribute` lines (`VB_PredeclaredId`
  and others). Export one from the VBE once and use it as the pattern rather
  than writing the header by hand.

## 2. Encoding: Windows-1252, not UTF-8

The VBE writes and reads `.bas` files in the system's ANSI code page,
**Windows-1252** on English Windows. It does not understand UTF-8. Two
consequences:

**a. Editing an exported module with a UTF-8 tool destroys characters.**
An en-dash is byte `0x96` in Windows-1252, and that byte is not valid UTF-8.
A tool that reads the file as UTF-8 replaces it with U+FFFD, the replacement
character, and writes the file back that way. The original character is gone
from the file, not just displayed wrongly.

This matters most where it is least visible. Accounting number formats use an
en-dash as the zero placeholder:

```
_(#,##0_);(#,##0);_("<en-dash>"_)        <- in the .bas file, <en-dash> is the single byte 0x96
```

If that byte is destroyed, every `If rng.NumberFormat = FMT_ACCT Then`
comparison quietly stops matching, and a format-toggle macro stops toggling
with no error at all. The safe way to write such a constant in new code is:

```vba
FMT_ACCT = "_(#,##0_);(#,##0);_(""" & ChrW(&H2013) & """_)"   ' a variable set at run time
```

`ChrW` cannot appear in a `Const`, so use a module-level variable set once,
or a function that returns the string.

**b. A file saved as UTF-8 imports as mojibake.** A UTF-8 en-dash is three
bytes (`E2 80 93`). Imported as Windows-1252, those bytes become three
characters: `â€“`.

**Rules:**
1. **Author new modules in pure ASCII.** Use `--` instead of an em-dash and
   straight quotes instead of curly quotes. If a string needs a non-ASCII
   character, build it at run time: `ChrW(&H2013)` is an en-dash. Then no
   tool can damage the file. Every template in this skill is pure ASCII.
2. **Before editing an existing export, check it for high bytes**:
   ```python
   data = open(path, "rb").read()
   print([ (i, hex(b)) for i, b in enumerate(data) if b > 127 ][:20])
   ```
   An empty list means the file is safe to edit normally. Otherwise, patch it
   with a script that reads and writes `encoding="cp1252"`, and do not use a
   general text editor or editing tool that assumes UTF-8.
3. `review_addin.py --bas-dir` FAILs a file containing U+FFFD bytes
   (`EF BF BD`, damage already done), a UTF-8 BOM, or UTF-8 multi-byte
   characters. It also WARNs about non-ASCII characters in the workbook's own
   modules, so you know which modules need care.

## 3. Line endings: check before you patch

The VBE exports with CRLF (`\r\n`). Files written by other tools are often
bare LF (`\n`). In testing, the VBE imported an LF-only module correctly, so
import is not the problem. **Patch scripts are.** A pattern written with
`\r\n` matches nothing in an LF file, and one written with `\n` leaves stray
`\r` characters in a CRLF file.

```python
raw = open(path, "rb").read()
crlf, lf = raw.count(b"\r\n"), raw.count(b"\n") - raw.count(b"\r\n")
print("CRLF" if crlf and not lf else "LF" if lf and not crlf else "MIXED")

text = open(path, encoding="cp1252", newline="").read()   # newline="" keeps \r\n intact
```

Without `newline=""`, Python's text mode converts line endings on read and
write, and the "patched" file comes back with different endings than it had.
Normalise once (to CRLF, for consistency with the VBE), then patch.

## 4. Getting the current code out of a workbook

The workbook is the source of truth. Exported files are snapshots. Read the
code straight from the file:

```
python -m oletools.olevba --code "MyAddIn.xlsm"                 # everything, as text
python -m oletools.olevba --code "MyAddIn.xlsm" > all-modules.txt
```

From Python:

```python
from oletools.olevba import VBA_Parser
vp = VBA_Parser("MyAddIn.xlsm")
for _, _, filename, code in vp.extract_macros():   # filename like "modRibbon.bas", "frmSettings.frm"
    print(filename, len(code))
vp.close()
```

olevba decodes the project's code page correctly, so en-dashes survive. It
reads `.xlam` files too, and does not need Excel. It also recovers UserForm
**code**. It does not recover the binary control layout, which is one more
reason to build forms in code (`reference/addin-storage.md`, section 6).

## 5. Importing: remove first, in dependency order

**File > Import File never replaces a module.** If `modReports` already
exists, importing `modReports.bas` creates `modReports1` beside it. Every
procedure then exists twice, and you get "Ambiguous name detected" pointing
somewhere unhelpful. Always:

1. Right-click the old module, choose **Remove modX...**, and answer **No**
   to "export before removing?" (the workbook copy is what you are replacing).
2. *File > Import File...* the new `.bas`.

`review_addin.py` FAILs a workbook that contains `modFoo1` next to `modFoo`.

**Order.** VBA compiles the whole project, so import order rarely matters for
correctness. It does matter for keeping each step checkable. Import shared
infrastructure first (`AddInStorage`, `AppStateManager`), then feature
modules, then `modRibbon`, then forms and `ThisWorkbook`. Run *Debug > Compile
VBAProject* at the end, and it must be clean. If a module needs another one
that is not there yet, compile tells you which.

**What cannot be imported, only pasted:** the code of `ThisWorkbook`, sheet
modules, and UserForms built in code. Ship those as
`Name (paste into ... module).txt` files with numbered instructions at the
top (see `templates/frmSettings.txt`).

**Scripted import** (only when the user wants it):

```python
vbp = wb.VBProject                      # needs "Trust access to the VBA project object model"
vbp.VBComponents.Remove(vbp.VBComponents("modReports"))
vbp.VBComponents.Import(r"build\modReports.bas")
```

`verify_vba.py` does exactly this on a throwaway copy, and FAILs if the module
lands under any name other than the one in its `VB_Name` line.

## 6. Handover: IMPORT.md

Every delivery includes `IMPORT.md`, built from `templates/IMPORT-template.md`.
Its job is to let the user do the import alone, one checkable step at a time.
It lists:
- every file, with its action (import, replace, paste, skip and why),
- the steps in order: back up, remove, import, paste, compile, ribbon,
  version bump, Save As `.xlam`,
- what changed and why, in plain language,
- what was verified (paste the PASS/FAIL output) and **what was not**,
- how to roll back.

Once the user confirms it is applied, put a banner at the top: "APPLIED in
v1.3 on 2026-10-15 - the workbook is now the source of truth, do not
re-import these files." Old `.bas` files sitting next to a workbook are an
invitation to re-import stale code.

## 7. Module hygiene checklist

- [ ] `Attribute VB_Name` matches the intended module name.
- [ ] `Option Explicit` is present.
- [ ] Pure ASCII, or the file is deliberately Windows-1252 and handled with
      cp1252-aware tooling.
- [ ] Ribbon callbacks are `Public`, live in standard modules, and are named
      with a prefix (`Ribbon_*`).
- [ ] No `Public` procedure name is repeated in another standard module.
- [ ] Add-in storage is reached only through `AddInStorage` (`ThisWorkbook`),
      never through `ActiveWorkbook`.
- [ ] No `MsgBox` inside a procedure a test needs to call. Split the UI
      (callback) from the work (plain procedure).
