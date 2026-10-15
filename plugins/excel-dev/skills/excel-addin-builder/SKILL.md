---
name: excel-addin-builder
description: Build and maintain an Excel VBA add-in (.xlam) - ribbon XML (customUI14), .bas modules, UserForms, the add-in's private settings storage, packaging, and verification before handover. Use when the user wants to add a ribbon button, tab, group, menu or splitButton; edit customUI or customUI14.xml; write or fix IRibbonControl callbacks (onAction, getLabel, getEnabled, getImage, onLoad, invalidate); create or import .bas modules; build a UserForm; store add-in settings inside the add-in (ThisWorkbook, not ActiveWorkbook); turn an .xlsm into an .xlam add-in; or says "my ribbon button does nothing", "my ribbon tab disappeared", "the import made Module1", or "my en-dashes turned into question marks". For making existing macros faster or refactoring them, use excel-macro-optimizer instead.
---

# Excel Add-in Builder

Build and maintain an Excel VBA add-in: its ribbon, its modules, its forms,
its private storage, and the checks that prove it works before the user
imports anything.

**Scope.** This skill owns the add-in's *structure*: ribbon XML and callbacks,
module layout and import, UserForms, settings storage, packaging as `.xlam`,
and verification. The sibling skill **excel-macro-optimizer** owns what goes
*inside* a macro: speed, `FastModeOn`/`FastModeOff`, array I/O, scope rules
(ActiveWorkbook / ActiveSheet / Selection). When a request is "make this
macro faster" or "refactor this macro", hand off to it. When it is "add a
button that runs it", stay here.

**Who you are working with.** Often a citizen developer who knows Excel well
and VBA reasonably, not a professional programmer. Explain *why* each step
exists, define jargon once (callback, relationship part, code module), and
keep each instruction small enough to do and check.

---

## The mental model (teach this first if the user is new to add-ins)

1. **Two workbooks are always in play.**
   `ThisWorkbook` is the add-in itself - its code, its ribbon and its private
   settings sheet. `ActiveWorkbook` is the user's file that the macros act on.
   Add-in storage must always be qualified to `ThisWorkbook`. Mixing them up is
   the classic add-in bug. It either errors, or it quietly writes into a
   same-named sheet in the user's file.

2. **The `.xlsm` is the source and the `.xlam` is a build output.**
   Develop in `MyAddIn.xlsm`, which you can open, see and debug. Then use
   *File > Save As > Excel Add-in (.xlam)* to produce what users install.
   Never hand-edit the `.xlam`, and never treat exported `.bas` files as
   current: they go stale the moment someone edits in the VBA editor.
   **Read the workbook itself** (see step 1 of the workflow).

3. **An `.xlsm` / `.xlam` is a zip file.** The parts this skill touches are:

   | Part in the zip | What it is |
   |---|---|
   | `xl/vbaProject.bin` | all the VBA (modules, forms, ThisWorkbook) |
   | `customUI/customUI14.xml` | the ribbon (Office 2010+ schema) |
   | `_rels/.rels` | must contain a relationship pointing at the ribbon part, or Excel ignores it |
   | `customUI/_rels/customUI14.xml.rels` + `customUI/images/*` | custom icons |
   | `xl/worksheets/*`, `xl/tables/*` | the hidden settings sheet and its tables |

4. **Most add-in failures are silent.** A bad ribbon attribute hides the whole
   tab without an error. A misspelled callback makes a button do nothing. A
   module imported twice becomes `Module1`. An en-dash edited as UTF-8 turns
   into `?`. That is why this skill checks before every handover. Reading the
   code and believing it works is not enough. See *Rules that fail silently*
   below.

---

## Workflow

### 1. Read the real current state - from the workbook

Ask for the path to the `.xlsm` (or `.xlam`). Copy it to a scratch folder
first if there is any chance it is open. Then:

```
python -m oletools.olevba --code "MyAddIn.xlsm"        # every module, form and ThisWorkbook
python scripts/review_addin.py "MyAddIn.xlsm"          # baseline: ribbon + VBA health
```

To read the ribbon XML or the storage tables, open the zip with Python's
`zipfile` and read `customUI/customUI14.xml` and `xl/tables/tableN.xml`. Table
positions change whenever the user edits the sheet, so read them from the
file and do not rely on notes.

Report baseline findings before changing anything. If the baseline already
FAILs, the user should know it was failing before you started.

### 2. Plan the change and confirm anything ambiguous

Decide what the change touches. It might be a new module, changes to an
existing module, a ribbon snippet, a storage setting, or a form. Confirm:
- **Which group or menu** a new button goes in, and its label and icon.
- **New module or existing one.** Prefer one module per feature, with ribbon
  callbacks kept thin (see `reference/ribbon-xml.md` - "thin callbacks").
- **Where data lives.** If it is a user-editable setting, store it on the
  add-in's storage sheet (`reference/addin-storage.md`), not in a `Const`.

### 3. Write the files into a build folder - never over the user's files

Create a `build\` folder (or a dated one) next to the workbook, and produce:

| File | Rules |
|---|---|
| `modX.bas` | Starts with `Attribute VB_Name = "modX"`, then `Option Explicit`. Pure ASCII and CRLF. See `reference/bas-modules.md`. |
| `frmX (paste into form module).txt` | Form code that **builds its controls at run time**, so no binary `.frx` is needed. Template: `templates/frmSettings.txt`. |
| `ThisWorkbook (paste).txt` / `Sheet (paste).txt` | Code for document modules. These cannot be imported, only pasted. |
| `ribbon - <feature>.xml` | A snippet with an XML comment saying exactly where to paste it. No `--` inside comments. |
| `IMPORT.md` | The handover steps, from `templates/IMPORT-template.md`. |

Starting points: `templates/` holds a complete working add-in skeleton (ribbon
template, ribbon module, storage, About and Settings forms, ThisWorkbook
events). All of it was compiled and run together.

### 4. Static check (no Excel)

```
python scripts/review_addin.py "MyAddIn.xlsm" --bas-dir build --pending --xml "build/ribbon - feature.xml"
```

`--pending` means the files are not imported yet, so callbacks are resolved
against them. The script checks that every callback in the snippet exists as a
Public procedure with the right number of parameters. It also checks ids
against the ribbon the snippet will be pasted into, property/callback clashes,
`.bas` encoding (UTF-8 or U+FFFD damage), and missing `VB_Name` headers. It
must say **PASS**.

### 5. Prove it in an isolated Excel

```
python scripts/verify_vba.py "MyAddIn.xlsm" --import build/modX.bas build/modTests.bas `
       --form "frmX=build/frmX (paste into form module).txt" `
       --run modTests.RunAll --ribbon Ribbon_DoX
```

This copies the workbook under a new name and starts its own hidden Excel. It
imports the files, forces a **full** compile (Debug > Compile, which is not
the same as running a macro), runs your test functions and button callbacks,
checks every `imageMso` name, and then closes only the Excel it started.
Write test functions that **return `"PASS ..."` / `"FAIL ..."`** and never
show a MsgBox. Details and traps: `reference/verify-in-isolated-excel.md`.

What it cannot prove is that the ribbon displays. A hidden Excel does not
load ribbon UI (in testing, `onLoad` never fired there). That check belongs to
the user, and the handover must say so.

### 6. Hand over - manual import by default

The default delivery is **files plus `IMPORT.md` with numbered manual steps**.
Many users want to see each step happen in the VBA editor and the RibbonX
editor rather than have a script edit their workbook. The steps are:

1. Back up the `.xlsm`.
2. In the VBA editor, **remove** each module being replaced, answering *No* to
   export. Do this before *File > Import File*, because Import never replaces:
   it adds `modX1` beside `modX`.
3. Import the new `.bas` files. Create forms and paste their `.txt`.
4. *Debug > Compile VBAProject*. It must be clean.
5. Close the workbook in Excel. Paste the ribbon snippet in the Office
   RibbonX Editor, click **Validate**, then save. If Excel still has the file
   open, its next save overwrites the editor's change.
6. Reopen the workbook and click the new button.
7. Bump the version in `modAbout`, save, then *Save As > Excel Add-in*.

Fill in `templates/IMPORT-template.md`. It also records what was verified,
what was not, and how to roll back.

**Scripted alternative (only if the user asks).** `VBProject.VBComponents.Import`
from Python, together with `scripts/inject_ribbon.py` for the ribbon, can do
steps 2-5 unattended. It needs *File > Options > Trust Center > Trust Center
Settings > Macro Settings > Trust access to the VBA project object model*.
Explain that trade-off. Even then, write to a copy and let the user swap it
in.

### 7. After the user imports

```
python scripts/review_addin.py "MyAddIn.xlsm" --bas-dir build
```

Without `--pending`, every file must now match its module in the workbook.
This confirms the import really happened, with nothing skipped and no stale
copy. Then add a banner to the top of `IMPORT.md` saying it was **APPLIED** on
that date, so nobody re-imports old files later.

---

## Rules that fail silently (memorise these)

| Symptom | Cause | Where |
|---|---|---|
| Whole custom tab missing | Malformed XML, duplicate `id`, `label` + `getLabel` on one control, wrong namespace, or the part is not registered in `_rels/.rels` | `reference/ribbon-xml.md` |
| Button does nothing / "Cannot run the macro" | Callback misspelled, `Private`, in a form or class module, defined in two modules, or wrong signature | `reference/ribbon-xml.md` |
| Dynamic label or enabled state never updates | No `onLoad` on `<customUI>`, so there is no `IRibbonUI` to `Invalidate`. Or VBA was reset, which loses it | `reference/ribbon-xml.md` |
| Icon blank | Unknown or wrong-case `imageMso`, or `image=` not embedded | `reference/ribbon-xml.md` |
| `modX1` appears; "Ambiguous name detected" | Imported without removing the old module first | `reference/bas-modules.md` |
| An en-dash or curly quote in a string became `?` or `â€“` (number-format toggles stop matching) | A Windows-1252 `.bas` was edited as UTF-8 | `reference/bas-modules.md` |
| Settings vanish after restart | Wrote to the `.xlam` but never called `ThisWorkbook.Save`. Excel closes add-ins without prompting | `reference/addin-storage.md` |
| Error 1004 "would move cells in a table" | `ListRows.Add` / `.Delete` with another table below that shares only some columns | `reference/addin-storage.md` |
| Dates come back as numbers like 46387 | `.Value2` used on a date cell | `reference/addin-storage.md` |
| `Application.OnTime "Proc"` fails from the add-in | A bare name resolves against the active workbook. Qualify it as `"'" & ThisWorkbook.Name & "'!Module.Proc"` | `templates/modRibbon.bas` |

---

## Scripts

All three are standalone Python 3 (Windows). Pass full paths. If `python` on
PATH is the Microsoft Store stub, call the real interpreter by its full path.

| Script | Opens Excel? | Use it to |
|---|---|---|
| `scripts/review_addin.py` | No | Static PASS/FAIL review of ribbon, callbacks, images, modules and `.bas` files. Run it first and last. Needs `oletools`. |
| `scripts/verify_vba.py` | Yes, its own hidden one | Import, compile, run tests and shims, and check icons in a throwaway copy. Needs `pywin32` and VBA project access trusted. |
| `scripts/inject_ribbon.py` | No | Scripted alternative to the RibbonX Editor. Writes `customUI14.xml`, its relationship and icons into a **new** file. |

Exit codes: 0 means pass, 1 means fail, 2 means could not run. Paste the
PASS/FAIL table into the handover, because it is the evidence.

## Templates

| File | What it is |
|---|---|
| `templates/customUI14-template.xml` | A tab with a group, buttons, separator, toggleButton, menu and splitButton. Shows dynamic getLabel/getEnabled/getImage/getPressed and onLoad |
| `templates/modRibbon.bas` | Every callback the template names, thin, plus `RefreshRibbon` and dialog launchers |
| `templates/add-in-storage.bas` | `AddInStorage`: settings as `Property Get` with fallbacks, a user-editable list, `WriteTableBlock` (Resize-based, restore on failure), `SetupStorage` |
| `templates/app-state-manager.bas` | `FastModeOn/Off/Reset` (owned by excel-macro-optimizer; included so a build is self-contained) |
| `templates/modAbout.bas` | Name, version, licence. One place to bump the version |
| `templates/frmSettings.txt` / `frmAbout.txt` | Forms that build their own controls at run time. Paste into a blank UserForm |
| `templates/ThisWorkbook.txt` | Application events that refresh the ribbon when the active workbook changes |
| `templates/IMPORT-template.md` | The handover document |

Rename before shipping: `ADDIN_NAME`, the tab label, the storage sheet and
table names (one edit block at the top of `AddInStorage`), and the settings.

## References

- `reference/ribbon-xml.md` - controls, callback signatures, dynamic
  updates, icons, and the silent-failure catalogue.
- `reference/bas-modules.md` - encoding, line endings, import order,
  extracting code, and safe patching.
- `reference/addin-storage.md` - the private settings sheet, typed reads,
  the Resize rule, saving, and forms without `.frx`.
- `reference/verify-in-isolated-excel.md` - how `verify_vba.py` works, the
  traps it handles, and how to write tests for it.

## Habits that keep an add-in healthy

- **Diagnose by experiment, not by reading.** When Excel misbehaves and the
  cause is not obvious, write a five-line throwaway test that tries the
  variations. Excel's rules often cannot be worked out from the source, and
  confident inference has been wrong many times where one experiment settled it.
- **Report what was not verified** as plainly as what was: "Ribbon display
  not tested - a hidden Excel does not load ribbons. Open the workbook and
  click the button."
- **One feature per module, one choke point per resource.** Storage goes
  through `AddInStorage`, ribbon state through `RefreshRibbon`, and the
  version through `modAbout`.
- **Never edit the user's original files in place.** Write to `build\`, and
  copy the workbook before reading it while it might be open.
