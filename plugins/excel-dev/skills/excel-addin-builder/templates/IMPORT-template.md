<!--
  Handover template. Fill every <angle-bracket> field, delete sections that do
  not apply, and delete this comment. Keep the steps numbered and small: the
  reader does them one at a time in the VBA editor and the RibbonX Editor.
  When the user confirms the import is done, add the APPLIED banner below.
-->

<!-- > **APPLIED - shipped in <Add-in name> v<x.y> on <yyyy-mm-dd>.**
> The workbook is now the source of truth. These files are a point-in-time
> record of this change and go stale the moment the workbook is edited.
> Do not re-import them. -->

# <Add-in name> - <feature or change name>

<One paragraph: what this adds or fixes, in the user's terms. What they will
see on the ribbon afterwards.>

**Back up `<MyAddIn>.xlsm` before starting.** Copy it somewhere safe. If an
import goes wrong, restoring that copy is the undo.

## Files

| File | Action |
|---|---|
| `<modFeature>.bas` | **import** - new module |
| `<AddInStorage>.bas` | **import** - replaces the existing `AddInStorage` (remove it first) |
| `<frmFeature> (paste into form module).txt` | **paste** into a new UserForm named `<frmFeature>` |
| `ThisWorkbook (paste).txt` | **paste** into `ThisWorkbook` (replace what is there) |
| `ribbon - <feature>.xml` | **paste** into `customUI14.xml` with the RibbonX Editor |
| `<file>` | **skip** - <why> |

## Steps

### A. VBA editor (Alt+F11, with `<MyAddIn>.xlsm` open)

1. If the add-in is ticked in *File > Options > Add-ins*, untick it and
   restart Excel, so the `.xlam` and the `.xlsm` are not both loaded.
2. In the Project Explorer, select the `<MyAddIn>.xlsm` project.
3. **Remove** each module being replaced: right-click `<AddInStorage>` >
   *Remove AddInStorage...* > **No** (do not export).
   *Why: File > Import File never replaces a module. It adds
   `AddInStorage1` beside the old one, and every procedure becomes
   ambiguous.*
4. *File > Import File...*, then select `<modFeature>.bas`. Repeat for each
   `.bas` in the table.
5. For each form: *Insert > UserForm*, press F4, set *(Name)* to
   `<frmFeature>`, double-click the form, delete anything in its code window,
   and paste the whole `.txt`. **Do not place any controls.** The form builds
   its own when it opens.
6. Double-click `ThisWorkbook`, select all, and paste
   `ThisWorkbook (paste).txt`.
7. *Debug > Compile VBAProject*. **It must finish with no message.** If it
   stops on an error, note the module and line and stop here.
8. Save the `.xlsm`, then close Excel.

### B. Ribbon (Office RibbonX Editor, with Excel closed)

9. Open `<MyAddIn>.xlsm` in the Office RibbonX Editor and select
   `customUI14.xml`.
10. Find `<the element named in the snippet's comment>` and paste the snippet
    `<immediately above / below / before </group>>`.
11. Click **Validate**. It must report the XML is well formed. *Why: one bad
    attribute hides the entire tab when Excel loads it, with no error
    message.*
12. Save, then close the editor.

### C. Try it

13. Open `<MyAddIn>.xlsm`. The `<tab label>` tab shows `<new control>` in
    `<group>`.
14. <The exact click-through: select X, click Y, expect Z.>
15. <Any setting to change in Settings, then quit Excel completely, reopen,
    and confirm it was kept.>

### D. Build the add-in

16. In `modAbout`, set `ADDIN_VERSION` to `"<x.y>"`. Save.
17. *File > Save As*, choose *Excel Add-in (\*.xlam)* as the type, and save
    over the installed copy (by default
    `%APPDATA%\Microsoft\AddIns\<MyAddIn>.xlam`).
18. Re-tick the add-in in *File > Options > Add-ins* if you unticked it in
    step 1.

## What changed, and why

| Was | Now | Why |
|---|---|---|
| <old behaviour> | <new behaviour> | <reason, in one line> |

## Checked before handover

Static review (`review_addin.py`, no Excel):

```
<paste the PASS/FAIL table>
```

Compiled and run in a separate, hidden Excel on a copy (`verify_vba.py`):

```
<paste the PASS/FAIL table>
```

**Not checked - please confirm by hand:**
- The ribbon displays and the new button responds. A hidden Excel does not
  load ribbons.
- <Form layout looks right on screen.>
- <Anything else not covered.>

## If something goes wrong

- **The tab disappeared**: the ribbon XML has an error. Reopen it in the
  RibbonX Editor and click Validate. Or restore the backup.
- **"Ambiguous name detected"**: a module was imported without removing the
  old one first. Remove the one ending in `1`.
- **A button does nothing**: send the output of
  `python review_addin.py "<MyAddIn>.xlsm"`.
- **Anything else**: restore the backup from the top of this page.
