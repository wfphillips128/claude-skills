# The add-in's private storage, settings, and forms

An add-in usually needs to remember things: settings the user can change, a
list they curate, and lookup tables. The robust place for that is **a
worksheet inside the add-in itself**. Code: `templates/add-in-storage.bas`
(module `AddInStorage`). Form: `templates/frmSettings.txt`.

---

## 1. Why a sheet inside ThisWorkbook

| Option | Problem |
|---|---|
| `Const` in a module | Changing it means opening the VBA editor, so users cannot. |
| Registry (`SaveSetting`) | Invisible, per-machine, and hard to inspect or back up. |
| A sheet in the **user's** workbook | Not there in the next workbook they open. Pollutes their file. |
| A separate config file | One more file to ship and locate. Path handling. |
| **A sheet in the add-in (`ThisWorkbook`)** | Travels with the add-in, is visible while developing in the `.xlsm`, and is hidden automatically in the `.xlam`. Tables make it structured. |

An `.xlam` has no visible window, so its sheets cannot be seen or selected.
They are fully readable and writable through code, though. Develop in the
`.xlsm`, where you can see and edit the storage sheet, then *Save As* the
add-in.

**Always `ThisWorkbook`, never `ActiveWorkbook`.** `ActiveWorkbook` is
whatever file the user is looking at. Code like
`ActiveWorkbook.Worksheets("Settings")` either errors (their file has no such
sheet) or, worse, silently reads and writes a same-named sheet in their file.
`ThisWorkbook` always means the workbook that contains the running code, which
is the add-in.

**One choke point.** Every read and write goes through `AddInStorage`. The
sheet and table names live in one edit block at the top of it. A rename is
then a one-line change, and no macro anywhere contains `Worksheets("...")`.

## 2. Layout

`AddInStorage.SetupStorage` (run once from the Immediate window in the
`.xlsm`) creates this, and leaves alone anything that already exists:

```
Sheet "AddInData"
     B            D                    E
2    Item         Setting              Value          <- table headers
3    Alpha        REPORT_TITLE         Monthly Report
4    Beta         PERIOD_END           12/31/2026
5    Gamma        DEFAULT_ROW_HEIGHT   15
6                 SHOW_TIPS            TRUE
     tblList      tblSettings
```

Tables (ListObjects) rather than loose ranges make the data self-describing,
let it grow without hardcoded addresses, and let code find it by name.

**Put tables side by side, not stacked**, so each can grow downward without
meeting another. Stacking is what produces the 1004 error in section 4.

## 3. Settings: read like constants, stored on the sheet

```vba
Public Property Get DEFAULT_ROW_HEIGHT() As Double
    DEFAULT_ROW_HEIGHT = SettingDouble("DEFAULT_ROW_HEIGHT", DEF_DEFAULT_ROW_HEIGHT)
End Property
```

- **Same name as the `Const` it replaces.** A `Public Property Get` in a
  standard module is visible everywhere a `Public Const` was, so every
  existing `ws.Rows.RowHeight = DEFAULT_ROW_HEIGHT` compiles unchanged.
  Migrating a constant means deleting its `Const` line and nothing else.
- **A fallback for everything.** A missing sheet, missing key, blank cell or
  wrong type (someone typed "fifteen") returns the private `DEF_*` default.
  A damaged settings sheet must never stop the add-in from running.
- **Typed readers** (`SettingDouble`, `SettingDate`, `SettingString`,
  `SettingBool`) do the `IsNumeric` / `IsDate` guarding in one place.
- **Cached.** Settings are read into a `Scripting.Dictionary` (created
  late-bound with `CreateObject`, so no library reference is needed) the
  first time they are used. **Every writer must call `InvalidateCache`.**
  If a user might edit the sheet by hand while developing, add this backstop
  to the storage sheet's own code module:
  ```vba
  Private Sub Worksheet_Change(ByVal Target As Range)
      On Error Resume Next
      If Not Intersect(Target, Me.ListObjects("tblSettings").Range) Is Nothing Then AddInStorage.InvalidateCache
  End Sub
  ```
  This is only a backstop. It cannot fire while a macro has
  `EnableEvents = False`, so the writers' own `InvalidateCache` call is the
  part that matters.
- **`.Value`, not `.Value2`, for reading settings.** `.Value2` returns a date
  as its serial number (12/31/2026 comes back as 46387). That does not error.
  It hands a number to code expecting a date, and shows up later as a wrong
  report. (excel-macro-optimizer recommends `.Value2` for bulk data for speed,
  which is right for numbers but wrong for anything that may hold a date.)

## 4. Writing lists: Resize, never ListRows.Add

```vba
AddInStorage.SetItemList Array("North", "South", "East")   ' replace the whole list
```

All row-count changes go through `WriteTableBlock`. It rewrites the whole
table as one block and sizes it with **`ListObject.Resize`**.

**Why not `ListRows.Add` / `ListRows.Delete`?** They shift cells down or up
inside the table's own columns. Excel moves another table below out of the
way only if it spans **exactly** the same columns. If the table below is
wider, narrower or offset, it would be split, so Excel refuses with run-time
error **1004: "This won't work because it would move cells in a table on your
worksheet."** (Tested: same-width below is moved; wider or offset below
raises 1004.) Storage sheets crowd tables together, so sooner or later this
bites. `Resize` only redeclares where the table starts and ends. No cell
moves, so neighbours are never at risk.

`WriteTableBlock` also:
- **checks every entry before touching the table.** A 40,000-character value
  cannot go in a cell (the limit is 32,767). Refusing it after the table was
  cleared would leave it empty, and the next save would make that permanent.
- **refuses to grow into a neighbour.** It returns False with a message naming
  the table in the way, and does not overwrite it.
- **keeps a copy and puts it back** if the write fails partway.
- handles the scalar trap: `Range.Value` of a single cell is a scalar, not a
  1x1 array.

**Storing long text or formula text** (templates, code snippets, formulas
kept as data):
- Set the column's number format to Text **before** writing values. Otherwise
  Excel tries to evaluate `=...` text as a formula on the way in.
- Excel was observed refusing to bulk-write text over roughly 8,200
  characters that starts with `=` ("Error 7 - Out of memory"). Check lengths
  before the table is cleared, the same way `WriteTableBlock` checks the cell
  limit.
- Do not turn on `WrapText` for long-text columns. Excel lays out every cell
  to work out the row height, rows stretch to the 409-point maximum, and the
  sheet becomes unusable. Pin `RowHeight` instead.

All writers return True/False and leave a readable reason in
`AddInStorage.LastError` rather than raising, so the form decides how to
report.

## 5. Save, or lose it

```vba
AddInStorage.SaveStorage        ' = ThisWorkbook.Save
```

Writing to a loaded `.xlam` marks it changed, but **Excel closes add-ins at
shutdown without asking to save.** Skip `SaveStorage` and the user's settings
silently vanish at the end of the session. The writers deliberately do not
save, so a form can batch ten writes and save once. The form's Save button
calls `SaveStorage`.

Test it the way a user would: change a setting, **quit Excel completely**,
reopen it, and check the setting is still there.

If the add-in is installed somewhere read-only (a network share without
write access), saving fails, and the settings last only until Excel closes.
Say so in the handover if that is how it will be deployed.

## 6. Forms without a .frx: build the controls in code

A UserForm exports as **two** files: `.frm` (code, as text) and `.frx` (the
control layout, as binary). You cannot write a `.frx` by hand, and a `.frm`
without it imports as an empty form. So the templates build their controls
when the form opens:

```vba
Private WithEvents cmdSave As MSForms.CommandButton     ' WithEvents = it can fire Click

Private Sub UserForm_Initialize()
    Set cmdSave = Me.Controls.Add("Forms.CommandButton.1", "cmdSave", True)
    With cmdSave
        .Caption = "Save": .Left = 296: .Top = 350: .Width = 84: .Height = 24
    End With
End Sub

Private Sub cmdSave_Click()            ' works because cmdSave was declared WithEvents
```

Install: in the VBE use *Insert > UserForm*, set *(Name)* in the Properties
window (F4), open its code window, and paste the `.txt`. Place no controls by
hand. The whole layout stays readable text that can be diffed, reviewed and
pasted. The cost is that the designer shows a blank form, and you only see
the controls when it runs.

Patterns the templates use, and why:

| Pattern | Why |
|---|---|
| The form edits **working copies** (arrays, a Collection), and writes to the sheet only on Save | So Cancel really cancels |
| Validate on Update (date, number, yes/no) | A typo on the sheet would silently fall back to the default. In the form it can be rejected with a message |
| Buttons call `Me.Hide`, not `Unload Me`. `QueryClose` turns the title-bar X into a hide | Unloading destroys the form's variables before the caller can read `f.Saved` |
| The caller creates an explicit instance: `Set f = New frmSettings: f.Show vbModal: ... : Unload f` | The default instance (`frmSettings.Show`) keeps state between uses and is hard to reason about |
| Save is not `.Default` | Otherwise Enter in a text box saves and closes the whole form |
| Titles come from `modAbout.ADDIN_NAME` | Rename the add-in in one place |

## 7. Facts that must NOT be editable

The licence text and version live in code (`templates/modAbout.bas`), not on
the storage sheet. Anything on the sheet is user-editable by design.

`LicenseText()` builds the text with one assignment per paragraph. VBA allows
at most 24 `& _` line continuations in one statement, and legal text is
exactly what grows past that.
