# VBA optimization patterns (before / after)

Recipes to apply when refactoring. Each says *why* it matters. Pair all of
these with the `FastModeOn`/`FastModeOff` skeleton in `SKILL.md`.

---

## 1. Batch range I/O through arrays — the #1 win

Every `.Cells(r, c)` read or write is a COM round-trip across the VBA↔Excel
boundary. On 100k rows that's hundreds of thousands of crossings. Read once,
work in memory, write once.

**Before** (slow — two COM calls per row):
```vba
Dim r As Long
For r = 2 To lastRow
    ws.Cells(r, 3).Value = ws.Cells(r, 1).Value * ws.Cells(r, 2).Value
Next r
```

**After** (fast — one read, one write):
```vba
Dim data As Variant, out() As Variant, r As Long, n As Long
data = ws.Range("A2:B" & lastRow).Value2   ' 2-D array, 1-based: data(row, col)
n = UBound(data, 1)
ReDim out(1 To n, 1 To 1)
For r = 1 To n
    out(r, 1) = data(r, 1) * data(r, 2)
Next r
ws.Range("C2").Resize(n, 1).Value2 = out   ' single write-back
```

Notes:
- Arrays from `.Value2` are **1-based**, indexed `(row, column)`.
- `.Value2` is faster than `.Value` and avoids Date/Currency coercion. Use
  `.Value` only when you specifically need dates as `Date` or currency typing.
- Write-back range must match the array's dimensions — use `.Resize(rows, cols)`.

---

## 2. Single-cell gotcha

If a range is exactly one cell, `.Value2` returns a **scalar**, not a 2-D array,
and your `UBound` blows up. Guard it:

```vba
Dim data As Variant
data = rng.Value2
If Not IsArray(data) Then
    Dim tmp(1 To 1, 1 To 1) As Variant
    tmp(1, 1) = data
    data = tmp
End If
```

---

## 3. Reliable last-row / last-column detection

Don't trust `UsedRange` (it drifts and includes formatted-but-empty cells) and
never hardcode row counts.

```vba
' Last used row in column A:
Dim lastRow As Long
lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

' Last used column in row 1:
Dim lastCol As Long
lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
```

Note the leading-dot qualification (`ws.Rows.Count`) — works for both .xls
(65,536) and .xlsx (1,048,576) without assuming a version.

---

## 4. Replace lookups-in-loops with a Dictionary — O(n²) → O(n)

Nested loops or `WorksheetFunction.Match` inside a loop rescan data every
iteration. Build a `Scripting.Dictionary` index once.

**Before** (slow — Match rescans the whole column every row):
```vba
For r = 2 To lastA
    res = Application.Match(ws.Cells(r, 1).Value, otherRange, 0)
    ...
Next r
```

**After** (fast — one pass to index, O(1) lookups):
```vba
Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
Dim keys As Variant: keys = wsLookup.Range("A2:A" & lastB).Value2
Dim vals As Variant: vals = wsLookup.Range("B2:B" & lastB).Value2
Dim i As Long
For i = 1 To UBound(keys, 1)
    dict(keys(i, 1)) = vals(i, 1)        ' last-wins; use If Not .Exists for first-wins
Next i

Dim src As Variant: src = ws.Range("A2:A" & lastA).Value2
For i = 1 To UBound(src, 1)
    If dict.Exists(src(i, 1)) Then ' ... use dict(src(i, 1))
Next i
```

`CreateObject("Scripting.Dictionary")` (late binding) needs no reference — good
for a portable add-in. Early binding (Tools ▸ References ▸ Microsoft Scripting
Runtime) gives IntelliSense if you prefer; only add the reference deliberately.

---

## 5. Kill Select / Activate

`Select`/`Activate` are slow, fragile, and depend on UI focus. Act on references.

**Before:**
```vba
Sheets("Data").Select
Range("A1").Select
Selection.Font.Bold = True
```
**After:**
```vba
wb.Worksheets("Data").Range("A1").Font.Bold = True
```

---

## 6. With blocks — resolve an object once

**Before** (re-resolves the range 4×):
```vba
ws.Range("A1").Font.Bold = True
ws.Range("A1").Font.Size = 12
ws.Range("A1").Interior.Color = vbYellow
ws.Range("A1").HorizontalAlignment = xlCenter
```
**After:**
```vba
With ws.Range("A1")
    .Font.Bold = True
    .Font.Size = 12
    .Interior.Color = vbYellow
    .HorizontalAlignment = xlCenter
End With
```

---

## 7. Cache resolved objects out of loops

**Before:** `Workbooks("Big.xlsx").Worksheets("Data")` re-resolved every pass.
**After:** `Set ws = Workbooks("Big.xlsx").Worksheets("Data")` once, then use `ws`.

---

## 8. Disable autofilter / sort / formatting churn during bulk work

Operations that trigger recalcs or redraws per change are best done with
`Calculation = xlCalculationManual` (handled by `FastModeOn`) and a single
`Application.Calculate` at the end if needed.

---

## 9. The three-scope shared-core pattern (eventual DRY goal)

Old macros often have three near-identical copies — one each for Workbook,
Sheet, and Selection. The clean end state: a **core sub that takes a `Range`**
(or a collection of ranges) and does the work, plus three thin wrappers that
resolve the target and call it. Refactor toward this *only when asked* — for
incremental work, optimize one version in place and keep its signature.

```vba
' Core: does the actual work, scope-agnostic. Fast mode is owned by the wrapper.
Private Sub Core_DoThing(ByVal rng As Range)
    Dim area As Range, data As Variant
    For Each area In rng.Areas          ' handles multi-area selections for free
        data = area.Value2
        ' ... process data, write back to area ...
    Next area
End Sub

' Wrapper 1 — Selection
Public Sub DoThing_Selection()
    On Error GoTo CleanExit
    If TypeName(Selection) <> "Range" Then Exit Sub
    Dim rng As Range
    Set rng = Application.Intersect(Selection, Selection.Worksheet.UsedRange)
    If rng Is Nothing Then Exit Sub
    FastModeOn
    Core_DoThing rng
CleanExit:
    FastModeOff
    If Err.Number <> 0 Then MsgBox Err.Description, vbExclamation
End Sub

' Wrapper 2 — Sheet
Public Sub DoThing_Sheet()
    On Error GoTo CleanExit
    FastModeOn
    Core_DoThing ActiveWorkbook.ActiveSheet.UsedRange
CleanExit:
    FastModeOff
    If Err.Number <> 0 Then MsgBox Err.Description, vbExclamation
End Sub

' Wrapper 3 — Workbook
Public Sub DoThing_Workbook()
    On Error GoTo CleanExit
    Dim ws As Worksheet
    FastModeOn
    For Each ws In ActiveWorkbook.Worksheets
        Core_DoThing ws.UsedRange
    Next ws
CleanExit:
    FastModeOff
    If Err.Number <> 0 Then MsgBox Err.Description, vbExclamation
End Sub
```

Benefits: fix a bug once, behavior stays identical across all three scopes, and
the multi-area `For Each area` loop means the selection version handles
non-contiguous ranges automatically.

---

## 10. Add-in private storage — reference `ThisWorkbook` through one accessor

The add-in keeps its **own** settings on a hidden worksheet and its LAMBDA
functions in an Excel table. Both live in **`ThisWorkbook`** (the add-in), not
the user's **`ActiveWorkbook`**. Qualifying these to `ActiveWorkbook` (or leaving
them bare) is a classic add-in bug — it errors or silently grabs a same-named
sheet in the user's file. Because some features are migrating from an existing
sheet to a new one, route every access through a single accessor so the switch
is one edit.

```vba
' --- One place that knows where add-in storage lives ---
Private Const REFERENCE_SHEET As String = "shtReference"      ' rename here only
Private Const COMPANY_TABLE   As String = "tblCompany"
Private Const LAMBDA_TABLE    As String = "tblAlonzoChurch"

' The add-in's hidden reference sheet — ALWAYS from ThisWorkbook, never Active.
Private Function ReferenceSheet() As Worksheet
    Set ReferenceSheet = ThisWorkbook.Worksheets(REFERENCE_SHEET)
End Function

' The LAMBDA-definitions table.
Private Function LambdaTable() As ListObject
    Set LambdaTable = ReferenceSheet().ListObjects(LAMBDA_TABLE)
End Function

' Read a user-customizable list from a table column into an array to loop over,
' replacing values that used to be hardcoded in the macro.
Public Function CompanyList() As Variant
    CompanyList = ReferenceSheet().ListObjects(COMPANY_TABLE) _
                    .ListColumns("Company").DataBodyRange.Value2
End Function
```

Hardcoded lists baked into a macro can't be customized by the people you share
the add-in with. Moving them to a table on the reference sheet makes them
user-editable with no code change. The full drop-in version of this — generic
column reader, blank-trimming, empty-table guard, typed setting readers and safe
writers — is `../excel-addin-builder/templates/add-in-storage.bas`.

Why this shape:
- **`ThisWorkbook`, captured in one function.** No macro hardcodes
  `Worksheets("...")`; a hidden-sheet rename or the existing→new migration is a
  one-line change instead of a find-and-replace across modules.
- **Read through the object, never `Activate`.** The sheet is hidden (and the
  whole add-in is hidden in `.XLAM` form) — activating it is unnecessary and
  fragile, and would flash the user's screen.
- **Batch-read the table** with `.DataBodyRange.Value2` into an array (same as
  any range, §1), not a `ListRows` loop.

**Migration note:** while a feature still points at the *existing* sheet, give it
its own constant/accessor (e.g. `SETTINGS_SHEET_OLD`) rather than overloading
one name — then retire it when the migration completes. Confirm with the user
which sheet a given macro should target before flipping it.

---

## Review checklist — flag these when auditing a macro

- [ ] `Option Explicit` present
- [ ] No `Select` / `Activate` / `Selection` (unless selection is genuine input)
- [ ] No bare `Range`/`Cells`/`Rows`/`Columns` — all qualified to a sheet/wb
- [ ] App state toggled **and restored on every exit path** (cleanup label)
- [ ] Cell-by-cell loops replaced with array I/O
- [ ] Lookups-in-loops replaced with a Dictionary
- [ ] Row counters typed `Long`, not `Integer`
- [ ] Last row/col detected dynamically, not hardcoded
- [ ] `.Value2` used unless Date/Currency typing is required
- [ ] Single-cell range handled (IsArray guard) if range size can be 1
- [ ] `On Error GoTo` with a single cleanup exit — no leaked state on error
- [ ] Targets (`ActiveWorkbook`/`ActiveSheet`/`Selection`) captured once at
      entry, not re-read in helpers
- [ ] **Add-in storage:** the add-in's own settings / LAMBDA table qualified to
      **`ThisWorkbook`**, never `ActiveWorkbook` or bare; reached through the
      single accessor + name constant (§10), not hardcoded per macro; settings
      sheet never `Activate`d/`Select`ed (it's hidden)
- [ ] **Selection scope:** `TypeName(Selection) = "Range"` guarded; whole
      column/row selections trimmed via `Intersect(sel, ws.UsedRange)`;
      multi-area (`Areas.Count > 1`) handled or explicitly unsupported
- [ ] Macro's scope (workbook/sheet/selection) preserved — not silently merged
- [ ] No volatile functions / recalcs inside loops
- [ ] 64-bit safe (`PtrSafe`/`LongPtr` on any `Declare` API calls)
