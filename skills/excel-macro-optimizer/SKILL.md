---
name: excel-macro-optimizer
description: Optimize, refactor, or write Excel VBA macros for speed, reliability, and maintainability — especially macros that live in an .XLAM add-in and run against the active workbook. Invoke when the user shares VBA/macro code to make faster or more robust, asks to fix slow Excel automation, or wants new add-in macros written from scratch. Triggers on "optimize this macro", "refactor VBA", "my macro is slow", "speed up Excel", "write a macro that...".
---

# Excel Macro Optimizer

Refactor Excel VBA for speed, reliability, and maintainability, and write new
macros from scratch. **Assume large datasets** unless the user says otherwise.

The target environment is the user's own **Excel add-in (`.XLAM`)**: macros are
invoked from the add-in's menu and act on objects in the **active workbook** the
user is looking at. This shapes every recommendation below — see *Add-in rules*.

**Default working mode: small, incremental, scope-preserving refactors.** The
user maintains this add-in macro by macro. A given macro usually exists at one
of **three scope levels** — and often there are three sibling versions of the
same operation:

1. **ActiveWorkbook** — act on every worksheet in the workbook
2. **ActiveSheet** — act on one sheet
3. **Selection** — act on what the user has selected (a *first-class input*,
   not disposable)

Older macros duplicate logic across these three rather than sharing a core sub.
**Do not impose a DRY rewrite unless the user asks.** The eventual goal is a
shared core that takes a `Range`/scope argument (see *Scope levels* below and
`reference/optimization-patterns.md` §9), but for now optimize the macro in
front of you, preserve its scope and entry signature, and at most *mention* the
shared-core opportunity as a future step.

## How to respond

1. **Propose one optimized rewrite** that preserves the same final outputs.
   Internal steps and structure may change for efficiency; results may not.
2. **After the code, explain in plain language:**
   - Performance improvements (what changed and why it's faster)
   - Tradeoffs
   - Assumptions made
   - Excel version compatibility (only if relevant)
3. **Teach, don't just deliver** — say *why* each change matters so the user
   learns the pattern for the next macro.

## Before changing assumptions, ASK

If the macro relies on something implicit, confirm before changing it:
- **Which scope is this?** Workbook, sheet, or selection. Never silently change
  it. If unclear from the code, ask.
- Reliance on the active sheet vs. a specific named sheet
- Existing calculation mode the user may depend on
- Excel version / platform (Windows vs Mac; 64-bit `PtrSafe`)
- Whether events firing during the run are intentional
- **Which storage sheet — existing or new?** Some of the add-in's settings /
  LAMBDA functionality is being migrated from an existing worksheet to a new
  one. When a macro touches the add-in's *own* storage, confirm whether it
  should read the **existing** or the **new** sheet before changing it.

Rules of engagement:
- **No speculative changes.** Optimize what's there; don't add features.
- **No new dependencies** (references, libraries) unless the user asks.
- **Preserve `Selection` as input when the macro is selection-scoped** — it's
  the user's chosen target, not incidental UI state.
- **Don't preserve cosmetic UI behavior** (scroll position, active cell,
  re-selecting after a run) *unless it affects the output*. It usually doesn't.
  Restoring the selection the user started with is a reasonable courtesy but not
  required.

## The optimization toolkit (in rough order of impact)

1. **Batch range I/O through arrays.** The single biggest win. Read a whole
   range into a `Variant` array, process in memory, write back in one shot.
   Each `.Cells(r, c)` touch is a COM round-trip; arrays cut thousands of them
   to two. See `reference/optimization-patterns.md`.
2. **Toggle application state** — `ScreenUpdating`, `Calculation`,
   `EnableEvents`, `DisplayStatusBar`, page breaks — and **restore it safely on
   every exit path**. Use the drop-in helper in
   `reference/app-state-manager.bas`.
3. **Eliminate `Select` / `Activate`.** Operate on object references directly.
4. **Replace lookups-in-loops** (`WorksheetFunction.Match`/`VLOOKUP`, nested
   `For` loops) with a `Scripting.Dictionary` keyed once — turns O(n²) into O(n).
5. **`With` blocks** to resolve an object once instead of per line.
6. **Fully qualify references** (`wb.Worksheets("X").Range(...)`), never bare
   `Range`/`Cells` from a module — they silently bind to whatever is active.
7. **`.Value2`** instead of `.Value` for read/write (skips Date/Currency
   coercion; faster).
8. **Kill volatility & redundant recalc** — avoid `INDIRECT`/`OFFSET`/`NOW` where
   avoidable; calculate once at the end, not inside loops.
9. **Right-size data types** — `Long` for row counters (`Integer` overflows at
   32,767, and worksheets have ~1M rows), `Boolean`/`Double` deliberately.
10. **Resolve objects once** — cache `Worksheets("X")`, `Workbooks(...)` in a
    variable instead of re-resolving the collection each iteration.

## Add-in rules (these are easy to get wrong)

An `.XLAM` runs in its own context but acts on *someone else's* workbook, so:

- **Capture targets at the entry point, per scope.** First lines bind the
  target once and pass references down — never let helper subs re-read
  `ActiveSheet`/`Selection`, which can change mid-run:
  - Workbook scope: `Set wb = ActiveWorkbook`, then loop `wb.Worksheets`.
  - Sheet scope: `Set ws = ActiveWorkbook.ActiveSheet`.
  - Selection scope: `Set sel = Selection` **before** anything else runs, and
    validate it (see *Selection-scope rules*).
- **Bare `Range`/`Cells`/`Rows` in an add-in module refer to nothing useful** —
  there's no host sheet. Always qualify to the captured workbook/sheet.
- **App-state changes are global.** Turning off `ScreenUpdating`/`EnableEvents`
  affects every open workbook. If the macro errors out with `EnableEvents =
  False`, events stay dead for the user's whole session. Restoration on error
  is non-negotiable — hence the cleanup-label pattern below.
- **Save and restore the *prior* state, not a hardcoded default.** The user may
  already be in manual calc; don't force them back to automatic. The helper
  module captures the real prior values.
- **Validate the active object.** Guard against no workbook open, a chart sheet
  where a worksheet is expected, a protected sheet, or an empty range.

### `ThisWorkbook` vs `ActiveWorkbook` — the add-in's own private storage

Two different workbooks are in play at all times, and conflating them is a
top-tier add-in bug:

- **`ActiveWorkbook`** — the *user's* workbook: the **data the macro operates
  on**. The three scope levels above all resolve from here. Default to this for
  anything the macro reads, transforms, or writes as output.
- **`ThisWorkbook`** — the **add-in itself** (`XL Edge`). It holds the add-in's
  *own* private resources, which the user's workbook does **not** contain:
  - a **hidden settings worksheet**, and
  - an **Excel table storing the add-in's LAMBDA functions**.

Rules:
- **Any read/write of the add-in's settings or LAMBDA table MUST be qualified to
  `ThisWorkbook`** — never `ActiveWorkbook`, never bare. `ActiveWorkbook` will
  either error (the user's file has no such sheet) or, worse, silently hit a
  same-named sheet the user happens to have open. `ThisWorkbook` always points
  at the add-in regardless of what's active.
- **Never `Activate`/`Select` the settings sheet to read it.** It's hidden — and
  in `.XLAM` form the whole add-in workbook is hidden. Go through the object
  reference: `ThisWorkbook.Worksheets(SETTINGS_SHEET).Range(...).Value2`.
- **Centralize the reference behind one accessor.** Because the storage sheet is
  mid-migration (some features use the existing sheet, some a new one), route
  every access through a single `Private Function` + name constant so the
  existing→new switch is one edit, not a hunt across macros. See
  `reference/optimization-patterns.md` §10.

## Scope levels

The three sibling versions resolve a target and then do the same work. When
refactoring *one* version, keep its scope; don't merge the three unasked.

**Selection-scope rules** (the trickiest, because `Selection` is unpredictable):
- **`Selection` may not be a range.** Guard `If TypeName(Selection) <> "Range"
  Then Exit Sub` (it could be a chart, shape, or drawing).
- **Multi-area selections.** A non-contiguous selection has `.Areas.Count > 1`.
  Either loop `For Each area In sel.Areas` or decide the macro only supports a
  single area and say so.
- **Whole-column / whole-row selections.** Selecting column A gives ~1,048,576
  cells. **Intersect with `UsedRange`** before reading into an array, or you'll
  pull a million blank rows: `Set rng = Intersect(sel, ws.UsedRange)`.
- Capture `Set sel = Selection` on the first line, before `FastModeOn` or
  anything that could move the selection.

**The eventual shared-core pattern** (offer, don't force): a core sub takes a
`Range` and three thin wrappers feed it from each scope. See
`reference/optimization-patterns.md` §9 for the full template.

## Mandatory skeleton for every macro

Single entry point, single exit, guaranteed cleanup. Pick the target-capture
block that matches the macro's scope.

```vba
' --- SELECTION SCOPE ---
Public Sub DoTheThing_Selection()
    On Error GoTo CleanExit

    If TypeName(Selection) <> "Range" Then Exit Sub
    Dim sel As Range, ws As Worksheet, rng As Range
    Set sel = Selection                          ' capture before anything moves it
    Set ws = sel.Worksheet
    Set rng = Application.Intersect(sel, ws.UsedRange)  ' trim blank rows/cols
    If rng Is Nothing Then Exit Sub              ' selection was entirely empty

    FastModeOn                                    ' from app-state-manager.bas
    ' --- work on rng (loop rng.Areas if multi-area matters) ---

CleanExit:
    FastModeOff                                   ' restores prior state, always runs
    If Err.Number <> 0 Then _
        MsgBox "Error " & Err.Number & ": " & Err.Description, vbExclamation, "DoTheThing"
End Sub

' --- SHEET SCOPE ---  Set ws = ActiveWorkbook.ActiveSheet, then act on ws.
' --- WORKBOOK SCOPE --- Set wb = ActiveWorkbook, then For Each ws In wb.Worksheets.
' (Same On Error / FastModeOn / CleanExit frame in all three.)
```

Always start modules with `Option Explicit`.

## References

- **`reference/app-state-manager.bas`** — drop this module into the add-in once.
  Provides `FastModeOn` / `FastModeOff` that capture and restore real prior
  state and are re-entrancy safe (nested macros won't restore early). Use it in
  every macro instead of hand-toggling `Application.*`.
- **`reference/add-in-storage.bas`** — drop this module into the add-in once.
  The single home for the add-in's *own* storage on the hidden **reference
  sheet**: `ReferenceSheet()`, `LambdaTable()`, `CompanyTable()` (all from
  `ThisWorkbook`), plus the list readers `TableColumnList(lo, col)` /
  `CompanyList()` / `HasItems()` and a `StorageReady()` guard. **Prefer a
  data-driven table over hardcoded lists** — e.g. `CompanyList()` replaces
  hardcoded company/department names so end users can customize them. Route
  every reference-sheet access through this module — never via `ActiveWorkbook`
  or a hardcoded `Worksheets("...")`. Sheet/table names live in one edit block
  at the top, including the legacy `*Old()` accessors (`ReferenceSheetOld`,
  `LambdaTableOld`) for the in-progress existing→new migration. Key|value
  settings can be added the same way when a settings table exists.
- **`reference/optimization-patterns.md`** — before/after recipes: array I/O,
  dictionary lookups, reliable last-row detection, single-cell array gotcha,
  the three-scope shared-core template (§9), and a review checklist.

Read the reference files when refactoring so the rewrites stay consistent with
the add-in's existing helpers.
