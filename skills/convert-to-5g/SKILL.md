---
name: convert-to-5g
description: Convert an existing Excel formula (or block of grid formulas) into a 5th-Generation (5G) dynamic-array LAMBDA function following Craig Hatmaker's Beyond Excel standards — self-documenting, spilling, with inline help. Use when the user says "convert to 5G", "make this a LAMBDA", "5G-ify this", "turn this formula into a LAMBDA", pastes an Excel formula and asks to modernize/vectorize it, or wants to replace a row-by-row corkscrew with a spilling function. Also covers writing the help text, the ErrMsgs/Errors/Messages validation block, and producing paste-ready output for the Advanced Formula Environment or Name Manager.
---

# Convert to 5G

Turn 4G formulas (one cell, copied across a grid) into 5G LAMBDA functions (one
formula, spilling across the whole timeline).

**This skill runs in paste-only mode.** Claude Code cannot see the user's Excel
selection or write to the names collection. The user pastes a formula; you return
paste-ready source. Never modify their `.xlsx` — reading it with `openpyxl` is fine,
writing to it is not. If live Excel interaction is genuinely needed, see
`reference/claude-for-excel.md`.

## Workflow

Hatmaker's original process asks 10 questions one at a time. Don't. Infer what you
can, then ask the rest in a single batch.

### 1. Get the formula and its context

The user pastes one or more formulas. If the workbook is on disk, **read it** rather
than asking — `openpyxl` gives you the surrounding layout, the named ranges, and the
precedent cells:

```python
import openpyxl
wb = openpyxl.load_workbook(path)              # formulas, not values
ws = wb["Sheet1"]
print(ws["O19"].value)                          # the formula text
print({k: str(wb.defined_names[k].value) for k in wb.defined_names.keys()})
```

Excel mangles stored formulas. Translate before reasoning about them — see
*Stored-name mangling* below.

### 2. Classify every referenced cell

Each cell the formula points at is one of three things:

| What's in the cell | What it becomes |
|---|---|
| A **value** | An argument to the LAMBDA |
| A **passthrough formula** (`=C3`, nothing else) | Follow the chain — `C3` becomes the argument, not this cell |
| A formula that **calculates** | This cell becomes the argument |

### 3. Work out the array shape

This is the step that separates a real 5G conversion from a cosmetic one.

- If the formula is **copied across a row or down a column** with relative references,
  you are looking at **one period of a timeline**. Convert that one period, then make
  the LAMBDA spill the whole timeline. Do not produce a scalar function the user has
  to copy across — that defeats the point.
- If a referenced cell is a **spill anchor**, pass the whole spill: `G23#`, not `G23`.
- If a referenced cell sits **inside** a spill range but isn't the anchor, trace back
  to the anchor and use `anchor#`.
- Horizontal timelines: prefer `SEQUENCE(1, n)`, `SCAN`, `HSTACK`, `VSTACK`, `DROP`,
  `INDEX`. Never hard-code a single period when the input spills.

### 4. Ask once, batched

Only what you genuinely cannot infer:

- **Function name** — including the library namespace (`fx.DebtSculptλ`). See
  naming rules in `reference/conversion-rules.md`.
- **Which arguments are optional, and each default.**
- **Name for the VERSION line** (skip if you already know it from a prior run).

Do not ask which language to use. Do not ask the user to "select cells and type C".

### 5. Pick the template tier

| Use | When |
|---|---|
| `reference/5g-template-full.md` | **Default for anything reusable.** Help + ErrMsgs + argument coercion + Errors + Messages + `CHOOSE(Return, Result, Help, Messages)`. This is what the user's own `fx.*` library uses. |
| `reference/5g-template-inline-help.md` | Quick one-off conversions, simple scalar helpers. `IF(ISOMITTED(...), TRIM(TEXTSPLIT(help)), LET(...))`. |

When in doubt, use the full template. It costs a few more lines and gives real
argument validation.

### 6. Emit all three output blocks

Non-negotiable — this is what makes paste-only mode workable. See *Output contract*.

### 7. Run `reference/checklist.md`

Every item, before you say it's done.

## Output contract

Every conversion returns these three blocks, in this order.

**① AFE / Excel Labs form** — indented, `/* */` header, `//` comments. This is the
readable, maintainable source and the one the user should keep.

**② Single-line form** — the identical formula flattened to one line with all comments
stripped. `//` comments are illegal in an Excel formula, so form ① can never be
installed as-is. Always supply both. Verify form ② contains no `//` and no `/*`.

Optional parameters are written `[Arg]` here — **one** bracket pair, never `[[Arg]]`.
Excel reads a leading `[` as the start of a structured table reference, so a doubled
bracket fails to parse on the first argument. The `[` and `]` counts must each equal
the parameter count exactly.

> **Name Manager cannot install a long function, and this is not a formula problem.**
> Its *Refers to* box has its own input limit: past roughly a thousand characters the
> text will not paste in at all. Excel never parses it, so no amount of shortening
> prose, stripping the help table or fixing syntax makes any difference. A ~130-char
> LAMBDA carrying every 5G feature pastes cleanly; a 5,400-char one will not enter the
> box. Name Manager also cannot create a **recursive** function at all, because a name
> cannot reference itself before it exists.
>
> So do not present form ② as "paste this into Name Manager" unless it is genuinely
> short. State the install route instead — see `reference/claude-for-excel.md`.

**③ Test plan** — concrete, not generic:
- The exact cell to paste the call into (pick empty space below/right of the original).
- The call with real arguments, using `#` for spill anchors.
- Which existing cells the output must equal, and the tolerance (usually $0.01).
- A one-line comparison formula, e.g. `=MAX(ABS(newRange - oldRange))`.

## Stored-name mangling

`openpyxl` and raw XML show Excel's internal prefixes. They are **not** what the user
types, and the optional-argument prefix is a useful signal:

| Stored | Means |
|---|---|
| `_xlfn.LAMBDA`, `_xlfn.LET`, `_xlfn.TEXTSPLIT` | Post-2019 function — strip `_xlfn.` |
| `_xlop.CFADS` | **Optional** LAMBDA parameter — declared `[CFADS]` |
| `_xlpm.Help` | LET name, or a required LAMBDA parameter |
| `_xlfn.ANCHORARRAY('Sheet'!$F$46)` | The spill reference `$F$46#` |
| `_xlfn.SINGLE(...)` | Implicit intersection `@` |

Because 5G declares *every* argument optional (so help text can display), a correctly
built 5G function shows `_xlop.` on all of its parameters.

## Hard rules

- Never write to the user's workbook. Read-only, always.
- The delimiters are real Unicode: `→` (U+2192) column, `¶` (U+00B6) row. Not `->`, not `|`.
- The function-name suffix is `λ` (U+03BB).
- All arguments are `[bracketed]` in the LAMBDA definition. Required vs optional is
  signalled in the **help text**, not the definition. (Rationale: an argument Excel
  thinks is required blocks the function from running at all, so the help text never
  displays.)
- If the converted function errors, say so and show the error. Do not iterate silently.

## References

- `reference/5g-template-full.md` — the full 5G structure, transcribed from working code
- `reference/5g-template-inline-help.md` — the lightweight template
- `reference/conversion-rules.md` — naming, array idioms, help-text mechanics
- `reference/checklist.md` — run before reporting done
- `reference/claude-for-excel.md` — when to hand off to a tool that's inside Excel

For building a **new** function out of existing 5G functions rather than converting a
grid formula, use the `compose-5g-lambda` skill instead.
