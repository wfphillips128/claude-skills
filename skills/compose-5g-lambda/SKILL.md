---
name: compose-5g-lambda
description: Build a NEW 5G Excel LAMBDA by composing or extending LAMBDAs that already exist in a workbook — the case where a working function needs a variant and naive edits break the recursion. Use when the user says "create a new LAMBDA based on", "extend this LAMBDA", "add a feature to my 5G function", "my recursive LAMBDA doesn't work", "make a variant of", or has a half-finished LAMBDA that returns wrong numbers or #NUM!. Includes inventorying an existing library from an .xlsx, choosing the right 5G source structure, and proving the maths with a Python shadow model tied out to a benchmark before any formula is handed over.
---

# Compose a 5G LAMBDA

Build a new 5G function out of ones that already work. This is harder than converting a
grid formula, because the failure mode is silent: a recursive LAMBDA that returns
plausible numbers that are wrong.

**Paste-only.** Never write to the user's workbook. Read it with `openpyxl`, hand back
source for them to paste.

**The core discipline: nothing ships until it ties out.** You cannot execute a LAMBDA
from here, so a Python shadow model stands in for Excel. A function that hasn't been
reconciled against a benchmark is not finished, and saying otherwise is the one failure
this skill exists to prevent.

## Workflow

### 1. Inventory the existing library — never guess a signature

```powershell
python `
  "scripts/dump_lambdas.py" "<workbook.xlsx>"
```

Prints every defined-name LAMBDA with its `_xlfn.`/`_xlop.`/`_xlpm.` prefixes stripped,
re-indented, plus a signature table marking which parameters are optional. Read the
output before writing a line of formula. Argument **order** and **count** matter — a
recursive self-call with the arguments shuffled compiles fine and returns garbage.

### 2. Write the spec before any formula

State plainly, and get the user to confirm anything you inferred:

- **Arguments** — name, scalar or per-period array, required or optional, default
- **What one period does** — the calculation chain for period *n*
- **Output rows** — the exact `VSTACK` order the function returns, and what each row is
- **Termination** — how the recursion stops
- **Errors** — which arguments get validated and the message for each
- **Self-containment** — confirm every input arrives through an argument. No cell,
  name or table references; the only outside call allowed is another 5g function
  in the same module. See `convert-to-5g/reference/5g-standard.md`.

If the new function differs from the existing one by a single concept (say, a debt
reserve offset), name that concept and describe precisely where it enters the chain.
Most broken variants come from inserting it at the wrong step.

### 3. Choose the source structure

See `reference/source-structures.md`. Five options; two matter most:

- **Recursive HSTACK with a period counter** — what the `fx.*` library uses. Each
  call computes one period's column and stacks itself for the next.
- **Convergence recursion** — for circularity (interest depends on principal depends on
  interest). Calls itself with the previous iteration's value until the change is
  below tolerance.

A function can use both: an outer period recursion whose interest step calls an inner
convergence recursion. That is the classic debt-sculpting shape.

Default to matching the structure the existing library already uses. Consistency beats
elegance in a library someone else maintains.

### 4. Compose, don't re-derive

Call the existing helper (`fx.InterestDebtIssuanceλ`) rather than inlining its
maths. If the helper genuinely can't serve the new case, say so explicitly and propose a
sibling helper with its own name — don't quietly change the shared one, which would
alter every function already calling it.

### 5. Build the Python shadow model

Copy `scripts/shadow_model_template.py`. Mirror the recursion **step for step** — same
variable names, same order, same rounding. It is a translation, not a reimplementation;
if you "improve" the maths here, the tie-out proves nothing.

Feed it the workbook's real inputs via `openpyxl`. Print the period-by-period schedule.

See `reference/shadow-model-guide.md`.

### 6. Tie out — this gates delivery

Compare the shadow output against a benchmark:

- A **benchmark PDF** (an existing sizing run from a financial advisor), or
- The **workbook's current grid** (the rows the LAMBDA is meant to replace), or
- A **hand calculation** for the first two or three periods

Print a variance table: period, model, benchmark, difference. State the tolerance
(`$0.01` on dollar amounts, `1e-9` on rates).

**A staged tie-out finds bugs faster than an all-at-once one.** Reconcile the base case
first — the new function with its new feature switched off should reproduce the existing
function exactly. Only then turn the feature on.

If any period is outside tolerance, go back to step 4. Do not deliver.

### 7. Deliver

Three blocks, same contract as `convert-to-5g`:

1. **AFE / Excel Labs form** — indented, commented
2. **Single-line form** — comments stripped (`//` is illegal in an Excel formula).
   Optional parameters are `[Arg]`, never `[[Arg]]`. Say how to install it: Name
   Manager only if the function is short — its *Refers to* box will not accept long
   text, whatever the formula says. Otherwise AFE or programmatic installation.
   Recursion is not a constraint either way. Verify by evaluating in a cell, not by
   checking Name Manager's display
3. **Paste plan** — the exact cell, the call with real arguments and `#` spill anchors,
   the cells the output must equal, the tolerance

Plus the **tie-out table** the function passed. That's the audit trail.

### 8. Report honestly

If it ties, say so with the variance table. If it doesn't, show the variance and say
which periods fail — do not hand over a function described as working when it isn't.

## Templates and rules

The 5G structure, naming, help-text mechanics and bracket conventions live in the
sibling skill and are not duplicated here:

- `../convert-to-5g/reference/5g-standard.md`
- `../convert-to-5g/reference/5g-template-full.md`
- `../convert-to-5g/reference/conversion-rules.md`
- `../convert-to-5g/reference/checklist.md`
- `../convert-to-5g/reference/testing.md`
- `../convert-to-5g/reference/hatmaker-library.md`

Read the full template before writing the function; run the checklist before delivering.

**Start with `hatmaker-library.md`.** Composing a new function is only worth doing
once you know it does not already exist — about 120 5G functions are published and
importable, and the catalogue also shows the naming families (`SumXλ` / `LabelXλ`,
and the `A`/`R`/`C` array/row/column suffixes) that a new variant should follow to
sit naturally beside them.

## References

- `reference/hatmaker-templates.md` — **the five published templates verbatim.**
  They exist on the Beyond Excel site only as screenshots; this is the code from
  the gist that backs them, with the performance ranking, the selection rules,
  and the bugs the published versions carry. Read it before choosing a structure.
- `reference/source-structures.md` — the five 5G source structures, written out
- `reference/pitfalls.md` — **read this first when debugging a broken variant**
- `reference/shadow-model-guide.md` — faithful Python mirroring and tie-out presentation
- `scripts/dump_lambdas.py` — inventory a workbook's LAMBDA library
- `scripts/shadow_model_template.py` — starting point for the shadow model

Both scripts are read-only against the workbook.
