# Pre-delivery checklist

Run every item before saying the conversion is done. If one fails, fix it — don't ship
with a caveat.

## Structure

- [ ] Function name identical everywhere: LAMBDA definition, help text FUNCTION/EXAMPLE
      line, and any recursive self-call
- [ ] Name ends in `λ` (U+03BB) and carries its library namespace if it belongs to one
- [ ] **All** arguments bracketed in the LAMBDA definition — `[Argument]`
- [ ] Help text FUNCTION/EXAMPLE line: required args **without** brackets, optional args
      **with**
- [ ] Every optional parameter has default-value logic in the LET section
      (`IF(ISOMITTED(X), default, X)`)
- [ ] The last calculation is named, and **the LET's final step is that bare
      name, not an expression** (Standard 10 — it's what makes the function
      debuggable). If dropped for stored-length reasons, say so explicitly
- [ ] No angle brackets `<>` left anywhere — those were template placeholders

## Self-contained (Standard 06)

The function has to run in a workbook that contains nothing but the function.

- [ ] No cell references — grep for `$`
- [ ] No table references — grep for `tbl`
- [ ] No named-range references
- [ ] No VBA or other scripting
- [ ] Nothing inside a `BYROW`/`BYCOL`/`MAP`/`SCAN` callback reaches out to a
      range — helpers get their data passed in as arguments
- [ ] Every name in the body is a LAMBDA parameter, a LET step, a native Excel
      function, or another 5g function **in the same module** (the one exception)

## Help text

- [ ] `&` present at the end of every concatenated line **except the last**
- [ ] `TEXTSPLIT` delimiters in the right order: `"→"` (column) then `"¶"` (row)
- [ ] Delimiters are the real characters — `→` U+2192, `¶` U+00B6 — not `->` or `|`
- [ ] Section order: `FUNCTION`, `DESCRIPTION`, `WEBPAGE` (if any), `VERSION`,
      `PARAMETERS`, `EXAMPLE`
- [ ] A `FUNCTION:` line exists and shows the syntax — required args bare,
      optional args bracketed. Without it the user cannot tell which is which
- [ ] Every argument appears under `PARAMETERS:` with `(Required)` or `(Optional)`
- [ ] Multi-row descriptions use a leading `→` on continuation lines
- [ ] Any space that must survive `TRIM` is U+2003 EM SPACE
- [ ] VERSION line carries the user's name and the current date
- [ ] REVISIONS block carries the user's name and the current date
- [ ] **Read the help table as the user will see it.** Is it this function's
      help, or pasted from a sibling and half-edited? Is it aligned? Spelled
      correctly?

## Full-template functions only

- [ ] `ErrMsgs` entries separated by `;` (vertical array), one per validated argument
- [ ] `Errors` built with `VSTACK` in the **same order** as `ErrMsgs`
- [ ] `Messages, FILTER(ErrMsgs, Errors, "")`
- [ ] `Return, IF(AND(Errors), 2, IF(OR(Errors), 3, IF(OR(ISERROR(Result)), 2, 1)))`
- [ ] Dispatch is `Output, CHOOSE(Return, Result, Help, Messages),` followed by
      `Output`
- [ ] Each validated argument is rebound to itself with its coercion

## Syntax

- [ ] All parentheses balanced — check `TEXTSPLIT`, `TRIM`, `IF`, `LET`, `LAMBDA`
- [ ] Booleans uppercase: `TRUE`, `FALSE`
- [ ] Function names uppercase: `INDEX`, `SCAN`, `REDUCE`, `VSTACK`, `HSTACK`
- [ ] `SEQUENCE(1, n)` used for horizontal `SCAN`/`REDUCE` output
- [ ] `DROP(array, , -1)` preferred over `TAKE(array, 1, n-1)`

## Conversion fidelity

- [ ] Every formula in the user's selection is represented in the output
- [ ] Passthrough references followed to the true input cell
- [ ] Spill anchors referenced with `#`
- [ ] If the original repeated across a timeline, the function **spills** the timeline —
      it isn't a scalar the user must copy across
- [ ] Output shape and row order match the original block

## Tests to recommend

Run them with `compose-5g-lambda/scripts/lambda_test_harness_template.py` when desktop
Excel is available; otherwise the test plan should name the ones that apply. See `testing.md`.

- [ ] **Inline help** — `=Funcλ()` with no arguments, read the table
- [ ] **Edge** — text in a numeric argument, negatives, invalid dates, a zero
      divisor. Invalid input must fail with an Excel error, not a plausible number
- [ ] **Orientation** — if it should work on both row and column input, test both;
      if not, validate the orientation and error out
- [ ] **Array vs range** — if it will be used *inside* another formula rather than
      spilled, test it there. `SUMIFS` and friends only accept ranges
- [ ] **Timeline boundaries** — start before the timeline starts, end after it ends
- [ ] **Length-1 timeline** — `DROP`/`TAKE`/`FILTER` to empty returns `#CALC!`
      (`compose-5g-lambda/reference/pitfalls.md` §20)
- [ ] **Differential** — a non-LAMBDA, non-dynamic-array twin over randomised
      inputs, compared with a tolerance. This is the one that finds shape bugs

## Delivery

- [ ] **Three blocks emitted**: AFE form, single-line form, test plan
- [ ] Single-line form contains **zero** `//` and zero `/* */` comments
- [ ] Single-line form is a single line
- [ ] Optional parameters written `[Arg]`, never `[[Arg]]`. **Count brackets in
      the code only, not inside string literals** — a `FUNCTION:` help line that
      marks its optional arguments legitimately puts `[` and `]` inside a
      string, so a raw count of the whole formula will exceed the parameter
      count. Strip string literals first, then `[` and `]` must each equal the
      parameter count
- [ ] **Install route matches the size**: Name Manager only if short, otherwise AFE or
      programmatic. Never tell the user to paste a long function into the *Refers to*
      box — it will not go in. Recursion is not a factor
- [ ] Verification is "evaluate it in a cell", not "check Name Manager shows it"
- [ ] Test plan names a specific target cell, a specific comparison range, and a
      tolerance
- [ ] The user's workbook was not modified
