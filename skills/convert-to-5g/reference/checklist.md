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
- [ ] The last calculation is named, and that name is the LET's return step
- [ ] No angle brackets `<>` left anywhere — those were template placeholders

## Help text

- [ ] `&` present at the end of every concatenated line **except the last**
- [ ] `TEXTSPLIT` delimiters in the right order: `"→"` (column) then `"¶"` (row)
- [ ] Delimiters are the real characters — `→` U+2192, `¶` U+00B6 — not `->` or `|`
- [ ] Multi-row descriptions use a leading `→` on continuation lines
- [ ] VERSION line carries the user's name and the current date
- [ ] REVISIONS block carries the user's name and the current date

## Full-template functions only

- [ ] `ErrMsgs` entries separated by `;` (vertical array), one per validated argument
- [ ] `Errors` built with `VSTACK` in the **same order** as `ErrMsgs`
- [ ] `Messages, FILTER(ErrMsgs, Errors, "")`
- [ ] `Return, IF(AND(Errors), 2, IF(OR(Errors), 3, IF(OR(ISERROR(Result)), 2, 1)))`
- [ ] Final line is `CHOOSE(Return, Result, Help, Messages)`
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

## Delivery

- [ ] **Three blocks emitted**: AFE form, Name Manager form, test plan
- [ ] Name Manager form contains **zero** `//` and zero `/* */` comments
- [ ] Name Manager form is a single line
- [ ] Test plan names a specific target cell, a specific comparison range, and a
      tolerance
- [ ] The user's workbook was not modified
