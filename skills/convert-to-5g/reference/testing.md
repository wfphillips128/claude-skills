# Testing a 5G function

Standard 13 is "thoroughly tested before releasing to the public." This is what
the course means by that, and it is a good deal more than one tie-out.

Source: <https://sites.google.com/site/beyondexcel/home/5g-modeling/5g-course/5g-testing-components>

Hatmaker on his own hit rate, which is the reason this file exists:

> I am always humbled by this process as it seems my tests always fail at least
> once for simple functions, and dozens of times for complex functions. When my
> functions finally do pass at least a thousand tests, I feel assured it is
> probably right.

## The core method: differential testing

Not "does the function look right" and not "does it match one benchmark", but:
**build a second, dumb implementation and make the two disagree.**

1. Catalogue every argument and decide the realistic range of each.
2. Generate random inputs with `RANDARRAY`, `RANDBETWEEN` or `RAND` — randomise
   the *dimensions* too, not just the values.
3. Write an **alternative solution that is neither a LAMBDA nor a dynamic
   array** — the ordinary grid formulas, copied across, that the 5G function is
   meant to replace.
4. Compare with an explicit tolerance: `=ABS(G5 - G19) <= 0.0001`.
5. `AND()` all the comparisons into one cell and **name that cell `Pass`**.
6. Recalculate hundreds of times. Any failure is a finding — in the function, in
   the alternative, or in the test design. Fix and repeat until it never fails.

The reason this is stronger than a benchmark tie-out: a benchmark is one point
in the input space, and the input that broke `fx.DebtSizeIssueλ` in practice was
a *shape* of input (a short live ladder, a flat coupon curve), not a value.
Randomising the dimensions is what finds those.

Hatmaker publishes a free add-in that automates the recalculate-and-count loop;
it is linked from the course page above, along with its user guide. It is the
reason step 5 names the cell `Pass` — that is the name the add-in looks for.

### When we cannot run Excel

Claude Code cannot execute a LAMBDA. The **Python shadow model** is the stand-in
for the non-dynamic-array twin: same role in the argument, different host. See
`compose-5g-lambda/reference/shadow-model-guide.md`.

The two are complementary, and the shadow model is weaker in exactly one way —
it is written by the same person, at the same time, from the same understanding
as the LAMBDA, so a conceptual error can appear identically in both. That is
precisely what happened with `Interest = Balance x COUPON(period)`, which tied
out perfectly against a flat coupon and was wrong by millions against a curve.
Randomised inputs catch that; a faithful translation does not.

So: shadow-model here, and recommend the randomised differential test in Excel
for anything going into a library.

## The six test classes

### 1. Inline help

The cheapest test there is. Type `=Funcλ()` with no arguments and read it.

- Is it help for *this* function, or was it pasted from another and not edited?
- Is it aligned — is a delimiter missing?
- Is the first line the function name and syntax?
- Is the webpage line right, if there is one?
- Is the version date current?
- Does every argument appear, with `(Required)` or `(Optional)` and a
  description?
- Is the last line an example?
- Is everything spelled correctly?

Copy-paste help text from a sibling function is the single most common defect
this catches.

### 2. Edge

Things the user should not do but will:

- invalid dates
- negative amounts where only positives make sense
- text in a numeric argument
- codes that are not in the valid list
- zero where a divisor is expected (`ParValueIncr` of 0)

> For invalid inputs, 5g Functions should fail with one of Excel's error
> conditions.

Failing loudly is the requirement. A function that quietly returns a plausible
number for a nonsense input is worse than one that returns `#VALUE!`.

### 3. Orientation

Dynamic arrays spill both ways. If the function is meant to work either way,
test it both ways — a function written around `COLUMNS()` and
`INDEX(x, 1, n)` will silently misbehave on a vertical input.

If it is *not* meant to work both ways, validate the orientation and error out.
Our templates do this with `IF(COLUMNS(Arg) <= 1, #VALUE!, Arg)`.

### 4. Array vs range

> In general, 5g functions are for ranges. But some can be used inside formulas
> and never spill into a range. This can cause some Excel functions to fail. The
> xIFS functions (SUMIFS, COUNTIFS, etc.) only work over ranges.

So if the function is intended to be used *inside* another formula rather than
spilled to the grid, test it there. `SUMIFS(Funcλ(...), ...)` fails in a way
that `=Funcλ(...)` in a cell never reveals.

### 5. Timeline boundaries

> When testing functions in timelines, we should consider how our function
> should respond when the start date is before the timeline start and the end
> date is after the timeline end.

For this library specifically: a delivery date late enough that the tranche
ladder runs past the payoff date, and a ladder entirely inside it. Both ends.

### 6. Regression

Not from the course, but earned here: **re-run the previous scenario after every
change.** A fix that ties out on the new case while moving the old one is not a
generalisation, it is a second bug cancelling the first.

## Reporting

Print the variance table — period, model, benchmark, difference — with a Total
row. A schedule can tie period by period and miss in total if a period was
dropped, and it can tie in total while individual periods offset.

State the tolerance used: `$0.01` on dollars, `1e-9` on rates, **exact** on par
amounts that have been snapped to an increment.

If it does not tie, say which periods fail and hand back the variance. Do not
describe a function as working when it is not.

## Certification

The course proposes a self-signed certificate — a declaration that a named
version of a library meets every 5g requirement, published alongside it:
<https://sites.google.com/site/beyondexcel/home/5g-modeling/5g-course/5g-self-certification>

There is no third-party certifier yet. Its practical value is as a checklist:

> A self signed certificate can provide a check list by which we check our own
> work.

`checklist.md` serves that purpose here.
