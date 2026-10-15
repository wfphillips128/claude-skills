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
in the input space, and the input that broke a real debt-sizing function in
practice was a *shape* of input (a short live ladder, a flat coupon curve), not
a value. Randomising the dimensions is what finds those.

Hatmaker publishes a free add-in that automates the recalculate-and-count loop;
it is linked from the course page above, along with its user guide. It is the
reason step 5 names the cell `Pass` — that is the name the add-in looks for.

### When Excel is available: run the formula itself

On Windows with desktop Excel 365, the LAMBDA can be evaluated for real — in a
separate, hidden Excel instance driven from Python, never the user's own.
`compose-5g-lambda/scripts/lambda_test_harness_template.py` does this: it lints
the source, installs it into a throwaway workbook, checks the stored text round
trips and fits the 8,192 limit, then compares every cell of seeded random cases
against an independent Python reference and exits non-zero on any failure. How
to adapt it, and the COM traps it already handles, are in
`compose-5g-lambda/reference/lambda-testing.md`.

This is the differential test with Python playing the grid twin, run hundreds of
times in seconds. Prefer it whenever Excel is on the machine: it tests the
formula that ships, not a translation of it. It found a `#CALC!` in its own
worked example on the first run — a one-period timeline — that reading the
formula had missed.

### When we cannot run Excel

Without desktop Excel (or on a machine where automating it is not allowed),
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
for anything going into a library — or run it yourself with the harness when
Excel is available.

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

For a debt schedule, say: a start date late enough that the repayment ladder
runs past the final payoff date, and a ladder entirely inside it. Both ends.

Include the **length-1 timeline**. `DROP`, `TAKE` and `FILTER` return `#CALC!`
when nothing is left, so a function can work for every length except one
(`compose-5g-lambda/reference/pitfalls.md` §20). Random cases rarely land on it.

### 6. Regression

Not from the course, but earned here: **re-run the previous scenario after every
change.** A fix that ties out on the new case while moving the old one is not a
generalisation, it is a second bug cancelling the first.

### 7. Sweep and classify, after a fix

When a fix changes a solver or a schedule, run the **whole scenario grid**
before and after — every start date, every mode — and diff the two runs
scenario by scenario. The result should split cleanly:

- every scenario that was correct before is **bit-identical** after, and
- every scenario that changed was one that was **failing** before, and now
  passes.

A changed scenario that was not failing, or a failing one that did not change,
means the fix is not the fix you think it is. In practice this turned "the new
case ties" into "N of M scenarios unchanged, the rest were exactly the breaching
ones, all now on the covenant" — a far stronger statement.

### 8. Recalculation and speed

Two properties a cell-by-cell comparison never sees:

- **Volatility.** A function containing `RANDARRAY` or `OFFSET` anywhere in its
  call chain recalculates on every edit in the workbook, even when that branch
  never runs (`pitfalls.md` §24). Test it with a no-argument VBA UDF that counts
  its own calls, placed in the same cell: edit an unrelated cell and the count
  must not move. Remove every other volatile source first, or the test is masked.
- **Speed at realistic size.** Time with calculation set to manual and
  `Range.Calculate()` on the cell; automatic mode through COM inflates timings.
  An O(n²) idiom (`pitfalls.md` §25) passes every small test and fails the user.

## Making the tests honest

A test is only worth its independence. Each of these defects passed every check
available at the time, because every check shared its blind spot.

**Compute the expected value by a different route.** A shadow model is a
translation, so it shares the LAMBDA's conceptual errors. A reference written as
a plain loop, from the definition rather than from the formula, does not. Two
routes that start from the same intermediate — the same sheet, the same rounded
values — are one route.

**Assert what the numbers must mean, not only that they add up.** An identity
like *closing = opening − payment* holds whatever the signs are. Add checks that
fail on an impossible answer: a balance that goes negative, a margin over 100%,
principal maturing after the payoff date.

**The reference can be the thing that is wrong.** Any disagreement is a finding
in the function, the reference, or the test design. Known cases: SciPy's
`beta.ppf`, in the version used, was off by about 100x in the extreme tails
(Beta(0.5, 2) at p = 1e-12) while Excel's `BETA.INV` was right, as `mpmath`
confirmed — so validate an
inverse function by **round trip** (`CDF(inverse(p)) = p`), not by comparing
against another inverse. And a dense reference grid built by concatenating
ranges carried duplicated endpoints that made every fit look infeasible — a bug
in the reference, not the LAMBDA.

**Make random cases reach the boundaries.** With an off-by-one planted at the
end of an asset's life, 23 of 61 random cases still passed — every one whose
timeline stopped before the asset's last period. Draw the dimensions so that a
good share of cases cross every boundary, and add the boundary cases by hand.

**Test what is installed, not what you edited.** A workbook holds its own copy
of every defined name. Change the source without reinstalling — or without
bumping a version stamp the installer keys on — and the test workbook keeps the
old definition, so a test can pass against a function that no longer exists, or
fail against one that was already fixed. Read the definition back from the
workbook and compare it with the source (normalised: see `lambda-testing.md`)
before trusting a result.

**Keep "does it tie?" separate from "does it matter?".** Pipeline checks should
be exact, because a non-zero difference means something is losing data and
today's size says nothing about tomorrow's. Whether a difference is *material*
is a separate judgement about the audience. State which question a number
answers, and never widen a tolerance to make a run pass.

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
