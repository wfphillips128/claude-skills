# Pitfalls when composing 5G LAMBDAs

Read this first when a variant returns wrong numbers, `#VALUE!`, or `#NUM!`. Sections
1-14 were found in real half-finished debt-sculpting LAMBDAs; sections 15-27 are traps
in the Excel engine itself, found while building and testing larger LAMBDA libraries.

> **When the error points nowhere, stop reading the formula and run an experiment.**
> Excel's rules are not derivable from the source text — it refuses things that look
> fine and accepts things that look identical. See §27.

Items marked *(from experience — not re-verified)* were observed while building real
libraries but not reproduced in the 2026-10 test pass. Treat them as strong hints and
confirm in your own Excel before relying on them.

---

## 1. Lookahead — period *n* cannot see period *n+1*

**The single biggest source of wrong-but-plausible results.**

```excel
NewDebtService, Payment + Interest,
DebtReserve,    MAX(NewDebtService),      // ← wrong
```

At period *n*, `NewDebtService` is a **scalar** — this period's value. `MAX()` of a
scalar is that scalar. The author meant *peak debt service across the whole schedule*,
but the schedule doesn't exist yet: periods *n+1* onward haven't been computed.

The result looks right at the peak period and is wrong everywhere else, so it survives
a spot check.

**Fixes, in order of preference:**

1. **Two-pass.** Run the recursion once to get the schedule, take the peak from the
   completed array, then run it again with the peak as an argument. Cheap, obvious,
   auditable.
2. **Pass the peak in as an argument.** If the user already knows it (from a prior
   sizing run), make it a parameter with a documented default.
3. **Solve it outside.** Compute the peak in a cell and reference it.

Do **not** try to make the recursion look forward. It can't.

The same trap wearing a different hat: `INDEX(DebtService, 1, MAX(Period))` — `Period`
is a scalar, so `MAX(Period)` is just `Period`.

---

## 2. The recursive call points at the wrong function

```excel
fx.DebtSculptIssuance2λ = LAMBDA( ...
    Result, IF(Period = ModelPeriods, NewColumn,
        HSTACK(NewColumn, fx.DebtSculptIssuanceλ( ... )))   // ← calls the OLD one
```

Copy a function to make a variant and the self-call still names the original. Excel
resolves it happily. Period 1 uses the new logic; every period after it uses the old
logic. This is the most under-diagnosed bug in the family — the output *is* a schedule,
it just isn't the one you asked for.

**Check every self-call by name after any copy-and-rename.** `dump_lambdas.py` prints
the recursion target for exactly this reason.

---

## 3. Arity mismatch on the recursive call

```excel
fx.InterestDebtIssuanceλ = LAMBDA([CFADS], [DSCR], [Principal], [InterestRate],
                                       [DBTRSVPCT], [DBTRSVRATE], [ParValueIncr], [DoNotUse],
    ...
    fx.InterestLRVλ(CFADS, DSCR, Principal, InterestRate,
                         DBTRSVPCT, DBTRSVRATE, ParValueIncr, -Repayment)  // ← 8 args
```

`InterestLRVλ` takes **five**. Because 5G brackets every argument as optional, Excel
does not reject the extra ones loudly — the convergence value lands in the wrong slot
and the iteration never settles, or settles on nonsense.

**Always confirm the callee's arity and argument order from `dump_lambdas.py` output.**

---

## 4. Mismatched output heights break `HSTACK`

`HSTACK` needs every column to have the same number of rows. If the variant's
`NewColumn` stacks five rows but the function it recurses into stacks four, you get
`#VALUE!` — or worse, `#N/A` padding that quietly propagates.

```excel
NewColumn, VSTACK(Principal, Interest, -Payment, ClosingDebt, DebtReserve),  // 5 rows
// but the recursive target returns 4
```

Adding an output row to a recursive function means adding it to **every** branch and to
the `SUGGESTED ROW IDS:` help text, and it changes the shape every consuming sheet
expects. Say so when proposing it.

---

## 5. State not carried through the recursion

```excel
Principal, 0 + 0,                    // placeholder never replaced
Interest,  0 * PeriodRate,           // always zero
...
fx.DebtSculptVariableDebtRsvλ(CFADS, DSCR, COUPON, DBTRSVPCT,
                                   DBTRSVRATE, MonthsPerPeriod, Period + 1)
                                   // ← no ClosingDebt passed
```

Two failures at once: scaffolding placeholders (`0 + 0`, `0 *`) left in, and the closing
balance not fed forward as the next period's opening balance. Every period restarts from
zero. The schedule looks superficially plausible because the CFADS-driven payments are
still there.

**The carried state belongs in the first argument slot**, and the self-call must pass
it. Grep the draft for `0 + 0`, `0 *`, and any argument in the signature that never
appears in the body.

---

## 6. Dead arguments

`fx.InterestDebtIssuanceλ` accepts `DBTRSVPCT` and `DBTRSVRATE`. Neither appears
anywhere in its calculation. The signature advertises reserve-awareness the body doesn't
implement — so a caller reasonably assumes the reserve is handled and it isn't.

Either use the argument, or remove it and update the help text. Flag dead arguments when
you find them; don't silently carry them into the new function.

---

## 7. Convergence tolerance smaller than the rounding step

```excel
Repayment, -FLOOR(CFADS / DSCR - Interest, ParValueIncr),
Result, IF(ABS(Payment + Repayment) < 0.01, MAX(Interest, 0), <recurse>)
```

`FLOOR` snaps the repayment to a par increment — often $5,000. The convergence test
asks for agreement within $0.01. If successive iterations alternate between two
increments, the test never passes and Excel returns `#NUM!` at its recursion limit.

**Test the quantity that actually converges smoothly** — the interest, or the
pre-rounding payment — and apply the increment rounding once, after convergence. Or set
a tolerance above the increment's effect and add an iteration cap.

---

## 8. No depth guard

The source library's convergence recursions have no iteration cap. Divergent inputs
produce `#NUM!` with no diagnosis. When building a new one, add a cap in its own
argument slot (see `source-structures.md`) — don't overload the convergence-value slot.

---

## 9. Sign conventions drift

The library's settled convention:

- `Payment` is **positive**, then negated for display: `VSTACK(..., -Payment, ...)`
- `ClosingDebt, Principal - Payment`
- In the convergence test, `Payment` and `Repayment` are **opposite-signed**

Mixing a `Repayment` that's already negative with a `Payment` that's positive breaks the
`ABS(Payment + Repayment)` test — it tests a sum that should be zero, so both terms must
point opposite ways. One sign flip and it never converges, or converges instantly on the
first iteration.

Write the sign of every flow in the spec before coding.

---

## 10. `INDEX` on whole arrays vs. already-sliced scalars

Pass whole arrays down the recursion; slice per period with `INDEX(Arg, 1, Period)`.
Passing an already-sliced scalar down means period *n+1* indexes into a single value and
gets `#REF!` — or, if it happens to be 1×1, silently reuses period *n*'s number forever.

---

## 11. Half-migrated error blocks

Fragments like:

```excel
Return, IF(AND(Errors), 2, IF(OR(Errors), 3, IF(OR(ISERROR(Result)), 2, 1))),
CHOOSE(Return, Result, Help, Messages)
```

pasted into a draft that never defines `Help`, `ErrMsgs`, `Errors` or `Messages`. Excel
reports an undefined-name error that points at the last line, not the missing
definitions. When adapting an existing function, carry the **whole** `LET` prologue or
none of it.

---

## 12. A LET name that looks like a cell reference

```excel
LET(
    Cl2,  Par - Cum,                 // CL2 is a cell address
    Op2,  HSTACK(Par, DROP(Cl2, , -1)),
    In2,  Op2 * COUPON,              // IN2, RSV2, CR2 likewise
    ...
```

Excel refuses any `LET`/`LAMBDA` name that parses as an A1 reference — up to three
letters forming a column at or before `XFD`, followed by a row number — or as an R1C1
form (`R12`, `C4`). It rejects the **entire binding list**, not just the offending name,
so the editor squiggles every identifier from that point on and the save fails with *"The
argument is invalid or missing or has an incorrect format."* The real error is the first
bad name; everything after it is collateral.

Easy to walk into when numbering a second pass — `Cl` → `Cl2`, `Op` → `Op2`. Note `Cl`,
`Op`, `In`, `Cr` are all perfectly legal; adding one digit breaks them. Use a
non-numeric suffix instead: `ClEnd`, `OpEnd`, `IntEnd`.

Watch for it especially with two- and three-letter abbreviations, since almost any
letter combination up to `XFD` is a real column.

More names refused for this reason: `h1`, `k1`, `d1`, `s2`, `Y1`, `M4`, `VOL2`,
`GAM1`, `LAM2` (`LAM` is a column). `k0` happens to be accepted — there is no row 0 —
but it is one edit away from `k1`. Keep digits out of names altogether.
`lambda_test_harness_template.py` lints for this before installing anything.

## 13. `MAX` / `MIN` collapse an array to a scalar

```excel
Prin, FLOOR(MAX(0, Raw), ParValueIncr),      // Raw is 1 x n - WRONG
```

`MAX(0, Raw)` returns **one number** — the largest element of `Raw`, or zero. Every
period then gets the same principal. The formula computes cleanly and the schedule
looks plausible, which is what makes it dangerous.

For an elementwise floor at zero, use `IF`:

```excel
Prin, FLOOR(IF(Raw < 0, 0, Raw), ParValueIncr),
```

`MAX`/`MIN` are fine on scalars — inside a `SCAN` or `REDUCE` callback, where you are
working one period at a time, `MAX(Interest, 0)` is correct. The trap is reaching for
the same idiom after you have lifted the calculation out to whole-array form.

## 14. A constant input hiding a structural error

The worst bugs survive testing because the test data is degenerate. A rate that is the
same in every period makes "aggregate × one rate" and "sum of per-item rates"
algebraically identical, so a function that confuses the two ties out perfectly — until
someone supplies a curve.

Real case: a debt sizer computed `Interest = Balance × COUPON(period)`. Against a flat
5% issue it was exact. Against a rising coupon curve it was wrong by millions a year,
with the sign flipping between early and late periods.

**Tie out against a case where the suspect input actually varies.** If only a constant
version is available, vary it yourself and check the result moves the way theory says
it should. And when a per-item attribute gets aggregated anywhere in the model, ask
explicitly whether the aggregate is still valid when the items differ.

---

# Excel-engine traps

Sections 15-27 are not about the logic of a schedule. They are places where Excel
itself does something other than what the formula appears to say. Most fail with only
the generic *"There's a problem with this formula"*, or with an error that points
nowhere near the cause, or — worst — with a plausible number.

## 15. A LET-bound LAMBDA named after a built-in function

```excel
LET(
    sec, LAMBDA(x, x + 1),     // meant: a helper called "sec"
    sec(2)                     // Excel calls SEC(2) and returns -2.40
)
```

When a name is followed by `(`, Excel resolves it as a **function call first**. A
helper that shares a built-in's name is never called. Measured in Excel 365:

| LET-bound helper called as | What Excel did |
|---|---|
| `t(2)` | returned `""` — the `T()` text function |
| `sec(2)` | returned `-2.40` — the secant |
| `var(2)` | returned `#DIV/0!` — the variance of one value |
| `row(2)` | refused, or `#VALUE!` |
| `na(2)`, `mid(2)`, `step(2)` | formula refused |

The silent cases are the dangerous ones. The same names used as plain *values*
(`LET(t, 3, t * 2)`) were accepted in that test, though some builds and entry routes
have refused `t`, `E`, `c`, `r` and `Step` even there — so avoid them; it costs nothing.

**Fix:** give LET-bound LAMBDAs a name no built-in can have. Ending them in `λ`
(`Stepλ`, `Solveλ`) does that and matches the function-naming convention.

## 16. Names that differ only by case

Excel names are case-insensitive. Binding `U` and `u` in the **same** `LET` is
refused with the generic error. A parameter `U` and a `LET` name `u` is accepted — and
the inner one silently shadows the outer, which is worse. One spelling per name.

## 17. A parameter that shadows a dotted function

```excel
LAMBDA(Beta, BETA.DIST(0.5, Beta, 2, TRUE))     // returns #FIELD!
```

With a parameter called `Beta` in scope, Excel reads `BETA.DIST(...)` as a data-type
*field access* on that parameter. The error is `#FIELD!`, which suggests nothing to do
with naming. The same trap waits behind `Norm`, `Binom`, `Gamma`, `Lognorm`,
`Percentile` — any name that is the first half of a dotted function. Grep for any
parameter or LET name that appears immediately before a `.`.

## 18. Unary minus binds tighter than `^`

```excel
-x^2          // with x = 3 this is 9, not -9: it means (-x)^2
EXP(-h^2/2)   // grows instead of decaying
```

Excel applies negation before exponentiation — the opposite of the maths convention.
A Gaussian-shaped term silently explodes. Always write `-(x^2)`. Grep before shipping:
`[(,=]\s*-\s*\w+\s*\^`.

## 19. A zero-argument LAMBDA must be called with `()`

A name defined as `LAMBDA(42)` is a function, not a value. `=Probe` in a cell returns
`#CALC!`; `="x" & Probe` returns `#VALUE!`. Only `=Probe()` returns 42. Easy to miss
on a constant-style helper (a version string, a default table).

## 20. An empty array is an error

`DROP`, `TAKE` and `FILTER` return `#CALC!` when nothing is left. The common idiom

```excel
OpenBal, HSTACK(Amount, DROP(CloseBal, , -1)),     // #CALC! when there is 1 period
```

works for every timeline except a one-period one, where `DROP` removes the only
column. (This idiom appears in `convert-to-5g/reference/conversion-rules.md` and
`source-structures.md` — it is fine there as long as the timeline can never be one
period long.) Stack first, then drop — the result always has at least one column:

```excel
OpenBal, DROP(HSTACK(Amount, CloseBal), , -1),
```

For `FILTER`, supply the `if_empty` argument. **Test the length-1 case explicitly**;
random tests rarely land on it.

## 21. Stacking pads with `#N/A`, and a bare error poisons the whole stack

`HSTACK` and `VSTACK` pad a short piece with `#N/A`. `HSTACK({1;2;3}, "")` gives
`""` in row 1 and `#N/A` below it, not a blank column. For a blank column, size it:
`IF(SEQUENCE(n), "")`.

A bare error value as a whole argument errors the **entire** stack:
`VSTACK({1;2}, NA())` is a single `#N/A`, not a column ending in `#N/A`. When you want
trailing `#N/A` cells (charts skip them), stack only real values and let `EXPAND`'s
default padding supply the `#N/A`s: `EXPAND(x, rows, 1)`. *(from experience — not re-verified)*

See also §4 — the padding is how a height mismatch hides.

## 22. `INDEX` into an in-memory array returns a 1×1 array, not a scalar

`TYPE(INDEX(VSTACK(5, 6), 1))` is `64` (array), not `1` (number). Values read from
worksheet *cells* are true scalars, so a test that feeds inputs as ranges never sees
the difference; it appears only when one block is chained into another in memory. In
one library this made an `IF(that >= 1, bigArray, 1)` return 1×1 instead of the full
array. A simple repro did not show the shrink, so treat coercion as cheap insurance
wherever a scalar is required:

```excel
SUM(INDEX(Model, 2, 1))          // a number
AND(INDEX(Flags, 1, 1))          // a boolean
MAP(Vec, LAMBDA(c, SUM(c)))      // each element of a vector
```

Test chained calls with in-memory inputs, not only with ranges.

## 23. Native functions that quietly mislead

- `CHISQ.INV`, `T.INV`, `F.INV` and `NEGBINOM.DIST` **truncate a non-integer degrees
  of freedom or count** (4.5 behaves as 4), with no error. For fractional degrees of
  freedom use the equivalent `GAMMA.INV(p, df/2, 2)` (chi-square) or `BETA.INV` /
  `BETA.DIST` forms. *(from experience — not re-verified)*
- `TEXT(n, "m0")` is a **month** format, not the letter m followed by a number. Use
  `"m" & n`. *(from experience — not re-verified)*
- For least squares, use `LINEST` (QR decomposition), not `MINVERSE` on the normal
  equations. On a 16-term fit the normal-equations route drifted about 1e-4 relative;
  `LINEST` held about 1e-11. `LINEST` also handles an exact fit (as many points as
  terms). *(from experience — not re-verified)*

## 24. Volatility is decided statically, across the whole LAMBDA chain

*(from experience — not re-verified)*

If `RANDARRAY` or `OFFSET` appears **anywhere** in a LAMBDA — or in any LAMBDA it
calls — every cell that calls it is volatile and recalculates on every edit in the
workbook, **even if the branch containing it never runs**. Passing an argument that
avoids the branch does not help; the flag is set from the formula text. (`RANDARRAY`
and `OFFSET` were measured; assume the same for any volatile function.)

**Fix:** keep the volatile generator in its own small function and pass its output in
as an argument, so the calculating function contains nothing volatile.

**Measure it, one source at a time.** A no-argument VBA UDF that counts its own calls,
placed in the same cell, shows whether an unrelated edit recalculates the cell. If
anything else in the chain is already volatile, removing one source changes nothing —
an earlier measurement concluded `OFFSET` "added no recalculation" for exactly that
reason, and was wrong. Remove every other volatile source first.

## 25. Performance traps

The timings below *(from experience — not re-verified)*. The ratios matter more than the seconds.

- **`MAP` over `SEQUENCE(T)` with `CHOOSEROWS(block, t)` inside is O(T²)** — the
  in-memory block is re-read on every call. At 2,000 rows it took 19 s;
  `BYROW(block, LAMBDA(r, ...))`, which hands each row straight in, took 0.11 s. The
  same trap applies to `REDUCE` + `VSTACK` accumulation.
- A `SCAN` whose step `INDEX`es into the whole array per step is slow at scale (6 s at
  10,000 elements for a ranking that a `SORT`/`UNIQUE`/`FREQUENCY` formulation does in
  a fraction of that).
- **Row-wise recursion without a loop:** one `SCAN` over
  `HSTACK(<text sentinel column>, block)` with
  `IF(ISTEXT(item), Start, f(acc, item))` restarts the accumulator at each row. It ran
  600,000 cells in about 0.15 s.
- **Time with manual calculation.** Timing through COM in automatic mode is inflated
  by repeated recalcs; set `Calculation` to manual and time `Range.Calculate()`.

## 26. Installing and shipping

- **`//` comments are illegal in an Excel formula.** AFE strips them on import; any
  other installer must too.
- **Indentation counts.** An installer that writes the indented AFE text verbatim pays
  for every indentation space against the 8,192-character limit. The tell: Excel Labs
  shows the formula on a single line. Flatten whitespace **outside string literals**
  before installing — a string-unaware flattener mangles the help text.
- **Excel rewrites what you install.** Verified: it drops the spaces in a LAMBDA's
  parameter list, stores an optional `[Arg]` as `_xlop.Arg`, adds `_xlfn.` / `_xlpm.` /
  `_xlws.` prefixes, and tightens whitespace inside `{array constants}`. Compare
  installed text against source only after normalising both, or the diff is noise.
  The prefixes are why the stored form is longer than the paste — about a third longer
  on a small function, and the 8,192 cap applies to the stored form.
- **`Name.Comment` raises an error above 255 characters.** *(from experience — not re-verified)*
- **One function can live in several places that do not sync** — the source file or
  gist, an add-in's library, an AFE module, and the workbook's defined name. AFE's
  save writes its own module text into the defined name, so a stale AFE module
  silently overwrites a good name. Keep one source of truth and delete the other
  copies, and check that the definition you are *testing* is the one you *edited*
  (see `lambda-testing.md`, *Test what is installed*).
- **AFE keeps the full module source — comments included — inside the workbook**, as
  base64 in a custom XML part. It is invisible in the grid, in Name Manager, and to a
  text search of the file. Anything private in an AFE comment travels with the
  workbook. Clear the module in the AFE pane before sharing; the defined names are a
  separate store and survive. *(from experience — not re-verified)*
- **Comments never enter the stored name.** A change confined to comments needs no
  reinstall; a change inside a string literal (help text) does. Compare stored lengths
  to tell which.

## 27. Diagnose by experiment, not by inference

When Excel refuses a formula or returns an error that points nowhere, every
hypothesis formed by reading the text — length, nesting depth, character content,
dependency order — tends to be disproved by another function in the same file that
shares the trait and works. Reading longer produces confident wrong answers.

Get an observation out of Excel instead:

- **Bisect.** Install a prefix of the `LET` (the first *k* bindings, returning the last
  one) and grow it until it fails. The first failing binding is the culprit; everything
  after it is collateral.
- **Write a throwaway probe** that tries five variations in one run — the real name vs
  a plain-ASCII name, one-step vs two-step install, wrapped vs bare. It kills four
  hypotheses at once. `lambda_test_harness_template.py`'s `OwnExcel` class makes such a
  probe a dozen lines.
- **If another tool succeeds, compare against it.** When Excel Labs imported a module
  that another importer could not, the difference exposed two real bugs (`//` comments
  left in, and indentation counted against the limit).
- **For size questions, bisect on size alone**: a tiny LAMBDA with every structural
  feature of the real one, then the real one. If the tiny one goes in, the content is
  not the problem.

When a user testing in real Excel reports something that contradicts your reasoning,
their observation wins. Correct it plainly and move on.

## Quick triage

| Symptom | Look at |
|---|---|
| `#NUM!` | 7, 8 — convergence never settles |
| `#VALUE!` | 4 arity/shape, 3 wrong argument count |
| `#REF!` | 10 — indexing a scalar |
| Undefined name | 11 — orphaned `Return`/`CHOOSE` |
| "Argument is invalid or missing or has an incorrect format" on save | **12 — a LET name that is a cell address** |
| Right at one period, wrong elsewhere | **1 — lookahead** |
| Every period gets the same value | 13 — `MAX`/`MIN` collapsed an array |
| Ties on one dataset, wrong on another | **14 — a constant input hid it** |
| Right in period 1, wrong after | **2 — self-call names the old function** |
| Balances never decline | 5 — state not carried |
| Feature has no effect | 6 — dead argument |
| Generic "There's a problem with this formula" | 12, 15, 16 — then bisect (27) |
| A helper returns `""`, a trig value, or a built-in's error | **15 — helper named like a built-in** |
| `#FIELD!` | 17 — parameter shadows a dotted function |
| A squared term explodes or has the wrong sign | 18 — `-x^2` is `(-x)^2` |
| `#CALC!` | 20 — empty array (often a one-period timeline); 19 — zero-arg LAMBDA called bare |
| `#N/A` inside a stacked block | 21 — padding, or a bare error in the stack |
| Right with range inputs, wrong when chained in memory | 22 — 1×1 array, not a scalar |
| Wrong only with fractional degrees of freedom | 23 — native function truncates |
| Every cell recalculates on any edit | 24 — volatility through the chain |
| Correct but slow at scale | 25 |
| Installed text differs from source | 26 — normalise before diffing; check for stale copies |
