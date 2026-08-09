# Pitfalls when composing 5G LAMBDAs

Read this first when a variant returns wrong numbers, `#VALUE!`, or `#NUM!`. Every
pattern below was found in real half-finished debt-sculpting LAMBDAs.

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
5% issue it was exact. Against a rising coupon curve it was wrong by up to $7m a year,
with the sign flipping between early and late periods.

**Tie out against a case where the suspect input actually varies.** If only a constant
version is available, vary it yourself and check the result moves the way theory says
it should. And when a per-item attribute gets aggregated anywhere in the model, ask
explicitly whether the aggregate is still valid when the items differ.

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
