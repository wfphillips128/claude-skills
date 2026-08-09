# 5G source structures

Five ways to build a schedule with a LAMBDA. Pick by what period *n* needs to know.

| Structure | Use when | Cost |
|---|---|---|
| **Corkscrew** | Classic open → movement → close balance roll-forward | Familiar, verbose |
| **Reduce and Stack** | Each period depends on the previous; you build one array | `REDUCE` + `VSTACK`, O(n²) copying |
| **Scan and Stack** | Same, but `SCAN` carries the accumulator natively | Cleaner than REDUCE for single-value carry |
| **Thunk and Scan** | Each period carries *several* values forward | Thunks (`LAMBDA(x)`) wrap arrays as single values |
| **Circularity** | A value depends on itself (interest ↔ principal) | Convergence recursion |

The `fx.*` library uses a sixth, informal one: **recursive HSTACK**. It behaves
like Scan-and-Stack but is written as direct self-recursion, which makes the
per-period logic easy to read.

---

## Recursive HSTACK with a period counter

Each call computes one period's column, then stacks itself for the next period.

```excel
Library.Scheduleλ = LAMBDA(
    [OpeningBalance], [CFADS], [DSCR], [COUPON], [MonthsPerPeriod], [DoNotUse],
    LET(
        // ... Help, ErrMsgs, coercion, Errors, Messages ...

        // Which period is this call? Omitted → 0 → period 1.
        Period,       IF(DoNotUse = 0, 1, DoNotUse),
        ModelPeriods, COLUMNS(CFADS),

        // Slice this period's scalars out of the whole-array arguments
        Rate,        COUPON / 12 * MonthsPerPeriod,
        PeriodRate,  INDEX(Rate,  1, Period),
        PeriodCFADS, INDEX(CFADS, 1, Period),
        PeriodDSCR,  INDEX(DSCR,  1, Period),

        // One period of arithmetic
        Principal,   OpeningBalance,
        Interest,    Principal * PeriodRate,
        Payment,     MIN(PeriodCFADS / PeriodDSCR - Interest, Principal),
        ClosingDebt, Principal - Payment,

        // The column this period contributes — order defines the output rows
        NewColumn, VSTACK(Principal, Interest, -Payment, ClosingDebt),

        // Stop at the last period, otherwise recurse with the counter advanced
        Result, IF(
            Period = ModelPeriods,
            NewColumn,
            HSTACK(
                NewColumn,
                Library.Scheduleλ(ClosingDebt, CFADS, DSCR, COUPON,
                                  MonthsPerPeriod, Period + 1)
            )
        ),
        Return, IF(AND(Errors), 2, IF(OR(Errors), 3, IF(OR(ISERROR(Result)), 2, 1))),
        CHOOSE(Return, Result, Help, Messages)
    )
);
```

### Rules

1. **`DoNotUse` is the counter.** The user never passes it. An omitted optional
   argument evaluates as `0` in arithmetic, so `IF(DoNotUse = 0, 1, DoNotUse)` yields
   period 1 on the outside call.
2. **Whole arrays go down; scalars come out with `INDEX`.** Pass `CFADS` entire on
   every recursive call and slice with `INDEX(CFADS, 1, Period)`. Never pass the
   already-sliced scalar down — the next period needs the whole array.
3. **The carried state is the first argument** (`ClosingDebt` → `OpeningBalance`).
4. **`ModelPeriods, COLUMNS(<a per-period array>)`** sets the length. Choose an
   argument guaranteed to span the full timeline.
5. **`NewColumn`'s `VSTACK` order is the function's output contract.** Document it
   under `SUGGESTED ROW IDS:` in the help text. Changing the order silently breaks
   every sheet already consuming it.
6. **Period *n* can only see periods 1..*n*.** See `pitfalls.md` — this is the single
   biggest source of wrong-but-plausible results.

---

## Convergence recursion (circularity)

For values defined in terms of themselves. Interest depends on the repayment, which is
sized off cash flow net of interest.

```excel
Library.Interestλ = LAMBDA(
    [CFADS], [DSCR], [Principal], [InterestRate], [ParValueIncr], [DoNotUse],
    LET(
        // Previous iteration's payment; on the first call DoNotUse = 0
        Payment,   IF(DoNotUse = 0, CFADS / DSCR, DoNotUse),
        Interest,  Principal * InterestRate,
        Repayment, -FLOOR(CFADS / DSCR - Interest, ParValueIncr),
        Result, IF(
            ABS(Payment + Repayment) < 0.01,          // settled?
            MAX(Interest, 0),
            Library.Interestλ(CFADS, DSCR, Principal, InterestRate,
                              ParValueIncr, -Repayment)   // feed this iteration forward
        ),
        Result
    )
);
```

### Rules

1. **`Payment` and `Repayment` are opposite-signed**, so `ABS(Payment + Repayment)`
   approaching zero means the iteration has settled. Get a sign wrong and the test
   never fires.
2. **The convergence value is passed in the counter slot** — `-Repayment` becomes the
   next call's `DoNotUse`, and thus its `Payment`.
3. **Tolerance must exceed the rounding granularity.** With `FLOOR(..., ParValueIncr)`
   at a $5,000 increment, a $0.01 tolerance can oscillate between two increments
   forever. Either test the *interest* delta instead, or make the tolerance larger than
   the increment's effect.
4. **There is no depth guard in the source library.** Excel's recursion limit produces
   `#NUM!`. If divergence is plausible, add an explicit iteration cap:

```excel
Iter, IF(ISOMITTED(IterCount), 1, IterCount),
Result, IF(
    OR(ABS(Payment + Repayment) < 0.01, Iter >= 100),
    MAX(Interest, 0),
    Library.Interestλ(..., -Repayment, Iter + 1)
),
```

Cap and tolerance are separate arguments — don't overload one slot with both.

---

## Nesting the two

The classic debt-sculpting shape: an outer period recursion whose interest step calls
an inner convergence recursion.

```
Scheduleλ  period 1 ─→ Interestλ  iterate to convergence ─→ interest for period 1
           period 2 ─→ Interestλ  iterate to convergence ─→ interest for period 2
           ...
```

Each period's circularity is solved independently. That's what makes the whole thing
tractable — and it's why period *n* still cannot see period *n+1*.

---

## Scan and Stack, for reference

When the carry is a single value and you don't want self-recursion:

```excel
CloseBal, SCAN(Amount, SEQUENCE(1, Term),
               LAMBDA(bal, n, bal - INDEX(Payments, 1, n))),
OpenBal,  HSTACK(Amount, DROP(CloseBal, , -1)),
Result,   VSTACK(OpenBal, Payments, CloseBal),
```

`SEQUENCE(1, Term)` — the 1-row form — keeps the output horizontal without a
`TRANSPOSE`. To carry more than one value forward, wrap them in a thunk
(`LAMBDA(VSTACK(a, b))`) and unwrap with `thunk()`; that's Thunk-and-Scan.
