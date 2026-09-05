# 5G source structures

Five ways to build a schedule with a LAMBDA. Pick by what period *n* needs to know.

| Structure | Use when | Relative speed |
|---|---|---|
| **Corkscrew** | Classic open → movement → close balance roll-forward | **Fastest** — ~400x Reduce-and-Stack |
| **Reduce and Stack** | Each period depends on the previous; period *n* needs prior columns | **Slowest** — the baseline |
| **Scan and Stack** | Same, but `SCAN` carries the accumulator natively | >= 2x Reduce-and-Stack |
| **Thunk and Scan** | Each period carries *several* values forward | >= 10x Scan-and-Stack |
| **Circularity** | A value depends on itself (interest ↔ principal) | Convergence, ~8 iterations typical |

Those multipliers are Hatmaker's own, from the course pages. They matter more
than they look: Reduce-and-Stack is the easiest to write and the one to reach
for when *converting a grid*, but it is two orders of magnitude off the pace of
a plain corkscrew. Write it first, and move only if it is actually slow.

Two rules that come with the table:

- **Do not use Scan-and-Stack or Thunk-and-Scan for a simple corkscrew.** They
  earn their complexity only when you must select current-period values with
  `INDEX`.
- **Thunk-and-Scan only when performance is a demonstrated problem** — it is the
  one structure a reviewer will not be able to read.

For the templates themselves, verbatim as published, see
`hatmaker-templates.md`. This file covers the two structures the `fx.*` library
actually uses, plus thunks.

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
`TRANSPOSE`.

---

## Thunks

The answer to "period *n* needs to carry more than one value forward", and the
way around Excel's *array of arrays* error.

### The problem

`SCAN` produces one row. That row is an array. Ask `SCAN` to output two rows and
Excel sees two arrays inside one array, which it cannot represent — so it throws
`#CALC!`, "array of arrays".

### What a thunk is

A function that holds an array. Excel treats a function as a single object, so
an *array of thunks* is legal where an array of arrays is not.

The canonical form, and the three-step progression that makes it click:

```excel
=LAMBDA(x, LAMBDA(x))          // #CALC! - "Cell formula result returns a function"
=LAMBDA(x, LAMBDA(x))(10)      // #CALC! - this returns LAMBDA(10), still a function
=LAMBDA(x, LAMBDA(x))(10)()    // 10
```

The first two error for the same reason `=SUM` errors in a cell: a cell displays
values, not functions. The empty parentheses are what make the thunk give up its
contents.

Hatmaker's own comment on this is worth keeping, because it sets the right
expectation:

> Easy right? Nope. At least not to me. I understand how it works but not
> completely why it works. I know people who do know why it works, but knowing
> why isn't as important as knowing how.

### Loading and unloading

```excel
    // 1. Declare the thunk. This does nothing yet - it just names a function.
    Thunk,  LAMBDA(x, LAMBDA(x)),

    // 2. Stack the rows, then wrap each COLUMN in its own thunk.
    //    The result is an array of thunks, not an array of arrays.
    Flows,  BYCOL(VSTACK(Gains, Losses), LAMBDA(col, Thunk(col))),

    // 3. SCAN over the thunks. Each callback receives ONE thunk.
    Close,  SCAN(Opening, Flows, LAMBDA(Acc, Flow,
                LET(
                    Pair, Flow(),                    // () unwraps it: a 2-row column
                    Gain, INDEX(Pair, 1, 1),
                    Loss, INDEX(Pair, 2, 1),
                    Acc + Gain - Loss
                ))),
```

**The naming convention is what makes thunk code readable**, and it is worth
following exactly:

| Name | What it is |
|---|---|
| `Flows` | the array of thunks (plural) |
| `Flow` | one thunk (singular) |
| `Flow()` | the array inside it |

### Why it is faster

Not because thunks are cheap, but because of what they replace:

> instead of using INDEX to pull from a larger array, we will use INDEX to pull
> from a much smaller array — a Thunk containing two values.

Every period's `INDEX` runs against a 2-cell column instead of the whole
timeline. That is where the >= 10x comes from.

### When not to

Only when performance is a demonstrated problem, and never for a simple
corkscrew. A reviewer who cannot read the function cannot check it — and the
course's answer to that ("they can test it, which is superior to inspection")
only holds if the tests in `convert-to-5g/reference/testing.md` actually exist.

---

## Recursion without a public counter

The `DoNotUse` / `Lo` / `Hi` / `Iter` pattern used throughout the `fx.*` library
puts recursion state in the function's own signature. It works, and it is what
the tested library does — but it has two costs:

- those arguments show up in the user's tooltip, against Standard 02's "clear
  and concise argument names"
- every recursive call re-passes *every* argument, and in stored form each one
  carries an `_xlpm.` prefix. On `fx.DebtSizeIssueλ` a single call site costs
  185 stored characters against a hard 8,192 limit.

Hatmaker's Circularity template avoids both by making the recursive function a
**LET step that is passed to itself**:

```excel
    // The sub-function's FIRST argument is the function to call next.
    Iterate, LAMBDA(fn, CashBeforeInt, BeginningDebt, InterestRate, AverageDebt,
                 LET(
                     Interest,    AverageDebt * InterestRate,
                     CFADS,       CashBeforeInt - Interest,
                     EndingDebt,  BeginningDebt - CFADS,
                     AverageDebt, AVERAGE(BeginningDebt, EndingDebt),
                     Delta,       ABS(Interest - (AverageDebt * InterestRate)),
                     Result,      IF(Delta < 0.01,
                                     Interest,
                                     fn(fn, CashBeforeInt, BeginningDebt,
                                        InterestRate, AverageDebt)),
                     Result
                 )),

    // Seed it by handing it to itself.
    Result,  Iterate(Iterate, CashBeforeInt, BeginningDebt, InterestRate, 0),
```

`fn(fn, ...)` is what keeps the reference alive down the chain. The public
signature stays clean, and the recursion needs no workbook-level name — which is
also what keeps it compliant with Standard 06, self-containment.

**Seed the circular quantity at zero.** `AverageDebt` starts at 0, so `Interest`
is 0 on the first pass and `Delta` is large. Each pass shrinks it; ~8 iterations
is typical for this class of problem.

**What the template does not give you:** no iteration cap and no divergence
handling. It assumes the delta shrinks monotonically. Where that is not
guaranteed, add a cap in its own argument slot as shown above — and note that
some problems do not converge at all under fixed-point iteration.
`fx.DebtSizeIssueλ` is the worked case: raising the balance raises interest,
which cuts principal by more than the balance rose, so it bisects on total
annual coupon instead. Verified divergent at 200 iterations before the switch.
