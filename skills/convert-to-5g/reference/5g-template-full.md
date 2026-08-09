# The full 5G structure

This is the template to use for anything reusable. It is **transcribed from working
code** — `fx.DebtSculptIssuanceλ` and `fx.InterestDebtIssuanceλ` as stored in
a real Beyond Excel workbook — not from documentation. (The beyondexcel Google Site
renders via JavaScript and cannot be fetched; working source is the authority.)

## Shape

```
LAMBDA( [args...],
  LET(
    Help,      TRIM(TEXTSPLIT( "...help text...", "→", "¶" )),
    ErrMsgs,   { "message 1" ; "message 2" ; ... },
    <arg>,     <coercion / validation, re-binding each argument to itself>,
    ...
    Errors,    VSTACK( OR(ISERROR(arg1)), OR(ISERROR(arg2)), ... ),
    Messages,  FILTER( ErrMsgs, Errors, "" ),
    <calculations>,
    Result,    <the answer>,
    Return,    IF(AND(Errors), 2, IF(OR(Errors), 3, IF(OR(ISERROR(Result)), 2, 1))),
    CHOOSE( Return, Result, Help, Messages )
  )
)
```

The final `CHOOSE` is the whole trick:

| `Return` | Condition | User sees |
|---|---|---|
| 1 | Everything valid | `Result` |
| 2 | **All** arguments bad (i.e. called with nothing) or the result errored | `Help` — the full inline help table |
| 3 | *Some* arguments bad | `Messages` — only the failing arguments' messages |

So typing `=MyFuncλ()` in a cell spills a readable help table. That is the 5G
self-documentation mechanism.

## Annotated template

```excel
/* FUNCTION NAME: Library.Nameλ
   DESCRIPTION:*//**What the function does, in one sentence*/
/* REVISIONS: Date         Developer   Description
              Mmm dd yyyy  UserName    Initial release
*/
Library.Nameλ = LAMBDA(
    // Every argument is bracketed so Excel treats it as optional.
    // Required-ness is communicated by the help text, not the signature.
    [Arg1],             // Required. Description.
    [Arg2],             // Required. Description.
    [OptArg],           // Optional. Description. Default: 0
    [DoNotUse],         // Internal recursion counter — never passed by the user

    LET(
        // ---- 1. Help table -------------------------------------------------
        // Column delimiter is → (U+2192). Row delimiter is ¶ (U+00B6).
        // Every line ends with & except the last.
        Help, TRIM(TEXTSPLIT(
            "DESCRIPTION:       →What the function does¶" &
            "SUGGESTED ROW IDS: →First output row¶" &
            "                   →Second output row¶" &
            "VERSION:           →UserName, Mmm dd, yyyy¶" &
            "PARAMETERS:→¶" &
            "Arg1               →(Required) Description¶" &
            "Arg2               →(Required) Description¶" &
            "OptArg             →(Optional) Description. Defaults to 0¶" &
            "DoNotUse           →This is required internally for counting period columns¶" &
            "→¶" &
            "EXAMPLE:           →=Library.Nameλ(Arg1#, Arg2#)",
            "→", "¶")),

        // ---- 2. One error message per validated argument, same order -------
        ErrMsgs, {
            "Arg1 must be a dynamic array holding ... for each period";
            "Arg2 must be a dynamic array holding ... for each period"
        },

        // ---- 3. Coerce / validate: rebind each argument to itself ----------
        // Array arguments: reject anything that isn't multi-column.
        Arg1, IF(COLUMNS(Arg1) <= 1, #VALUE!, Arg1),
        Arg2, IF(COLUMNS(Arg2) <= 1, #VALUE!, Arg2),
        // Scalar arguments: require a single number.
        // OptArg, IF(AND(ISNUMBER(OptArg), COLUMNS(OptArg) = 1), OptArg, #VALUE!),
        // Optional with a default:
        OptArg, IF(ISOMITTED(OptArg), 0, OptArg),
        // Recursion counter: omitted means "first period".
        Period, IF(DoNotUse = 0, 1, DoNotUse),

        // ---- 4. Error vector, aligned 1:1 with ErrMsgs ---------------------
        Errors,   VSTACK(OR(ISERROR(Arg1)), OR(ISERROR(Arg2))),
        Messages, FILTER(ErrMsgs, Errors, ""),

        // ---- 5. Calculations ----------------------------------------------
        ModelPeriods, COLUMNS(Arg1),
        PeriodArg1,   INDEX(Arg1, 1, Period),
        // ... one named step per line, last calculation named Result
        Result, <expression>,

        // ---- 6. Dispatch ---------------------------------------------------
        Return, IF(AND(Errors), 2, IF(OR(Errors), 3, IF(OR(ISERROR(Result)), 2, 1))),
        CHOOSE(Return, Result, Help, Messages)
    )
);
```

## Notes that matter

**Rebinding arguments to themselves** (`Arg1, IF(COLUMNS(Arg1) <= 1, #VALUE!, Arg1)`)
is deliberate. It converts a bad argument into an error *value* that flows through
`ISERROR` into the `Errors` vector, without a separate variable name.

**`ErrMsgs` uses `;` separators** — that's a vertical array constant, one message per
row, matching `Errors` built with `VSTACK`. Order must line up exactly or the wrong
message shows.

**`Errors` need not cover every argument.** In the source library only the array
arguments are validated; `OpeningBalance` is simply defaulted:
`OpeningBalance, IF(ISNUMBER(OpeningBalance), OpeningBalance, 0)`.

**`#VALUE!` is written literally** inside the formula. That is valid Excel.

**`Period, IF(DoNotUse = 0, 1, DoNotUse)`** works because an omitted optional argument
behaves as `0` in arithmetic. `InterestDebtIssuanceλ` uses the more explicit
`DoNotUse, IF(ISOMITTED(DoNotUse), 0, DoNotUse)` first. Either is fine — be consistent
within a library.

## Worked example — the real thing, formatted

`fx.InterestDebtIssuanceλ`, exactly as stored, re-indented. It solves a
circularity: interest depends on the principal repayment, which depends on interest.
It converges by calling itself with the previous iteration's payment.

```excel
fx.InterestDebtIssuanceλ = LAMBDA(
    [CFADS], [DSCR], [Principal], [InterestRate],
    [DBTRSVPCT], [DBTRSVRATE], [ParValueIncr], [DoNotUse],
    LET(
        Help, TRIM(TEXTSPLIT(
            "DESCRIPTION:   →Calculates debt sculpting interest using method presented by Lance Rubin¶" &
            "VERSION:       →Aug 08 2023¶" &
            "PARAMETERS:→¶" &
            "CFADS          →(Required) Cash Flow Available for Debt Services for this period¶" &
            "DSCR           →(Required) Debt Service Coverage Ratio for this period¶" &
            "Principal      →(Required) Principal balance for this period¶" &
            "InterestRate   →(Required) Interest's Percentage Rate for this period¶" &
            "DoNotUse       →This is required internally for calculating rates¶" &
            "→¶" &
            "EXAMPLES:→¶" &
            "Result         →Formula¶" &
            "222.90         →=InterestDebtIssuanceλ( 6666.37, 3.50, 90000, 3%/12)",
            "→", "¶")),
        ErrMsgs, {
            "CFADS must be a single number holding cash available for debt service for this period";
            "DSCR must be a single number holding debt service coverage ratio for this periods";
            "Principal must be a single number holding this period's opening debt balance";
            "InterestRate must be a single number holding this period's interest rate"
        },
        CFADS,        IF(AND(ISNUMBER(CFADS),        COLUMNS(CFADS) = 1),        CFADS,        #VALUE!),
        DSCR,         IF(AND(ISNUMBER(DSCR),         COLUMNS(DSCR) = 1),         DSCR,         #VALUE!),
        Principal,    IF(AND(ISNUMBER(Principal),    COLUMNS(Principal) = 1),    Principal,    #VALUE!),
        InterestRate, IF(AND(ISNUMBER(InterestRate), COLUMNS(InterestRate) = 1), InterestRate, #VALUE!),
        DoNotUse,     IF(ISOMITTED(DoNotUse), 0, DoNotUse),
        Errors,   VSTACK(OR(ISERROR(CFADS)), OR(ISERROR(DSCR)),
                         OR(ISERROR(Principal)), OR(ISERROR(InterestRate))),
        Messages, FILTER(ErrMsgs, Errors, ""),

        Payment,   IF(DoNotUse = 0, CFADS / DSCR, DoNotUse),
        Interest,  Principal * InterestRate,
        Repayment, -FLOOR(CFADS / DSCR - Interest, ParValueIncr),
        Balance,   Principal + Repayment,
        Result, IF(
            ABS(Payment + Repayment) < 0.01,
            MAX(Interest, 0),
            fx.InterestDebtIssuanceλ(CFADS, DSCR, Principal, InterestRate,
                                          DBTRSVPCT, DBTRSVRATE, ParValueIncr, -Repayment)
        ),
        Return, IF(AND(Errors), 2, IF(OR(Errors), 3, IF(OR(ISERROR(Result)), 2, 1))),
        CHOOSE(Return, Result, Help, Messages)
    )
);
```

Note the convergence test `ABS(Payment + Repayment) < 0.01` — `Payment` is the previous
iteration's value, `Repayment` this one's, and they're opposite-signed, so the sum
approaching zero means it has settled. **This pattern has no depth guard**; if the
inputs make it diverge, Excel hits its recursion limit and returns `#NUM!`. See
`compose-5g-lambda/reference/pitfalls.md`.
