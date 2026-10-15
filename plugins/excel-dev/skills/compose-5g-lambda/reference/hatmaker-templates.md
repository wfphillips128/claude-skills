# Hatmaker's published templates, verbatim

The 5g Creator's Course publishes five source-structure templates. **Every one
of them appears on the site as a screenshot**, so scraping the pages gives you
the prose and none of the code. All six pages point at one gist for the actual
download:

<https://gist.github.com/CHatmaker/3e6cba6414fadfabfa3725255da2ba74>
(raw: `https://gist.githubusercontent.com/CHatmaker/3e6cba6414fadfabfa3725255da2ba74/raw/`)

That gist holds four of them — Base, Corkscrew, REDUCE and STACK, Circularity.
SCAN and STACK and Thunk and SCAN have course pages but no gist entry; what is
below for those two is reconstructed from the prose and **marked as such**.

Fetched 05 Sep 2026. Everything in the four verbatim blocks is exactly as
published, **including its bugs** — see the last section before you copy any of
it.

## Choosing between them

| Template | Use when | Speed |
|---|---|---|
| **Corkscrew** | A plain opening → inflow → outflow → closing roll-forward | Fastest. ~400x REDUCE and STACK |
| **REDUCE and STACK** | Converting an existing grid schedule; period *n* needs prior columns | Slowest — the baseline |
| **SCAN and STACK** | A REDUCE and STACK function that is too slow | >= 2x REDUCE and STACK |
| **Thunk and SCAN** | A SCAN and STACK function that is still too slow; or you hit "array of arrays" | >= 10x SCAN and STACK |
| **Circularity** | A value defined in terms of itself | n/a — convergence, ~8 iterations typical |

Two rules that come with that table:

- **Do not use SCAN-and-STACK or Thunk-and-SCAN for a simple corkscrew.** Both
  pages say so explicitly. They earn their complexity only when you need to
  select current-period values with `INDEX`.
- **Thunk and SCAN only when performance is a demonstrated problem.**
  > This is one of those highly technical solutions that few people will
  > understand. For that reason, I recommend using this only when performance is
  > an issue.

REDUCE and STACK is the one to reach for in `convert-to-5g`:

> REDUCE and STACK most closely mimics traditional financial modeling
> calculation structures as it calculates one column at a time — just like
> traditional formulas — and can retrieve prior column balances.

## The five templates are calculation cores, not finished functions

Four of the five deliberately strip the help block:

> To focus on the calculations, we will remove all lines related to inline help.
> **We expect any function derived from this function to add customized inline
> help.**

So the shape is: take the **Base** template, and drop the calculation core into
its Procedure section. A template pasted as-is is not a 5G function.

---

## Base Template (verbatim)

The one that carries the full 5G scaffolding. Note the help section order —
`FUNCTION:` first — and that this is the lightweight `IF(OR(ISOMITTED(...)))`
tier, not the `ErrMsgs`/`Errors`/`Messages` tier our full template uses.

```excel
/*  FUNCTION NAME:  <enter function's name here with λ>
    DESCRIPTION:*//**<enter a brief description of the function's purpose>*/
/*  REVISIONS:      Date    Developer           Description
               <Mmm dd yyyy><Your name>         <What changed>
*/
FunctionNameλ = LAMBDA(
//  Parameter Declarations
    [Argument1],                // All arguments must be enclosed in square brackets
    [Argument2],                // Enter as many arguments as needed
    [Argument3],
//  Check required arguments
    IF( OR( ISOMITTED( Argument1),
            ISOMITTED( Argument2)
        ),
    //  If required arguments omitted display inline help text
        TRIM(TEXTSPLIT(
                "FUNCTION:      →FunctionNameλ( [Argument1], [Argument2], Argument3)¶" &
                "DESCRIPTION:   →<Copy the function's description to here>¶" &
                "WEBPAGE:       →<enter function's webpage URL here (if any)¶" &
                "VERSION:       →<Author's initials>, Mmm-dd-yyyy¶" &
                "PARAMETERS:    →¶" &
                " Argument1     →(Required) <argument description>¶" &
                " Argument2     →(Required) <argument description>¶" &
                " Argument3     →(Optional) <default value> <argument description>¶",
                "→", "¶" )
        ),
    //  If required arguments provided, process request
        LET(
        //  Set Defaults
            Argument3,  IF(OR(ISOMITTED(Argument3), Argument3=""), "default value", Argument3),
        //  Process Request
            LetStep1,   "enter formulas starting here",
            Result,     "enter last formula here",
            Result
        )
    )
);
```

The eight sections the course names, in order: Documentation Block, Function
Declaration, Argument Declarations, Omitted Required Arguments Test, Inline
Help, Defaults and Named Constants, Procedure, Function Result.

Note the bracket asymmetry, which is easy to get wrong: **in the signature,
bracket everything. In the `FUNCTION:` help line, bracket only the genuinely
optional ones.** The signature lies to Excel so help can display; the help text
tells the user the truth.

## Corkscrew Template (verbatim)

```excel
/*  FUNCTION NAME:  <enter name here>λ
    DESCRIPTION:*//**<this comment displays in name manager>*/
/*  REVISIONS:      Date        Developer       Description
                    <changed>   <your name>     <what changed>
*/
Corkscrewλ = LAMBDA(                                    //Change this name
//  Function Arguments
    [Opening],
    [Gains],                                            //Set arguments as needed
    [Losses],
//  Help
    LET(Help,   TRIM(TEXTSPLIT(
                        "FUNCTION:      →FunctionNameλ( Argument1, Argument2, [Argument3])¶" &
                        "DESCRIPTION:   →Copy description from comment block at top.¶" &
                        "WEBPAGE:       →URL¶" &
                        "VERSION:       →copy date from last revision in REVISIONS¶" &
                        "PARAMETERS:    →¶" &
                        "Argument1      →(Required) Description of required argument1¶" &
                        "Argument2      →(Required) Description of required argument2¶" &
                        "Argument3      →(Optional) Description of optional argument3 with what default is¶" &
                        "EXAMPLES:      →Formula (ID is assumed to be the module's name)¶" &
                                        "→=ID.FunctionNameλ( 123, 456)",
                    "→", "¶" )
                ),
    //  Check inputs - Omitted required arguments (do not include optional arguments)
        Help?,      OR( ISOMITTED( Argument1),          //In this example, Argument1 is an opening balance
                        ISOMITTED( Argument2 )          //In this example, Argument2 is an array to be processed
                    ),
    //  Pre calcs
        Flows,  Gains - Losses,                         //Change as needed
    //  Calculations
        CLose,  SCAN( Opening, Flows, SUM),             //Calculate Close
        Open,   HSTACK( Opening, DROP( Close, , -1)),   //Create Open
        Result, VSTACK( Opening, Gains, Losses, Close), //Stack rows
    //  Return Result
        IF( Help?, Help, Result)
    )                                                   // Close LAMBDA's LET
);                                                      // Close LAMBDA
```

Why a corkscrew needs `SCAN` at all — the clearest statement of it anywhere:

> A corkscrew calculation creates a pseudo circular equation when performed by
> dynamic arrays alone. I say pseudo because the value by value calculation flow
> is not circular. But Excel views an array as a single object, and not
> individual values within the array. So to Excel, a closing balance array that
> depends on an opening balance array and an opening balance array that depends
> on a closing balance array, is circular. To get around this situation, we
> perform all the calculations in one formula using SCAN().

Two notes worth carrying:

- `SCAN(Opening, Flows, SUM)` uses an **eta-reduced LAMBDA** — where the callback
  would be `LAMBDA(a, b, SUM(a, b))`, you may write just `SUM`.
- Both flows are positive and the outflow is subtracted. If your convention is
  outflows-as-negative, add rather than subtract.

## REDUCE and STACK Template (verbatim)

```excel
/*  FUNCTION NAME:  <enter name here>λ
    DESCRIPTION:*//**<this comment displays in name manager>*/
/*  REVISIONS:      Date        Developer       Description
                    <changed>   <your name>     <what changed>
*/
ReduceAndStackλ = LAMBDA(      //Change this name
//  Function Arguments
        [Opening],      // All arguments must have [ ] - For this template assume this is the opening amount
        [Array#1],      // Add as many arguments as needed. Not all need to be arrays.
        [Array#2],
//  Help
    LET(Help,   TRIM(TEXTSPLIT(
                        "FUNCTION:      →FunctionNameλ( Argument1, Argument2, [Argument3])¶" &
                        "DESCRIPTION:   →Copy description from comment block at top.¶" &
                        "WEBPAGE:       →URL¶" &
                        "VERSION:       →copy date from last revision in REVISIONS¶" &
                        "PARAMETERS:    →¶" &
                        "Opening        →(Required) Opening Balance¶" &
                        "Array1         →(Required) Description of required argument2¶" &
                        "Array2         →(Optional) Description of optional argument3 with what default is¶" &
                        "EXAMPLES:      →Formula (ID is assumed to be the module's name)¶" &
                                        "→=ID.FunctionNameλ( 123, 456)",
                    "→", "¶" )
                ),
    //  Check inputs - Omitted required arguments (do not include optional arguments)
        Help?,      OR( ISOMITTED( Opening),        //In this example, an opening balance
                        ISOMITTED( Array1)          //In this example, Argument2 is an array to be processed
                    ),
    //  Set internal named constants if needed

    //  Set defaults for optional arguments if needed
        Array2,     IF( ISOMITTED( Array2), <fill in default>, Array2),

    //  Calculate things that only need to be calculated once here
        Columns,    COLUMNS( <Array1>),
        Periods,    SEQUENCE( , Columns),

    //  Procedure
//  Process each column
        Result,     REDUCE(  Opening, Periods,
                        LAMBDA( Acc, Period,
                            LET(
                            //  Get Opening Balance
                                Open,   TAKE( Acc, -1, -1),
                            //  Get current period's values
                                Value1, INDEX( Array1, Period),
                                Value2, INDEX( Array2, Period),
                            //  Calculate close
                                Close,  Open + Value1 - Value2                          // Change as needed
                            //  Vertically stack values into a column
                                Stack,  VSTACK( Open, Value1, Value2, Close),           // Change as needed
                            //  Horizontally stack columns together
                                Result, IF( Period = 1, Stack, HSTACK( Acc, Stack)),
                            //  Return REDUCE's Result
                                Result
                            )                                                           // Close REDUCE's LAMBDA's LET
                        )                                                               // Close REDUCE's LAMBDA
                    ),                                                                  // Close REDUCE
        IF( Help?, Help, Result)
    )                                                                                   // Close LET
);                                                                                      // Close LAMBDA
```

**The accumulator read is the trick to remember:**

> whether the accumulator contains one value, as would be the case for REDUCE's
> initial value, or one or more columns from REDUCE's LAMBDA, the opening
> balance is usually found in the bottom right corner of the accumulator.
> `TAKE( Acc, -1, -1)` takes the bottom right row of single numbers or arrays.

The course's cell-to-LET-step mapping, which is the actual conversion recipe for
`convert-to-5g`:

| Grid formula | Becomes LET step |
|---|---|
| `J10: =I13` — this period's opening from last period's closing | `Open` |
| `J11: =J7` — this period's gain | `Gain` |
| `J12: =J8` — this period's loss | `Loss` |
| `J13: =J10 + J11 - J12` | `Close` |
| — | `Stack` — the values as one column, matching the grid's column |
| — | `Result` — the columns as one array, matching the whole section |

And the constraint: **it does not handle circularity.**

> This template shines when model sections are complex, as long as the
> traditional formulas do not involve circular equations.

## Circularity Template (verbatim)

The important one. This is how you recurse **without a public counter
argument** — the sub-function is a LET step, and it is passed to itself.

```excel
/*  FUNCTION NAME:  <enter name here>λ
    DESCRIPTION:*//**<this comment displays in name manager>*/
/*  REVISIONS:      Date        Developer       Description
                    <changed>   <your name>     <what changed>
*/
AverageBalanceInterestλ = LAMBDA(  //Change this name
//  Parameter Declarations
        [CashBeforeInt],    // Add as many arguments as needed.
        [BeginningDebt],    // All arguments must have [ ]
        [InterestRate],     // Change these names to suit

//  Help
    LET(Help,   TRIM(TEXTSPLIT(
                        "FUNCTION:      →FunctionNameλ( Argument1, Argument2, [Argument3])¶" &
                        "DESCRIPTION:   →Copy description from comment block at top.¶" &
                        "WEBPAGE:       →URL¶" &
                        "VERSION:       →copy date from last revision in REVISIONS¶" &
                        "PARAMETERS:    →¶" &
                        "Argument1      →(Required) Description of required argument1¶" &
                        "Argument2      →(Required) Description of required argument2¶" &
                        "Argument3      →(Optional) Description of optional argument3 with what default is¶" &
                        "EXAMPLES:      →Formula (ID is assumed to be the module's name)¶" &
                                        "→=ID.FunctionNameλ( 123, 456)",
                    "→", "¶" )
                ),
    //  Check inputs - Omitted required arguments (do not include optional arguments)
        Help?,      OR( ISOMITTED( CashBeforeInt),     //Change this to match argument 1's name
                        ISOMITTED( BeginningDebt),     //Change this to match argument 2's name
                        ISOMITTED( InterestRate )      //Change this to match argument 3's name
                    ),
    //  Set internal named constants if needed

    //  Set defaults for optional arguments if needed

    //  Iterative function definition (fn) - Change the parameters and LET() statement to suit
        Iterate,    LAMBDA( fn, CashBeforeInt, BeginningDebt, InterestRate, AverageDebt,
                            LET(Interest,       AverageDebt * InterestRate,
                                CFADS,          CashBeforeInt - Interest,
                                EndingDebt,     BeginningDebt - CFADS,
                                AverageDebt,    Average( BeginningDebt, EndingDebt),
                                Delta,          ABS( Interest - (AverageDebt * InterestRate)),
                                Result,         IF( Delta < 0.01,
                                                    Interest,
                                                    fn( fn,
                                                        CashBeforeInt, BeginningDebt, InterestRate, AverageDebt
                                                    )
                                                ),
                                Result
                        )
                    ),
    //  Procedure - Change parameters to suit
        Result,     Iterate( Iterate, CashBeforeInt, BeginningDebt, InterestRate, 0),
    //  Return result or Help
        CHOOSE( Help? + 1, Result, Help)
    )
);
```

The circularity it solves, stated as an equation:

> `Interest = Opening Debt + AVERAGE(Opening Debt, Opening Debt + Interest) * Interest Rate`

**How `fn` works.** `Iterate` is a LET step holding a LAMBDA whose *first*
argument is `fn` — the function to call for the next iteration. The outer call
seeds it by passing `Iterate` to itself: `Iterate(Iterate, ...)`. Each recursive
step does `fn(fn, ...)`, keeping the reference alive. This is why a 5G function
can recurse while staying self-contained: the recursion needs no workbook-level
name and no public argument.

**Convergence.** Seed the circular quantity at 0 (`AverageDebt` starts at 0, so
`Interest` is 0 on the first pass and `Delta` is large). Each pass shrinks the
delta.

> It typically takes about 8 iterations for function like this to achieve a
> small enough delta to stop calling itself and return a value.

**What the template does not give you.** No iteration cap, no divergence
handling — termination is purely `Delta < 0.01`. It assumes monotonic
convergence. Where divergence is plausible, add an explicit cap in its own
argument slot; see `source-structures.md`. `fx.DebtSizeIssueλ` is the worked
case of a problem where fixed-point iteration *does* diverge and bisection was
needed instead.

Note also `CHOOSE(Help? + 1, Result, Help)` — a third dispatch idiom, alongside
`IF(Help?, Help, Result)` and our `CHOOSE(Return, Result, Help, Messages)`.

---

## SCAN and STACK (reconstructed — no gist source)

Prose only on the site. The differences from REDUCE and STACK, per the page:

> The major difference here is we use SCAN to only calculate Close. From Close
> we can derive Open.

> The important thing to remember about SCAN is the opening balance is always in
> SCAN's LAMBDA's accumulator.

Three pre-calcs rather than two — `Columns`, `Periods`, and `Flows`, because

> Processing the flows outside of the SCAN function speeds processing.

Shape:

```excel
    Columns,  COLUMNS(Array1),
    Periods,  SEQUENCE(1, Columns),
    Flows,    Gains - Losses,
    Close,    SCAN(Opening, Periods, LAMBDA(Acc, Period,
                  Acc + INDEX(Flows, 1, Period))),
    Open,     HSTACK(Opening, DROP(Close, , -1)),
    Result,   VSTACK(Open, Gains, Losses, Close),
```

**A copy-paste artifact to not propagate:** the SCAN and STACK page still says
"This section is for performing calculations that can be performed once and not
for every iteration of **REDUCE()**" and refers to "the **REDUCE and STACK**
template" in its own pre-calcs and initial-value paragraphs. Hatmaker copied the
section across and did not finish updating it.

## Thunk and SCAN (reconstructed — no gist source)

See `source-structures.md` for the full thunk explanation and worked example.
The structural difference:

> The major difference here is we will retrieve period values from an array of
> Thunks instead of using INDEX against period arrays directly. This results in
> much faster calculations because instead of using INDEX to pull from a larger
> array, we will use INDEX to pull from a much smaller array — a Thunk
> containing two values.

```excel
    Thunk,   LAMBDA(x, LAMBDA(x)),
    Flows,   BYCOL(VSTACK(Gains, Losses), LAMBDA(col, Thunk(col))),
    Close,   SCAN(Opening, Flows, LAMBDA(Acc, Flow,
                 LET(Pair, Flow(),
                     Acc + INDEX(Pair, 1, 1) - INDEX(Pair, 2, 1)))),
    Open,    HSTACK(Opening, DROP(Close, , -1)),
    Result,  VSTACK(Open, Gains, Losses, Close),
```

The page's own template block is titled `ScanAndStackλ` — another copy-paste
error from the previous page.

---

## Bugs in the published templates

These are in the gist as downloaded. Do not copy them forward.

| Template | Bug |
|---|---|
| Corkscrew | `CLose,` defined with a capital L, then referenced as `Close` — undefined name |
| Corkscrew | `Result, VSTACK( Opening, Gains, Losses, Close)` — should be `Open`, not `Opening`; as written the first row is a scalar, not the opening row |
| REDUCE and STACK | No comma after `Close, Open + Value1 - Value2` — the binding list will not parse |
| REDUCE and STACK | `INDEX( Array1, Period)` — the two-argument form. Our convention is the explicit `INDEX(Array1, 1, Period)` |
| REDUCE and STACK | `SEQUENCE( , Columns)` — valid but obscure; prefer `SEQUENCE(1, Columns)` |
| REDUCE and STACK | `COLUMNS( <Array1>)` and `<fill in default>` — unreplaced placeholders |
| Circularity | `Average(` — should be `AVERAGE()` |

The `<angle bracket>` placeholders throughout are meant to be replaced. Strip
every one before delivering.

## Sources

- Course index — <https://sites.google.com/site/beyondexcel/home/5g-modeling/5g-course>
- Basic — <https://sites.google.com/site/beyondexcel/home/5g-modeling/5g-course/5g-source-structure-basic>
- Corkscrew — <https://sites.google.com/site/beyondexcel/home/5g-modeling/5g-course/5g-source-structure-corkscrew>
- REDUCE and STACK — <https://sites.google.com/site/beyondexcel/home/5g-modeling/5g-course/5g-source-structure-reduce-and-stack>
- SCAN and STACK — <https://sites.google.com/site/beyondexcel/home/5g-modeling/5g-course/5g-source-structure-scan-and-stack>
- Thunk and SCAN — <https://sites.google.com/site/beyondexcel/home/5g-modeling/5g-course/5g-source-structure-thunk-and-scan>
- Circularity — <https://sites.google.com/site/beyondexcel/home/5g-modeling/5g-course/5g-source-structure-circularity>
- Templates gist — <https://gist.github.com/CHatmaker/3e6cba6414fadfabfa3725255da2ba74>
