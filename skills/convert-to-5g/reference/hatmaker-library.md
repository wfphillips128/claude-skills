# Hatmaker's published 5G library

What already exists, where to import it from, and where the published material
differs from the conventions in `conversion-rules.md`.

Harvested from <https://sites.google.com/site/beyondexcel> on 15 Aug 2025 using
the `webscrape` skill; the Creator's Course pages and templates gist were added
05 Sep 2026. The site renders via JavaScript, so it needs a real browser to
fetch — plain HTTP returns an empty shell. Re-harvest with:

```
firecrawl.py scrape "https://sites.google.com/site/beyondexcel/home/5g-modeling/5g-component-libraries/<library>"
```

**The course, not the catalogue, is where the standard is written down.** See
`5g-standard.md` for the 13 published standards, and
`compose-5g-lambda/reference/hatmaker-templates.md` for the five source-structure
templates verbatim. Note that the templates exist on the site only as
screenshots — the real code is one gist:
<https://gist.github.com/CHatmaker/3e6cba6414fadfabfa3725255da2ba74>

## Read this before writing a new function

Roughly 120 5G functions already exist and are free to import. Writing a new
`Corkscrewλ` or `ScheduleValuesλ` from scratch wastes effort and produces
something less tested than the original. Check the catalogue below first; if a
function exists, import the library and use it. If something close exists,
extend it rather than starting over — that is what `compose-5g-lambda` is for.

## The five importable libraries

Each URL is a GitHub gist that Microsoft's free **Advanced Formula Environment**
add-in can import directly into a workbook.

| Library | Import URL | Functions |
|---|---|---|
| Array Essentials | `gist.github.com/CHatmaker/d0829c4bde303f5f8602c21158831323` | 42 |
| Financial Starter | `gist.github.com/CHatmaker/472f65269519a0937b03f741e9271334` | 26 |
| Financial Ratios | `gist.github.com/CHatmaker/19caa8b52144de03edf0e08de6765cfc` | 38 |
| Dates and Scheduling | `gist.github.com/CHatmaker/fcbd2ffab850a4affc80ac0553cf732a` | 11 |
| Debt | `gist.github.com/CHatmaker/28031fac3d3e91dc5997c80fd4a180ec` | 7 |

All of Hatmaker's gists: <https://gist.github.com/CHatmaker>

The site states that whole libraries should be imported rather than individual
functions — the entire Financial library adds under 11 KB and testing showed no
calculation penalty. Importing complete libraries also makes version upgrades a
single step. **The gists are more current than the site pages**, which say
outright that a page's code is "as of this writing" and point at the gist for
the latest.

## Function catalogue

### Array Essentials

*Miscellaneous* — `IsBetweenλ` `Instanceλ` `IsInListλ` `RangeToDaλ`
(RangeToDaλ converts static actuals or inputs to dynamic arrays.)

*Array basics* — `GetDiagonalsλ` `FlipArrayλ` `RotateArrayλ`

*Lookup and filter* — `BYCOLContainsλ` `FilterContainsλ` `MLookupλ` `XYLookupλ`
(`MLookupλ` matches 1+ criteria across 1–4 rows or columns; `XYLookupλ` matches
criteria in a column **and** a header — the two-dimensional case.)

*On-change suite* — `OnChangeλ` `OnChangeCountλ` `OnChangeResetλ`

*Running totals* — `RunTotColsλ` `RunTotRowsλ`

*Rolling suite* — `Rollingλ` `RollAvgλ` `RollMaxλ` `RollMinλ` `RollSumλ`
(`Rollingλ` is a helper in the shape of `BYCOL` for rolling windows; the rest
fix the applied function.)

*Aware suites* — three parallel families distinguished by suffix:

| Suffix | Meaning | Driver | Members |
|---|---|---|---|
| `A` | Array-aware — processes the whole array | `Arrayλ` | `AndAλ` `AvgAλ` `MaxAλ` `MinAλ` `OrAλ` `SumAλ` |
| `R` | Row-aware — knows which array row it sits beside | `ProcessRowλ` | `AndRλ` `AvgRλ` `MaxRλ` `MinRλ` `OrRλ` `SumRλ` |
| `C` | Column-aware | `ProcessColλ` | `AndCλ` `AvgCλ` `MaxCλ` `MinCλ` `OrCλ` `SumCλ` |

Row-awareness removes the need to select a row with `CHOOSEROWS()` first. The
driver function takes an optional `Function` argument (defaulting to `SUM`); the
suffixed members fix it. This suffix convention is worth imitating when building
a family of related functions.

### Financial Starter

*General* — `Allocateλ` `AvgBalIntλ` `Cumulativeλ` `Growthλ`
`CompoundInterestλ` `EffAnnRateλ` `EffPrdRateλ` `Inflationλ` `IntOnIntλ`
`IRRλ` `Movementλ` `PMTAλ` `Reversalλ` `SumContainsλ` `SumPeriodsλ`

*Amortisation* — `Amortizeλ` `SumAmortizeλ` `LabelAmortizeλ`

*Corkscrews* — `Corkscrewλ` `CorkscrewAλ` `CorkscrewRxOλ`
`CorkscrewReversalλ` `CorkscrewAvgBalIntλ`

*Depreciation* — `Depreciateλ` `SumDepreciateλ` `LabelDepreciateλ`

Note the recurring `X` / `SumX` / `LabelX` triad: the calculation, its
aggregation, and the row labels it produces. Follow that when building a
function whose output is a multi-row block.

### Financial Ratios

*Liquidity* — `CurrentRatioλ` `QuickRatioλ` `CashRatioλ`
`OperatingCashFlowRatioλ`

*Efficiency* — `ReceivablesTurnoverRatioλ` `InventoryTurnoverRatioλ`
`WorkingCapitalTurnoverRatioλ` `AssetTurnoverRatioλ` `DSIRatioλ`
`OperatingRatioλ`

*Leverage* — `DebtRatioλ` `DSCRλ` `DebtToCapitalRatioλ` `EquityMultiplierλ`
`DebtToEquityRatioλ` `DebtToAssetRatioλ` `InterestCoverageRatioλ` `EquityRatioλ`

*Margins* — `GrossMarginλ` `EBITDAMarginλ` `OperatingMarginλ` `PretaxMarginλ`
`NetProfitMarginλ` `CashFlowMarginλ`

*Returns* — `ROAλ` `ROEλ` `ROIλ` `ROICλ`

*Market* — `PriceEarningsRatioλ` `PriceToBookRatioλ` `PriceToSalesRatioλ`
`PriceToCashRatioλ` `BVPSλ` `CAPERatioλ` `DividendPayoutRatioλ`
`DividendYieldRatioλ` `EPSλ` `RetentionRatioλ`

### Dates and Scheduling

`CountDOWλ` `CountDOWInMonthλ` `IsBetweenλ` `IsLeapYearλ` `IsOccurrenceDateλ`
`OverLapDaysλ` `PeriodLabelλ` `Periodsλ` `ScheduleRatesλ` `ScheduleValuesλ`
`Timelineλ`

`ScheduleValuesλ(Timeline, Starts, Values)` is the standard answer to "values
typed across a timeline that rarely change" — enter only the change dates and
let it spill the rest. Start dates must be sorted ascending.

### Debt

`PMTAλ` `GetBalanceλ` `TLAλ` `TLAActλ` `TLBλ` `TLBActλ` `Installmentλ`

Term Loan A and Term Loan B schedules, each with an `Act` variant that folds in
actuals, plus a constant-payment installment schedule. The library page also
maps the generated row labels onto financial-statement lines (Debt Opening,
Draw, Repayment, Extra Payment, Closing; Interest Opening, Charge, Payment,
Closing) — useful when the output has to slot into a three-statement model.

Note `IsBetweenλ` and `PMTAλ` each appear in two libraries. Importing both
libraries will collide; keep one definition.

## Two template styles, and which to use

`5g-template-full.md` documents the `LET` / `Help` / `ErrMsgs` structure
transcribed from a working workbook. The **published** simple functions use a
different, lighter shape — a `/* */` header comment above a plain LAMBDA:

```excel
/*
FUNCTION NAME: ScheduleValuesλ
DESCRIPTION:   Schedules values in a timeline from a schedule in a table.
EXAMPLE:       =ScheduleValuesλ(Timeline, tblRates(Start), tblRates(Values))
ARGS:
  Timeline   A horizontal range or dynamic array holding the model's timeline
  Starts     A table column or array of dates when values take effect
  Values     Table column or array of values that change with dates
GUILTY PARTIES: Craig Hatmaker 2022
*/
ScheduleValuesλ = LAMBDA(
    // Parameter Declarations
    Timeline,
    Starts,
    Values,
    // Function starts here
    XLOOKUP(Timeline, Starts, Values, "", -1)
);
```

Both are genuine Hatmaker patterns serving different targets, so the choice is
not stylistic:

- **Comment-header style** works in the Advanced Formula Environment and in
  gists, where comments are preserved. It is the right shape for a library
  function distributed as source, and it is much less work.
- **Inline-help style** (`LET` + `Help` + `CHOOSE`) survives the Name Manager,
  which discards comments, and it is the only one that makes `=Funcλ()` spill a
  help table to the user in the grid. Everything in `conversion-rules.md` about
  bracketing every argument exists to serve that behaviour.

Choose by where the function will live and whether in-grid help matters. A
function going into a shared workbook for non-technical users wants inline help;
a function published as library source does not need it.

## Where the published material differs from `conversion-rules.md`

These are conflicts, not corrections. `conversion-rules.md` derives from working
code the user has tested, and it says so; a website is not automatically more
authoritative. Flag them rather than silently reconciling.

| Topic | `conversion-rules.md` | Published site | Reading |
|---|---|---|---|
| Namespace | `fx.` prefix, e.g. `fx.Timelineλ` | Bare names in the catalogue, e.g. `Timelineλ` | **No longer read as a conflict.** Hatmaker's own templates write calls as `=ID.FunctionNameλ( 123, 456)`, glossed *"ID is assumed to be the module's name"*, and Standard 02 requires names be unique only *within a module*. The dot-prefix is his convention too; the catalogue pages simply list the bare part. Keep `fx.`. |
| λ position | Suffix: `ScheduleValuesλ` | Suffix everywhere, **except** one `EXAMPLE:` line reading `=λScheduleValues(...)` | Suffix is the rule. The leading-λ instance is a stale artefact in his own docs. |
| Validation block | `ErrMsgs` / `Errors` / `Messages` required for anything reusable | Absent from published simple functions | Correlates with the template split above, not with disagreement about quality. |
| Spelling | — | `IsOccurrenceDateλ` in the catalogue, but the page URL is `lambdaisoccurancedate` | The function name is the double-r spelling; the URL misspelling is historical. |

## Related pages worth fetching if a task needs them

- `home/5g-modeling` — the 5G overview
- `home/5g-modeling/5g-mb-start` and `/5g-mb-comparison` — model-build tutorial
- `home/excel-library/n-fold-cartesian-formula` — multi-dimensional modelling,
  the largest single page on the site
- `home/5g-modeling/5g-component-libraries/5g-array-essentials-library/5g-row-aware-suite`

Individual function pages follow
`home/5g-modeling/5g-component-libraries/<library>/5g-<name>λ`, with λ
percent-encoded as `%CE%BB`.
