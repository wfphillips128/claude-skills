# The Python shadow model

Claude Code can't execute a LAMBDA. The shadow model is the substitute: a Python
translation of the recursion, run against real inputs, reconciled to a benchmark. It
turns "this formula looks right" into "this formula ties to the penny".

## The rule that makes it worth anything

**Translate, don't reimplement.** Same variable names, same order of operations, same
rounding, same sign conventions as the LAMBDA. If you write cleaner or smarter maths in
Python, a passing tie-out proves the Python is right and tells you nothing about the
formula you're shipping.

Concretely:

| LAMBDA | Python |
|---|---|
| `PeriodRate, INDEX(Rate, 1, Period)` | `period_rate = rate[period - 1]` |
| `FLOOR(x, incr)` | `math.floor(x / incr) * incr` |
| `ROUND(x, 2)` | `round(x, 2)` — but see *Rounding* below |
| `MAX(Interest, 0)` | `max(interest, 0.0)` |
| `IF(Period = ModelPeriods, ..., recurse)` | the loop's terminating condition |

Excel is 1-based, Python 0-based. Write `period` 1-based and index with `period - 1` so
the code reads like the formula. Off-by-one here is the most common translation bug.

## Structure

Write the recursion as a **loop**, not Python recursion — you want the intermediate
state visible for the variance table.

```python
"""Shadow model for <Functionλ>. Mirrors the LAMBDA step for step."""
import math

def schedule(opening_balance, cfads, dscr, coupon, months_per_period, par_incr):
    n = len(cfads)
    rows = []                       # one dict per period
    principal = opening_balance
    for period in range(1, n + 1):
        i = period - 1
        rate = coupon[i] / 12 * months_per_period[i]
        interest = solve_interest(cfads[i], dscr[i], principal, rate, par_incr[i])
        payment = math.floor((cfads[i] / dscr[i] - interest) / par_incr[i]) * par_incr[i]
        closing = principal - payment
        rows.append(dict(period=period, opening=principal, interest=interest,
                         payment=payment, closing=closing))
        principal = closing          # carry state — mirrors the recursive first arg
    return rows

def solve_interest(cfads, dscr, principal, rate, par_incr, tol=0.01, max_iter=100):
    """Mirrors the convergence recursion in Interestλ."""
    payment = cfads / dscr
    for _ in range(max_iter):
        interest = principal * rate
        repayment = -math.floor((cfads / dscr - interest) / par_incr) * par_incr
        if abs(payment + repayment) < tol:
            return max(interest, 0.0)
        payment = -repayment
    raise RuntimeError("did not converge — the LAMBDA would return #NUM!")
```

Note the `max_iter` guard raising rather than looping forever. **A shadow model that
fails to converge is a finding**, not a bug in the shadow model — it means the LAMBDA
returns `#NUM!` for these inputs. Report it.

## Getting real inputs

Read them from the workbook. Never retype numbers.

```python
import openpyxl
wb = openpyxl.load_workbook(PATH, data_only=True)   # values, not formulas
ws = wb["Bond Model"]
cfads = [ws.cell(row=r, column=10).value for r in range(18, 36)]   # column J
```

Two loads serve different purposes:

- `data_only=True` → the **cached values** Excel last calculated. What you feed the
  shadow model, and what you reconcile against.
- `data_only=False` (default) → the **formula text**. What you read to understand the
  existing logic.

`data_only=True` returns `None` if the workbook was never opened and saved by Excel.
If you get `None`, say so — don't substitute zeros.

## Tolerances

| Quantity | Tolerance |
|---|---|
| Dollar amounts | `0.01` |
| Rates, ratios | `1e-9` |
| Par amounts snapped to an increment | exact — any difference is a real break |

State the tolerance in the output. Don't quietly widen it to make a run pass; if you
need a wider tolerance, that's a finding to report.

## Tie out against more than one dataset

One passing benchmark proves the function works *on that benchmark*. If a second is
available — a different scenario, rate set, or structure — run it, and prefer the one
where inputs **vary** over the one where they are constant. Constant inputs collapse
distinct formulas into the same arithmetic and let structural errors through untouched
(see `pitfalls.md` §14).

When you do generalize a function, re-run the original case as a **regression**. A
change that ties out on the new scenario but moves the old one is not a
generalization — it is a second bug cancelling the first.

## Stage the tie-out

Reconciling everything at once tells you *that* something is wrong. Staging tells you
*what*.

1. **Base case.** Turn the new feature off. The shadow model must reproduce the
   existing function / existing grid exactly. If it doesn't, the translation is wrong —
   fix that before touching the new feature.
2. **First period only.** Hand-check period 1 against the benchmark. Most sign and
   ordering errors surface here.
3. **Full schedule, feature off.**
4. **Full schedule, feature on**, against the benchmark that includes it.
5. **Both scenarios** if there are variants (e.g. tax-exempt and taxable).

## Presenting the tie-out

A variance table, every period, no elisions:

```
Tie-out: shadow model vs benchmark (tax-exempt)   tolerance $0.01

Period  Date        Model Par      Benchmark Par   Diff
     1  2027-01-01   12,345,000.00   12,345,000.00    0.00
     2  2027-07-01   12,890,000.00   12,890,000.00    0.00
   ...
    18  2035-07-01    9,115,000.00    9,115,000.00    0.00
                     -------------   -------------   -----
    Total           215,430,000.00  215,430,000.00    0.00

PASS — 18 of 18 periods within tolerance
```

Include the total row. A schedule can tie period by period and miss the total if a
period was dropped, and it can tie in total while individual periods offset.

If it fails, show the failing periods with their differences and say plainly which step
you suspect. Never report a partial tie-out as success.

## Extracting a benchmark from PDF

```powershell
python -m pip install pypdf
```

```python
from pypdf import PdfReader
text = "\n".join(p.extract_text() for p in PdfReader(path).pages)
```

Then parse the schedule rows with a regex and **print what you parsed** before comparing
— a mis-parsed benchmark produces a confident, wrong tie-out. Check the parsed row count
against the expected period count and the parsed total against the PDF's printed total.

If the PDF is a scan with no text layer, `extract_text()` returns empty or garbage. Say
so and ask for the numbers another way rather than guessing.

## Keeping it honest

- The shadow model is a **check**, not the deliverable. The LAMBDA is what ships.
- If you change the LAMBDA after a passing tie-out, the tie-out is void. Re-run it.
- Keep the script — it's the audit trail for the number that ends up in a bond sizing.
