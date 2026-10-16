# Building LAMBDAs with Claude Code

What the two LAMBDA skills do, how they fit together, and how to get from an
idea to a tested function in your workbook.

## The 5G standard in one paragraph

Craig Hatmaker's **5th-Generation (5G)** modelling replaces grids of
copied-down formulas with LAMBDA functions that **spill** their whole result,
**document themselves** (call the function with no arguments and it returns its
own help), and **validate their inputs** (bad arguments return a readable
message, not `#VALUE!`). One function, one cell, no corkscrew rows. The
methodology is his — see [Beyond Excel](https://sites.google.com/site/beyondexcel/)
and his gists; these skills are an implementation of it.

## The two skills

| Skill | Use it when | It gives you |
|---|---|---|
| `convert-to-5g` | you **have** a working formula or block of formulas | one 5G LAMBDA that reproduces it, with help text and input checks |
| `compose-5g-lambda` | you want a **new** function, or a variant of one you already have | a new LAMBDA built from existing ones without breaking them — recursion and binding traps handled — plus a test |

`compose-5g-lambda` reads `convert-to-5g`'s templates and checklist, so both
are installed together (they're in the same plugin).

## The loop

1. **Describe or paste.** "Convert this to 5G: `=…`" or "write a 5G LAMBDA
   that …". Give a few input/expected-output pairs if you have them — they
   become the test.
2. **Claude writes the function** to the template: parameters, help block,
   error checks, then the calculation in a `LET`.
3. **Claude tests it before handing it over.** The test harness in
   `compose-5g-lambda/scripts/` evaluates the LAMBDA in a separate, hidden
   Excel and compares every result with a Python reference implementation,
   within a tolerance. You see a PASS/FAIL table. It also checks the function
   survives being stored (read back = written, and under Excel's 8,192-character
   formula limit).
4. **You paste it** into the Advanced Formula Environment (Excel Labs) or Name
   Manager, as the skill instructs, and use it.

Why test against Python rather than read the formula? Because Excel's
behaviour can't be fully predicted from the text. Lessons in the skills'
`pitfalls` reference come from exactly that: functions that looked right and
weren't. A run is cheaper than an argument.

**See it done end to end:** [Roll Your Own LAMBDA](roll-your-own-lambda/) is a
worked example. It covers the description, the prompt, the test results, and a
workbook to follow along in.

## Extending a function you already have

Point `compose-5g-lambda` at the existing function (paste it, or let it read
the workbook) and say what the variant should do. It keeps the original
intact and builds the new one around it — the common failure it prevents is an
"edit" that silently breaks a recursive helper.

## Finished libraries to learn from

These are complete, tested LAMBDA libraries. They are **not** installed by
this repo — import them yourself if you want them, through the AFE (Modules >
Import from URL):

- [Monte Carlo LAMBDAs](https://gist.github.com/wfphillips128/f91bff77212ab2c3d8f55a4f0a51b8b6) —
  distributions, copulas, risk measures, chart data; reproducible, non-volatile
  random numbers.
- [Time-series statistics LAMBDAs](https://gist.github.com/wfphillips128/96171c91caba82aa708b12b9f8f575ce) —
  simulators, ARIMA/GARCH fits and forecasts, normality and unit-root tests,
  regression diagnostics.

They document each function with comment headers rather than the 5G inline
help block. Ask `compose-5g-lambda` to "make a 5G variant of …" one of them and
it will wrap it to the standard.

## Credit

The 5G methodology, its templates and naming conventions are Craig Hatmaker's
([Beyond Excel](https://sites.google.com/site/beyondexcel/)). Go to the source
for the full treatment.
