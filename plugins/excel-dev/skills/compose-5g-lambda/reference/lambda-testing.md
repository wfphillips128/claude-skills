# Testing a LAMBDA in real Excel

The shadow model (`shadow-model-guide.md`) proves the **maths**. This proves the
**formula** — the exact text that ships — by evaluating it in Excel and comparing
every cell against an independent Python reference.

Use it whenever the machine has desktop Excel. Excel's rules cannot be derived
from the formula text (`pitfalls.md` §27), so a formula that has only been read,
however carefully, has not been tested.

`scripts/lambda_test_harness_template.py` is the working starting point. It ships
with a small worked example — a straight-line depreciation schedule — that passes.
Replace the example with your function.

## What building the template itself turned up

Worth knowing before you trust a reading of any formula. On its first run against
its own worked example, which had been written carefully and read twice:

| Finding | What it means |
|---|---|
| `#CALC!` on every one-period case | `HSTACK(Cost, DROP(Closing, , -1))` drops the only column when there is one period. Fixed by stacking first: `DROP(HSTACK(Cost, Closing), , -1)`. See `pitfalls.md` §20 |
| The round-trip check failed on correct text | Excel rewrote the parameter list (spaces removed, `[Cost]` stored as `_xlop.Cost`). The comparison now normalises; see *Round trip* below |
| A planted off-by-one passed 23 of 61 cases | Every case whose timeline stopped before the asset's last period passed. Random cases must reach the boundaries |

## Requirements and safety

- Windows, desktop Excel 365 (LAMBDA and dynamic arrays), Python with `pywin32`.
- **Never touch the user's Excel.** The harness starts its own instance with
  `win32com.client.DispatchEx`, which always creates a new process, and records
  that process id. It never uses `GetObject` or `Dispatch`, which can attach to
  an Excel already running. On exit it quits its own instance and, only if that
  one process is still alive after 15 seconds, ends it. It never closes, kills or
  reads any other Excel process.
- Check before and after a run: `Get-Process EXCEL` in PowerShell. The list after
  should equal the list before. Another tool or agent may be running its own Excel
  at the same time; a process you did not start is not yours to end.
- **Never open the user's workbook.** The harness works in a new workbook saved
  to a temp folder, deleted afterwards (`--keep-files` to inspect it).
- **A corrupt makepy cache** (`module 'win32com.gen_py...' has no attribute
  'CLSIDToClassMap'`) breaks early binding. Do not delete a shared cache; force
  late binding before dispatching:
  `win32com.client.gencache.GetClassForCLSID = lambda clsid: None`. The template
  does this unconditionally.
- **Ticked add-ins load into a `DispatchEx` instance too.** That is usually
  harmless, but an add-in that shows a dialog on load, or any `MsgBox`, blocks a
  hidden Excel forever. Run with `--visible` if a run hangs.
- **Run alone.** Heavy concurrent Excel automation on the same machine turns COM
  calls into transient failures (`RPC failed`, `Member not found`). The template's
  `retry()` absorbs most of them; it cannot absorb a machine that is saturated.

## What the harness checks

| Step | Check | Why |
|---|---|---|
| 1 Lint | Names that are cell addresses; a LET-bound LAMBDA named like a built-in and called; names differing only by case; a parameter heading a dotted function; `-x^2`; `[[Arg]]`; comment text left in; length | Each fails in Excel with a generic or misleading error, or silently (`pitfalls.md` §12, §15-18) |
| 2 Install | `Workbook.Names.Add(Name, RefersTo)` with the flattened, comment-free text | The same route an installer uses. COM wants user syntax (`[Arg]`, no prefixes) |
| 3 Round trip | `RefersTo` read back equals what was written; the definition **stored in the saved file** equals it too; the stored length is under 8,192 | Proves Excel holds the text you think it holds, and measures the real limit |
| 4 Help | `=Func()` spills a table starting `FUNCTION:` | 5G inline help, test class 1 in `testing.md` |
| 5 Values | Fixed, boundary, invalid-input and seeded random cases against the Python reference, within a stated tolerance | The differential test |
| 6 Reopen | Save, close, reopen, recalculate one case | The definition survives as an ordinary defined name |

Exit code `0` all passed, `1` any check failed, `2` the harness could not run.
The summary prints a PASS/FAIL count per section and lists every failure.

## Adapting the template

Copy it, then edit the four marked blocks:

1. **`DEMO_SOURCE`** — the LAMBDA in AFE form, comments and indentation included,
   or leave it and pass `--source module.txt`. Several `Name = LAMBDA(...);`
   definitions may be given; all are installed, in whatever order makes them
   resolve, so helpers can sit in the same module.
2. **`FUNC`** — the name the value tests call.
3. **`reference()`** — the independent Python implementation. Return a list of
   rows; raise `ValueError` for inputs the LAMBDA must reject with an Excel error.
4. **`cases()`** — `(label, args)` pairs. `None` omits an argument; lists become
   array constants `{a,b,c}`; nested lists become 2-D `{a,b;c,d}`.

```powershell
python lambda_test_harness_template.py --lint-only            # no Excel
python lambda_test_harness_template.py                        # 50 random cases
python lambda_test_harness_template.py --cases 500 --seed 7 --quiet-cases
python lambda_test_harness_template.py --source my_module.txt --tol 1e-6
```

Keep the template **pure ASCII**: build the lambda character and the help
delimiters with `chr(0x03BB)`, `chr(0x2192)`, `chr(0x00B6)`. Non-ASCII characters
and backslashes written through a shell heredoc get mangled — an em dash once
truncated a heredoc outright, and `\b` became a backspace character inside a
Python raw string. Write scripts to a file with an editor, then run the file.

## Writing the reference

- **Write it differently from the LAMBDA.** A plain loop from the definition, not
  a line-by-line port of `SCAN` / `HSTACK`. A port shares the formula's
  misunderstandings; that is the shadow model's job, not this one's.
- **Decide invalid inputs in the reference.** If the reference raises, the
  harness expects an Excel error, and a plausible number for a nonsense input
  fails the test. That is test class 2 in `testing.md`.
- **When the two disagree, any of three things can be wrong**: the LAMBDA, the
  reference, or the test. Validate inverse functions by round trip rather than
  against another library's inverse (`testing.md`, *Making the tests honest*).

## Choosing cases

- A **hand case** you can check on paper, printed in full.
- **Both boundaries** of every dimension, including the **length-1** timeline and
  the case where an optional argument is omitted.
- **Invalid inputs**: zero divisors, negatives where only positives make sense,
  fractions where integers are required, text in a numeric argument.
- **Seeded random cases whose shape varies**, not just their values. Draw the
  dimensions so that a good share of cases cross each boundary. Seeded, so a
  failure reproduces exactly.
- **Inputs that vary where the real data might not.** A constant rate hides a
  per-period error (`pitfalls.md` §14).

## Tolerance

The template compares `|expected - actual| <= tol * max(1, |expected|)` — relative
above 1, absolute below. The default `1e-9` suits arithmetic that Excel and Python
do in the same double precision. The worked example's worst cell differed by
1.5e-9 on a value near a million — about 1e-15 relative, pure floating-point
noise, yet enough to fail a purely absolute 1e-9 test. That is why the
comparison is relative above 1.
Use `0.01` on currency only where the function itself rounds, and **exact** on
amounts snapped to an increment. State the tolerance in the output, and never
widen it to make a run pass — a wider tolerance is a finding to report.

## COM mechanics that bite

| Trap | What happens | What the template does |
|---|---|---|
| `Range.Formula` on a dynamic-array formula | Excel inserts the implicit-intersection operator, the formula becomes `=` + at-sign + `SEQUENCE(3)`, and returns one cell | Writes `Range.Formula2` |
| Reading a spilled result | `cell.Value` gives only the anchor | `HasSpill` then `SpillingToRange.Value` |
| Error cells | Arrive as large negative integers (`#N/A` is -2146826246), not strings | Reads `cell.Text` for errors; treats such integers as errors |
| Writing inside another formula's spill range | Silently breaks that spill into `#SPILL!`, which surfaces later somewhere unrelated | Clears the sheet between cases; one formula at a time |
| Busy Excel under late binding | Not "call rejected" but `AttributeError: <unknown>.Name`, or `TypeError: 'bool' object is not callable` when a no-argument method already ran during the lookup | `retry()` recognises all three shapes |
| `Workbook.Activate` in a hidden instance | Does not reliably change `ActiveWorkbook` | Holds explicit workbook and sheet objects |
| Printing the lambda character | A cp1252 console raises `UnicodeEncodeError` | Reconfigures stdout to UTF-8 |
| A crashed run | Leaves an orphan `EXCEL.EXE` | `try/finally` quit, then ends only its own recorded PID |

## Round trip

Excel does not store the text you give it. Verified on the worked example:

- spaces in the LAMBDA's parameter list are removed;
- an optional parameter `[Cost]` is stored as `_xlop.Cost` (no brackets);
- functions gain `_xlfn.`, worksheet-only ones `_xlfn._xlws.`, and every
  parameter reference and `LET` name gains `_xlpm.`;
- whitespace inside `{array constants}` is tightened.

So the comparison normalises both sides — prefixes stripped, `_xlop.X` turned
back into `[X]`, whitespace outside string literals removed — and then requires
every token and every string literal to match exactly. Removing whitespace is
safe here only because a self-contained 5G function holds no range references:
between two ranges a space is Excel's intersection operator.

**The 8,192 limit applies to the stored form.** The harness measures it by
reading `xl/workbook.xml` out of the saved file. On the worked example 1,017
written characters were stored as 1,359 — a third longer. The estimator in
`scripts/stored_length.py` put the same function at 1,390, about 2% high; good
for planning, but measure when you are within a few hundred characters of the cap.

## Volatility and timing

Not in the template, because most functions need neither — but when they do:

- **Volatility**: add a no-argument VBA UDF that increments a counter, call it in
  the same cell as the function, edit an unrelated cell, and check the counter
  did not move. Needs *Trust access to the VBA project object model*.
  `pitfalls.md` §24 explains why volatility hides.
- **Timing**: set `Application.Calculation` to manual, then time
  `Range.Calculate()` on the cell. In automatic mode, COM timings include
  repeated recalculations.

## When a case fails

1. Copy the failing arguments (printed under the case) into the fixed list in
   `cases()`, then re-run with `--cases 0 --visible --keep-files` and look at the
   cell in Excel.
2. Swap the function's last `LET` step name for an earlier step (5G Standard 10)
   and re-run to see that step's value. That is what the bare final name is for.
3. If Excel refuses the formula, bisect the `LET` bindings (`pitfalls.md` §27).
4. If the reference is suspect, check it against a hand calculation or a third
   route before changing the LAMBDA.

Fix, then re-run the whole suite — not just the case that failed.
