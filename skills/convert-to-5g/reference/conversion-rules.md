# 5G conversion rules

Craig Hatmaker's conventions, plus what the working library actually does.

## Naming

- Function names end in **`λ`** (U+03BB, Greek small letter lambda).
- Library functions carry a dot namespace: `fx.DebtSculptIssuanceλ`,
  `fx.Timelineλ`. The prefix groups the library; the part after the dot is the
  function.
- `LET` step names are PascalCase and describe the *thing*, not the operation:
  `ClosingDebt`, `PeriodRate`, `NewColumn`. Name the last calculation `Result`.
- Excel table names start with `tbl` plus a very short name.

## Argument brackets — the rule people get wrong

**In the LAMBDA signature, bracket every argument**, required or not:

```excel
LAMBDA([CFADS], [DSCR], [COUPON], [DoNotUse], ...)
```

Why: if Excel considers an argument required, calling the function with nothing
returns `#VALUE!` before your code runs — so the inline help never displays. Making
everything optional lets `=MyFuncλ()` reach the `CHOOSE`/`ISOMITTED` branch and spill
the help table.

**Required-ness is declared in the help text**, two places:
- The `FUNCTION:` / `EXAMPLE:` line — required args bare, optional args `[bracketed]`.
- The `PARAMETERS:` block — `(Required)` or `(Optional) Defaults to X`.

You can verify this on any workbook: Excel stores optional LAMBDA parameters with the
`_xlop.` prefix. A correct 5G function shows `_xlop.` on all of them.

## Help text mechanics

- Delimiters are real Unicode: **`→`** (U+2192) between columns, **`¶`** (U+00B6)
  between rows. Not `->`. Not `|`.
- `TRIM(TEXTSPLIT(text, "→", "¶"))` — column delimiter first, row delimiter second.
- Every concatenated line ends with `&` **except the last**. A missing `&` is the most
  common reason a 5G function won't parse.
- Pad the label column with spaces so the source lines up; `TRIM` cleans it in the
  output.
- A parameter description too long for one row continues on the next with a **leading
  `→`** so it lands in the correct column:

```excel
" PositiveOnly? →(Optional) TRUE: Outflows are entered as positive values ¶" &
"→ that this function will subtract from the opening balance. ¶" &
"→ FALSE: Outflows are entered as negative values. Default is FALSE.",
```

- `"→¶"` on its own produces a blank spacer row.
- Useful sections beyond the required ones: `SUGGESTED ROW IDS:` (label each row the
  function returns), `EXAMPLE:` / `EXAMPLES:`.

## Array idioms

**Horizontal `SCAN`/`REDUCE` — use `SEQUENCE(1, n)`, not `SEQUENCE(n)`.** The 1-row
form produces horizontal output directly, avoiding a `TRANSPOSE`:

```excel
// inefficient
CloseBal, SCAN(Amount, SEQUENCE(Term), LAMBDA(bal, n, formula)),
CloseRow, TRANSPOSE(CloseBal),

// elegant
CloseBal, SCAN(Amount, SEQUENCE(1, Term), LAMBDA(bal, n, formula)),
```

**Prefer `DROP` over `TAKE`** when removing elements from the end — it avoids
referencing a length variable:

```excel
// inefficient
OpenBal, HSTACK(Amount, TAKE(CloseBal, 1, Term - 1))
// elegant
OpenBal, HSTACK(Amount, DROP(CloseBal, , -1))
```

**`INDEX` on a single-row array treats the second argument as the column number.**
`INDEX(rowArray, n)` returns the nth column. The library writes it explicitly as
`INDEX(Rate, 1, Period)` — clearer, do the same.

**Fixed multi-column range arguments** get expanded so sizes stay consistent:
`EXPAND(argument, 1, columns, 0)`.

**Spill references.** A dynamic-array anchor is passed as `G23#`, never `G23`. If a
referenced cell is inside a spill but isn't the anchor, trace back to the anchor.
Stored in XML this appears as `_xlfn.ANCHORARRAY(...)`.

## Excel syntax conventions

- Booleans uppercase: `TRUE`, `FALSE`.
- Function names uppercase: `INDEX()`, `SCAN()`, `REDUCE()`, `PMT()`.
- Dates in help text: `Mmm-dd-yyyy` (e.g. `May-15-2026`) or `Mmm dd, yyyy`. Use the
  current date for the REVISIONS and VERSION lines.

## Multiple selected formulas

When the user selects a block containing several different formulas:
- Identify all of them.
- Combine them into **one** LAMBDA, returning the rows stacked with `VSTACK` in the
  same order they appear on the sheet.
- Create parameters for the union of unique inputs.
- Document what each output row is under `SUGGESTED ROW IDS:` in the help text.
- Match the example output's shape and formats.

## Worksheet hygiene (when the user is building the sheet)

- Use cell **Styles** rather than applying colours directly.
- Remove gridlines; no borders; no wrapped rows.
- Apply the default table style; don't restyle table headers.
- Data validation: put the argument name in the Input Message **Title** and the
  argument's help description in the Input Message **body**.
- Treat the function's help text as authoritative for defaults and validation lists.
  Don't infer ranges.
