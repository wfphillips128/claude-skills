# 5G conversion rules

Craig Hatmaker's conventions, plus what the working library actually does.

## Naming

- Function names end in **`λ`** (U+03BB, Greek small letter lambda). It goes at
  the *end*, and IntelliSense means nobody has to type it.
- Library functions carry a dot namespace: `fx.DebtSculptIssuanceλ`,
  `fx.Timelineλ`. The prefix groups the library; the part after the dot is the
  function. Hatmaker's own templates use the same shape —
  `=ID.FunctionNameλ( 123, 456)`, where *"ID is assumed to be the module's
  name"*. Names need only be unique **within a module**.
- `LET` step names are PascalCase and describe the *thing*, not the operation:
  `ClosingDebt`, `PeriodRate`, `NewColumn`. Name the last calculation `Result`.
- Excel table names start with `tbl` plus a very short name.

### Concise versus conveys-purpose

These two goals fight, and the course is explicit that the job is to balance
them rather than pick one. Its worked example is Excel's own `DB()` —
depreciation, declining balance — which is concise and conveys nothing:

| Candidate | Problem |
|---|---|
| `DB()` | Too concise. Conveys no purpose. |
| `DEPRECIATIONDECLININGBALANCE()` | Conveys purpose, unreadable. |
| `DepreciateDecliningBalance()` | Readable, still not concise. |
| `DepnDecBal()` | Concise, but there is no standard abbreviation for either word. |
| **`DepreciateDB()`** | **The answer.** "Depreciate" supplies the context that lets `DB` be read as declining balance. |

So: camel case for readability, standard abbreviations where they exist, and
**context to disambiguate an abbreviation** where they do not.

### Argument names

- Same rules, but users never type them, so conciseness matters less than
  clarity.
- **Be consistent across the library.** Decide once whether it is `Opening` or
  `OpeningBalance`, and use that everywhere. Arguments recur across functions in
  a way function names do not.
- Avoid single letters unless the context makes them obvious — `MAKEARRAY`'s
  `R` and `C` are fine; a lone `D` is not.

### LET step and constant names

- Concise and unique within the `LET`. **Use a mnemonic once a name would run
  past about ten characters.**
- Nested `LET`s can see outer step names.
- Name your constants, and gloss the mnemonic on the same line. Data becomes
  information:

```excel
    MPY,        12,                     //Months Per Year
    Allocation, AnnualAmount / MPY,
```

  A bare `12` in a formula tells a maintainer nothing; `MPY` tells them what the
  12 *is*.

### The last LET step must be a name

`LET` allows a formula as its final step. 5G does not:

```excel
    Return, IF(AND(Errors), 2, IF(OR(Errors), 3, 1)),
    Output, CHOOSE(Return, Result, Help, Messages),
    Output                                  // a name, not the expression
```

The reason is debugging — swap `Output` for any other step name and that step's
value spills to the grid. See Standard 10 in `5g-standard.md`.

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

- Delimiters are real Unicode: **`→`** (U+2192, ALT-26) between columns, **`¶`**
  (U+00B6, ALT-0182) between rows. Not `->`. Not `|`. They were chosen because
  they are unlikely to appear in help text — the default `TEXTSPLIT` delimiters
  break on any description containing a comma. To type them: hold ALT and use
  the numeric keypad, or Insert > Symbol into a cell and copy the result.
- `TRIM(TEXTSPLIT(text, "→", "¶"))` — column delimiter first, row delimiter second.
- Every concatenated line ends with `&` **except the last**. A missing `&` is the most
  common reason a 5G function won't parse.
- Pad the label column with spaces so the source lines up; `TRIM` cleans it in the
  output. The padding exists purely to make the source resemble the rendered
  table — format it to look like what the user will see.
- **`TRIM` eats indentation you meant to keep.** Where a space has to survive —
  indenting a continuation line under its parent, say — use U+2003 EM SPACE,
  which `TRIM` does not remove. It shows in the source as an outlined block.

### Section order

The published standard's order, and what each line is for:

| Line | Required? | Content |
|---|---|---|
| `FUNCTION:` | Yes | Name and syntax. **The only place the user learns which arguments are required** — bracket the optional ones here, since the signature brackets everything. |
| `DESCRIPTION:` | Yes | One or two sentences. Copy from the documentation block. |
| `WEBPAGE:` | Optional | URL for a page about this function. Omit the line if there is none. |
| `VERSION:` | Yes | Author's initials and the date of the most recent change. |
| `PARAMETERS:` | Yes | Each argument with `(Required)` or `(Optional)` and a description. |
| `EXAMPLE:` / `EXAMPLES:` | Where practical | A real call. |

`SUGGESTED ROW IDS:` is ours, not the standard's — it labels the rows a
multi-row function returns and is worth keeping for anything that `VSTACK`s an
output block.
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

Caution: if `Term` can be 1, `DROP(CloseBal, , -1)` is empty and the whole result is
`#CALC!`. `DROP(HSTACK(Amount, CloseBal), , -1)` gives the same answer for every
length, including 1 (`compose-5g-lambda/reference/pitfalls.md` §20).

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
