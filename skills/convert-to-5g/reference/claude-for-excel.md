# Handing off to Claude for Excel

Claude Code runs outside Excel. Claude for Excel (the Anthropic add-in) runs *inside*
it, so it can see the active selection and touch the workbook. Neither replaces the
other — this is a division of labour.

## What each side is good at

| | Claude Code (this skill) | Claude for Excel |
|---|---|---|
| Sees your selection | No — you paste | Yes |
| Reads the whole workbook | Yes, via `openpyxl` (formulas + defined names) | Within the add-in's reach |
| Writes to the names collection | No | Yes |
| Evaluates a LAMBDA | No | Yes — it's in a live Excel |
| Runs Python to prove the maths | Yes | No |
| Reads PDF benchmarks, ties out schedules | Yes | No |
| Holds long reference material | Yes — these skill files | Limited |

Rough rule: **design and verify here, place and test there.**

## When to hand off

Hand off when the remaining work is *interaction with a live workbook*:

- Adding the finished function to the names collection
- Confirming the function actually spills and returns the right shape
- Iterating on a `#VALUE!` / `#NUM!` that only reproduces in Excel
- Reading a selection whose layout is hard to describe in text
- Applying styles, data validation, table formatting

Stay here when the work is *reasoning or proof*:

- Deriving the recursion
- Reconciling against a benchmark schedule or PDF
- Composing from an existing library (use `compose-5g-lambda`)
- Anything needing a Python shadow model

## Handoff prompts

Give the user these verbatim, filled in.

**To add the function.** Pick the route by size and by whether it recurses. Name
Manager only works for a short, non-recursive function — its *Refers to* box refuses
long text outright, and a name cannot reference itself before it exists.

*Short (under ~1,000 chars) and not recursive:*

> Add a workbook-scoped defined name called `<FullName>` whose formula is the LAMBDA
> below. Paste it exactly — do not reformat, do not add line breaks, do not add
> comments. Then confirm the function evaluates in a cell.
>
> ```
> <the single-line form>
> ```

*Long or recursive — use Excel Labs' Advanced Formula Environment:*

> Install the free [Excel Labs](https://www.microsoft.com/en-us/garage/profiles/excel-labs/)
> add-in, open the Advanced Formula Environment, and paste the commented form ① into
> the Modules pane. AFE strips the comments on import, so they cost nothing.

*Or install it programmatically, which also works for recursive functions:*

```python
# COM / VBA object model — wants USER syntax: [Arg] brackets, no _xlfn. prefixes
wb.Names.Add(Name="fx.Exampleλ", RefersTo="=" + single_line_form)

# openpyxl — the opposite: wants the STORED form, prefixes intact
wb.defined_names[name] = DefinedName(name, attr_text=stored_form)
```

Verify by **evaluating the function in a cell**, never by checking that Name Manager
displays it. Name Manager draws an all-optional LAMBDA without its `LAMBDA(` head or
its commas — a display bug that looks exactly like a corrupted definition.

**To test it:**

> In cell `<TestCell>` enter `=<FullName>(<real args, using # for spill anchors>)`.
> It should spill `<rows>` rows by `<cols>` columns. Then in `<CompareCell>` enter
> `=MAX(ABS(<TestRange> - <OriginalRange>))` and tell me the value — it should be
> under 0.01.

**To get the help table:**

> In an empty cell enter `=<FullName>()` with no arguments. It should spill the
> function's help table. Paste back what it shows.

**To diagnose an error:**

> The function returns `<error>`. Please: (1) select the cell and read out the exact
> error, (2) use Formulas → Evaluate Formula and tell me the first step that errors,
> (3) paste the values of `<the suspect inputs>`.

## Coming back

When Claude for Excel reports a result, bring it back here with the actual numbers and
error text — not a paraphrase. This skill can then correct the source and reissue both
output forms.
