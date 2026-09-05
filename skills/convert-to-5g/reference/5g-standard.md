# The 5g Standard

The published specification, from Craig Hatmaker's 5g Function Creator's Course:
<https://sites.google.com/site/beyondexcel/home/5g-modeling/5g-course/5g-standards>

This is the *why* behind the templates. The other reference files say what to
type; this one says what a 5G function is obliged to do, and it is the thing to
check a finished function against.

The course's framing is worth keeping in mind, because every standard follows
from it:

> When we create functions for ourselves we are free to figure things out as we
> go, use cryptic naming, neglect structure, and neglect testing — because when
> it fails, we can fix it. **These freedoms can create disaster for our
> clients.** When we create for others we must make sure things are figured out
> beforehand, because we won't have access to the function once handed to
> others.

## The 13 standards

Quoted, lightly condensed where the original repeats itself.

01. **Designed for reuse**, preferably across a wide variety of scenarios.

02. **Look and feel of Excel's own functions**, which means:
    - unique, clear, concise names — not confusable with native functions
    - a tooltip on entry (desktop Excel only)
    - clear argument names
    - inline help

03. **Inline help appears when any required argument is missing**, and must
    contain:
    - function name and syntax
    - brief description
    - *(optional)* webpage / URL
    - version date
    - argument names with `(Required)` or `(Optional)` and a description

04. **All arguments enclosed in square brackets** so Excel treats them as
    optional, which is what lets the help display at all.

05. **100% native Excel functions and operators** — so it runs in any workbook.

06. **100% self-contained** — see the next section.

07. **A documentation block at the top of the source**, containing the function
    name, a description formatted for Name Manager's comment box, and a revision
    log.

08. **`LET()` steps have meaningful and concise names.**

09. **`LET()` step formulas are short and concise** where practical.

10. **The last `LET` step is a step *name*, not a formula** — see below.

11. **Stored in a GitHub Gist** for one-step import by the Advanced Formula
    Environment.

12. **New versions are backward compatible** with previous versions.

13. **Thoroughly tested** before release. See `testing.md`.

## Standard 06 — self-contained

The rule the skills previously never stated, and the one most likely to produce
a function that works on the machine it was built on and nowhere else.

> Self contained means that everything the function needs is contained within
> the function except the values brought in through the function's arguments.
> This means the function cannot contain any cell references, named references,
> table references, VBA references, etc.

**The one exception:**

> 5g Functions may reference other 5g functions that are in the same module
> (aka library).

So `fx.DebtSizeIssueλ` calling `fx.InterestDebtIssuanceλ` is fine. Either of
them reading `tblInputs` is not.

### The course's lab

Non-compliant — `tblItems` is baked into the calculation:

```excel
=LAMBDA(Columns,
    LET(Array, SEQUENCE(ROWS(tblItems), Columns),
         SCAN(0, Array, LAMBDA(Acc, Ele, Acc + Ele))
    )
)(2)
```

Compliant — the table moved out to an argument:

```excel
=LAMBDA(Table, Columns,
    LET(Array, SEQUENCE(ROWS(Table), Columns),
        SCAN(0, Array, LAMBDA(Acc, Ele, Acc + Ele))
    )
)(tblItems, 2)
```

### The helper-function case

The same trap appears one level down, inside `BYROW`/`BYCOL`/`MAP` callbacks. A
range reference buried in the helper is just as fatal as one in the body. The
fix is to wrap the helper in a LAMBDA that takes the range as an argument and
passes it through.

### How to check it

Grep the finished source for:
- `tbl` — table references
- `$` — absolute cell references
- any bare name that is not a LAMBDA parameter, a LET step, a native function,
  or a 5g function in the same module

## Standard 10 — the last LET step must be a name

> At the end of the procedure, LET allows us to have a formula as its last step;
> however, **to make debugging functions easier, 5g requires the last LET step to
> be the name of the last formula.** This allows us, when debugging, to change
> this last name to any of the other LET step names to see what they produce.

So this is non-compliant:

```excel
Return, IF(AND(Errors), 2, IF(OR(Errors), 3, 1)),
CHOOSE(Return, Result, Help, Messages)      // a formula
```

and this is compliant:

```excel
Return, IF(AND(Errors), 2, IF(OR(Errors), 3, 1)),
Output, CHOOSE(Return, Result, Help, Messages),
Output                                      // a name
```

The payoff is real and cheap: a maintainer swaps `Output` for `Errors` or
`Parts` and sees that step's value in the grid, without restructuring anything.
It costs one extra stored name — which matters only if the function is near the
8,192-character stored limit.

## Standard 12 — backward compatibility

> Updates must maintain backward compatibility to prevent errors that would be
> caused by incompatibility in existing implementations.

For this library that means, concretely:

- **Never reorder or remove arguments.** Add new ones at the end, optional, with
  a default that reproduces the old behaviour.
- **Never change the `VSTACK` row order of the output.** It is the output
  contract; every sheet consuming the function depends on it positionally.
- If the new behaviour genuinely cannot be reached without breaking one of
  those, that is a **new function with a new name**, not a new version.

This is the standard that argues against renaming `fx.*` functions to match some
other convention: a rename breaks every existing call.

## Where our templates diverge from the standard

Recorded rather than silently reconciled — see the same policy in
`hatmaker-library.md`.

| Standard | Our full template | Reading |
|---|---|---|
| 03 — help contains function name and syntax | Help starts at `DESCRIPTION:`; there is no `FUNCTION:` line | A genuine gap. Add the `FUNCTION:` line; it is what tells the user which arguments are required. |
| 03 — optional webpage | Absent | Add `WEBPAGE:` when a page exists, omit the line otherwise. |
| 10 — last LET step is a name | Ends on `CHOOSE(...)` | A genuine gap, cheap to fix unless stored length is tight. |
| 02 — clear argument names | `DoNotUse`, `Lo`, `Hi`, `Iter` appear in the signature and so in the tooltip | Recursion state leaking into the public signature. The `Iterate(Iterate, ...)` pattern in `compose-5g-lambda/reference/source-structures.md` avoids it entirely. Weigh against the cost of changing a tested function. |
| 13 — thoroughly tested | Single-benchmark tie-out | See `testing.md` — the standard expects randomised differential testing over hundreds of scenarios. |

## Guiding principles, for judgement calls

When a decision is not covered above, these are what the standards were derived
from. Value comes from **frequent reuse**, which requires functions that apply
to many similar scenarios, need no special skill to use, have clear names, feel
like Excel, are self-contained, and import easily. Functions must be **error
free**, which requires version information, easy updating, backward
compatibility, and thorough testing. And they must be **maintainable by
somebody else**, which requires structure and documentation.

## Sources

- Standards — <https://sites.google.com/site/beyondexcel/home/5g-modeling/5g-course/5g-standards>
- Self-containment — <https://sites.google.com/site/beyondexcel/home/5g-modeling/5g-course/5g-self-contained>
- Documentation — <https://sites.google.com/site/beyondexcel/home/5g-modeling/5g-course/5g-documentation>
- Inline help — <https://sites.google.com/site/beyondexcel/home/5g-modeling/5g-course/5g-inline-help>
- Course index — <https://sites.google.com/site/beyondexcel/home/5g-modeling/5g-course>

Scraped 05 Sep 2026 with the `webscrape` skill; the site renders via JavaScript
and needs a real browser.
