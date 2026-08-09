# The lightweight 5G template (inline help)

Use for quick one-off conversions and simple scalar helpers. For anything reusable,
use `5g-template-full.md` instead.

The difference: this template gates on `ISOMITTED` and returns *either* help *or* the
answer. It has no argument validation and no per-argument error messages.

```excel
/* FUNCTION NAME: Functionλ
   DESCRIPTION:*//**Description of what the function does*/
/* REVISIONS: Date         Developer    Description
              Mmm dd yyyy  UserName     Copyright/Initial Release
*/
Functionλ = LAMBDA(
    // Function Arguments — all bracketed, so Excel lets the function run empty
    [Required1],  // Description
    [Required2],  // Description
    [Optional1],  // Description (default: value)

    // Check required arguments
    IF( OR( ISOMITTED(Required1),
            ISOMITTED(Required2)
          ),
        // Display inline help text
        TRIM(TEXTSPLIT(
            "FUNCTION:    →Functionλ( Required1, Required2, [Optional1])¶" &
            "DESCRIPTION: →Description of what the function does¶" &
            "WEBPAGE:     →https://5gmodelling.com¶" &
            "VERSION:     →UserName, Mmm-dd-yyyy¶" &
            "PARAMETERS:  →¶" &
            "  Required1  →(Required) Description¶" &
            "  Required2  →(Required) Description¶" &
            "  Optional1  →(Optional) Default: value. Description",
            "→", "¶")),
        // Main function logic
        LET(
            // Set Defaults
            Optional1, IF(ISOMITTED(Optional1), DefaultValue, Optional1),
            // Calculations
            Result, formula,
            Result
        )
    )
);
```

## Rules specific to this template

- The **FUNCTION line** in the help text is where required/optional is declared:
  required arguments appear **without** brackets, optional ones **with**. The LAMBDA
  signature above always brackets everything.
- The `IF(OR(ISOMITTED(...)))` guard lists only the *truly required* arguments.
- Name the last calculation in the `LET`, then use that name as the return step.
- Placeholders in angle brackets (`<description>`) are instructions — strip the angle
  brackets from generated output.

## Choosing between the templates

| | Inline help | Full 5G |
|---|---|---|
| Bad argument type | `#VALUE!` with no explanation | Named message telling the user which argument and why |
| Called empty | Help table | Help table |
| Called with some args | Help table (unhelpful) | Only the failing arguments' messages |
| Lines of boilerplate | ~12 | ~30 |
| Right for | Scratch helpers, scalar math | Anything shared, anything in a library |
