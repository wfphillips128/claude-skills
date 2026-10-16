# Roll Your Own LAMBDA

A worked example of the `excel-dev` skills: describe the Excel LAMBDA you want in
plain English, and Claude Code builds it, tests it in Excel, and checks it
against a known answer.

| File | What it is |
|---|---|
| [Roll Your Own LAMBDA.pdf](Roll%20Your%20Own%20LAMBDA.pdf) | The article (9 pages): how to write the description, the prompt to paste into Claude Code, what Claude does with it, the finished function, and two follow-on exercises. The function's source is attached inside the PDF. |
| [Roll Your Own LAMBDA Example.xlsx](Roll%20Your%20Own%20LAMBDA%20Example.xlsx) | The starting workbook. 10,000 simulated trials in column B, an existing statistics function in F3, and an empty I3 for the one you build. |
| [Roll Your Own LAMBDA Example - Completed.xlsx](Roll%20Your%20Own%20LAMBDA%20Example%20-%20Completed.xlsx) | The answer: `Example1` installed and working at I3, with a tie-out cell at L3 confirming it matches F3:G35. |

**To follow along:** download the starting workbook into a folder of its own,
open a terminal there, run `claude`, and paste the prompt from section 4 of the
PDF. You need the `excel-dev` plugin installed; see [SETUP.md](../SETUP.md).

The workbooks use the author's Monte Carlo LAMBDA library (the `fx.Risk…`
functions) to generate the trials. You don't need it for anything else; it is
also published as a [gist](https://gist.github.com/wfphillips128/f91bff77212ab2c3d8f55a4f0a51b8b6).
