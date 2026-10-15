# Claude Code skills

[Claude Code](https://claude.com/claude-code) skills for **building** Excel
LAMBDA functions to the 5G standard, writing and refactoring VBA, and building
`.xlam` add-ins -- plus a few everyday workflow helpers. They are used daily on
real workbooks rather than written as demos.

This repo gives Claude Code the *capability*. It does not install any finished
LAMBDA library -- see [Finished LAMBDA libraries](#finished-lambda-libraries).

## Install

This repo is a Claude Code **plugin marketplace**. In a terminal:

```
claude plugin marketplace add wfphillips128/claude-skills
claude plugin install excel-dev@wfphillips128-skills
claude plugin install claude-workflow@wfphillips128-skills      # optional
```

Then restart Claude Code (or run `/reload-plugins`). Skills trigger on their
own when relevant; you can also call one directly, e.g. `/excel-dev:convert-to-5g`.

**Updating:** `claude plugin marketplace update` -- see [CHANGELOG.md](CHANGELOG.md)
for what changed.

**Without plugins:** clone the repo and run `install.ps1` (Windows PowerShell),
which copies the skills into `~/.claude/skills` and backs up any skill of the
same name first. Re-run it after `git pull` to update.

**Full setup** -- prerequisites, Python packages, Anthropic's document skills,
MCP servers, a starter `CLAUDE.md`, and prompts to try: [docs/SETUP.md](docs/SETUP.md).

## Plugins

### `excel-dev`

Build new Excel LAMBDAs to the 5G standard, refactor and write VBA, and build .xlam add-ins (ribbon XML, .bas modules, verification).

| Skill | What it does |
|---|---|
| [`compose-5g-lambda`](plugins/excel-dev/skills/compose-5g-lambda/) | Build a NEW 5G Excel LAMBDA by composing or extending LAMBDAs that already exist in a workbook — the case where a working function needs a variant and naive edits break the recursion. |
| [`convert-to-5g`](plugins/excel-dev/skills/convert-to-5g/) | Convert an existing Excel formula (or block of grid formulas) into a 5th-Generation (5G) dynamic-array LAMBDA function following Craig Hatmaker's Beyond Excel standards — self-documenting, spilling, w |
| [`excel-addin-builder`](plugins/excel-dev/skills/excel-addin-builder/) | Build and maintain an Excel VBA add-in (.xlam) - ribbon XML (customUI14), .bas modules, UserForms, the add-in's private settings storage, packaging, and verification before handover. |
| [`excel-macro-optimizer`](plugins/excel-dev/skills/excel-macro-optimizer/) | Optimize, refactor, or write Excel VBA macros for speed, reliability, and maintainability — especially macros that live in an .XLAM add-in and run against the active workbook. |
| [`/opti`](plugins/excel-dev/commands/opti.md) | Slash command -- see the file. |

### `claude-workflow`

Everyday Claude Code helpers: export a conversation to HTML, end-of-session housekeeping, read the latest screenshot, scrape web pages with Firecrawl.

| Skill | What it does |
|---|---|
| [`copyconvo`](plugins/claude-workflow/skills/copyconvo/) | >- Convert a Claude Code conversation transcript (.jsonl) into a clean, self-contained HTML file that anyone can read in a browser on any machine — no Claude Code or setup required. |
| [`housekeep`](plugins/claude-workflow/skills/housekeep/) | End-of-session wrap-up. |
| [`ss`](plugins/claude-workflow/skills/ss/) | Read the user's most recent screenshot(s) and act on them. |
| [`webscrape`](plugins/claude-workflow/skills/webscrape/) | | Pull content and design off live websites using Firecrawl — scrape a single page to clean markdown, crawl or map a whole site, search the web, or capture a page's full design system (screenshot, HTM |

## Prerequisites

| What | Needed by |
|---|---|
| Microsoft Excel for Microsoft 365 on Windows (dynamic arrays and LAMBDA; win32com automation needs Windows). | `compose-5g-lambda`, `convert-to-5g`, `excel-addin-builder`, `excel-macro-optimizer` |
| Excel Labs add-in (its Advanced Formula Environment is where LAMBDAs are edited and gists imported). | `compose-5g-lambda`, `convert-to-5g` |
| Optional: a Firecrawl account (free tier is fine) with the key in the FIRECRAWL_API_KEY environment variable, for the webscrape skill. | `webscrape` |
| Python 3.11+ with the packages in requirements.txt. On Windows, a bare `python` can resolve to the Microsoft Store stub; use the full path to a real interpreter if so. | `compose-5g-lambda`, `copyconvo`, `excel-addin-builder`, `webscrape` |
| Office RibbonX Editor, to paste ribbon XML into an .xlsm/.xlam. | `excel-addin-builder` |
| Optional: ShareX (or any tool that saves screenshots to a folder) for the ss skill. | `ss` |

Python packages: `pip install -r requirements.txt`.

## Related

| Skill | Where |
|---|---|
| [`ibcs-excel-charts`](https://github.com/wfphillips128/Notation-Charts-for-Excel-via-Claude-Skill) | Lives in its own repository. |
| [`panel-charts`](https://github.com/wfphillips128/Panel-Charts-for-Excel-via-Claude-Skill) | Lives in its own repository. |
| `skill-creator`, `xlsx`, `docx`, `pdf`, `pptx` | Anthropic's own -- install from [anthropics/skills](https://github.com/anthropics/skills) (see [SETUP.md](docs/SETUP.md)). |

## Finished LAMBDA libraries

The skills here teach Claude to build and test LAMBDAs. For ready-made
libraries to use or to learn from, see the author's gists:

- [Monte Carlo LAMBDAs](https://gist.github.com/wfphillips128/f91bff77212ab2c3d8f55a4f0a51b8b6)
- [Time-series statistics LAMBDAs](https://gist.github.com/wfphillips128/96171c91caba82aa708b12b9f8f575ce)

Import either through Excel Labs' Advanced Formula Environment. See
[docs/GETTING-STARTED-LAMBDA.md](docs/GETTING-STARTED-LAMBDA.md).

## The 5G methodology

`convert-to-5g` and `compose-5g-lambda` follow Craig Hatmaker's
[Beyond Excel](https://sites.google.com/site/beyondexcel/) 5th-Generation
methodology -- self-documenting, spilling LAMBDA functions with inline help.
The methodology is his; these skills are just an implementation of it.

## Licence

MIT -- see [LICENSE](LICENSE).
