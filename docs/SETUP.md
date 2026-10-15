# Setup

Everything you need to give Claude Code the same Excel capabilities as the
setup this repo comes from: building new LAMBDAs to the 5G standard, writing
and refactoring VBA, and building `.xlam` add-ins. Work through it in order;
each step says what it is for, so you can skip what you don't need.

Assumes **Windows** and **Excel for Microsoft 365**. The LAMBDA skills work on
a Mac too; the scripts that drive Excel (win32com) and the `ss` skill are
Windows-only.

## 1. Prerequisites

| What | Why | Get it |
|---|---|---|
| Claude Code | runs the skills | <https://claude.com/claude-code> |
| Excel for Microsoft 365 | dynamic arrays and LAMBDA | — |
| Excel Labs add-in | its **Advanced Formula Environment** (AFE) is where LAMBDAs are edited, and it imports gists | Excel > Home > Add-ins > search "Excel Labs" |
| Python 3.11+ | the skills' helper scripts | <https://www.python.org/downloads/> or `winget install --id Python.Python.3.12 -e` |
| Office RibbonX Editor | pastes ribbon XML into an add-in (`excel-addin-builder`) | <https://github.com/fernandreu/office-ribbonx-editor/releases> |
| *Optional:* ShareX | the `ss` skill reads your newest screenshot | Microsoft Store |
| *Optional:* Firecrawl account | the `webscrape` skill | free tier at <https://firecrawl.dev> |
| *Optional:* Node.js LTS | the MCP servers in step 5 | `winget install --id OpenJS.NodeJS.LTS -e` |

**Python on Windows:** a bare `python` can resolve to the Microsoft Store stub,
which does nothing and exits. If a skill's script "does nothing", call your
real interpreter by its full path. Tell Claude that path once in your
`CLAUDE.md` (step 6) and it will use it.

**Excel automation:** the verification scripts open their own hidden Excel and
import VBA into a scratch copy. That needs *File > Options > Trust Center >
Trust Center Settings > Macro Settings > Trust access to the VBA project object
model*. Only the scripted checks need it; everything else works without.

## 2. Install the skills

```
claude plugin marketplace add wfphillips128/claude-skills
claude plugin install excel-dev@wfphillips128-skills
claude plugin install claude-workflow@wfphillips128-skills    # optional
```

Restart Claude Code (or run `/reload-plugins`). Check with `/plugin` — both
plugins should be listed and enabled.

**Updates:** `claude plugin marketplace update`, then restart. Read
[CHANGELOG.md](../CHANGELOG.md) to see what changed.

**No plugins?** Clone the repo and run `.\install.ps1` from it. It copies the
skills into `%USERPROFILE%\.claude\skills` and backs up any skill of the same
name first (to `skills\.backups\`). After `git pull`, run it again.

## 3. Python packages

```
python -m pip install -r requirements.txt
```

(from the cloned repo, or copy the file's few lines). `pywin32` is what lets
the scripts drive Excel; `oletools` reads VBA out of a workbook without
opening Excel; `openpyxl` reads and writes `.xlsx`.

## 4. Anthropic's own skills

Spreadsheet, Word, PDF and PowerPoint handling — and `skill-creator`, for
writing your own skills — come from Anthropic. Install them from the source
rather than from a copy:

```
claude plugin marketplace add anthropics/skills
claude plugin install document-skills@anthropic-agent-skills
```

Some of those skills call outside programs. Install the ones you'll use:

| Program | Used for | Install |
|---|---|---|
| LibreOffice | recalculating formulas in generated `.xlsx`; Word/PowerPoint to PDF | `winget install --id TheDocumentFoundation.LibreOffice -e` |
| Poppler | PDF page images, slide thumbnails | `winget install --id oschwartz10612.Poppler -e` |
| Pandoc | Word document conversion | `winget install --id JohnMacFarlane.Pandoc -e` |
| Tesseract | OCR on scanned PDFs | `winget install --id UB-Mannheim.TesseractOCR -e` |

## 5. MCP servers (optional)

MCP servers give Claude extra tools. Two are worth having:

```
claude mcp add --scope user --transport http context7 https://mcp.context7.com/mcp
claude mcp add --scope user playwright -- npx -y @playwright/mcp@latest
```

- **context7** fetches current documentation for libraries (openpyxl, pandas,
  win32com…) so Claude isn't working from stale memory.
- **playwright** gives Claude a real browser, e.g. to check a web page.

Both need Node.js (step 1).

## 6. A starter `CLAUDE.md`

`CLAUDE.md` in `%USERPROFILE%\.claude\` is read at the start of every session.
It is where you tell Claude how you work. [examples/CLAUDE.md](../examples/CLAUDE.md)
is a cleaned-up version of the one this setup uses. Copy what fits into your
own and fill in the `<placeholders>` — especially your Python path.

The habits in it matter more than they look: *validate by actually running
it*, *make totals tie out*, *explain why*. They are why the LAMBDA work ends
in a function that has been tested, not one that merely looks right.

## 7. Excel side

1. Install **Excel Labs** (step 1) and open it: Home > Excel Labs > Advanced
   Formula Environment.
2. New LAMBDAs that Claude writes are pasted into the AFE's **Modules** (or
   into Name Manager for short ones). The 5G skills explain which and why.
3. For finished libraries to build on or learn from, import a gist into the
   AFE (Modules > Import from URL) — see
   [GETTING-STARTED-LAMBDA.md](GETTING-STARTED-LAMBDA.md).

## 8. Try it

Small inputs for these are in [samples/](samples/). Each shows what a working
setup does.

| Prompt | Skill | What you should get |
|---|---|---|
| "Convert this formula to 5G" and paste `samples/formula-to-convert.txt` | `convert-to-5g` | A single LAMBDA with an inline help block, an error-check block, and paste instructions for the AFE |
| "Write a new 5G LAMBDA that builds a loan amortization schedule from rate, periods and principal, then test it" | `compose-5g-lambda` | A spilling LAMBDA, plus a test run against a Python reference showing PASS |
| `/opti` with `samples/slow-macro.bas` | `excel-macro-optimizer` | A faster version that keeps the same output, with the reason for each change |
| "Add a ribbon group with two buttons to my add-in" | `excel-addin-builder` | customUI14 XML, a `.bas` module with matching callbacks, import steps, and a static check that passes |

If a skill doesn't trigger, name it: "use the convert-to-5g skill to…".
