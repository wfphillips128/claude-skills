# Global preferences

<!--
  Starter CLAUDE.md for the Excel/finance setup in this repo.
  Copy what fits into %USERPROFILE%\.claude\CLAUDE.md (create it if missing)
  and replace every <placeholder>. Delete sections that don't apply.
-->

User-level instructions for Claude Code. These apply across all projects. A
project's own `CLAUDE.md` overrides anything here when they conflict.

## About my work
I mostly use Claude Code for:
- **Excel modelling** — LAMBDA functions (5G standard), VBA macros and add-ins.
- **Finance / accounting reporting** — turning accounting data into
  statements, models and Excel/CSV outputs.
- **Data analysis / scripting** — ad-hoc wrangling and one-off automation.

Assume a finance/analysis mindset: get the numbers right, make outputs tie out,
and prefer clear, auditable results over clever code.

<!-- If you are not a professional programmer, say so — it changes how Claude explains: -->
Explain the *why* behind changes and teach the underlying pattern so I learn
it — don't just hand over code. Define jargon briefly.

## Response style
- **Concise and direct.** Lead with the answer.
- Explain reasoning when it matters (trade-offs, risks, non-obvious choices).
- Use tables and short lists for options or comparisons.

## Autonomy
- Act freely on read-only work and routine edits — just do it and report back.
- **Confirm before destructive or outward-facing actions**: deleting or
  overwriting files I didn't create, force operations, or anything that sends
  data to an external service.

## Tooling defaults
- **Python** for scripts, data work and Excel (`openpyxl`, `pandas`,
  `pywin32`). Write portable, well-commented scripts.
- **PowerShell** for Windows automation.
- Don't introduce other stacks without a reason.

## Environment notes
- <Windows 11>, shell is <Windows PowerShell 5.1> — use PowerShell syntax.
- Python is at: `<C:\full\path\to\python.exe>`. Bare `python` may resolve to
  the Windows Store stub — use the full path when a command must run reliably.
- Excel: <Microsoft 365, version>. Excel Labs (Advanced Formula Environment) is installed.

## Working conventions
- **Validate by actually running it.** For Excel output, open/recalculate and
  confirm totals reconcile before reporting done. For a LAMBDA, test it in
  Excel against an independent calculation. For VBA, compile and smoke-test
  it in a separate Excel instance before handing it over.
- **Diagnose Excel by experiment.** When something fails in Excel and the
  cause isn't obvious, get an observation out of Excel (a small test macro or
  formula) rather than reasoning about the text.
- For financial outputs, always include a reconciliation or balance check and
  say if anything doesn't tie.
- Match the style of surrounding code when editing existing files.
- When a task is finished and verified, say so plainly; if something failed or
  was skipped, say that too.

## Secrets
- API keys live in user environment variables, set through the Windows GUI
  dialog — never pasted into chat, never `setx` (it lands in shell history).
- Verify a key by its length or prefix, never by printing it.
