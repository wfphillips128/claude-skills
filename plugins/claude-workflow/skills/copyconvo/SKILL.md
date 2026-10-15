---
name: copyconvo
description: >-
  Convert a Claude Code conversation transcript (.jsonl) into a clean,
  self-contained HTML file that anyone can read in a browser on any machine —
  no Claude Code or setup required. Use this whenever the user types
  /copyconvo, or asks to export, share, save, or "make readable" a conversation
  or chat transcript, wants to send a conversation to a coworker, or mentions
  turning a .jsonl session log into something human-friendly. Triggers even if
  they don't say "HTML" — sharing or exporting a Claude Code conversation is
  the cue.
---

# copyconvo — export a conversation to readable HTML

Claude Code stores every session as a `.jsonl` transcript (one JSON object per
line) under `~/.claude/projects/<encoded-cwd>/`. Those files are machine logs —
full of metadata, tool calls, and token counts — so they're unpleasant to read
raw and only render nicely back inside Claude Code. This skill turns one into a
single `.html` file that shows just the dialogue (tool calls tucked into
expandable sections), opens in any browser, and needs no software to view. That
makes it the right format for sharing a conversation with a coworker.

## The script does the work

[scripts/convert.py](scripts/convert.py) handles all the parsing and rendering.
Run it with the user's Python; don't reimplement it inline.

```powershell
python `
  "scripts/convert.py" `
  "<input.jsonl>" -o "<output.html>"
```

Flags:
- `--no-tools` — omit tool calls, results, and thinking; keep only the
  user/assistant dialogue. Good for the cleanest read.
- `--title "..."` — override the page title (otherwise it uses the
  transcript's own title, then the filename).

## Parsing the arguments (all optional)

Everything after `/copyconvo` is optional. Read whatever the user provides and
fall back to asking when it's ambiguous — the default with no arguments is to
list recent conversations and let the user pick. The text can contain, in any
combination:

- **A path** to a `.jsonl` file → use it directly, skip the picking step.
- **A leading integer N** (a bare number) → target the N most recent *completed*
  conversations instead of asking. Skip the current live session (its file is
  still being written, so it would export incomplete). If N is 1 and the only
  candidate is the live session, ask rather than export it half-written.
- **Free text** (e.g. "gl mapping", "the board deck one") → treat as a search
  term: pre-filter recent transcripts to those whose opening prompt, title, or
  cwd matches, then pick/confirm from that shorter list.
- **"no tools" / "dialogue only" / "clean"** → pass `--no-tools`.

A leading token is only a count if it's a bare integer; `/copyconvo gl` means
search text "gl", not a count.

## Workflow

1. **Find the transcript.** Apply the arguments above. If the user gave a path,
   use it. If they gave a count or search text, resolve it to candidate files. If
   they gave nothing, list the most recent `.jsonl` files and confirm which one
   they mean. Transcripts live in `~/.claude/projects/`, one subfolder per
   working directory (the folder name is the cwd with separators turned into
   dashes). Each file's first `user` line shows the opening prompt and its `cwd`,
   which helps identify it. Remember the newest file is usually *this* live
   session — don't silently export it; prefer completed conversations unless the
   user explicitly wants the current one.

2. **Pick the output location.** Default to the same folder as the input with a
   `.html` extension, unless the user wants it somewhere specific (e.g. a shared
   drive or their Desktop to hand off).

3. **Decide on tool calls.** Default is to include them (collapsed). If the user
   just wants the conversation to read, or it's going to a non-technical
   coworker, prefer `--no-tools`. When unsure, ask briefly.

4. **Run the script** and report the output path. When more than one transcript
   was selected (a count > 1), run the script once per file — it takes a single
   input — and list every `.html` produced. Mention they can double-click the
   `.html` to open it, and that it's fully self-contained (safe to email or
   copy to another machine).

## Notes

- The script skips non-conversation lines (queue ops, attachments, snapshots)
  automatically and handles both old-style string content and newer block
  content, so it works across Claude Code versions.
- Long tool results are truncated in the HTML to keep the file readable; the
  dialogue text is never truncated.
- The output is one file with CSS inlined — there are no external assets to
  ship alongside it.
