---
name: housekeep
description: End-of-session wrap-up. The user is about to close VS Code and end the Claude Code session, and wants all internal housekeeping done (memory, project notes, skill docs, handoff notes, background tasks) followed by a clear verdict on whether it is safe to close without losing work. Invoke when the user types /housekeep, or says they are wrapping up, closing the session, shutting down for the day, or asks "is it safe to close?".
---

# /housekeep — wrap up the session safely

The user is closing soon. Your job has two parts:

1. **Housekeeping**: get everything worth keeping out of this conversation and
   into durable places (memory, project notes, skill docs).
2. **A verdict**: tell them plainly whether they can close VS Code now, or
   what is still blocking.

Treat this as wrap-up only. **Do not start new feature work**, and do not
"quickly finish" an open task unless it takes seconds and the user asked for
it this session. Anything unfinished goes into a handoff note (step 6).

Do not export the conversation. `/copyconvo` does that, and the user runs it
on demand.

---

## Step 1 — Inventory the session

Go back through the whole conversation (including any summary of earlier
context) and list the following for yourself:

- **Files** you created, edited, moved or deleted, with full paths
- **Decisions** made, and **corrections or preferences** the user gave you
- **Lessons**: bugs found, approaches that failed, non-obvious gotchas
- **Open items**: tasks started but unfinished, promised follow-ups,
  unanswered questions
- **Running work**: background shells, subagents, Monitors, `/loop` or cron
  jobs

This inventory drives every step below. If the session did almost nothing,
say so and keep the remaining steps short.

## Step 2 — Settle running work

Closing VS Code kills anything still running in this session.

- Background shells, subagents and Monitors: if one is close to done, wait for
  it. Otherwise stop it (`TaskStop`) and record what it was doing in the
  handoff note.
- Use `CronList` to check for session-scoped cron or `/loop` jobs. Tell the
  user that these stop when the session closes, so the work isn't silently
  lost.
- Cloud routines (`/schedule`) keep running after close. List them for
  information only; they do not block closing.

## Step 3 — Verify the work is on disk and works

- `Edit`/`Write` save to disk immediately, so your edits are already saved.
  Check that no edit was left half-applied. For example, a rename should be
  complete everywhere, and a multi-file change should not stop partway.
- If you changed any **script**, re-run it once and confirm it still runs.
  For finance outputs, confirm the reconciliation or tie-out checks pass.
  If it fails, it is a blocker. Report it, and don't patch it in a hurry
  at close time.
- **Unsaved editor tabs are the one thing you cannot see.** The user may have
  unsaved changes in VS Code tabs (a dot on the tab). Always include a reminder
  to check, or to run *File → Save All*.

## Step 4 — Update internal documentation

Work through these in order. Before you write anything, check whether an
existing note already covers it. If one does, update that note rather than
adding a duplicate.

| Target | Rule |
|---|---|
| **Memory** (the session's memory directory + `MEMORY.md` index) | Save decisions, user feedback/preferences (with the *why*), project status, and lessons that a future session would need. Follow the memory conventions in the system prompt exactly: one fact per file, frontmatter, `[[links]]`, one-line pointer in `MEMORY.md`, relative dates converted to absolute. Update stale memories, and delete ones that proved wrong. Don't save what the code, files or CLAUDE.md already record. |
| **Project notes** (e.g. `workflow-notes.md`, project `CLAUDE.md`, a README) | Update **only if** this session changed something they describe. Match their existing style. |
| **Global `~/.claude/CLAUDE.md`** | **Never edit it here.** If the session suggests a change, propose the exact wording in the report and let the user decide. |
| **Skills edited this session** | Update that skill's own docs (SKILL.md wording, references, changelog/version notes if it has them) so they match what the skill now does. Confirm a pre-edit backup exists at `~/.claude/skills/.backups/<skill>-YYYY-MM-DD/`. If none was taken, say so in the report. Don't create one now from the already-edited files, because that would back up the new version, not the original. |

**Never touch `~/.claude/skills/synced/`.** Claude's skill sync manages it,
and any edits there get overwritten.

## Step 5 — Account for loose files

- **Scratchpad**: the session scratchpad directory is deleted with the
  session. List anything in it that the user might want, such as generated
  outputs or useful scripts, and offer to move it to the project. Ignore
  throwaway intermediates.
- **Git**: if the working directory (or any directory you edited) is a git
  repo, run `git status` and **report** uncommitted changes. **Never commit
  automatically.** Uncommitted changes are safe on disk, so they don't block
  closing. Just list them.

## Step 6 — Write a handoff note if anything is unfinished

If there are open items from Step 1, write or update a single `project`-type
memory so the next session can pick up cold:

- Name it `handoff-<short-topic>` and add it to `MEMORY.md`
- Include: the goal; what's done; the exact next step; relevant file paths;
  any decision still waiting on the user; today's date (absolute)
- Write it for a reader who has none of this conversation's context

Also look for **existing** `handoff-*` memories that this session finished.
Delete them, or fold the outcome into a regular memory, so stale handoffs
don't pile up.

## Step 7 — Report and give the verdict

Keep the report short, using this structure:

```
## Housekeeping done
- Memory: <created/updated files, one line each — or "nothing new">
- Project notes: <files updated — or "none needed">
- Skill docs: <skill + what changed, backup status — or "n/a">
- Background tasks: <stopped/finished/none>

## Carry-forward
- <open item> → handoff note: <memory file>      (or "Nothing open")
- <uncommitted git changes / scratch files offered / proposed CLAUDE.md edit>

## Before you close
- Check VS Code for unsaved tabs (File → Save All).

✅ **Safe to close.** Everything from this session is saved; the
conversation itself is kept by Claude Code and can be resumed later.
```

…or, if something blocks:

```
⚠️ **Not yet.** <one-line reason> — <what you need from the user, or what
you're waiting on>
```

**What blocks closing** (gives ⚠️): a task still running that you couldn't
stop; a script or tie-out that now fails; a half-applied edit; a question
whose answer must be recorded before close; scratch files the user hasn't
decided on yet.

**What doesn't block**: uncommitted git changes, open items that already have
a handoff note, cloud routines, proposed CLAUDE.md edits.

The ✅/⚠️ line must be the **last line** of the report, and it must be
honest. If a step was skipped or failed, say so above it. Never give ✅ while
a blocker is outstanding.
