---
name: ss
description: Read the user's most recent screenshot(s) and act on them. Invoke when the user types /ss, optionally followed by a count and an instruction (e.g. "/ss explain this", "/ss 3 summarize"). Grabs the N newest screenshots from the configured screenshot folder and performs the requested action.
---

# /ss — Read recent screenshots

## Finding the screenshot folder

Use the first of these that exists:

1. `$env:SCREENSHOT_DIR` — set this if your screenshots live anywhere else.
2. `$env:USERPROFILE\Documents\ShareX\Screenshots` — ShareX's default.
3. `$env:USERPROFILE\Pictures\Screenshots` — Windows' own default (Win+PrtScn).

ShareX files its captures into dated subfolders, so the search must recurse.

## Parsing the arguments

The text after `/ss` has two optional parts:

1. **A leading number** = how many of the most recent screenshots to grab. If absent, default to `1`.
2. **The rest of the text** = the action to perform on the screenshot(s). If absent, just describe/explain the screenshot(s).

Examples:
- `/ss explain this` → grab 1 newest, explain it
- `/ss 3 summarize` → grab 3 newest, summarize them together
- `/ss 4` → grab 4 newest, describe them
- `/ss` → grab 1 newest, describe it

Only treat a leading token as the count if it is a bare integer. `/ss explain` means count=1, action="explain".

## Steps

1. Find the N most recent screenshots, newest first. Run this PowerShell (substitute `<N>`):

```powershell
$dirs = @(
  $env:SCREENSHOT_DIR,
  "$env:USERPROFILE\Documents\ShareX\Screenshots",
  "$env:USERPROFILE\Pictures\Screenshots"
) | Where-Object { $_ -and (Test-Path $_) }

if (-not $dirs) { "No screenshot folder found. Set SCREENSHOT_DIR."; exit }

Get-ChildItem -Path $dirs[0] -File -Recurse -Include *.png,*.jpg,*.jpeg |
  Sort-Object LastWriteTime -Descending |
  Select-Object -First <N> |
  Select-Object -ExpandProperty FullName
```

   If the folder is empty or missing, tell the user no screenshots were found and stop.

2. Read each returned path with the Read tool so you can see the images.

3. Perform the requested action on the screenshot(s). When multiple are grabbed, consider them together (e.g. summarize/compare across all of them) unless the action implies otherwise.

4. Briefly note which file(s) you used (filename + timestamp) so the user knows which screenshots were picked up.
