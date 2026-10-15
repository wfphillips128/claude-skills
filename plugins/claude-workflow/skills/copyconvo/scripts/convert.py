#!/usr/bin/env python3
"""
copyconvo - Convert a Claude Code conversation .jsonl transcript into a
readable, self-contained HTML file.

A Claude Code transcript is a JSON-lines file: one JSON object per line. Most
lines are bookkeeping (queue-operation, attachment, file-history-snapshot,
ai-title, last-prompt, summary). The actual conversation is in lines with
type "user" or "assistant", whose `message.content` is either a plain string
(older format) or a list of typed blocks (text / tool_use / tool_result /
thinking / image).

This script keeps only the conversation, renders it as chat bubbles, and tucks
tool calls and results into collapsible <details> blocks so a human reader sees
the dialogue first. The output is a single .html file with all CSS inlined, so
it opens in any browser on any machine with no dependencies.

Usage:
    python convert.py INPUT.jsonl [-o OUTPUT.html] [--no-tools] [--title "..."]

    --no-tools   Omit tool calls/results entirely (dialogue only).
    --title      Override the page title (defaults to the transcript's
                 ai-title, then the filename).
"""

import argparse
import html
import json
import sys
from pathlib import Path


def _text_from_content(content):
    """A tool_result's content may itself be a string or a list of blocks."""
    if isinstance(content, str):
        return content
    if isinstance(content, list):
        parts = []
        for block in content:
            if isinstance(block, dict):
                if block.get("type") == "text":
                    parts.append(block.get("text", ""))
                elif "text" in block:
                    parts.append(block.get("text", ""))
            elif isinstance(block, str):
                parts.append(block)
        return "\n".join(parts)
    return str(content)


def parse_transcript(path):
    """Read the .jsonl and return (title, [turns]).

    Each turn is a dict: {"role", "blocks"} where blocks is a list of
    ("text"|"tool_use"|"tool_result"|"thinking", payload) tuples.
    """
    title = None
    turns = []

    with open(path, "r", encoding="utf-8") as fh:
        for line in fh:
            line = line.strip()
            if not line:
                continue
            try:
                rec = json.loads(line)
            except json.JSONDecodeError:
                continue  # skip anything that isn't valid JSON

            rtype = rec.get("type")

            # Grab the human-friendly title if the transcript carries one.
            if rtype == "ai-title" and rec.get("aiTitle") and not title:
                title = rec["aiTitle"]
                continue

            if rtype not in ("user", "assistant"):
                continue

            message = rec.get("message") or {}
            role = message.get("role", rtype)
            content = message.get("content")

            blocks = []
            if isinstance(content, str):
                if content.strip():
                    blocks.append(("text", content))
            elif isinstance(content, list):
                for block in content:
                    if not isinstance(block, dict):
                        continue
                    btype = block.get("type")
                    if btype == "text":
                        if block.get("text", "").strip():
                            blocks.append(("text", block["text"]))
                    elif btype == "thinking":
                        if block.get("thinking", "").strip():
                            blocks.append(("thinking", block["thinking"]))
                    elif btype == "tool_use":
                        payload = {
                            "name": block.get("name", "tool"),
                            "input": block.get("input", {}),
                        }
                        blocks.append(("tool_use", payload))
                    elif btype == "tool_result":
                        blocks.append(
                            ("tool_result", _text_from_content(block.get("content", "")))
                        )
                    elif btype == "image":
                        blocks.append(("text", "[image]"))

            if blocks:
                turns.append({"role": role, "blocks": blocks})

    return title, turns


def _esc(text):
    return html.escape(str(text))


def render_html(title, turns, include_tools=True):
    out = []
    out.append("<!DOCTYPE html>")
    out.append('<html lang="en"><head><meta charset="utf-8">')
    out.append('<meta name="viewport" content="width=device-width, initial-scale=1">')
    out.append(f"<title>{_esc(title)}</title>")
    out.append("<style>" + CSS + "</style>")
    out.append("</head><body>")
    out.append('<div class="wrap">')
    out.append(f"<h1>{_esc(title)}</h1>")
    out.append(
        f'<p class="meta">{len(turns)} messages · '
        f'rendered by copyconvo</p>'
    )

    for turn in turns:
        role = turn["role"]
        label = "You" if role == "user" else "Claude"
        out.append(f'<div class="turn {role}">')
        out.append(f'<div class="role">{_esc(label)}</div>')
        out.append('<div class="bubble">')

        for btype, payload in turn["blocks"]:
            if btype == "text":
                out.append(f'<div class="text">{_esc(payload)}</div>')
            elif btype == "thinking":
                if include_tools:
                    out.append(
                        '<details class="thinking"><summary>thinking</summary>'
                        f'<pre>{_esc(payload)}</pre></details>'
                    )
            elif btype == "tool_use":
                if include_tools:
                    try:
                        args = json.dumps(payload["input"], indent=2, ensure_ascii=False)
                    except (TypeError, ValueError):
                        args = str(payload["input"])
                    out.append(
                        f'<details class="tool"><summary>'
                        f'⚙️ tool call: {_esc(payload["name"])}</summary>'
                        f'<pre>{_esc(args)}</pre></details>'
                    )
            elif btype == "tool_result":
                if include_tools:
                    text = payload
                    if len(text) > 4000:
                        text = text[:4000] + "\n… [truncated]"
                    out.append(
                        '<details class="tool result"><summary>'
                        '↩️ tool result</summary>'
                        f'<pre>{_esc(text)}</pre></details>'
                    )

        out.append("</div></div>")

    out.append("</div></body></html>")
    return "\n".join(out)


CSS = """
* { box-sizing: border-box; }
body { margin: 0; background: #f4f4f5; color: #1a1a1a;
  font-family: -apple-system, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
  line-height: 1.55; }
.wrap { max-width: 820px; margin: 0 auto; padding: 32px 20px 80px; }
h1 { font-size: 22px; margin: 0 0 4px; }
.meta { color: #71717a; font-size: 13px; margin: 0 0 28px; }
.turn { margin: 0 0 20px; }
.role { font-size: 12px; font-weight: 600; text-transform: uppercase;
  letter-spacing: .04em; color: #71717a; margin: 0 0 6px; }
.bubble { padding: 14px 18px; border-radius: 12px; border: 1px solid #e4e4e7; }
.turn.user .bubble { background: #eef2ff; border-color: #c7d2fe; }
.turn.assistant .bubble { background: #ffffff; }
.text { white-space: pre-wrap; word-wrap: break-word; }
.text + .text { margin-top: 10px; }
details { margin-top: 10px; border: 1px solid #e4e4e7; border-radius: 8px;
  background: #fafafa; padding: 6px 10px; }
details.thinking { background: #fdf6ec; border-color: #f0e0c0; }
details.result summary { color: #166534; }
summary { cursor: pointer; font-size: 13px; color: #4338ca; font-weight: 500; }
pre { white-space: pre-wrap; word-wrap: break-word; margin: 8px 0 4px;
  font-family: "Cascadia Code", Consolas, "Courier New", monospace;
  font-size: 12.5px; color: #27272a; overflow-x: auto; }
"""


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__,
                                     formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("input", help="Path to the .jsonl transcript")
    parser.add_argument("-o", "--output", help="Output .html path "
                        "(defaults to the input name with .html)")
    parser.add_argument("--no-tools", action="store_true",
                        help="Omit tool calls/results and thinking")
    parser.add_argument("--title", help="Override the page title")
    args = parser.parse_args(argv)

    in_path = Path(args.input)
    if not in_path.is_file():
        print(f"error: file not found: {in_path}", file=sys.stderr)
        return 1

    title, turns = parse_transcript(in_path)
    if not turns:
        print("error: no user/assistant messages found in transcript",
              file=sys.stderr)
        return 1

    page_title = args.title or title or in_path.stem
    out_path = Path(args.output) if args.output else in_path.with_suffix(".html")

    page = render_html(page_title, turns, include_tools=not args.no_tools)
    out_path.write_text(page, encoding="utf-8")

    print(f"Wrote {out_path}  ({len(turns)} messages)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
