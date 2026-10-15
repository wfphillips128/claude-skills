#!/usr/bin/env python3
"""
firecrawl.py — a small, credit-aware client for the Firecrawl v2 REST API.

Why this exists rather than the vendor CLI: Firecrawl is a plain REST service,
so a Python client avoids a Node toolchain entirely. More importantly, this
adds three things the generic tooling does not:

  1. A credit ledger. Every run reports credits before and after, straight from
     the account. On a 1000-credit/month plan an unbounded crawl is the one
     real way to lose money, so `crawl` refuses to exceed a page budget without
     an explicit --yes.
  2. A local cache. Responses are keyed by endpoint + request body, so
     re-running a command costs nothing. Iterating on a design capture or
     re-reading a harvest should never re-bill.
  3. A design-capture mode. Markdown throws away everything that makes a page
     worth emulating; `capture` keeps the screenshot, the raw HTML, the actual
     stylesheets, and a distilled token summary.

Usage:
    python firecrawl.py credits
    python firecrawl.py search "query" [-n 5]
    python firecrawl.py map    <url> [--limit 500]
    python firecrawl.py scrape <url> [-o out.md] [--formats markdown,html,screenshot]
    python firecrawl.py crawl  <url> -o <dir> [--limit 50] [--yes]
    python firecrawl.py capture <url> -o <dir>

Auth: set FIRECRAWL_API_KEY, or pass --api-key.
Requires: requests (already present in the user's Python 3.12).
"""

import argparse
import hashlib
import json
import os
import re
import sys
import time
from collections import Counter
from pathlib import Path
from urllib.parse import unquote, urljoin, urlparse

try:
    import requests
except ImportError:
    sys.exit("Missing dependency: run  python -m pip install requests")

API_BASE = "https://api.firecrawl.dev/v2"

# The cache lives under the user's home rather than the working directory, so a
# page fetched while working in one project is still free when it comes up in
# another. A cwd-relative cache silently re-bills every time you change folder.
DEFAULT_CACHE = str(Path.home() / ".firecrawl-cache")

# A crawl over this many pages needs --yes. Sized so a routine docs harvest
# runs unattended but a misaimed crawl at a big site stops and asks.
CRAWL_CONFIRM_THRESHOLD = 60

# Windows consoles default to cp1252 and will crash on page titles containing
# characters like the lambda in Hatmaker's function names.
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")


# --------------------------------------------------------------------------
# Client
# --------------------------------------------------------------------------

class Firecrawl:
    def __init__(self, api_key, cache_dir=DEFAULT_CACHE, use_cache=True, quiet=False):
        self.api_key = api_key
        self.cache_dir = Path(cache_dir)
        self.use_cache = use_cache
        self.quiet = quiet
        self.session = requests.Session()
        self.session.headers.update({
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json",
        })
        self.cache_hits = 0
        self.api_calls = 0

    def log(self, msg):
        if not self.quiet:
            print(msg, file=sys.stderr)

    # -- cache -------------------------------------------------------------

    def _cache_path(self, endpoint, body):
        blob = json.dumps({"e": endpoint, "b": body}, sort_keys=True)
        return self.cache_dir / (hashlib.sha256(blob.encode()).hexdigest()[:24] + ".json")

    def _cache_read(self, endpoint, body):
        if not self.use_cache:
            return None
        p = self._cache_path(endpoint, body)
        if p.exists():
            try:
                return json.loads(p.read_text(encoding="utf-8"))["response"]
            except (json.JSONDecodeError, KeyError):
                return None  # corrupt entry: just refetch
        return None

    def _cache_write(self, endpoint, body, response):
        if not self.use_cache:
            return
        self.cache_dir.mkdir(parents=True, exist_ok=True)
        self._cache_path(endpoint, body).write_text(
            json.dumps({
                "endpoint": endpoint,
                "body": body,
                "fetched_at": time.strftime("%Y-%m-%dT%H:%M:%S"),
                "response": response,
            }, ensure_ascii=False, indent=1),
            encoding="utf-8",
        )

    # -- transport ---------------------------------------------------------

    def _request(self, method, endpoint, body=None, cacheable=True, timeout=180):
        if method == "POST" and cacheable:
            cached = self._cache_read(endpoint, body)
            if cached is not None:
                self.cache_hits += 1
                self.log(f"  [cache] {endpoint} (0 credits)")
                return cached

        url = f"{API_BASE}{endpoint}"
        self.api_calls += 1
        try:
            if method == "POST":
                r = self.session.post(url, json=body, timeout=timeout)
            else:
                r = self.session.get(url, timeout=timeout)
        except requests.RequestException as e:
            raise SystemExit(f"Network error calling {url}: {e}")

        if r.status_code == 401:
            raise SystemExit(
                "Firecrawl rejected the API key (401).\n\n"
                "The stored key is missing, wrong, or was deleted in the dashboard.\n"
                "To store a new one without writing it into PowerShell history:\n\n"
                '    $k = Read-Host "Paste key"; '
                '[Environment]::SetEnvironmentVariable("FIRECRAWL_API_KEY",$k,"User")\n\n'
                "Typing the key directly on the command line would persist it to\n"
                "ConsoleHost_history.txt, which is plaintext and survives reboots.\n"
                "A shell opened before the variable was set will not see it."
            )
        if r.status_code == 402:
            raise SystemExit("Out of Firecrawl credits (402). Check `firecrawl.py credits`.")
        if r.status_code == 429:
            raise SystemExit("Rate limited by Firecrawl (429). Wait a moment and retry.")
        if not r.ok:
            raise SystemExit(f"Firecrawl error {r.status_code} on {endpoint}: {r.text[:400]}")

        data = r.json()
        if method == "POST" and cacheable:
            self._cache_write(endpoint, body, data)
        return data

    # -- endpoints ---------------------------------------------------------

    def credits(self):
        """Read the account credit balance. Never cached — it must be live."""
        d = self._request("GET", "/team/credit-usage", cacheable=False, timeout=30)
        return d.get("data", {})

    def search(self, query, limit=5):
        return self._request("POST", "/search", {"query": query, "limit": limit})

    def map_site(self, url, limit=500):
        return self._request("POST", "/map", {"url": url, "limit": limit})

    def scrape(self, url, formats=None, only_main=True):
        formats = formats or ["markdown"]
        return self._request("POST", "/scrape", {
            "url": url,
            "formats": formats,
            "onlyMainContent": only_main,
        })

    def crawl(self, url, limit=50, only_main=True, poll_every=5, max_wait=1800):
        """Start an async crawl and poll to completion, following pagination."""
        body = {
            "url": url,
            "limit": limit,
            "scrapeOptions": {"formats": ["markdown"], "onlyMainContent": only_main},
        }
        started = self._request("POST", "/crawl", body)

        # A cached crawl already holds its finished pages.
        if started.get("status") == "completed" and "data" in started:
            return started

        job_id = started.get("id")
        if not job_id:
            raise SystemExit(f"Crawl did not return a job id: {started}")

        self.log(f"  crawl job {job_id} started; polling...")
        waited = 0
        while waited < max_wait:
            time.sleep(poll_every)
            waited += poll_every
            status = self._request("GET", f"/crawl/{job_id}", cacheable=False, timeout=120)
            state = status.get("status")
            self.log(f"  {state}: {status.get('completed', 0)}/{status.get('total', '?')} pages")
            if state == "completed":
                pages = list(status.get("data", []))
                # Large crawls page the results; follow `next` until exhausted.
                nxt = status.get("next")
                while nxt:
                    r = self.session.get(nxt, timeout=120)
                    if not r.ok:
                        break
                    chunk = r.json()
                    pages.extend(chunk.get("data", []))
                    nxt = chunk.get("next")
                status["data"] = pages
                self._cache_write("/crawl", body, status)
                return status
            if state == "failed":
                raise SystemExit(f"Crawl failed: {status}")
        raise SystemExit(f"Crawl still running after {max_wait}s. Job id: {job_id}")


# --------------------------------------------------------------------------
# Design token extraction
# --------------------------------------------------------------------------

HEX_RE = re.compile(r"#(?:[0-9a-fA-F]{3,4}|[0-9a-fA-F]{6}|[0-9a-fA-F]{8})\b")
# Matches both legacy rgb(1, 2, 3) and modern rgb(1 2 3 / var(--x)) syntax.
# Utility frameworks emit the modern form almost exclusively, and it carries
# the site's real palette — an extractor that only knows the comma form will
# report a handful of incidental blacks and whites and miss the design.
RGB_FUNC_RE = re.compile(r"\brgba?\(([^)]*(?:\([^)]*\))?[^)]*)\)", re.I)
HSL_FUNC_RE = re.compile(r"\bhsla?\(([^)]*(?:\([^)]*\))?[^)]*)\)", re.I)
NUM_RE = re.compile(r"-?[\d.]+")
FONT_FAMILY_RE = re.compile(r"font-family\s*:\s*([^;}\n]+)", re.I)
FONT_SIZE_RE = re.compile(r"font-size\s*:\s*([^;}\n]+)", re.I)
RADIUS_RE = re.compile(r"border-radius\s*:\s*([^;}\n]+)", re.I)
CSS_VAR_RE = re.compile(r"(--[a-zA-Z0-9_-]+)\s*:\s*([^;}\n]+)")
STYLESHEET_RE = re.compile(
    r'<link[^>]+rel=["\']?stylesheet["\']?[^>]*>', re.I)
HREF_RE = re.compile(r'href=["\']([^"\']+)["\']', re.I)
STYLE_BLOCK_RE = re.compile(r"<style[^>]*>(.*?)</style>", re.I | re.S)


# Utility-first CSS puts the palette in class names rather than in declarations,
# so "this site is slate + teal" is recoverable from the markup even when the
# stylesheet is a wall of generated utilities.
TW_COLORS = ("slate|gray|grey|zinc|neutral|stone|red|orange|amber|yellow|lime|"
             "green|emerald|teal|cyan|sky|blue|indigo|violet|purple|fuchsia|"
             "pink|rose")
TW_UTIL_RE = re.compile(
    r"\b(?:bg|text|border|ring|from|via|to|fill|stroke|decoration|outline|"
    r"shadow|accent|caret|divide|placeholder)-(" + TW_COLORS + r")-(\d{2,3})\b")
CLASS_ATTR_RE = re.compile(r'class=["\']([^"\']*)["\']', re.I)

# Tailwind's --tw-* properties are internal machinery (opacity slots, transform
# accumulators), not design decisions. Left in, they crowd out the real tokens.
FRAMEWORK_VAR_PREFIXES = ("--tw-",)


def normalize_hex(h):
    """Expand #abc to #aabbcc and drop alpha, so shades group correctly."""
    h = h.lower()
    body = h[1:]
    if len(body) in (3, 4):
        body = "".join(c * 2 for c in body[:3])
    return "#" + body[:6]


def rgb_to_hex(parts):
    try:
        r, g, b = (max(0, min(255, int(round(float(x))))) for x in parts[:3])
    except (ValueError, TypeError):
        return None
    return f"#{r:02x}{g:02x}{b:02x}"


def hsl_to_hex(parts):
    """Convert hsl() components to hex so they merge into one palette."""
    try:
        h = float(parts[0]) % 360 / 360.0
        s = float(parts[1]) / 100.0
        light = float(parts[2]) / 100.0
    except (ValueError, TypeError, IndexError):
        return None

    def channel(p, q, t):
        t %= 1
        if t < 1 / 6:
            return p + (q - p) * 6 * t
        if t < 1 / 2:
            return q
        if t < 2 / 3:
            return p + (q - p) * (2 / 3 - t) * 6
        return p

    if s == 0:
        r = g = b = light
    else:
        q = light * (1 + s) if light < 0.5 else light + s - light * s
        p = 2 * light - q
        r, g, b = channel(p, q, h + 1 / 3), channel(p, q, h), channel(p, q, h - 1 / 3)
    return "#%02x%02x%02x" % tuple(max(0, min(255, round(c * 255))) for c in (r, g, b))


def parse_color_functions(css):
    """Pull hex values out of every rgb()/hsl() form, plus a translucency flag."""
    found, translucent = Counter(), set()

    for match in RGB_FUNC_RE.finditer(css):
        inner = match.group(1)
        # Alpha may be a var() reference, so read only the leading channels.
        head, sep, tail = inner.partition("/")
        nums = NUM_RE.findall(head)
        if len(nums) < 3:
            continue
        hx = rgb_to_hex(nums)
        if not hx:
            continue
        found[hx] += 1
        # Only a literal alpha below 1 means the colour is really translucent.
        # A var() alpha is usually a framework opacity slot defaulting to 1, so
        # treating it as transparency would mislabel most of the palette.
        alpha_src = tail if sep else (nums[3] if len(nums) > 3 else "")
        if "var(" not in alpha_src:
            a = NUM_RE.findall(alpha_src)
            if a:
                try:
                    val = float(a[0])
                    if ("%" in alpha_src and val < 100) or ("%" not in alpha_src and val < 1):
                        translucent.add(hx)
                except ValueError:
                    pass

    for match in HSL_FUNC_RE.finditer(css):
        head = match.group(1).split("/")[0]
        nums = NUM_RE.findall(head)
        if len(nums) >= 3:
            hx = hsl_to_hex(nums)
            if hx:
                found[hx] += 1
    return found, translucent


def extract_utility_palette(raw_html):
    """Count Tailwind-style colour utilities, e.g. text-slate-200."""
    hits = Counter()
    for attr in CLASS_ATTR_RE.findall(raw_html or ""):
        for name, shade in TW_UTIL_RE.findall(attr):
            hits[f"{name}-{shade}"] += 1
    return hits


def collect_css(raw_html, page_url, log=print):
    """Return (combined_css_text, [source_labels]).

    Stylesheets are static files, so plain HTTP fetches them — no Firecrawl
    credits are spent here.
    """
    css_parts, sources = [], []

    for block in STYLE_BLOCK_RE.findall(raw_html or ""):
        css_parts.append(block)
    if css_parts:
        sources.append(f"{len(css_parts)} inline <style> block(s)")

    hrefs = []
    for tag in STYLESHEET_RE.findall(raw_html or ""):
        m = HREF_RE.search(tag)
        if m:
            hrefs.append(urljoin(page_url, m.group(1)))

    for href in hrefs[:12]:  # a dozen sheets is plenty; avoids pathological pages
        try:
            r = requests.get(href, timeout=30, headers={"User-Agent": "Mozilla/5.0"})
            if r.ok and "css" in r.headers.get("Content-Type", "css"):
                css_parts.append(r.text)
                sources.append(href)
        except requests.RequestException:
            log(f"  (could not fetch stylesheet {href})")
    return "\n".join(css_parts), sources


def extract_tokens(css, raw_html):
    """Distil a design system out of the CSS actually served by the page."""
    colors = Counter()
    for h in HEX_RE.findall(css):
        colors[normalize_hex(h)] += 1
    func_colors, translucent = parse_color_functions(css)
    colors.update(func_colors)

    fonts = Counter()
    for f in FONT_FAMILY_RE.findall(css):
        fonts[f.strip().strip('"\'')] += 1

    sizes = Counter(s.strip() for s in FONT_SIZE_RE.findall(css))
    radii = Counter(s.strip() for s in RADIUS_RE.findall(css))

    # Custom properties are the closest thing to the site's declared design
    # system — when a site defines them, they beat frequency-counted guesses.
    cssvars, skipped_vars = {}, 0
    for name, val in CSS_VAR_RE.findall(css):
        v = val.strip()
        if name.startswith(FRAMEWORK_VAR_PREFIXES) or not v:
            skipped_vars += 1
            continue
        if name not in cssvars and len(v) < 120:
            cssvars[name] = v

    return {
        "colors": colors,
        "translucent": translucent,
        "utility": extract_utility_palette(raw_html),
        "fonts": fonts,
        "sizes": sizes,
        "radii": radii,
        "cssvars": cssvars,
        "skipped_vars": skipped_vars,
    }


def write_design_tokens(path, url, title, tokens, sources, css_len):
    L = []
    L.append(f"# Design tokens — {title or url}")
    L.append("")
    L.append(f"Source: {url}")
    L.append(f"CSS analysed: {css_len:,} chars from {len(sources)} source(s)")
    L.append("")
    L.append("Counts are how often each value appears in the stylesheets. High")
    L.append("frequency usually means a core token; a long tail of near-identical")
    L.append("shades usually means one-off overrides worth ignoring.")
    L.append("")

    if tokens["utility"]:
        L.append("## Palette by utility class")
        L.append("")
        L.append("This page uses utility-first CSS, so the palette is named directly")
        L.append("in the markup. This is the clearest statement of the design's colour")
        L.append("intent — read it before the raw hex counts below.")
        L.append("")
        families = Counter()
        for token, n in tokens["utility"].items():
            families[token.rsplit("-", 1)[0]] += n
        L.append("Dominant colour families: "
                 + ", ".join(f"**{f}** ({n})" for f, n in families.most_common(5)))
        L.append("")
        L.append("| Utility colour | Uses |")
        L.append("|---|---|")
        for token, n in tokens["utility"].most_common(20):
            L.append(f"| `{token}` | {n} |")
        L.append("")

    L.append("## Colour palette")
    L.append("")
    L.append("Every hex, rgb() and hsl() value found, normalised to hex and merged.")
    L.append("")
    if tokens["colors"]:
        L.append("| Colour | Uses | Notes |")
        L.append("|---|---|---|")
        for c, n in tokens["colors"].most_common(28):
            note = "used with transparency" if c in tokens.get("translucent", ()) else ""
            L.append(f"| `{c}` | {n} | {note} |")
    else:
        L.append("_No colours found._")
    L.append("")

    L.append("## Typography")
    L.append("")
    if tokens["fonts"]:
        L.append("Font stacks, most used first:")
        L.append("")
        for f, n in tokens["fonts"].most_common(10):
            L.append(f"- `{f}` ({n})")
    else:
        L.append("_No font-family declarations found._")
    L.append("")
    if tokens["sizes"]:
        L.append("Font sizes (the type scale):")
        L.append("")
        L.append("`" + "`, `".join(s for s, _ in tokens["sizes"].most_common(18)) + "`")
        L.append("")

    if tokens["radii"]:
        L.append("## Border radius")
        L.append("")
        L.append("`" + "`, `".join(s for s, _ in tokens["radii"].most_common(10)) + "`")
        L.append("")

    if tokens["cssvars"]:
        L.append("## CSS custom properties")
        L.append("")
        L.append("The site's own declared design system — prefer these over the")
        L.append("frequency counts above when they disagree.")
        if tokens.get("skipped_vars"):
            L.append("")
            L.append(f"({tokens['skipped_vars']} framework-internal or empty properties "
                     "hidden, so the real tokens stay visible.)")
        L.append("")
        L.append("| Property | Value |")
        L.append("|---|---|")
        for k, v in list(tokens["cssvars"].items())[:60]:
            L.append(f"| `{k}` | `{v}` |")
        L.append("")

    L.append("## Stylesheet sources")
    L.append("")
    for s in sources:
        L.append(f"- {s}")
    L.append("")

    Path(path).write_text("\n".join(L), encoding="utf-8")


# --------------------------------------------------------------------------
# Helpers
# --------------------------------------------------------------------------

def slugify_url(url, fallback="page"):
    """Turn a page URL into a readable, filesystem-safe stem."""
    p = urlparse(url)
    path = unquote(p.path).strip("/")
    stem = path.replace("/", "__") if path else (p.netloc or fallback)
    stem = re.sub(r"[^\w\-.=]+", "-", stem, flags=re.UNICODE).strip("-")
    return (stem or fallback)[:120]


def parse_formats(spec):
    """'markdown,html,screenshot' -> the mixed string/object list the API wants."""
    out = []
    for f in [x.strip().lower() for x in spec.split(",") if x.strip()]:
        if f == "screenshot":
            out.append({"type": "screenshot", "fullPage": True})
        elif f == "html":
            out.append("rawHtml")
        else:
            out.append(f)
    return out


def resolve_key(args):
    key = args.api_key or os.environ.get("FIRECRAWL_API_KEY")
    if not key:
        sys.exit(
            "No Firecrawl API key found.\n\n"
            "Set it once, at user level:\n"
            '    setx FIRECRAWL_API_KEY "fc-..."\n'
            "then open a NEW terminal (setx does not affect the current one),\n"
            "or pass --api-key on this command."
        )
    return key


def report_credits(fc, before, label="Credits"):
    """Tie out spend against the account, not against our own assumptions."""
    try:
        after = fc.credits().get("remainingCredits")
    except SystemExit:
        return
    if before is None or after is None:
        return
    spent = before - after
    print(f"\n{label}: {spent} used this run | {after} remaining "
          f"({fc.cache_hits} cache hit(s), {fc.api_calls} API call(s))")


# --------------------------------------------------------------------------
# Commands
# --------------------------------------------------------------------------

def cmd_credits(fc, args):
    d = fc.credits()
    print(f"Remaining credits : {d.get('remainingCredits')}")
    print(f"Plan credits      : {d.get('planCredits')}")
    print(f"Billing period    : {d.get('billingPeriodStart')} -> {d.get('billingPeriodEnd')}")


def cmd_search(fc, args):
    before = fc.credits().get("remainingCredits")
    r = fc.search(args.query, args.num)
    results = (r.get("data") or {}).get("web") or []
    if not results:
        print("No results.")
    for i, item in enumerate(results, 1):
        print(f"{i}. {item.get('title')}")
        print(f"   {item.get('url')}")
        if item.get("description"):
            print(f"   {item['description'][:160]}")
    report_credits(fc, before)


def cmd_map(fc, args):
    before = fc.credits().get("remainingCredits")
    r = fc.map_site(args.url, args.limit)
    links = r.get("links") or []
    print(f"{len(links)} URL(s) found on {args.url}\n")
    for item in links:
        u = item.get("url") if isinstance(item, dict) else item
        print(u)
    if args.output:
        Path(args.output).write_text(
            "\n".join(i.get("url") if isinstance(i, dict) else i for i in links),
            encoding="utf-8")
        print(f"\nWritten to {args.output}")
    report_credits(fc, before)


def cmd_scrape(fc, args):
    before = fc.credits().get("remainingCredits")
    formats = parse_formats(args.formats)
    r = fc.scrape(args.url, formats, only_main=not args.full_page_content)
    data = r.get("data") or {}
    md = data.get("markdown") or ""
    title = (data.get("metadata") or {}).get("title") or ""

    if args.output:
        out = Path(args.output)
        out.parent.mkdir(parents=True, exist_ok=True)
        out.write_text(
            f"---\ntitle: {title}\nurl: {args.url}\n---\n\n{md}", encoding="utf-8")
        print(f"Wrote {len(md):,} chars of markdown to {out}")
    else:
        print(md)

    if data.get("screenshot"):
        print(f"\nScreenshot URL (expires): {data['screenshot']}")
    print(f"\nTitle: {title}", file=sys.stderr)
    print(f"Markdown length: {len(md):,} chars", file=sys.stderr)
    report_credits(fc, before)


def cmd_crawl(fc, args):
    before = fc.credits().get("remainingCredits")
    outdir = Path(args.output)

    # Look before spending: map is cheap and tells us the real page count.
    if not args.skip_map:
        m = fc.map_site(args.url, limit=max(args.limit, 500))
        found = len(m.get("links") or [])
        planned = min(found, args.limit)
        print(f"Site map: {found} URL(s) found; crawl limit is {args.limit} "
              f"-> about {planned} page(s), ~{planned} credit(s).")
        if planned > CRAWL_CONFIRM_THRESHOLD and not args.yes:
            sys.exit(
                f"\nStopping: {planned} pages exceeds the {CRAWL_CONFIRM_THRESHOLD}-page "
                f"safety threshold.\nRe-run with --yes to proceed, or lower --limit."
            )
        if before is not None and planned > before:
            sys.exit(f"\nStopping: {planned} pages needed but only {before} credits remain.")

    status = fc.crawl(args.url, limit=args.limit, only_main=not args.full_page_content)
    pages = status.get("data") or []
    outdir.mkdir(parents=True, exist_ok=True)

    index = ["# Crawl index", "", f"Source: {args.url}", f"Pages: {len(pages)}", "",
             "| # | Title | File | URL |", "|---|---|---|---|"]
    seen = set()
    for i, page in enumerate(pages, 1):
        meta = page.get("metadata") or {}
        url = meta.get("sourceURL") or meta.get("url") or f"page-{i}"
        title = (meta.get("title") or "").replace("|", "\\|")
        stem = slugify_url(url, f"page-{i}")
        if stem in seen:
            stem = f"{stem}-{i}"
        seen.add(stem)
        fname = f"{stem}.md"
        (outdir / fname).write_text(
            f"---\ntitle: {title}\nurl: {url}\n---\n\n{page.get('markdown') or ''}",
            encoding="utf-8")
        index.append(f"| {i} | {title} | [{fname}]({fname}) | {url} |")

    (outdir / "index.md").write_text("\n".join(index) + "\n", encoding="utf-8")
    print(f"\nWrote {len(pages)} page(s) + index.md to {outdir}")
    print(f"Firecrawl reported creditsUsed = {status.get('creditsUsed')}")
    report_credits(fc, before)


def cmd_capture(fc, args):
    """Capture everything needed to emulate a page's design."""
    before = fc.credits().get("remainingCredits")
    outdir = Path(args.output)
    outdir.mkdir(parents=True, exist_ok=True)

    r = fc.scrape(
        args.url,
        formats=["markdown", "rawHtml", "links", {"type": "screenshot", "fullPage": True}],
        only_main=False,  # design lives in the chrome, not just the article body
    )
    data = r.get("data") or {}
    meta = data.get("metadata") or {}
    title = meta.get("title") or args.url
    raw_html = data.get("rawHtml") or ""

    (outdir / "page.html").write_text(raw_html, encoding="utf-8")
    (outdir / "content.md").write_text(data.get("markdown") or "", encoding="utf-8")
    print(f"Saved page.html ({len(raw_html):,} chars) and content.md")

    shot_url = data.get("screenshot")
    if shot_url:
        try:
            img = requests.get(shot_url, timeout=90)
            if img.ok:
                (outdir / "screenshot.png").write_bytes(img.content)
                print(f"Saved screenshot.png ({len(img.content):,} bytes)")
        except requests.RequestException as e:
            print(f"Could not download screenshot: {e}")

    css, sources = collect_css(raw_html, args.url)
    if css:
        (outdir / "styles.css").write_text(css, encoding="utf-8")
        tokens = extract_tokens(css, raw_html)
        write_design_tokens(outdir / "design-tokens.md", args.url, title,
                            tokens, sources, len(css))
        print(f"Saved styles.css ({len(css):,} chars) and design-tokens.md "
              f"({len(tokens['colors'])} colours, {len(tokens['fonts'])} font stacks, "
              f"{len(tokens['cssvars'])} CSS variables)")
    else:
        print("No CSS found — the page may load styles via JS-injected sheets.")

    (outdir / "metadata.json").write_text(
        json.dumps({"url": args.url, "title": title, "metadata": meta,
                    "links": data.get("links") or []}, ensure_ascii=False, indent=2),
        encoding="utf-8")
    print(f"\nDesign bundle written to {outdir}")
    report_credits(fc, before)


# --------------------------------------------------------------------------

def main():
    p = argparse.ArgumentParser(
        description="Credit-aware Firecrawl v2 client.",
        formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--api-key", help="Overrides FIRECRAWL_API_KEY.")
    p.add_argument("--cache-dir", default=DEFAULT_CACHE)
    p.add_argument("--no-cache", action="store_true",
                   help="Force a live refetch and bill for it.")
    sub = p.add_subparsers(dest="cmd", required=True)

    sub.add_parser("credits", help="Show plan and remaining credits.")

    s = sub.add_parser("search", help="Search the web.")
    s.add_argument("query")
    s.add_argument("-n", "--num", type=int, default=5)

    s = sub.add_parser("map", help="List a site's URLs (cheap; no page content).")
    s.add_argument("url")
    s.add_argument("--limit", type=int, default=500)
    s.add_argument("-o", "--output")

    s = sub.add_parser("scrape", help="Scrape one URL to markdown.")
    s.add_argument("url")
    s.add_argument("-o", "--output")
    s.add_argument("--formats", default="markdown",
                   help="Comma list: markdown,html,links,screenshot,summary")
    s.add_argument("--full-page-content", action="store_true",
                   help="Keep nav/footer instead of main content only.")

    s = sub.add_parser("crawl", help="Crawl a site into one .md per page.")
    s.add_argument("url")
    s.add_argument("-o", "--output", required=True)
    s.add_argument("--limit", type=int, default=50)
    s.add_argument("--yes", action="store_true", help="Skip the page-budget prompt.")
    s.add_argument("--skip-map", action="store_true", help="Skip the pre-flight cost check.")
    s.add_argument("--full-page-content", action="store_true")

    s = sub.add_parser("capture", help="Design bundle: screenshot, HTML, CSS, tokens.")
    s.add_argument("url")
    s.add_argument("-o", "--output", required=True)

    args = p.parse_args()
    fc = Firecrawl(resolve_key(args), args.cache_dir, use_cache=not args.no_cache)

    {
        "credits": cmd_credits, "search": cmd_search, "map": cmd_map,
        "scrape": cmd_scrape, "crawl": cmd_crawl, "capture": cmd_capture,
    }[args.cmd](fc, args)


if __name__ == "__main__":
    main()
