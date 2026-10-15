---
name: webscrape
description: |
  Pull content and design off live websites using Firecrawl — scrape a single
  page to clean markdown, crawl or map a whole site, search the web, or capture
  a page's full design system (screenshot, HTML, CSS, colour palette, fonts).
  Firecrawl runs a real browser, so it succeeds on JavaScript-rendered sites
  that plain fetching returns empty. Use this skill whenever the user wants
  content from a URL or a site — "scrape this page", "crawl these docs", "get
  the content of", "pull the text from this link", "what's on this site",
  "copy this site's design", "I want my page to look like this one", "build a
  reference set from this documentation", "research this company's website", or
  any mention of Firecrawl. Reach for it even when the user does not say
  "scrape" — if the task needs live web page content, this is the tool. Also
  use it when a normal fetch of a page came back empty, blocked, or
  JavaScript-only.
---

# Webscrape

A credit-aware Firecrawl client. Everything runs through one script:

```
scripts/firecrawl.py
```

Run it with the user's real Python, which is reliable on this machine where
bare `python` can resolve to the Windows Store stub:

```powershell
python `
  "scripts/firecrawl.py" <command> ...
```

## Auth

The key lives in the user-level environment variable `FIRECRAWL_API_KEY`.
`setx` writes it permanently but does **not** affect an already-open terminal,
so if a command reports a missing key in a shell that was open when the key was
set, either open a new terminal or pass `--api-key` for that call.

Never write the API key into a file inside a project folder, and never echo it
into output the user might share.

## Choosing a command

The commands differ mainly in what they cost and what they throw away. Pick by
what you actually need to end up with.

| Need | Command | Cost |
|---|---|---|
| Find pages when you have no URL | `search "query"` | ~1 credit per result |
| See a site's shape before committing | `map <url>` | ~1 credit total |
| One page as clean markdown | `scrape <url>` | 1 credit |
| Many pages as a reference set | `crawl <url> -o dir` | 1 credit per page |
| A page's *design*, not its text | `capture <url> -o dir` | 1 credit |

```bash
firecrawl.py credits                                  # plan + balance
firecrawl.py search "craig hatmaker 5g lambda" -n 5
firecrawl.py map https://example.com -o urls.txt
firecrawl.py scrape https://example.com -o page.md
firecrawl.py scrape https://example.com --formats markdown,html,screenshot
firecrawl.py crawl https://example.com/docs -o ./harvest --limit 50
firecrawl.py capture https://example.com -o ./design-ref
```

Useful flags: `--no-cache` forces a paid refetch, `--limit` bounds a crawl,
`--yes` clears the crawl safety threshold, `--full-page-content` keeps nav and
footer instead of main content only.

## Spending credits responsibly

The account is on a free plan: **1000 credits per month**, and the balance does
not roll over. That is enough for real work but not enough to be careless, so
two habits matter.

**Map before you crawl.** `map` costs about one credit and tells you exactly
how many pages a crawl would bill for. `crawl` does this automatically and
refuses to exceed 60 pages without `--yes`, but when planning aloud with the
user, run `map` first so you can quote a real number instead of guessing.

**Trust the cache.** Every response is cached in `.firecrawl-cache/` keyed by
the request, so re-running a command is free and instant. This means you can
re-derive a design bundle or re-read a harvest as many times as you like — only
the first fetch bills. Reach for `--no-cache` only when the page has genuinely
changed since the last fetch; the cache never expires on its own.

Every command prints a credit ledger when it finishes:

```
Credits: 1 used this run | 1015 remaining (0 cache hit(s), 3 API call(s))
```

Report that line to the user after any multi-page job. The count comes from the
account balance before and after, not from an internal estimate, so it is the
real number.

## Workflow: capturing a design to emulate

When the user wants a page to look like some other page, markdown is useless —
it discards exactly the information that matters. `capture` writes a bundle:

| File | What it gives you |
|---|---|
| `screenshot.png` | Full-page render — the ground truth for layout |
| `design-tokens.md` | Palette, type scale, fonts, radii, CSS variables |
| `styles.css` | Every stylesheet, concatenated |
| `page.html` | Raw markup — structure and class names |
| `content.md` | Text content |
| `metadata.json` | Title, description, links |

Read `design-tokens.md` first, then look at `screenshot.png`. The token file
leads with a **utility-class palette** when the site uses Tailwind-style CSS,
because a site that writes `text-slate-200` and `bg-teal-400` in its markup is
telling you its colour intent directly — that is far more legible than counting
hex codes. On sites that do not use utility classes, fall back to the raw
palette table and the CSS custom properties, which are the site's own declared
tokens.

A caution worth carrying into any emulation job: reproducing a site's *layout,
type scale and spacing rhythm* is ordinary design practice, but its logo,
wordmark, photography, copy and distinctive illustration are the author's. Lift
the system, not the identity.

## Workflow: harvesting a site into a reference set

For turning documentation or a reference site into material another skill can
use:

1. `map` the site and look at the URL list. Decide what is actually worth
   having — a site's tag archives and pagination pages usually are not.
2. **Prefer targeted scrapes over a broad crawl when the map shows the pages you
   want are a small, identifiable subset.** A crawl walks the site in its own
   order and stops when it hits `--limit`; if the site has a large blog archive,
   the budget can be spent entirely inside it before the crawl ever reaches what
   you were after. On a real 40-page crawl of a reference site, 38 pages were
   blog posts and the target function library was never reached — the 10 pages
   that mattered then cost 10 more credits to fetch by URL. Scraping a filtered
   list from the map output would have cost a quarter as much and taken less
   time.

   Use a broad crawl when you genuinely want the whole site, or when the map is
   small enough that the limit cannot bite. Otherwise filter the map, then
   scrape those URLs in a loop.
3. Where a crawl is right, run it with a `--limit` you can justify, writing to a
   folder in the **project**, not into a skill.
4. Read `index.md` in the output, which lists every page with its title and
   source URL.
5. **Curate rather than paste.** A harvest is raw input, not a deliverable.
   Distil what is genuinely reusable into a written reference and cite the
   source URLs; dumping fetched pages into a skill folder makes it bigger
   without making it better.

When harvested material contradicts something a skill already asserts, say so
explicitly rather than silently overwriting — the existing text may encode a
deliberate decision, and a published website is not automatically more
authoritative than working code someone tested.

## When a page comes back thin or empty

Firecrawl already runs a real browser, so a JavaScript-only site is not by
itself a problem. If a result still looks wrong:

- Empty or nearly empty markdown → retry with `--full-page-content`. Main
  content detection can be too aggressive on unusual layouts.
- Need the markup, not the prose → `--formats html`.
- Blocked, paywalled, or login-gated → Firecrawl cannot help; tell the user
  plainly instead of retrying and burning credits.
- A crawl that returns far fewer pages than `map` found → the site links those
  pages only from JS navigation. Scrape the specific URLs from the map output
  instead.

Errors surface as plain messages, not tracebacks: 401 means the key, 402 means
credits exhausted, 429 means slow down.

## References

- `references/api-reference.md` — verified v2 endpoints, request shapes and
  response fields, including endpoints this script does not yet wrap
  (`parse`, `monitor`, the research index). Read it before extending the
  script or calling the API directly.
- `references/workflows.md` — worked end-to-end examples of the design capture
  and site harvest workflows, with real commands and outputs.
