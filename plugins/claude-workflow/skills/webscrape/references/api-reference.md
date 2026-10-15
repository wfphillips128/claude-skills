# Firecrawl v2 API reference

Base URL: `https://api.firecrawl.dev/v2`
Auth header: `Authorization: Bearer fc-...`

Everything below was verified against the live API on 15 Aug 2026, not copied
from documentation. Where the official docs and this file disagree, re-verify
before trusting either — the docs at <https://docs.firecrawl.dev> are the
upstream source of truth for anything not listed here.

## Contents

- [Endpoints wrapped by the script](#endpoints-wrapped-by-the-script)
- [Endpoints not yet wrapped](#endpoints-not-yet-wrapped)
- [Formats](#formats)
- [Errors](#errors)
- [Credit costs](#credit-costs)

---

## Endpoints wrapped by the script

### `GET /team/credit-usage`

Live account balance. Not cached by the script, because a stale balance defeats
the point.

```json
{"success": true,
 "data": {"remainingCredits": 1015, "planCredits": 1000,
          "billingPeriodStart": "2026-08-15T15:22:07.345Z",
          "billingPeriodEnd": "2026-09-15T15:22:07.345Z"}}
```

Note `remainingCredits` can exceed `planCredits` — signup bonuses land on top
of the monthly allowance.

### `POST /search`

```json
{"query": "craig hatmaker 5g lambda", "limit": 5}
```

Results are nested under `data.web` — an array of `{title, url, description}`.
Other verticals (news, images) appear as sibling keys under `data`, so read
`data.web` explicitly rather than assuming `data` is itself the list.

### `POST /map`

```json
{"url": "https://sites.google.com/site/beyondexcel", "limit": 500}
```

Returns `{"success": true, "links": [{"url": "...", "title": "...?"}]}`.
`links` holds objects, not bare strings.

Cheap and content-free — this is the right way to size a crawl before paying
for one. The count is not perfectly deterministic between runs (34 then 35 on
consecutive calls against the same site), so treat it as a close estimate.

### `POST /scrape`

```json
{"url": "https://example.com",
 "formats": ["markdown", "rawHtml", "links", {"type": "screenshot", "fullPage": true}],
 "onlyMainContent": true}
```

Response `data` contains one key per requested format, plus `metadata`:

```
markdown, rawHtml, html, links, summary, screenshot, metadata
```

`screenshot` is a **signed Google Storage URL with an expiry**, not image
bytes — download it immediately or lose it.

`metadata` includes `title`, `sourceURL`, `statusCode`, `language`, `favicon`,
`creditsUsed`, `cacheState`, `cachedAt`, `proxyUsed`.

`onlyMainContent: true` strips nav and footer. Good for reading, wrong for
design work — the chrome is part of the design.

### `POST /crawl` — asynchronous

Starting a crawl returns a job, not results:

```json
{"success": true, "id": "01a00617-...", "url": "https://api.firecrawl.dev/v2/crawl/01a00617-..."}
```

Request body:

```json
{"url": "https://example.com/docs", "limit": 50,
 "scrapeOptions": {"formats": ["markdown"], "onlyMainContent": true}}
```

### `GET /crawl/{id}` — poll

```json
{"status": "scraping", "completed": 2, "total": 2, "creditsUsed": 2,
 "data": [{"markdown": "...", "metadata": {"title": "...", "sourceURL": "..."}}],
 "expiresAt": "...", "createdAt": "...", "completedAt": "...", "duration": 0}
```

`status` moves `scraping` → `completed`, or `failed`. Note that `completed` can
already equal `total` while `status` is still `scraping`; wait for the status
field, not the counters.

Large result sets paginate via a `next` URL at the top level — follow it until
absent or pages will be silently missing. Results expire (`expiresAt`), which
is a further reason the script caches them locally.

---

## Endpoints not yet wrapped

Add these to `scripts/firecrawl.py` if a task needs them.

| Endpoint | Purpose |
|---|---|
| `POST /interact` | Clicks, form fills, navigation on a live page — for content behind an interaction |
| `POST /parse` | Upload a **local** document (PDF, DOCX, XLSX, up to 50 MB) as `multipart/form-data` → markdown. A document with a public URL should go through `/scrape` instead |
| `POST /monitor` | Recurring checks on a page or search, diffed against the last snapshot, with webhook/email/Slack notification. Better than repeat scraping whenever a URL needs watching more than once |
| `GET /search/research/papers` | Scientific paper index — search, inspect, full-text passages, citation expansion |
| `GET /search/research/github` | GitHub issues, PRs, discussions and READMEs |
| `POST /support/ask` | Diagnose a failing call; takes `{question, jobId}`, returns prose plus `fixParameters` |
| `POST /support/docs-search` | "How does Firecrawl do X" answered from current docs, with citations |

The research index endpoints work without an `Authorization` header, on a
rate-limited keyless tier.

---

## Formats

Passed as `formats` on `/scrape` and inside `scrapeOptions` on `/crawl`. Mixed
strings and objects in one array is valid and normal.

| Value | Returns |
|---|---|
| `"markdown"` | Clean markdown (default) |
| `"rawHtml"` | Unmodified HTML |
| `"html"` | Cleaned HTML |
| `"links"` | Array of links found |
| `"summary"` | AI summary |
| `{"type": "screenshot", "fullPage": true}` | Signed URL to a PNG |
| `{"type": "json", "schema": {...}}` | Structured extraction against a schema |

The script's `--formats` flag takes a comma list and maps `screenshot` to the
full-page object form and `html` to `rawHtml`.

---

## Errors

| Code | Meaning | Response |
|---|---|---|
| 401 | Bad or missing key | Check `FIRECRAWL_API_KEY`; `setx` does not affect an open shell |
| 402 | Out of credits | Check the balance; wait for the billing period or upgrade |
| 429 | Rate limited | Back off; the free tier is throttled |
| 5xx | Firecrawl side | Retry once, then stop |

---

## Credit costs

Roughly one credit per page or per search result. A screenshot does not appear
to bill separately from the scrape that requested it.

The reliable way to know actual spend is to read `/team/credit-usage` before
and after a job and take the difference — which is what the script's ledger
line does. Do not rely on estimates when reporting cost to the user.
