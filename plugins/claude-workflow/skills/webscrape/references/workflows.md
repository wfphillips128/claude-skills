# Worked workflows

Both examples below are real runs, with real output. `PY` and `FC` stand for the
Python path and the script path from SKILL.md.

---

## 1. Capture a design to emulate

**Situation:** the user is building a personal site and points at an example
whose look they want.

```bash
$PY $FC capture "https://brittanychiang.com" -o ./design-ref/chiang
```

Output:

```
Saved page.html (65,921 chars) and content.md
Saved screenshot.png (192,751 bytes)
Saved styles.css (28,994 chars) and design-tokens.md
  (30 colours, 6 font stacks, 5 CSS variables)

Credits: 1 used this run | 1015 remaining
```

### Reading the result

Open `design-tokens.md` first. For this site it led with:

```
Dominant colour families: slate (116), teal (78), yellow (2), red (1)

| Utility colour | Uses |
| slate-200      | 64   |
| teal-300       | 39   |
| teal-400       | 38   |
| slate-800      | 14   |
```

That single line — *slate and teal* — is the design brief. The raw palette
table then supplies the actual values: `#e2e8f0` for the body text, `#0f172a`
for the background, `#5eead4` and `#2dd4bf` for accents.

Then open `screenshot.png` to see how those tokens are deployed: where the
accent colour is spent, how much whitespace carries the layout, what the type
hierarchy does. Tokens tell you *what* the palette is; the screenshot tells you
*how much of each* to use, which is usually the part people get wrong when
emulating a design.

### Why not just read the markdown

`content.md` for this page is 8,971 characters of résumé prose with no colour,
no spacing, no type scale. Everything that makes the page worth copying lives
in the other five files.

### Comparing several references

Capture each into its own folder under one parent, then read the token files
side by side. Shared choices across three sites the user likes are a much
stronger signal than anything one site does.

```bash
$PY $FC capture "https://site-one.com"   -o ./design-ref/one
$PY $FC capture "https://site-two.com"   -o ./design-ref/two
```

Re-reading any of these later is free — the scrape is cached.

### Scope

Layout, spacing rhythm, type scale and palette are ordinary things to learn
from. Logos, wordmarks, photography, illustration and copy are not. Take the
system, leave the identity.

---

## 2. Harvest a site into a reference set

**Situation:** a reference site holds material that should inform a skill.

### Step 1 — map first, and look at what came back

```bash
$PY $FC map "https://sites.google.com/site/beyondexcel" -o urls.txt
```

```
35 URL(s) found
```

Read the list before crawling. Structure usually reveals itself in the URL
paths — here, a `home/excel-library/lambda-*` group turned out to be the
function-by-function library, which was the actual objective, while
`home/vba-coding/*` and `home/about/project-updates/*` were adjacent material.

Thirty-five pages is about thirty-five credits, which is worth stating to the
user before spending it.

### Step 2 — crawl within a justified limit

```bash
$PY $FC crawl "https://sites.google.com/site/beyondexcel" -o ./harvest --limit 40
```

The script maps first as a pre-flight check, prints the projected page count and
cost, and stops if the job would exceed 60 pages or the remaining balance.

### Step 3 — read the index, not the folder

`index.md` in the output directory lists every page with title, filename and
source URL. Use it to decide what to read; do not read forty files in sequence
hoping something useful appears.

### Step 4 — curate

A harvest is raw input. Turning forty scraped pages into a skill means writing
a reference that states what was learned and cites the URLs it came from — not
copying the pages across. A skill that grows by forty files has become harder
to use, not more capable.

Where harvested material contradicts what a skill already says, surface the
conflict rather than resolving it silently. A published web page is not
automatically more authoritative than working code someone has actually tested,
and the existing wording may encode a deliberate decision.

---

## 3. Find something when there is no URL

```bash
$PY $FC search "craig hatmaker 5g lambda" -n 5
```

```
1. 5G Modeling - Beyond Excel
   https://sites.google.com/site/beyondexcel/home/5g-modeling
2. 5g Functions LAMBDA for Excel for Financial Models (gist)
   https://gist.github.com/CHatmaker/5dcb3de4ad9766d1baf0b95338ebdbf3
```

Search costs about a credit per result, so keep `-n` small. Then `scrape` the
one or two URLs that matter rather than crawling everything the search found.

---

## Cost reference from these runs

| Action | Credits |
|---|---|
| `credits` | 0 |
| `search -n 5` | ~2 |
| `map` (35 URLs) | 1 |
| `scrape` one page | 1 |
| `capture` one page | 1 |
| Any repeat of a cached request | **0** |

The cache is what makes iteration affordable: the design-token extractor was
revised and re-run three times against the same page for one credit total.
