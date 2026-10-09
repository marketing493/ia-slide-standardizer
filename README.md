# IA Battlecard Builder

A competitive battlecard builder for Impact Analytics. Pick a competitor, pick an IA
product, pick a length, and Claude researches the public record and writes the deck.

- **Battlecard Builder** at `/` is the product. Three inputs, a chat assistant, and a
  PPTX download.
- **Library** at `/library` is the shared gallery. Every generated card lands there, so
  the team browses one place instead of passing decks around.
- **Slide Standardizer** at `/standardizer` is the original tool, kept but no longer the
  front door.

Every deck follows the Impact Analytics brand guide and opens unchanged in PowerPoint,
Keynote and Google Slides.

## How generation works

1. **Research.** Claude Opus 5 searches the public record with the `web_search` server
   tool and writes a sourced brief. The brief streams to the browser while it is written,
   so a run that takes minutes never looks stalled and never trips a worker timeout.
2. **Structure.** A second call turns the brief into the card payload. The shape is given
   to the model as a contract in the prompt and parsed on the way back, tolerating a code
   fence or a sentence of preamble, with one retry if it does not parse. `normalize()` is
   the validator: it coerces every field and drops anything unrecognised.

   This deliberately does **not** use `output_config.format`. That compiles the schema into
   a grammar, and `CARD_SCHEMA` has twelve nested object shapes, which the API rejects with
   "The compiled grammar is too large". Splitting the schema would only move the ceiling, so
   the request carries no grammar at all and a test asserts it never will.
3. **Merge.** Where a hand researched card exists in `content/battlecards/`, its sourced
   sections win. The model fills gaps rather than overwriting research.
4. **Build.** The card is normalised, trimmed to the chosen depth, checked against the
   brand writing rules, and rendered to PPTX.

**The honesty design is the point.** Claude receives the IA product facts as ground truth
and is instructed never to state a competitor capability, number, customer or price it did
not source. Where the public record is silent, the capability rating comes back `unknown`,
which the deck renders as "Unclear" with "Ask the question, do not assert the gap". A
battlecard that overstates a rival's gap loses the deal the moment the buyer corrects it,
so the builder would rather hand a seller a question than a guess.

### What a run costs

Every run reports its own token spend, live during generation and again in the result,
computed from `response.usage` at the published Claude Opus 5 rates ($5 per million input,
$25 per million output, cache writes 1.25x input, cache reads 0.1x).

| Path | Token cost |
|---|---|
| Building a card already in the library, at any depth | **$0**, no API call |
| Economy run: low effort, 3 searches | roughly **$0.20** |
| Standard run: high effort, 8 searches | roughly **$0.55 to $0.75** |

**Web search bills a separate per-search fee** that does not appear in `usage`, so the
figure the app shows is a floor rather than the whole invoice. Check the Anthropic pricing
page for the current per-search rate.

The economy checkbox drops effort to `low`, cuts searches from eight to three and tightens
`max_tokens`. The brief comes back thinner, and the card has more Unclear ratings, which is
the honest tradeoff of searching less.

### Depths

| Depth | Slides | What it is |
|---|---|---|
| Summary | 6 to 8 | The one page card plus the essentials |
| Standard | 11 to 13 | The working card for most meetings |
| Full technical | 17 to 20 | Every section, including the capability matrix and pricing |
| Head to head only | 4 to 10 | The comparison on its own, as a deck of its own |

Depth caps rows per section, so a summary stays a summary. A card that names its own
sections, such as the curated files in `content/`, is never trimmed.

### Head to head only

The comparison matrix as its own deck, for the buyer who asks for a straight comparison. It is
built from the ratings alone, so it never says more than the matrix does:

1. **Scoreboard.** Impact Blue cover with a one sentence verdict, one mark per capability in
   lead order, and the count of rows where we lead, where it is level, where they lead, and
   what is still unverified.
2. **The field at a glance.** Every capability as a tile with a Harvey ball per vendor, in four
   lanes by who leads. Lanes take width by how much they hold.
3. **The matrix.** The detailed rows with what to say.
4. **How to play the matrix.** One move per capability, taken from the last instruction in its
   note: lead with these, compete on how, reframe, prove before you claim.

Pick it before generating, or use **Head to head only** on the result panel or in the library to
build it from a card that already exists, with no new research.

## Running it

```bash
pip install -r requirements.txt
python3 app.py                                  # http://127.0.0.1:5000
python3 build_battlecards.py                    # build every card in content/
python3 -m pytest tests -q
```

### Environment

| Variable | Effect when unset |
|---|---|
| `ANTHROPIC_API_KEY` | No research and no assistant. Saved cards still build. |
| `DATABASE_URL` | No shared library. Cards still generate and download. |
| `IA_AUTH_USER`, `IA_AUTH_PASSWORD` | Every route is open. Unsafe with real content. |

Nothing here is required to run the app. Each missing piece disables its own feature and
says so in the UI rather than failing.

### Access control

Set `IA_AUTH_USER` and `IA_AUTH_PASSWORD` to gate every route behind HTTP basic
auth. Battlecards hold competitive intelligence and unpublished customer numbers,
so a deployment with real content needs both. `/healthz` and `/static/` stay open
so the platform probe keeps working. With the variables unset the app runs open
and logs a warning. No credential lives in this repository.

Docker and Render configuration already exist in `Dockerfile` and `render.yaml`.

## Battlecard Builder

### What it produces

A battlecard deck of up to 16 sections. Long sections paginate on their own, so seven
objections become four slides without any manual work.

### How text is set

A generated card writes long: matrix notes of 400 characters, talk track blocks of 800, a
citation on every claim. `battlecards/typeset.py` makes that fit without shrinking it into
unreadable type or letting it run off a card.

- Citations leave the slide face. "Source: ..." sentences, ", per Guide.pdf pages 6 and 20"
  clauses and bare URLs go to the speaker notes. Proof slots show a short linked label such
  as `example.com` or `Planning Guide, pp. 6 to 20`.
- Every block is measured before it is drawn, counting the line spacing, and set at the
  largest legible size that fits. Sibling cards on a slide share one size.
- A block that still does not fit is cut at a sentence boundary. The full text goes to the
  speaker notes under IN FULL, so a detail moves, it is never lost.
- Lists, matrix rows and card sets paginate, spread evenly across pages (ten advantages at
  three a page become 3, 3, 2, 2), numbered continuously, with "2 of 4" in the kicker.

`tools/deckcheck.py` is the arbiter. It renders a deck through LibreOffice, reads every
word's position from the PDF and reports collisions, text past the body area and text that
escapes its card:

```bash
python3 tools/deckcheck.py path/to/deck.pptx
```

`tests/test_typeset.py` builds a deck from deliberately long synthetic content and runs the
same check when LibreOffice is installed.

| Section | Slide |
|---|---|
| `cover` | Title card on Impact Blue with the grid pattern and the white logo |
| `how_to_use` | Working rules plus the win theme |
| `snapshot` | Company facts grid and recent moves |
| `positioning` | Their claim, our claim, and the wedge between them |
| `strengths_weaknesses` | An honest read of both sides |
| `why_we_win` | Numbered advantage cards, each with a proof line |
| `comparison` | Head to head matrix: a Harvey ball per vendor, an edge bar per row, rows ordered from our lead to theirs, and a running score |
| `objections` | They say, we say, proof |
| `landmines` | Trap questions with why each one lands |
| `discovery` | Questions grouped by theme |
| `proof_points` | Stat cards with a source link or an internal citation |
| `talk_track` | Positioning statement, pitch, opener, trap |
| `dos_donts` | Selling hygiene |
| `pricing` | Commercial models plus ground rules |
| `next_steps` | Numbered steps and linked resources |
| `one_pager` | The dense printable card |

### Using the web builder

Open `/battlecard`, name the competitor, pick the Impact Analytics product, then press
**Load starter card**. The starter fills every section with structure and prompts. Competitor
facing fields carry research prompts rather than claims, because nothing in this repository
asserts a fact about a rival. Edit, press **Review copy** to check the brand writing rules,
then press **Build PPTX**.

Import and export JSON so a team can version a card in git.

### Using the API

```bash
# Is Claude reachable, and what depths exist
curl localhost:5000/api/battlecard/status

# Research a competitor and stream the card back as server sent events
curl -N -X POST localhost:5000/api/battlecard/generate \
  -H 'Content-Type: application/json' \
  -d '{"competitor":"o9 Solutions","ia_product":"AttributeSmart","depth":"summary"}'

# Presets: products, solutions, competitor list, section list
curl localhost:5000/api/battlecard/presets

# A prefilled card to edit
curl -X POST localhost:5000/api/battlecard/scaffold \
  -H 'Content-Type: application/json' \
  -d '{"competitor":"Example Rival","ia_product":"PriceSmart"}' > card.json

# Brand and completeness check, no file written
curl -X POST localhost:5000/api/battlecard/validate \
  -H 'Content-Type: application/json' -d @card.json

# Build the deck
curl -X POST localhost:5000/api/battlecard/build \
  -H 'Content-Type: application/json' -d @card.json
```

`build` returns the download URL, the slide count, brand warnings and the Google Slides
compatibility report.

### AttributeSmart cards

`battlecards/attributesmart.py` holds a deep card set for AttributeSmart against
Oracle Retail, Blue Yonder, RELEX Solutions and o9 Solutions. The four rendered
payloads live in `content/battlecards/`, so they version in git and rebuild on
demand with `build_battlecards.py`.

Two rules govern that module. The Impact Analytics side is specific, drawn from
the AttributeSmart NRF 2026 deck and cited to it. The competitor side states only
what public material supports, with a source on every card; where the record is
silent the capability rating is `Unclear` and the seller is told to ask rather
than handed a guess. `test_unsourced_competitor_capabilities_stay_unknown` enforces
that. A card that overstates a rival's gap loses the deal the moment the buyer
corrects it.

**Category note.** AttributeSmart *generates* attributes from images, labels and
text. Those four vendors mostly *consume* attributes for planning, and Oracle is
the only one shipping a directly competing extraction product. The vendors that
actually compete on generation are listed in `DIRECT_RIVALS`: Lily AI, Vue.ai,
Pixyle AI, Syndigo, Salsify and Akeneo.

## Teaching the builder

Public research only reaches so far. What a rival's engineer admitted in a bake off, the
objection that keeps landing, the gap you watched them fail to cover: none of that is
searchable, and it decides deals. `/intel` is where the team puts it, and
`battlecards/intel.py` feeds it to Claude as ground truth on every card from then on.

Write the note however you would say it out loud. Claude splits it into separate claims
and tags each one, you correct it, and nothing is stored until you approve it. Adding a
claim never edits a card that already exists; it changes what the next card is written
from.

### Uploading documents

The documents that matter here are the ones a search cannot reach: a win loss
report, an RFP response, a rival's datasheet handed over in a meeting, a call
transcript. Drop them on `/intel` and `battlecards/docs.py` reads PDF,
PowerPoint, Word, Excel, CSV, text and transcript files, page by page or slide by
slide, so a claim can cite "Their datasheet.pdf, page 4" rather than pointing at
nothing.

The file is parsed in the request that uploaded it and then discarded. Only the
claims somebody approves are stored, so a competitor document under an agreement
never lands on a disk here. Reading it does send its text to the Anthropic API,
which the page says plainly next to the drop zone.

A long document is read in full, in passes of about 24,000 characters, four
passes at a time, streamed back so the page shows progress instead of hanging.
There is no page cap: a 136 page administration guide is read end to end in a
few minutes, and the cost so far is shown as it goes. A file with almost no text is refused with the reason, because a scan
with no text layer is the most likely real failure and the least obvious one.

Claims read from a document default to the `documented` tier, and `intel`
refuses to record a file citation as `verified`: that tier means a link a buyer
could open. Where the document is itself reporting a rumour, "we believe they
will launch X", the claim comes back as `hearsay`.

### Teaching the builder about our own products

The honesty rules forbid a card from claiming any Impact Analytics capability
that is not in its ground truth, and outside AttributeSmart that ground truth
used to be one catalog line per product. So our side of most cards had almost
nothing to say, however much was known about the rival.

Choose **Impact Analytics** on the Teach page, pick the product, and upload the
product deck, the technical documentation or a case study. What is approved is
filed against `Impact Analytics` and that product, and `intel.ia_block()` adds it
to the product's ground truth, so it becomes the mechanism, the advantages, the
proof points and the talk track of every card for that product. Facts about the
other products follow under a fence, for the suite story.

Our own material may also teach about a rival: a row that names one is filed
against that rival, as our assessment, and never as `verified`. The reverse is
blocked. In a rival's document, "we beat Impact Analytics on assortment" is
their claim about us, so every row from a rival's material stays with that rival
whatever the model proposes.

`content/intel/impact-analytics.json` seeds the public facts: each product's
published mechanism, the Briscoe deployment, the published case study results,
the funding, the 2025 Gartner Competitive Landscape inclusion and the named
references. Undated results are recorded with the missing year stated, because
statistics must be from 2025 or later to be a proof point.

### How a taught claim reaches the card

Everything taught about a rival goes into the prompt, not just the claims keyed
to the product being written. A claim taught against another IA product still
describes the same rival, so it appears under a fence that says it may shape the
win theme, the talk track and the questions to ask, but not this card's
capability comparison.

Where a curated card exists, the model is shown it as verified ground truth
while it writes, and the finished card is the **union** of the two: every curated
row survives, and the model's rows are added where they say something new. This
replaced an earlier merge that overwrote eleven sections outright, which meant a
generated card for one of the four curated rivals came out identical to the hand
written one no matter how much the team had taught the builder. Company facts are
the one exception and run the other way: a hand sourced fact outranks anything a
research pass produced, because "Not found" must never beat "Dallas, Texas."

### Confidence decides what a card may do

The tier on a claim is the whole mechanism. `intel.CONFIDENCE` holds the rule for each
one, and that wording goes into the prompt verbatim:

| Tier | Means | What a card may do |
| --- | --- | --- |
| `verified` | A public page, filing or release backs it. The source must be an http link. | State it as fact and cite the source. |
| `documented` | A file we hold backs it, cited by name and page. Not findable online. | State it as fact and cite the document. Never imply the buyer can look it up, and leave it off a customer facing card. |
| `field` | A colleague saw it first hand in a live deal. | Use it, attributed as a field note with the date, never as the competitor's own published claim. Never name the account. |
| `hearsay` | Heard from a buyer, an analyst or the market. Unconfirmed. | Never assert it. It becomes a question the seller asks on the call. |

When Claude cannot tell which tier a note belongs in, it picks `hearsay`. Overstating
confidence is the one mistake that reaches a slide as a false claim.

Two kinds carry a house rule that outranks the tier, in `intel.GUARDED`: a competitor
price is never printed on a slide, and an account is never named unless the claim is
verified with a public source. A claim older than nine months is flagged stale and reaches
the prompt as something to re-confirm rather than a current fact.

### Attaching Render Postgres

`render.yaml` already declares the database and wires `DATABASE_URL` into the web
service, so a deployment created as a **Blueprint** gets it automatically. A
service created by hand in the dashboard never read that file, which is the usual
reason the Teach page reports that nothing can be saved.

Check which situation you are in at `/healthz`. `"library": false` means no
reachable database; `"intel"` counts the claims currently loaded, committed seeds
included. When it says false, `/api/intel/diagnose` says why: it connects once
and reports the host, database, user, and the error, with the credential
stripped out. It is behind the same auth as everything else.

To attach one to an existing service: create a Postgres instance in the Render
dashboard **in the same region as the web service**, copy its *Internal Database
URL*, and add it to the service as `DATABASE_URL`. The service redeploys and
`store.init()` and `intel.init()` create their tables on boot.

A free Render Postgres **expires 30 days after it is created**, with a 14 day
grace period before the data is deleted, and one free instance is allowed per
account. So treat the database as the convenient store and `content/intel/*.json`
as the durable one: claims committed there survive any expiry, need no database,
and are reviewable as a diff. The Teach page's "Download to commit" button writes
that file for you.

### A free database that does not expire

The database does not have to be Render's. `DATABASE_URL` is a plain Postgres
connection string, so any provider works.

| Provider | Free storage | The catch |
| --- | --- | --- |
| **Neon** | 0.5 GB per project, permanent | Compute parks after 5 minutes idle and wakes on the next connection |
| Supabase | 500 MB | The project pauses after a week with no database activity and needs resuming by hand |
| Render | 1 GB | Expires 30 days after creation |

Neon is the one that stays. Create a project, copy the connection string, and set
it as `DATABASE_URL` on the web service. Nothing else changes: the tables are
created on boot either way.

Either endpoint works, pooled or direct. The pooled one suits this app, which
opens a connection per operation and closes it.

Parking the compute is what makes it free, and `store._connect()` is built for
it. A parked endpoint refuses the first connection outright while it wakes, so a
longer timeout would not help and the fix is to try again: three attempts,
backing off, which turns a wake into a slow save rather than a failed one.

Waiting is only right when something is coming up, so the retry is fenced in
two directions. A message that names a wrong password, a missing database or an
unresolvable host raises on the first attempt, because retrying only delays the
news. And only a **write** waits: boot, a read and the page's own status check
each try once, so an unreachable database costs a page load milliseconds rather
than half a minute. `DB_CONNECT_TIMEOUT` raises the per attempt timeout if a
provider is slower still.

TLS is required for any host that is not local. A development URL carries
credentials, so the host is parsed rather than prefix matched, which is how
`postgresql://postgres:pw@127.0.0.1/db` keeps working against a local server
that has no certificate.

### Where it lives

Claims the team adds go to Postgres, beside the card library. Claims committed to
`content/intel/*.json` ship with the app, work with no database, and are read only at
runtime, so anything worth keeping permanently belongs there. `GET /api/intel/export`
returns the database contents in that file shape, so a snapshot can be committed and read
as a diff in git. Retiring a claim is a soft delete: the record of what the team once
believed survives.

If the intel store fails, generation carries on without it. A card matters more than the
extra context.

## The shared library

`battlecards/store.py` keeps every generated card in Postgres: the card JSON, the research
brief behind it, the slide count, how many capabilities came back unverified, and who built
it. The gallery at `/library` lists them newest first with the solution accent on each card,
filters by competitor text or IA product, and opens a drawer with the talk track, the
capability matrix and the proof points. Any card rebuilds to PPTX at a different depth
without regenerating, so a colleague's full technical card becomes your summary in one
click.

Render provisions the database from `render.yaml`:

```yaml
databases:
  - name: ia-battlecards-db
    plan: free
    databaseName: battlecards
    user: battlecards
```

`DATABASE_URL` is wired to it through `fromDatabase`, so a Blueprint apply needs no manual
step. The table is created on first boot, and `init()` is safe to call repeatedly. Render
hands out `postgres://` URLs and psycopg wants `postgresql://`, so the store rewrites the
scheme; a test covers that.

Run the library tests against a real Postgres by setting `DATABASE_URL`. They skip
without one, so the suite passes either way:

```bash
DATABASE_URL=postgresql://localhost/battlecards python3 -m pytest tests -q
```

### Using it from Python

```python
from battlecards import scaffold, normalize, validate, build_presentation, audit

card = normalize(scaffold('Example Rival', 'PriceSmart'))
print(validate(card)['warnings'])
build_presentation(card).save('battlecard.pptx')
print(audit('battlecard.pptx'))
```

## The interface

`static/css/ia.css` is the shared design system both pages use. It implements the brand
guide's visual devices rather than approximating them:

- **Spectral Light for display and headings**, the guide's named web alternate for ABC Otto.
  The serif headline is what separates IA from the usual sans-serif SaaS look, so it carries
  every heading. Inter Tight takes H5 and below, per the hierarchy.
- **The clarifying glass effect.** Impact Blue with the grid pattern forms the base layer,
  and the form panel is a sharp edged glass module that overlaps it and clarifies what sits
  beneath. `backdrop-filter` does the middle blurring layer.
- **Grain** over blue for tactility, generated by an SVG turbulence field so no image asset
  ships for it.
- **The primary grid pattern** at 10 percent opacity, always bleeding off every edge.
- **Solution colours one at a time.** Picking a product sets `--accent`, which drives the
  hero rule, the step marks and the selected depth wash. The gallery is the one place many
  appear together, and there each is a small mark rather than a field of colour.
- Italic single words for emphasis, never whole phrases.

Both pages are verified at 400px, 760px and 1440px with no horizontal overflow, and honour
`prefers-reduced-motion`.

## Brand compliance

The builder encodes the brand guide rather than approximating it.

**Colour.** Only the official palette reaches a slide: Impact Blue `#264CD7`, Off-White
`#F4F4F6`, Black `#1C1B1B`, White, Accent Orange `#FF6F1C`, and Gray 1 through 3. The four
use-case colours map to solutions, and a card uses exactly one of them, since the guide bans
mixing solution colours in a single composition. Picking an IA product picks the solution,
which picks the accent. Accent Orange appears in one place only, the `Gap` rating in the
matrix, which keeps it an accent. `tests/test_battlecard.py` scans the generated XML and fails
on any colour outside this set.

**Typography.** Inter Tight everywhere, set on every run and in the theme font scheme, with
Lato for the footer. Both were confirmed against the house deck, which uses 1357 Inter Tight
runs and 813 Lato runs. ABC Otto is licensed and rarely installed, so the guide specifies
Inter Tight for PPTX. Spectral is available as the serif headline option through
`options.serif_headings`.

**House template devices**, measured from the IA AttributeSmart deck and reproduced here: the
10 by 5.625in canvas, two tone titles that split a Black phrase from an Impact Blue emphasis,
whitespace rather than a rule under the title, white rounded cards with a soft shadow and no
edge stripe, blue pill headers, and stat cards with an inner tinted panel under an italic
label. The Data and Intelligence accent follows the shipping theme at `#BFD1F5`, which
disagrees with `visual-specs.md`; the theme wins.

**Logo.** The bundled logo is the primary horizontal mark, Impact Blue plus Black on
transparency. On Impact Blue fills the builder generates the white variant the guide requires,
recolouring opaque pixels and preserving the shape, the spacing and full opacity. Clear space
and the 0.75in minimum width are respected.

**Pattern.** The cover carries the primary grid pattern, thin white lines at 10 percent opacity
with grain, bleeding off every edge. It ships as one cached PNG rather than hundreds of
hairline shapes.

**Voice.** `battlecards/schema.py` strips em dashes and en dashes on the way in, then flags
sentences that end with a preposition, `FAQ` used as a heading, and statistics dated before
2025. Warnings surface in the UI and in the `validate` response. A proof point may cite an
internal source that names its year, such as a deck title; a card marked
`meta.distribution: customer` still demands a public link.

## Google Slides compatibility

Google Slides imports a subset of OOXML. The builder stays inside it:

- Exact 16:9 canvas, 9144000 by 5143500 EMU, which is the 10 by 5.625in size the IA
  template uses, so a slide pastes into a house deck without rescaling.
- Every slide is drawn on the blank layout with explicit geometry and fills. Nothing inherits
  from a master or a theme.
- Every run names its typeface on the latin, east asian and complex script slots.
- No shrink on overflow. Google Slides ignores PowerPoint's `fontScale`, so
  `battlecards/typeset.py` measures wrapped text and picks a size that fits before the file
  is written.
- The theme names Impact Blue as the hyperlink colour, so linked sources match the text
  around them instead of the office default blue.
- Tables carry explicit cell fills, borders and fonts, and the theme table style is removed so
  Slides cannot substitute its own banding.
- Preset shapes only. No custom geometry, no 3-D, no glow, no reflections. An outer shadow
  is the one effect used, because Google Slides both imports it and exposes it in its own
  editor.
- Bullets use `buChar` with an explicit hanging indent rather than an inherited list style.
- Links are limited to `http` and `https`. A `javascript:` URL never reaches a slide.

`battlecards/compat.py` audits any PPTX against these rules and returns findings. The builder
runs it on every deck, and the result rides along in the build response. To open a deck: upload
the file to Google Drive, then open it with Google Slides.

## Layout

```
battlecards/
  brand.py       palette, type scale, geometry, Google Slides safe drawing primitives
  patterns.py    cached grid and dot overlays
  schema.py      data model, normalisation, validation, brand copy rules
  library.py     IA portfolio, capability and question banks, competitor presets, scaffold
  typeset.py     measuring, fitting, trimming, citations out of body copy
  builder.py     slide builders, pagination, the deck assembler
  compat.py      Google Slides compatibility audit
  ai.py          Claude Opus 5 research, structured card generation, chat assistant
  store.py       the shared Postgres library, optional
  attributesmart.py  deep AttributeSmart card set, IA facts plus sourced rivals
  auth.py        optional HTTP basic auth over every route
  service.py     glue for the Flask routes
content/battlecards/
  *.json            versioned battlecard payloads
build_battlecards.py  build every card from content/
static/css/
  ia.css            the shared design system
templates/
  battlecard.html   the builder UI, the landing page
  library.html      the shared gallery
  index.html        the original standardizer, at /standardizer
tools/
  deckcheck.py      render a deck and report collisions and overflow
tests/
  test_battlecard.py
  test_typeset.py
```
