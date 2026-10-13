# Product Brief Generator

A web app that turns a **Product Development (PD) sheet** into a **Product
Brief** (`.docx`) — plus the three downstream QA/marketing documents that
reuse the same data — in minutes instead of ~4 days. Built for CD Commerce's
R&D → QA workflow.

The app itself is static and does all its work client-side (in the
browser): PD sheets are parsed with [SheetJS](https://sheetjs.com/), the
Product Brief is generated with [docx.js](https://docx.js.org/), and the
downstream spreadsheets are written with SheetJS too. No PD sheet data is
ever uploaded anywhere — this matters given PD sheets contain
supplier/commercial data. The only server-side piece is a small function
that password-protects the whole site (see "Password protection" below) —
it never sees your PD sheets, it just checks a password before handing over
the app's files.

## Which PD sheet layout this targets

This is built around the **current (post-optimized) PD sheet**: one `Final
Product Specifications` tab with our target specs — and a competitor
comparison column — built in from the start, rather than the older two-stage
process (a separate `Initial Product Specifications` tab, finalized later).
Two things follow from that:

- The parser reads each feature row's **first spec column as "ours"** and
  ignores the competitor comparison columns for the Product Brief fields
  (Material, Color, etc.) — those columns are for sourcing decisions, not
  the brief.
- It also pulls two things unique to the post-optimized layout that the old
  layout didn't have: the **"Review Analyses" table** under "Concerns to be
  Discussed with Supplier" (customer complaints about competitor products,
  ranked by mention count, with a recommended spec/action for each) is used
  to pre-fill starter rows in the QC Acceptable/Not Acceptable table, and the
  **Quality Inspection / PO Information** tab becomes per-product QC notes
  that carry into the Pre-QC Check.

Older two-stage PD sheets (LTF, SUP) are still supported as a **legacy
fallback** — the parser detects when the "Final" tab is just an unfilled
template and reads "Initial" instead — but new PD sheets should follow the
post-optimized layout; that's what this tool is optimized for.

## What it does today

1. **Import** — upload a PD sheet `.xlsx`. Reads `Overview`, `Final Product
   Specifications`, and `Quality Inspection/PO Information` — plus any
   embedded product photos on the `Overview` and spec tabs (see "Images"
   below). If a brief is already loaded (opened from history, or already
   parsed once this session), Step 1 shows a note reminding you that
   re-uploading here only refreshes spec-derived fields — anything you've
   entered in Step 3 (supplier, PO, compliance) is never touched by a
   re-import, so you can safely re-upload a PD sheet later to backfill
   updated specs without losing what you've already filled in.
2. **Review & Edit** — pre-fills Article No. (preferring the fuller SKU
   variant list from the spec sheet's component column over the Overview
   tab's shorter base identifier, when both exist), Item, a concise Material
   summary, Color, Inclusions, Material & Workmanship Instructions, Technical
   Specifications, and Packaging. The Typical Production Mistakes/QC table
   auto-seeds starter rows from the PD sheet's review-analysis concerns when
   it's still empty — a first draft, not a finished answer (see "Known
   limitations"). You review and correct — the parser is deliberately not
   "magic": PD sheets still vary enough that a human pass is the right
   safety net, but it removes the retyping/reformatting work.
3. **PO & Compliance** — the fields that only exist at PO time (supplier
   name/contact/email, PO number, compliance test reports, approval
   contacts) plus quality-inspection setup (brand name, sample count,
   inspection dates, product-specific QC notes) that feeds Step 5.
4. **Export & Save** — download a `.docx` Product Brief, save the product to
   its history (see "Product history" below), or export raw data as JSON.
5. **Downstream Documents** — generate, from the *same* brief data:
   - **Pre-inspection Briefing Form** (`.xlsx`)
   - **Pre-QC (Quality Control) Check** (`.xlsx`)
   - **Marketing Guide Sheet** (`.xlsx`) — pre-fills Item/Brand/Description/
     Article No./Material/Size and seeds "Important Features to Highlight"
     from the brief's USPs/ESPs; benefit copy and target-group sections are
     left as clearly marked `TODO` prompts, since that's original
     copywriting a machine shouldn't be inventing on your behalf.

### Product history (one record per product, not per save)

A dedicated **Product History** tab sits in the main navigation (separated
from the 1-5 wizard steps, since it's a standing view you can open anytime,
not a sequential step) — it keeps **one card per product**, not one per
save. Generating a brief five times for the same product updates that one
card and adds five entries to its history, rather than creating five
separate cards. Step 4's "Saved products" list and the "Open saved brief"
modal in the top bar show the same data for convenience mid-workflow. Re-saving the same product (matched by a "Product Key" — auto-
suggested from Article No., editable if you'd rather set it yourself)
updates that product's current brief and appends a lightweight entry to its
own history, instead of creating a new top-level card. Each card shows a
"View history" toggle listing every past save with a one-click **Restore**,
so you can roll the current brief back to an earlier revision without
losing the history itself (restoring is its own new fact, not an erasure).

This is what makes "re-upload the PD sheet to backfill later info" safe in
practice: open a product from history (loads its current brief, PO/supplier
info included), go back to Step 1, upload the newer PD sheet, and only the
spec-derived fields refresh — then save again under the same product key
and it's still one card, just with one more revision in its history.

Versioning follows CD Commerce's own convention loosely (minor bump for
edits/corrections, major bump for a real restructure) — a "Suggest next"
button in Step 4 offers `1.0 → 1.1`-style increments, but never overwrites
what you've typed without your say-so.

### Why `.docx` and not Google Docs directly
Generating a native Google Doc requires a signed-in Google API call (OAuth),
which needs a backend. Plain `.docx` avoids that: it opens with full
formatting in Google Docs (Drive → right-click the file → **Open with →
Google Docs**, or **File → Open** from within Docs), and in Word. This keeps
the app fully static and deployable to GitHub Pages/Vercel with zero backend.
Google Drive/Docs API integration is a natural Phase 3 (see below).

### Document format

The generated Product Brief matches CD Commerce's official template
(`CD_Commerce_-_Product_Brief_Template.docx`): landscape A4, the branded
orange header banner (`site/img/header-banner.jpg`, fetched at export time
and embedded in the page header so it repeats on every page) instead of a
text title, page numbers in the footer, and this section order: Date/PO
Number → To: (supplier) → Article info → Product Individualization (with
per-component sub-rows, split automatically from the Material & Workmanship
Instructions field) → II. Compliance → III. Typical Production Mistakes/QC
(pre-seeded, see above) → sign-off. If the template changes again,
`site/js/docx-generator.js` is the only file that needs to change — swap
`site/img/header-banner.jpg` for a new banner image and adjust the section
order/labels to match.

### Images

PD sheets carry embedded product photos that SheetJS (the library used for
reading cell data) can't see — it doesn't expose embedded drawings. So
`js/xlsx-images.js` reads the `.xlsx` a second way: as a raw zip, walking the
OOXML relationship chain (`workbook.xml` → sheet → its drawing → each
picture) by hand to find every embedded image and which sheet/row it's
anchored to.

- Images on the **Overview** tab (next to "Product Images:") become the
  **hero product image(s)**, shown right under the Article info table.
- Images on the **Final/Initial Product Specifications** tab become
  **reference images**, shown in their own row inside the Product
  Individualization table.
- Every image is downscaled client-side (canvas, max 700px on the long edge,
  JPEG ~72% quality) before use — PD sheet photos are often 1-2MB each at
  full camera resolution, which is both too large to embed cleanly in a
  `.docx` and too much to keep in `localStorage` for saved products.
- Step 1 shows a thumbnail gallery of everything found, with a single
  "Include these images" checkbox (Step 5's exported `.docx` and the
  downstream `.xlsx` docs respect it) — there's no per-image picker; if a
  sheet has images that don't belong in the brief, uncheck the box and skip
  them all for that import.
- This doesn't attempt to recreate every image in a fully custom illustrated
  brief (close-up QC photos, packaging carton diagrams, logo-placement
  mockups) — only what actually exists as embedded pictures in the PD sheet
  itself. Anything beyond that is still a manual step in Word/Google Docs
  after export.

## Project structure

```
product-brief-app/
├── site/                    # the static app
│   ├── index.html
│   ├── css/style.css
│   └── js/
│       ├── xlsx-parser.js      # PD-sheet -> structured data parser
│       ├── xlsx-images.js      # extracts embedded PD-sheet images (raw zip parse), incl. pictures anchored on the Brief Input tab
│       ├── brief-input.js      # maps the Product Brief Input tab onto the brief (components, QC rows, compliance, cartons)
│       ├── defaults.js         # supplier directory + sticky field defaults
│       ├── state.js            # in-memory brief state
│       ├── storage.js          # local (this browser) product history, one record per product
│       ├── shared.js           # syncs that local copy with the shared team database
│       ├── docx-generator.js   # builds the Product Brief .docx
│       ├── xlsx-generators.js  # builds the 3 downstream .xlsx documents
│       └── app.js              # UI wiring (5-step wizard)
├── api/
│   ├── gate.js               # password gate (see "Password protection" below)
│   ├── _auth.js              # the login check, shared by gate.js and store.js
│   └── store.js              # shared version history (see "Shared version history")
├── vercel.json                # routes page requests through the gate (not /api/*)
└── README.md
```

No build step needed to run the app itself — it's plain HTML/CSS/JS loading
SheetJS, JSZip, and docx.js from CDN. The only thing that isn't purely
static is `api/gate.js`, a small Node function that password-protects
everything else.

## Password protection

The whole site sits behind a password (HTTP Basic Auth), enforced
server-side by `api/gate.js` — every request is routed through it via
`vercel.json`'s rewrite rule before any file is served. This works on
Vercel's free **Hobby** plan; you don't need Vercel's paid "Password
Protection" add-on ($20/month/project on Pro) to get real protection, though
that's a fine alternative if you'd rather manage it from the dashboard
instead of an env var.

**Setup (one-time, after the first deploy — see "Deploying" below):**

1. In the Vercel dashboard, open the project → **Settings → Environment
   Variables**.
2. Add two variables (for Production — add to Preview too if you want
   preview deployments protected the same way):
   - `SITE_USER` — e.g. `cdcommerce`
   - `SITE_PASSWORD` — pick something you wouldn't mind sharing over a
     private channel (Slack DM, not a public doc)
3. Redeploy (**Deployments** tab → latest deployment → **⋯ → Redeploy**) so
   the function picks up the new env vars.

Anyone opening the URL now gets their browser's native username/password
prompt before seeing anything.

**What this is and isn't:**
- It's real, server-side protection — unlike a JS password prompt baked into
  the page, there's nothing to bypass by reading page source, since the
  server won't hand over the HTML/JS/CSS at all without valid credentials.
- It's HTTP Basic Auth, so credentials are protected in transit by Vercel's
  HTTPS, but there's no "log out" button — closing the browser or using a
  private/incognito window is the practical equivalent.
- Changing the password means updating the env var and redeploying, not an
  in-app setting.
- If several people need it, share one username/password rather than
  individual logins — this isn't a full auth system, just a door lock.

## Running locally

Just open `site/index.html` directly in a browser — password protection only
applies once deployed to Vercel (that's where the serverless function runs),
so locally you'll see the app unprotected. To test the gate itself locally,
use the Vercel CLI (`npm i -g vercel`, then `vercel dev` from the project
root) with `SITE_USER`/`SITE_PASSWORD` set in a local `.env` file.

## Deploying

### 1. Push to GitHub

```bash
cd product-brief-app
git init
git add .
git commit -m "Initial Product Brief Generator"
git branch -M main
git remote add origin https://github.com/<your-org>/product-brief-generator.git
git push -u origin main
```

### 2. Deploy on Vercel

- Go to [vercel.com/new](https://vercel.com/new), import the GitHub repo.
- Framework preset: **Other** (it's static — no build command needed).
- Root directory: repo root (where `vercel.json` lives).
- Deploy. Vercel will give you a `*.vercel.app` URL immediately.
- Then follow "Password protection" above to lock it down.

Every push to `main` auto-redeploys.

## Marketing Guide Sheet: matching the real structure exactly

Checked two real reference documents directly - SUP and FKT, two
completely different product categories (an inflatable sports product and
a pool chemical) - and found they share an identical section structure and
color convention that the generator wasn't following at all:

- **ITEM / BRAND NAME / DESCRIPTION** at the top, then an **INCLUSIONS**
  table, then **CERTIFICATIONS (Link to the certificates)** and "Other
  Compliance related informations" (both genuinely blank - that data
  doesn't exist in a PD sheet), then a **PRODUCT SPECIFICATIONS** section -
  a second, more detailed restatement of Article No./Material/Color/Size
  plus whatever category-specific specs exist (load capacity, operating
  pressure, coverage, pH, etc. - now pulled from any `techSpecs` entries
  beyond the one used for Size), then Packaging size.
- **IMPORTANT FEATURES TO HIGHLIGHT is split into two sub-groups** in both
  reference documents: "These features are also offered by other
  competitors" and "These features are underutilized. Can be positioned as
  USPs." This maps directly onto data the parser already extracts
  separately - ESPs and USPs - which were previously being flattened
  together into one undifferentiated list, losing exactly the distinction
  the real template is built around.
- **Color convention**, confirmed identical in both files: strong orange
  (`#FF9900`) marks top-level section dividers (ITEM, BRAND NAME, PRODUCT
  SPECIFICATIONS, IMPORTANT FEATURES TO HIGHLIGHT, Weaknesses in our
  product., Target group, Sample Picture); a lighter orange (`#F6B26B`)
  marks sub-field labels within a section. **Green is never used anywhere
  in this document** - the earlier version incorrectly used it for the
  Inclusions/Features headers, borrowed from the Pre-inspection/Pre-QC
  table-header convention, which doesn't apply here.
- **"Weaknesses in our product."** (exact wording in both references, not
  "Known weaknesses / trade-offs...") is left genuinely blank now rather
  than reusing the PD sheet's review-analysis data - that data describes
  *competitor* product complaints (correctly used to seed the Product
  Brief's QC table), not an honest self-assessment of this specific
  design's own limitations, which a PD sheet can't capture.
- Renamed "Sample Pictures" to "Sample Picture" (singular, matching both
  references).

Verified against realistic data for both a post-optimized-sheet product
(dimensions pulled from the Material & Workmanship text fallback) and a
legacy-sheet product (dimensions from an explicit `techSpecs` entry), plus
confirmed zero green fill anywhere in the output and exact-match styling
on every section divider and sub-field across the whole sheet.

## Marketing Guide Sheet: column layout (third reference, BRS)

A third real reference (an office chair, "BRS") confirmed a layout detail the
generator had wrong: everything from ITEM through PRODUCT SPECIFICATIONS puts
its label in **column B**, but from **IMPORTANT FEATURES TO HIGHLIGHT**
onward (both feature sub-groups, Weaknesses, Target group, Sample Picture)
the whole section sits one column to the right, labels and content both
starting in **column C**. The generator now does this, and "Technical
Drawing Link" was replaced by "Artworks/Image/Color" (present in two of the
three references; the other was a one-off). Verified by generating a sheet
and checking each row's first filled column and fill color. Not yet done: a
full side-by-side pass against BRS for fields only it has (Amazon size tier /
FBA fees, shipping carton weight, Planned Sales and Profit, Competitor).

## Marketing Guide Sheet: blank instead of "TODO" placeholders

Every cell with no real source data used to show a repeated "TODO: benefit
copy" / "TODO: expectation-setting note" / "TODO: who is this for?" string
— noisy across a dozen+ rows, and meant deleting filler text before typing
real content. Changed so any field with no PD-sheet-derivable data is
genuinely blank instead: the section headers and structure stay intact
(so the sheet still shows exactly what needs filling in), but nothing
fabricated goes in the actual cells. The one exception is the "Mention all
the product's features..." line, which stays as-is since it's verbatim
instructional text from the reference template, not a generated
placeholder — and the Sample Pictures note when the brief actually has
extracted images, since that's real derived information, not a TODO.

Also fixed while in there: the "Known weaknesses" content row(s) were
missing their border/wrap styling entirely (a leftover gap from the
previous formatting pass) - now styled consistently with every other
content row.

## Marketing Guide Sheet field-mapping fixes

Two content bugs surfaced on a real generated file: **SIZE** was coming out
blank for any PD sheet without an explicit "Technical Specifications"
sub-table (i.e. every current KPM/AKP/SSC-style sheet), because
`findDimensionSpec` only ever looked in `brief.techSpecs` - it never fell
back to the "DIMENSIONS" block that lives inside Material & Workmanship
Instructions on those sheets instead. Fixed with a fallback that pulls it
out of there when `techSpecs` doesn't have it.

Separately, **PACKAGING SIZE** was dumping the *entire* packaging
write-up (carton material, weight, shipping fees, everything) into a field
that the reference document uses for just the carton dimension (e.g. "90 x
40 x 22 cm"). Fixed by extracting just the "Packaging Dimensions — ..."
line out of the full packaging text, falling back to the whole thing only
if that specific line isn't present (so the field is never silently
empty).

## Downstream doc formatting (colors, fonts, borders)

The three downstream `.xlsx` files used to come out completely unstyled —
plain black text, default column widths, no borders — because the library
generating them (`xlsx.full.min.js`, plain SheetJS Community Edition)
**deliberately does not support writing cell styles at all**, by design,
to stay lightweight. This isn't a bug to work around in our own code; it's
a hard wall in that specific library.

Fixed by switching the CDN script to **`xlsx-js-style`**, a maintained fork
with the exact same SheetJS 0.18.5 core and API (confirmed: our existing PD
sheet parsing code needed zero changes and produces identical output) that
adds style-writing on top. Checked the actual formatting (fonts, fills,
borders, column widths, merged cells) of the real reference documents with
openpyxl and matched it exactly:

- Title bars and top-level fields: bold, orange `#FF9900` fill
- Table header rows (Test/s, Methods, Inclusions, etc.): bold, green
  `#00FF00` fill
- Marketing Guide Sheet sub-section headers (Description, Technical
  Specifications, etc.): bold, lighter orange `#F6B26B` fill
- Body text: wrapped, top-aligned, thin borders throughout
- Column widths and cell merges matching the reference files' exact
  proportions (e.g. a wide ~35-58 character label/content column, not
  Excel's cramped ~8.4 character default)

Verified by generating real files with realistic data and checking every
section's styling with an independent strict parser (not just "it opens
without erroring") — title, field rows, table headers, and custom
product-specific test rows all confirmed correct.

## Improving the downstream docs

Compared the actual reference Pre-inspection Briefing Form, Pre-QC Check,
and Marketing Guide Sheet (the real, manually-filled SUP examples) against
what the generators produced. The biggest gap: **the test/check content had
zero editability** — every row was hardcoded in `xlsx-generators.js`, with
no way to add the product-specific tests real inspections need (the SUP
reference form has "Inflation test," "Stability and leak test," "Air
retention test," "Deflation test" — none of which a stunt scooter or a
mattress would ever need, and an inflatable board absolutely does).

Fixed by adding, in Step 3:
- **Additional Pre-inspection Tests** and **Additional Pre-QC Tests** —
  editable add/remove tables (same pattern as the Technical Specifications
  and QC Acceptable/Not Acceptable tables already had). The standard
  baseline checks (package completeness, measurement, visual, packaging/
  printing) stay fixed and automatic; anything product-specific goes here
  and gets appended after that baseline in the generated `.xlsx`.
- **Production Status at Inspection** — the reference forms include a line
  like "70% finished, 20% packed," which is inspection-day status that
  can't come from a PD sheet or exist before the inspection happens. Added
  as its own optional field, since it's a genuinely different category of
  data from everything else in Step 3.

Also enriched the Marketing Guide Sheet with data the brief already has but
wasn't using:
- A full **Technical Specifications** block (every spec row, not just the
  single dimension line) — marketing needs the complete numbers to write
  accurate claims.
- **"Known weaknesses / trade-offs to acknowledge"** now pulls the real
  PD-sheet review-analysis content (the same market complaints that seed
  the Product Brief's QC table) instead of a bare `TODO` placeholder —
  reframed as things not to overclaim rather than QC criteria.
- **Sample Pictures** now points to the Product Brief's actual extracted
  images when present, instead of a blind "attach photos" prompt.

What's still intentionally left as `TODO`: benefit copy per inclusion,
target group, and the "features also offered by competitors" vs.
"differentiators" split the real reference document makes. The real
reference Marketing Guide Sheet runs to dozens of rows of original,
product-specific copywriting (e.g. explaining *why* a rounded nose vs. a
pointed nose matters to different buyers) — that's genuine marketing
judgment a deterministic generator shouldn't fabricate, the same reasoning
documented elsewhere in this README for not inventing persuasive copy.

## Speeding up the process further

Beyond parsing the PD sheet, three things now cut down repeat typing across
briefs:

- **Supplier directory** — saving a brief with a supplier name filled in
  automatically remembers that supplier (name, address, phone, email,
  contact person). Step 3 offers a dropdown ("Used this supplier before?")
  that fills all five fields in one click on the next product from the same
  factory, instead of retyping them. No separate "save supplier" step —
  it learns automatically.
- **Sticky Approval Contacts** — this field is nearly identical across every
  brief in practice (the same 2-3 names), so the last value used pre-fills
  new briefs automatically. Never overwrites something already typed or
  loaded from history.
- **Compliance quick-insert buttons** — one click appends a citation for
  REACH, RoHS, GPSR, PPWR, or POP (the regulations that show up in almost
  every brief seen so far) instead of typing the full regulation name and
  number out. Deliberately doesn't presume a report number or test cost,
  since those genuinely vary per product/supplier and shouldn't be
  auto-filled with a guess.
- **"Download everything" button** — one click bundles the Product Brief
  and all three downstream documents into a single `.zip`, instead of four
  separate download clicks.

The supplier directory and sticky defaults are per-browser (localStorage) and
are not part of the shared history.

## Fourth fix: signature line label running into the underline

The "Supplier" / "Customer - CD Commerce GmbH" labels were appearing
immediately after the underscores on the same line instead of centered
below them. Root cause: a `"\n"` character inside a single text run isn't a
line break in Word — it's silently dropped — so
`p("________________________\nSupplier")` always rendered as one line, not
two. Fixed by using two separate (centered) paragraphs instead of one
string with an embedded newline. Worth knowing for future edits to this
file: any text meant to appear on its own line needs its own `Paragraph`
(or `multiline()`/`numberedLines()`, which already do this correctly) — a
literal `\n` inside a single `p(...)` call will not create a line break.

## Third reverse-engineering pass

Comparing another freshly-generated brief found the "Product Image(s)"
heading could get stranded alone at the bottom of a page, with its actual
photo pushed to the top of the next page — leaving a label with nothing
under it and an orphaned image with nothing above it. Fixed by setting
`keepNext` (a standard docx paragraph property) on that heading and the
other three standalone section headings ("I. PRODUCT INDIVIDUALIZATION",
"II. COMPLIANCE", "III. TYPICAL PRODUCTION MISTAKES/QC"), so Word/Google
Docs always push a heading and what comes right after it to the next page
together rather than splitting them.

## Second reverse-engineering pass (after the fixes above shipped)

Comparing a freshly-generated brief against the manual one again turned up
three more polish items, all fixed:

- **Packaging** wasn't numbered like the rest of the Material & Workmanship
  breakdown — now uses the same auto-numbered list format.
- **The "To:" block** showed dangling labels ("Tel/Fax:" with nothing after
  it) and blank gap rows when Step 3 hadn't been filled in yet. Empty
  optional lines are now skipped entirely rather than rendered blank.
- **Page orientation metadata** said "portrait" even though the page
  rendered as landscape (harmless visually, but incorrect for strict
  readers). Fixed — note that `docx.js` swaps width/height internally
  whenever you set `orientation: LANDSCAPE`, so the values passed into
  `size` have to be pre-swapped to end up correct; this tripped up the fix
  on the first attempt and is called out in a comment in
  `docx-generator.js` so it doesn't regress.

A note on testing hygiene: validating the orientation fix involved
reinstalling the `docx` npm package for local testing, which silently
pulled the *latest* version instead of the `8.5.0` pinned in the app's CDN
link — producing files that failed to open in a strict parser. That was a
test-environment artifact, not a real bug, but it's a reminder that any
future local testing against this codebase must pin `docx@8.5.0` exactly
(matching the CDN URL in `index.html`) or risk false failures.

## Product Brief Input tab (the brief is built from the PD sheet)

If the PD sheet has a tab whose name contains **"Brief Input"** (currently
*Product Brief Input*), the tool builds the Product Brief from it instead of
making the team retype it in Steps 2 and 3. Everything is found by its
headings and column titles, not fixed row numbers, so extra rows are fine —
but **keep these headings and column titles**: `PART A`, `SKU TABLE`,
`PART B`, `PART C`, and the column headers (SKU / Color / Item carton…;
Component / Instruction / SKU-specific detail / Image / Image caption /
Applies to / Status; Parameter / Acceptable / Not acceptable / Defect photo /
Test method).

**How each part is used**
- **Part A** → Item, Description (the technical one), Material, Inclusions
  (a leading "1 x " is dropped), the compliance text (test-cost clause with
  amounts in bold, requirement list, standard warranty sentence, test report
  number). *Product category* is stored but doesn't drive anything yet.
- **SKU table** → Article No. (one SKU per line), Color lines as
  `<SKU suffix> - <colour>` (SSC01-SW + Black → "SW - Black"), and the carton
  sizes the approved brief states under Packaging: an "individual carton
  dimensions … weight" line is added, and a master-carton row that says
  "(dimensions and weight per SKU table)" gets the real numbers. If SKUs have
  different cartons, the lines become per-SKU a./b./c. entries.
- **Part B** → one row of the Product Individualization table per component,
  each instruction numbered (restarting at 1 per component). *SKU-specific
  detail* prints under its instruction — lettered a./b./c. if the instruction
  ends with ":", plain lines otherwise. The picture anchored in the **Image**
  cell prints directly under that instruction, with its caption (Step 2 has a
  switch to hide captions). A component named exactly **Packaging** is the
  Packaging row.
- **Yellow highlighting**: any sentence containing "sent separately" (the
  approved briefs highlight these), and every row whose Status is **Pending**.
- **Part C** → the QC table. The sheet says default rows **Odor, Finish,
  Printing, Packaging** are added automatically but doesn't give their
  wording, so the tool uses neutral placeholder wording of its own (not copied
  from any brief). A Part C row about the same thing overrides the default and
  keeps its own position; defaults not covered are added (Odor first, the rest
  last). Defect photos print under *Not acceptable*.
- **Compliance** is built only from what the sheet says (test-cost clause,
  requirements, test report number). No standard wording, pictures or contact
  names are added by the tool — anything a brief needs beyond that must be in
  the sheet or typed in Step 3. The box understands `**bold**`.

**Nothing from any sample brief is built into the tool.** The approved SSC
brief was used only as a guide to layout, formatting and structure. The tool
never reads content from a sample brief; every fact, name, number and sentence
in a generated brief comes from the uploaded PD sheet or Step 3. The only
fixed wording is structural: the section headings, the sign-off / alterations
paragraphs that were already in the template, and two sentences *generated
from the SKU table's numbers* ("The individual carton dimensions should be …").
- Because the brief now carries its own pictures, the PD sheet's product /
  reference gallery is **off by default** when this tab is present.

**Other changes made to match the approved brief**: Arial 10 pt; plain (not
bold or shaded) labels; ✓ / ✗ icons in the QC header; QC rows never split
across pages; the "Alterations… directly to" sentence and its contacts in one
paragraph (names come from Step 3 — nothing is pre-filled); and the sign-off "shown on page 1-N" now uses the
document's real page count (it was hard-coded to 1-22).

**Not in the PD sheet, so still typed in Step 3**: supplier name / address /
contact / phone / email, PO number, date.

**Known limits**
- Wording is printed exactly as typed. Differences from an approved brief
  ("must" vs "should", "Aluminium" vs "Aluminum", a colour name) are
  differences in the sheet, not something the tool can know.
- A "SKU-specific detail" line is printed as typed — including layout hints
  like "(top to bottom)". Put hints like that in the caption column (and turn
  captions off) or leave them out.
- Pictures are only as sharp as the copies in the sheet (Google exports them
  as small thumbnails, ≤ 360 px here). Pictures with a transparent background
  can arrive with a **black** background from Google's export — the SSC
  bar-end-plug picture does; insert such pictures flattened (JPEG).
- Pictures are kept on the current brief only (saved history snapshots and
  the shared team history hold the text, not the pictures); restoring an
  older version reattaches pictures wherever the instruction is unchanged.
- Re-importing the sheet replaces components and QC rows with the sheet's
  version (Step 2 edits to those are overwritten) but never overwrites a
  Compliance text you edited.
- The default QC rows (Odor, Finish, Printing, Packaging) use neutral
  placeholder wording. Add a row of the same kind to Part C — or give us the
  official default wording — to use yours.

## Shared version history (everyone on the team sees the same products)

Saved products and their version history can live in a small shared database
instead of each person's own browser, so anyone who logs in sees the same
cards, the same revisions, and **who saved each one**. Until the database is
connected the app behaves exactly as before (history stays in your own
browser), so nothing breaks if you skip this.

### One-time setup (about 5 minutes, no code)

1. In Vercel, open this project and go to the **Storage** tab (the exact
   wording and layout of Vercel's dashboard changes from time to time; what
   you want is a **Redis** database from **Upstash**, available through
   Vercel's Marketplace on the free plan).
2. Create the database and **connect it to this project** for all
   environments. Vercel adds the connection settings for you.
3. Check **Settings → Environment Variables**: you should see either
   `KV_REST_API_URL` + `KV_REST_API_TOKEN`, or `UPSTASH_REDIS_REST_URL` +
   `UPSTASH_REDIS_REST_TOKEN` (the app accepts either pair). You don't need
   to copy anything anywhere.
4. **Deployments → Redeploy** (new environment variables only take effect on
   a fresh deployment).
5. Open the app. On **Product History** (or Step 4) the status line should
   turn green: *"Shared with your team."* If anything saved on your browser
   before this, a button appears — **"Share N products saved only on this
   browser"** — click it once to add them to the shared history.

### Using it

- Step 4 has a **Your name** box. It's remembered on that browser and shown
  next to each save ("by Leah"), so the team can see who changed what.
- **Save this product** saves locally and uploads in one go. If you're
  offline, the save is kept and uploads itself on the next sync.
- The app syncs when it opens, when you open Step 4 / Product History / the
  "Open saved brief" window, when you click **Refresh**, and when you return
  to the tab after a minute or more. It doesn't poll constantly, which keeps
  database usage small (an idle refresh is a single database command).
- **Two people saving the same product at the same time** both keep their
  saves — each saved version is stored as its own entry, so one can't
  overwrite the other.
- **Delete is for everyone.** The confirmation says so when shared history is
  on. Anything saved but not yet shared is never wiped by a sync.
- **What each product keeps:** its latest 25 saves (the same limit as before).
  The "current" version of a product is its newest save.
- **Not shared:** product photos (re-import the PD sheet to bring them back
  on your own computer), the supplier directory, and remembered defaults.

### How it works (for whoever maintains this)

- `api/store.js` is a small serverless function that talks to Upstash Redis
  over plain HTTPS (no npm packages, so the project stays build-free). It
  uses the same Basic Auth login as the site (`api/_auth.js`).
- Each saved version is its own field in a Redis hash (`pbg:p:<key>`), added
  with `HSETNX` so it can never overwrite another. `pbg:o:<key>` orders them
  to trim the oldest beyond 25, and `pbg:updated` records when each product
  last changed (`ZADD GT`, which only moves forward). A first version of this
  stored each product as one blob and silently lost a save when two people
  saved at the same instant — caught by a multi-browser test and replaced
  with this layout.
- `js/shared.js` merges the database's history into the browser's local copy
  by entry id (nothing is lost on either side) and retries failed uploads.
- Each save is capped at 400 KB (real briefs are roughly 3-40 KB without
  photos).

### Honest caveats

- This was built and tested against a **simulated** Upstash database, not a
  live one: the API, the sync logic across several simulated browsers
  (simultaneous saves, offline saves, deletion, old pre-sharing data), and
  the on-screen flow were all exercised, but no one has yet connected it to a
  real Vercel + Upstash account from this side. If the status line stays
  yellow or red after setup, the message on it says why; the usual causes are
  forgetting to redeploy, or the variables being attached to only some
  environments.
- Upstash's free plan has monthly usage limits and Vercel's integration terms
  can change; check their current pricing page. This app keeps its use low,
  but a very large team saving constantly could outgrow a free plan.
- Anyone with the site login can read, add to, and delete any product. There
  are no per-person permissions, and the "Your name" label is typed in, not
  verified.
- History order uses each computer's clock, so a badly wrong clock could
  place a save out of order.

## Reverse-engineering pass against a real manually-written brief

A side-by-side comparison against a hand-written brief for the same product
(SSC Stunt Scooter) — rendering both to images and checking page by page —
found two real bugs and a handful of formatting gaps, now fixed:

- **Article No. and Color were silently losing data.** Both were single-line
  `<input>` fields, but their values routinely contain multiple SKU/color
  variants joined with newlines (e.g. "SSC01-SW\nSSC01-SG\nSSC01-SR").
  Browsers strip embedded newlines from single-line inputs, so the exported
  brief showed "SSC01-SWSSC01-SGSSC01-SR" all run together. Fixed by making
  both `<textarea>` fields, same as Material.
- **A stray "Feature | DESCRIPTION" row was leaking into every brief.** The
  PD sheet's own literal column-header row for its component table
  (`COMPONENT | FEATURE | DESCRIPTION`) was being misparsed as if it were a
  real data row. Fixed with an explicit skip in `xlsx-parser.js`.
- **Formatting now matches the reference template exactly**: the header
  banner is full-bleed (page-anchored, edge-to-edge) rather than sitting
  inside the content margins; the "To:" block is borderless plain text
  rather than a bordered table; table borders are solid black (not gray);
  label-cell shading is barely-there off-white (not a strong gray); "I."
  was added to "PRODUCT INDIVIDUALIZATION" (II. and III. were already
  correct); and Product Inclusions plus the Material & Workmanship
  breakdown now render as real auto-numbered lists (1. 2. 3., restarting at
  1 for each row) instead of plain unnumbered lines.

What the comparison also confirmed **isn't a tool gap**: the hand-written
brief includes an engineering/negotiation judgment call not in any PD sheet
("load capacity is 120kg, but for compliance purposes claim 100kg"),
specific real product photos placed at exact points in a hand-written
15-point manufacturing spec, and fully custom compliance content (lab
costs, specific regulations, carton label mockups). None of that exists in
a PD sheet at any stage — it's downstream engineering/legal/QA work that
happens after the brief is drafted, not something a parser can extract.

## Known limitations (be aware of these)

- **PD sheets still need a human review pass.** Even on the post-optimized
  layout, wording and column position vary a little sheet to sheet — Step 2
  is a real review step, not a formality.
- **Technical Specifications table** only auto-populates when the sheet has
  an explicit "Technical Specifications" sub-table (LTF/SUP legacy style
  does; KPM/AKP/SSC-style post-optimized sheets fold specs into the Features
  list instead). For those, add spec rows manually in Step 2 — the raw
  "Detected sections" panel in Step 1 has everything you need to copy from.
- **The QC Acceptable/Not Acceptable table is a starter draft, not a
  finished answer.** It's auto-seeded from the PD sheet's review-analysis
  concerns when the table is still empty, but the wording is mechanically
  derived (issue → "Not Acceptable", recommended fix → "Acceptable") rather
  than editorially rewritten — check and tighten the wording before
  exporting, the same way you'd review any auto-filled field.
- **Component grouping in Material & Workmanship Instructions follows the PD
  sheet's own field names** (Material, Dimensions, Color, Features), not a
  hand-picked grouping by physical part (e.g. "Deck & Grip Tape", "Wheels &
  Bearings") the way a person might write it. Getting genuinely human-style
  reorganization would need an LLM rewrite step, which means a backend
  (see Phase 3) — out of reach for a static, no-backend site. The content is
  the same either way; only the editorial grouping differs, and it's fully
  editable in Step 2 either way.
- **Marketing Guide Sheet is a scaffold, not a finished document.** Benefit
  copy, target group, and photos are original writing/assets that stay as
  `TODO` prompts — auto-generating persuasive marketing copy from a spec
  sheet would produce generic, unreliable text, which is worse than an
  honest blank.
- **Shared history needs a one-time database setup.** Until a database is
  connected in Vercel (see "Shared version history"), product history stays
  in each person's own browser (`localStorage`), exactly as it used to. It's
  still a drafting aid, not a replacement for the Product Brief Shared Drive
  / ClickUp Masterlist as the system of record.
- **Product photos aren't shared.** They're large and come straight from the
  PD sheet, so only the text of each saved version goes to the shared
  history. Re-import the PD sheet to bring photos back on your computer.
- **Product Key matching is a heuristic, not guaranteed unique.** It
  defaults to the first SKU in Article No. (or the Item name if that's
  blank). If a PD sheet's SKU format changes between uploads for what's
  really the same product, the auto-suggested key might not match the
  existing card — check Step 4's "Product Key" field before saving if
  you're unsure, and edit it to match an existing product's key if needed.

## What the PD sheet can't give the tool, no matter how good the parser gets

Comparing hand-written briefs against what the app produces from the same
PD sheet (done for the SUP and SSC products) turned up a consistent
pattern: every remaining gap is missing *source data*, not a parsing
limitation. None of the following exist anywhere in a PD sheet, across
every layout seen so far — they only show up once someone is further down
the pipeline (supplier picked, PO cut, compliance testing underway):

- Supplier identity: company name, address, contact person, phone, email
- PO number, PO date
- Compliance test report numbers, the specific regulations tested against,
  and their costs (these come from a lab, after testing — a PD sheet is
  built before a supplier is even chosen)
- Sample count, Pre-QC date, 3rd-party inspection date
- The Acceptable/Not Acceptable wording in the QC table in its final,
  polished form (the PD sheet's review-analysis table gets you a rough
  draft — see "Known limitations" above — but not the finished phrasing)

One existing PD sheet tab is already named **"Quality Inspection & PO
Info"** — the "PO Info" half currently goes unused. If that tab (or a new
one) got a structured section filled in once the PO/supplier/compliance
info exists later in the pipeline, the tool could auto-fill all of Step 3
on the next PD sheet re-upload, e.g.:

```
PO & SUPPLIER INFORMATION
PO Number:              PO Date:
Supplier Company Name:
Supplier Address:
Supplier Contact Person: Email: Phone:
Brand Name:  Sample Count:  Pre-QC Date:  3rd-Party Inspection Date:

COMPLIANCE
Test Report No.:  Regulations tested:
Approval Contacts:
```

This would need a matching parser update in `xlsx-parser.js` to actually
read it — not built speculatively against a layout that doesn't exist yet,
but straightforward to add once a real sample sheet with this section
exists to test against.

## Phase 3 ideas (not built yet)

1. **Google Drive / Docs API integration** — instead of a local `.docx`
   download, write directly into the `01 Product Brief by SKU` folder
   structure, following the naming convention (`Product Brief_[SKU]_
   [Description]_v[Major].[Minor]`) and versioning rules already defined in
   the Quality Process Optimization brief. Requires a small backend (a
   Vercel serverless function) to hold Google service-account credentials —
   the static-site architecture here can absorb that without a rewrite.
2. **ClickUp sync** — push new/updated briefs as tasks in the Product Brief
   Masterlist automatically (ClickUp API), replacing the manual step.
3. ~~Shared version history~~ — built, see "Shared version history".
4. **Parse a "PO & Supplier Information" PD sheet section**, once one
   exists in a real sample file (see above) — would let Step 3 auto-fill on
   re-upload the same way Step 2 already does.

The codebase is intentionally modular (`xlsx-parser.js` / `xlsx-images.js` /
`storage.js` / `docx-generator.js` / `xlsx-generators.js` are independent of
the UI) so each of these can be added without a rewrite.
