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

Step 4's "Saved products" list keeps **one card per product**, not one per
save. Re-saving the same product (matched by a "Product Key" — auto-
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
│       ├── xlsx-images.js      # extracts embedded PD-sheet images (raw zip parse)
│       ├── state.js            # in-memory brief state
│       ├── storage.js          # localStorage product history (one record per product)
│       ├── docx-generator.js   # builds the Product Brief .docx
│       ├── xlsx-generators.js  # builds the 3 downstream .xlsx documents
│       └── app.js              # UI wiring (5-step wizard)
├── api/
│   └── gate.js               # password gate (see "Password protection" below)
├── vercel.json                # routes every request through the gate
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
- **Product history is per-browser.** Saved products live in that browser's
  `localStorage`, not a shared database — they won't show up on a
  colleague's machine. Good enough for drafting; not a replacement for the
  Product Brief Shared Drive / ClickUp Masterlist as the system of record.
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
3. **Shared version history** — swap `js/storage.js`'s localStorage calls for
   calls to a small database (or Google Sheets as a lightweight DB) so
   product history is visible across the team, matching the "any team
   member can identify which version is used on which PO" success
   indicator, and resolving the "per-browser" limitation above.
4. **Parse a "PO & Supplier Information" PD sheet section**, once one
   exists in a real sample file (see above) — would let Step 3 auto-fill on
   re-upload the same way Step 2 already does.

The codebase is intentionally modular (`xlsx-parser.js` / `xlsx-images.js` /
`storage.js` / `docx-generator.js` / `xlsx-generators.js` are independent of
the UI) so each of these can be added without a rewrite.
