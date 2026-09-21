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
  ranked by mention count, with a recommended spec/action for each) becomes
  "Known Market Complaints & Preventive Actions" in the brief, and the
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
   below).
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
4. **Export & Save** — download a `.docx` Product Brief, save a version
   locally (browser storage, keyed by SKU + version), or export raw data as
   JSON.
5. **Downstream Documents** — generate, from the *same* brief data:
   - **Pre-inspection Briefing Form** (`.xlsx`)
   - **Pre-QC (Quality Control) Check** (`.xlsx`)
   - **Marketing Guide Sheet** (`.xlsx`) — pre-fills Item/Brand/Description/
     Article No./Material/Size and seeds "Important Features to Highlight"
     from the brief's USPs/ESPs; benefit copy and target-group sections are
     left as clearly marked `TODO` prompts, since that's original
     copywriting a machine shouldn't be inventing on your behalf.

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
  `.docx` and too much to keep in `localStorage` for saved versions.
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

### Why `.docx` and not Google Docs directly
Generating a native Google Doc requires a signed-in Google API call (OAuth),
which needs a backend. Plain `.docx` avoids that: it opens with full
formatting in Google Docs (Drive → right-click the file → **Open with →
Google Docs**, or **File → Open** from within Docs), and in Word. This keeps
the app fully static and deployable to GitHub Pages/Vercel with zero backend.
Google Drive/Docs API integration is a natural Phase 2 (see below).

### Document format

The generated Product Brief matches CD Commerce's official template
(`CD_Commerce_-_Product_Brief_Template.docx`): landscape A4, the branded
orange header banner (`site/img/header-banner.jpg`, fetched at export time
and embedded in the page header so it repeats on every page) instead of a
text title, page numbers in the footer, and this section order: Date/PO
Number → To: (supplier) → Article info → Product Individualization (with
per-component sub-rows, split automatically from the Material & Workmanship
Instructions field) → Known Market Complaints / Quality Inspection Notes →
II. Compliance → III. Typical Production Mistakes/QC → sign-off. If the
template changes again, `site/js/docx-generator.js` is the only file that
needs to change — swap `site/img/header-banner.jpg` for a new banner image
and adjust the section order/labels to match.

## Project structure

```
product-brief-app/
├── site/                    # the static app (unchanged from before)
│   ├── index.html
│   ├── css/style.css
│   └── js/
│       ├── xlsx-parser.js      # PD-sheet -> structured data parser
│       ├── xlsx-images.js      # extracts embedded PD-sheet images (raw zip parse)
│       ├── state.js            # in-memory brief state
│       ├── storage.js          # localStorage version tracking
│       ├── docx-generator.js   # builds the Product Brief .docx
│       ├── xlsx-generators.js  # builds the 3 downstream .xlsx documents
│       └── app.js              # UI wiring (5-step wizard)
├── api/
│   └── gate.js               # password gate (see "Password protection" below)
├── vercel.json                # routes every request through the gate
└── README.md
```

No build step needed to run the app itself — it's plain HTML/CSS/JS loading
SheetJS and docx.js from CDN. The only thing that isn't purely static is
`api/gate.js`, a small Node function that password-protects everything else.

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
- Root directory: repo root (where `index.html` lives).
- Deploy. Vercel will give you a `*.vercel.app` URL immediately; add a custom
  domain if you want one under CD Commerce's domain.

Every push to `main` auto-redeploys.

## Known limitations (be aware of these)

- **PD sheets still need a human review pass.** Even on the post-optimized
  layout, wording and column position vary a little sheet to sheet — Step 2
  is a real review step, not a formality.
- **Technical Specifications table** only auto-populates when the sheet has
  an explicit "Technical Specifications" sub-table (LTF/SUP legacy style
  does; KPM/AKP-style post-optimized sheets fold specs into the Features
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
- **Versioning is per-browser.** Saved versions live in that browser's
  `localStorage`, not a shared database — they won't show up on a
  colleague's machine. Good enough for drafting; not a replacement for the
  Product Brief Shared Drive / ClickUp Masterlist as the system of record.

## What the PD sheet can't give the tool, no matter how good the parser gets

Comparing a hand-written brief against what the app produces from the same
PD sheet (done for the SUP and SSC products) turned up a consistent pattern:
every remaining gap is missing *source data*, not a parsing limitation.
None of the following exist anywhere in a PD sheet, across every layout
seen so far — they only show up once someone is further down the pipeline
(supplier picked, PO cut, compliance testing underway):

- Supplier identity: company name, address, contact person, phone, **email**
- PO number, PO date
- Compliance test report numbers, the specific regulations tested against,
  and their costs (these come from a lab, after testing — a PD sheet is
  built before a supplier is even chosen)
- Sample count, Pre-QC date, 3rd-party inspection date
- The Acceptable/Not Acceptable wording in the QC table in its final,
  polished form (the PD sheet's review-analysis table gets you a rough
  draft — see "Known limitations" above — but not the finished phrasing)

If CD Commerce wanted the tool to auto-fill more of this, the realistic
options are: (a) accept Step 3 as a short manual step (current design), or
(b) pull from a different source entirely once one exists — a PO record, a
supplier database, or a lab-testing tracker — rather than expecting the PD
sheet to carry data it was never meant to hold.

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
   versions are visible across the team, matching the "any team member can
   identify which version is used on which PO" success indicator.

The codebase is intentionally modular (`xlsx-parser.js` / `storage.js` /
`docx-generator.js` / `xlsx-generators.js` are independent of the UI) so
each of these can be added without a rewrite.
