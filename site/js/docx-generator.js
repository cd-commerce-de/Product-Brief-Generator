/* docx-generator.js
 * Builds a .docx matching CD Commerce's official Product Brief template
 * (CD_Commerce_-_Product_Brief_Template.docx): landscape A4, a full-width
 * orange branded header banner instead of a text title, "Date / PO Number"
 * as the first line, a "To:" supplier block, the Article info table, a
 * unified Product Individualization table (with per-component sub-rows),
 * "II. COMPLIANCE" before "III. TYPICAL PRODUCTION MISTAKES/QC", and sign-off.
 *
 * Uses the `docx` library (loaded globally via CDN). Plain .docx opens
 * cleanly in Google Docs (File > Open > Upload, or right-click > Open with
 * Google Docs in Drive) - no special export settings needed.
 *
 * IMPORTANT width note: docx.js's Table defaults `columnWidths` to 100 DXA
 * per column (about 1.7mm) if it isn't set explicitly - individual cell
 * `width` values are NOT enough on their own to size columns; the table's
 * own `columnWidths` grid is what most renderers (incl. Google Docs)
 * actually use. Every table below sets `columnWidths` in DXA (twentieths of
 * a point) explicitly for this reason.
 */
(function (global) {
  const {
    Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
    HeadingLevel, WidthType, BorderStyle, VerticalAlign, Header, Footer,
    AlignmentType, ImageRun, PageNumber, HorizontalPositionRelativeFrom,
    VerticalPositionRelativeFrom, TextWrappingType, LevelFormat, PageOrientation,
  } = docx;

  // Body text in the approved briefs is Arial 10pt (size is in half-points).
  const FONT = "Arial";
  const BODY_SIZE = 20;

  // Reference template renders Product Inclusions and the Material &
  // Workmanship Instructions breakdown as auto-numbered lists (1. 2. 3.),
  // not plain unnumbered lines - this id ties paragraphs below to the
  // numbering definition registered on the Document itself.
  const NUMBERING_REF = "pb-numbered-list";

  // Landscape A4, 1in (1440 twip) margins, matching the reference template.
  const PAGE_WIDTH = 16834;
  const PAGE_HEIGHT = 11909;
  const MARGIN = 1440;
  const CONTENT_WIDTH = PAGE_WIDTH - MARGIN * 2; // 13954 DXA

  function dxa(pct) {
    return Math.round((CONTENT_WIDTH * pct) / 100);
  }

  // Matches the reference template's actual table borders exactly: solid
  // black, 1pt (size is in eighths of a point, so 8 = 1pt).
  const BORDER = {
    top: { style: BorderStyle.SINGLE, size: 8, color: "000000" },
    bottom: { style: BorderStyle.SINGLE, size: 8, color: "000000" },
    left: { style: BorderStyle.SINGLE, size: 8, color: "000000" },
    right: { style: BorderStyle.SINGLE, size: 8, color: "000000" },
  };
  const NO_BORDER = {
    top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
    bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
    left: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
    right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  };

  function p(text, opts = {}) {
    return new Paragraph({
      children: [new TextRun({ text: text || "", bold: !!opts.bold, size: opts.size || BODY_SIZE, font: FONT })],
      spacing: { after: opts.after ?? 120 },
      heading: opts.heading,
      alignment: opts.align,
      // Stops a section heading from being stranded alone at the bottom of
      // a page while its content (image, table, paragraph) gets pushed to
      // the next one - Word/LibreOffice push the whole pair together instead.
      keepNext: opts.keepNext,
    });
  }

  function multiline(text, opts = {}) {
    const lines = (text || "").split("\n");
    return lines.map((line) => p(line, opts));
  }

  // Same idea as multiline(), but each non-empty line becomes an
  // auto-numbered (1. 2. 3.) list item instead of a plain paragraph.
  //
  // Each call gets its OWN numbering definition (registered into
  // `numberingConfigs`, collected per-document and passed to `new
  // Document()` at the end) rather than sharing one global reference -
  // Word/docx.js numbers all paragraphs under the same reference as one
  // continuous sequence, so sharing a reference across table rows would
  // count 1-2-3-4-5-6... straight through instead of restarting at 1 for
  // each row, the way the reference template's per-row lists do.
  function makeNumberedLinesFactory() {
    const numberingConfigs = [];
    let counter = 0;

    // A fresh numbering definition per list so every list restarts at 1.
    // Level 0 is "1." and level 1 is "a." - the same two levels, with the
    // same indents, as the numbered lists in the approved briefs.
    function newListRef() {
      const reference = `${NUMBERING_REF}-${counter++}`;
      numberingConfigs.push({
        reference,
        levels: [
          {
            level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.START,
            style: { paragraph: { indent: { left: 720, hanging: 360 } } },
          },
          {
            level: 1, format: LevelFormat.LOWER_LETTER, text: "%2.", alignment: AlignmentType.START,
            style: { paragraph: { indent: { left: 1440, hanging: 360 } } },
          },
        ],
      });
      return reference;
    }

    function numberedLines(text, opts = {}) {
      const lines = (text || "").split("\n").filter((l) => l.trim());
      if (!lines.length) return [];
      const reference = newListRef();
      return lines.map((line) => new Paragraph({
        children: [new TextRun({ text: line, bold: !!opts.bold, size: opts.size || BODY_SIZE, font: FONT })],
        spacing: { after: opts.after ?? 40, line: 240 },
        numbering: { reference, level: 0 },
      }));
    }
    return { numberedLines, newListRef, numberingConfigs };
  }

  function cell(children, widthDxa, opts = {}) {
    return new TableCell({
      borders: opts.noBorder ? NO_BORDER : BORDER,
      width: { size: widthDxa, type: WidthType.DXA },
      columnSpan: opts.columnSpan,
      verticalAlign: VerticalAlign.TOP,
      children: Array.isArray(children) ? children : [children],
      // The reference template uses only a barely-there off-white (f8f8f8)
      // on label cells, not a strong gray - matching that exactly rather
      // than the heavier shading used earlier.
      shading: opts.shaded ? { fill: "F8F8F8" } : undefined,
      margins: opts.noMargins ? undefined : { top: 100, bottom: 100, left: 100, right: 100 },
    });
  }

  function table(columnWidthsDxa, rows) {
    return new Table({
      width: { size: CONTENT_WIDTH, type: WidthType.DXA },
      columnWidths: columnWidthsDxa,
      rows,
    });
  }

  // Two-column "label | value" table (header block, Article info). Pass
  // `plain: true` for a borderless, unshaded version (used for "To:") -
  // the reference template renders that block as plain indented text, not
  // a bordered table; a borderless table with the same column widths gives
  // the same visual result without needing real paragraph indentation.
  function labelValueTable(pairs, labelPct = 25, plain = false) {
    const labelW = dxa(labelPct);
    const valueW = CONTENT_WIDTH - labelW;
    const rows = pairs.map(
      ([label, value, opts = {}]) => new TableRow({
        children: [
          cell(p(label, { bold: plain }), labelW, plain ? { noBorder: true } : {}),
          cell(Array.isArray(value) ? value : multiline(value, { after: 20 }), valueW, { ...opts, noBorder: plain }),
        ],
      })
    );
    return table([labelW, valueW], rows);
  }

  // Splits the (already user-edited, plain-text) Material & Workmanship
  // Instructions field back into per-component rows for the table, using
  // the same "LABEL\ncontent" block shape the parser produces (blocks
  // separated by a blank line). Mirrors the reference template's per-
  // component rows (e.g. "Mulled wine kettle" | details, "Drip tray..." |
  // details) instead of one large undifferentiated cell.
  function splitIntoSubRows(text) {
    const blocks = (text || "").split(/\n\s*\n/).map((b) => b.trim()).filter(Boolean);
    return blocks.map((block) => {
      const lines = block.split("\n");
      const first = lines[0].trim();
      const rest = lines.slice(1).join("\n").trim();
      const looksLikeLabel = first.length > 0 && first.length <= 60;
      if (looksLikeLabel && lines.length > 1) {
        return [toTitleCase(first), rest];
      }
      return ["Notes", block];
    });
  }

  function toTitleCase(s) {
    if (s === s.toUpperCase() && /[A-Z]/.test(s)) {
      return s.replace(/\w\S*/g, (w) => w.charAt(0).toUpperCase() + w.slice(1).toLowerCase());
    }
    return s;
  }

  function dividerRow(label) {
    return new TableRow({
      children: [cell(p(label, { bold: true }), CONTENT_WIDTH, { columnSpan: 2 })],
    });
  }

  // Lays out a set of already-downscaled images (dataUrl + natural width/
  // height from xlsx-images.js) inline at a fixed display width, preserving
  // each one's own aspect ratio. Docx.js wraps multiple inline ImageRuns in
  // one paragraph the same way inline text wraps, so this reads as a simple
  // photo grid without needing a nested table.
  function imageParagraphs(images, targetWidth) {
    if (!images || !images.length) return [];
    const children = [];
    images.forEach((im, i) => {
      const h = Math.max(1, Math.round(targetWidth * (im.height / im.width)));
      children.push(new ImageRun({ data: im.dataUrl, transformation: { width: targetWidth, height: h } }));
      if (i < images.length - 1) children.push(new TextRun({ text: "   " }));
    });
    return [new Paragraph({ children, spacing: { after: 150 } })];
  }

  // ---------- rich text helpers ----------
  function run(text, o = {}) {
    return new TextRun({
      text, font: FONT, size: o.size || BODY_SIZE, bold: !!o.bold, color: o.color,
      highlight: o.highlight ? "yellow" : undefined,
    });
  }

  // "**bold**" markers inside a line -> alternating plain/bold runs.
  function inlineRuns(text, base = {}) {
    return String(text || "").split("**").map((part, i) => run(part, { ...base, bold: base.bold || i % 2 === 1 }));
  }

  // One image per paragraph, left-aligned, at the picture's own pixel size
  // (capped) - the approved briefs stack several pictures in one cell.
  function pictureParagraph(im, o = {}) {
    const maxW = o.maxWidth || 480;
    const w = Math.max(1, Math.min(im.width, maxW));
    const h = Math.max(1, Math.round(w * (im.height / im.width)));
    return new Paragraph({
      children: [new ImageRun({ data: im.dataUrl || im.data, transformation: { width: w, height: h } })],
      spacing: { before: 60, after: o.after ?? 80 },
      indent: o.indent !== undefined ? { left: o.indent } : undefined,
      alignment: o.align,
      keepNext: o.keepNext,
    });
  }

  // ---------- Product Individualization table ----------
  // One row per component. Each instruction is a numbered line (restarting
  // at 1 in every component); "SKU-specific detail" goes underneath as a./b./c.
  // when the instruction ends with ":" and as plain lines otherwise; pictures
  // sit directly under the instruction they belong to, with their caption.
  // Instructions that say a file "will be sent separately" (and any row whose
  // status is Pending) are highlighted yellow, as in the approved briefs.
  function componentCell(items, listRef, showCaptions) {
    const out = [];
    items.forEach((it) => {
      const segs = PBBriefInput.splitHighlights(it.text, it.status);
      out.push(new Paragraph({
        children: segs.map((sg) => run(sg.text, { highlight: sg.hl })),
        numbering: { reference: listRef, level: 0 },
        spacing: { after: 40, line: 240 },
      }));
      const detail = String(it.skuDetail || "").split("\n").map((l) => l.trim()).filter(Boolean);
      const lettered = /:\s*$/.test(String(it.text || "").trim());
      detail.forEach((l) => out.push(new Paragraph({
        children: [run(l)],
        numbering: lettered ? { reference: listRef, level: 1 } : undefined,
        indent: lettered ? undefined : { left: 720 },
        spacing: { after: 20, line: 240 },
      })));
      const pics = it.images || [];
      pics.forEach((im) => out.push(pictureParagraph(im, { indent: 720 })));
      if (pics.length && it.caption && showCaptions) {
        out.push(new Paragraph({ children: [run(it.caption)], indent: { left: 720 }, spacing: { after: 80, line: 240 } }));
      }
    });
    return out.length ? out : [p("")];
  }

  function productIndividualizationTable(brief, numberedLines, newListRef) {
    const useComponents = Array.isArray(brief.components) && brief.components.length > 0;
    // The approved briefs use a narrow label column for components and the
    // same ~11% for "Product Inclusions".
    const labelW = dxa(useComponents ? 11 : 25);
    const valueW = CONTENT_WIDTH - labelW;
    const rows = [];

    rows.push(new TableRow({
      children: [
        cell(p("Product Inclusions"), labelW),
        cell(numberedLines(brief.inclusions), valueW),
      ],
    }));

    rows.push(dividerRow("Material and workmanship instructions"));

    if (useComponents) {
      brief.components.forEach((c) => {
        const items = (c.items || []).filter((it) => (it.text || "").trim());
        if (!items.length) return;
        rows.push(new TableRow({
          children: [
            cell(p(c.name || ""), labelW),
            cell(componentCell(items, newListRef(), brief.showCaptions !== false), valueW),
          ],
        }));
      });
    } else {
      splitIntoSubRows(brief.materialInstructions).forEach(([label, content]) => {
        rows.push(new TableRow({
          children: [
            cell(p(label), labelW),
            cell(numberedLines(content), valueW),
          ],
        }));
      });
      if (brief.techSpecs && brief.techSpecs.length) {
        rows.push(new TableRow({
          children: [
            cell(p("Technical Specifications"), labelW),
            cell(brief.techSpecs.map((r) => p(`${r.label}: ${r.value}`, { after: 40 })), valueW),
          ],
        }));
      }
      rows.push(new TableRow({
        children: [
          cell(p("Packaging"), labelW),
          cell(numberedLines(brief.packaging), valueW),
        ],
      }));
      if (brief.includeImages && brief.images && brief.images.feature && brief.images.feature.length) {
        rows.push(new TableRow({
          children: [
            cell(p("Reference Images"), labelW),
            cell(imageParagraphs(brief.images.feature, 130), valueW),
          ],
        }));
      }
    }

    return table([labelW, valueW], rows);
  }

  // ---------- QC table ----------
  // PARAMETER | ACCEPTABLE (green tick) | NOT ACCEPTABLE (red cross): bold
  // centred headers, bold parameter names, centred text, defect photos
  // stacked under the "not acceptable" text.
  function qcTable(rows, assets) {
    const w1 = dxa(24.6), w2 = dxa(37.7), w3 = CONTENT_WIDTH - w1 - w2;
    const icon = (a) => a ? [new Paragraph({
      alignment: AlignmentType.CENTER, spacing: { after: 0 },
      children: [new ImageRun({ data: a.data, transformation: { width: 39, height: 39 } })],
    })] : [];
    const head = (t, a) => cell([p(t, { bold: true, align: AlignmentType.CENTER, after: 40 }), ...icon(a)], t === "PARAMETER" ? w1 : (t === "ACCEPTABLE" ? w2 : w3));
    const header = new TableRow({
      cantSplit: true,
      children: [head("PARAMETER"), head("ACCEPTABLE", assets && assets.check), head("NOT ACCEPTABLE", assets && assets.cross)],
    });
    const centred = (text) => String(text || "").split("\n").map((l) => p(l, { align: AlignmentType.CENTER, after: 40 }));
    const body = (rows || []).map((r) => {
      const notAcc = centred(r.notAcceptable);
      (r.images || []).forEach((im) => notAcc.push(pictureParagraph(im, { align: AlignmentType.CENTER, maxWidth: 260 })));
      // A row (text + its photos) never splits across pages - the photo
      // stays with the limits it illustrates, as in the approved briefs.
      return new TableRow({
        cantSplit: true,
        children: [
          cell(p(r.parameter, { bold: true }), w1),
          cell(centred(r.acceptable), w2),
          cell(notAcc, w3),
        ],
      });
    });
    return table([w1, w2, w3], [header, ...body]);
  }

  // ---------- compliance ----------
  // Lines may use **bold**; blank lines become small gaps.
  function complianceParagraphs(text) {
    const out = [];
    String(text || "").split("\n").forEach((line) => {
      if (!line.trim()) { out.push(p("", { after: 60 })); return; }
      out.push(new Paragraph({ children: inlineRuns(line), spacing: { after: 80 } }));
    });
    return out;
  }

  // The two QC header icons (format only) bundled with the app.
  async function loadAssets() {
    const spec = {
      check: ["img/qc-check.png", 39, 39],
      cross: ["img/qc-cross.png", 39, 39],
    };
    const out = {};
    for (const [key, [path, width, height]] of Object.entries(spec)) {
      try {
        const resp = await fetch(path);
        out[key] = { data: new Uint8Array(await resp.arrayBuffer()), width, height };
      } catch (e) { /* a missing picture must never block the export */ }
    }
    return out;
  }

  // Full-width branded banner (logo lockup on the CD Commerce brand
  // terracotta) fetched from the app's own assets and embedded in the page
  // header, exactly as in the reference template - replaces a text title.
  async function makeHeader() {
    const resp = await fetch("img/header-banner.jpg");
    const buf = new Uint8Array(await resp.arrayBuffer());
    const NATIVE_W = 2048, NATIVE_H = 146;
    // Full-bleed: the reference template anchors this image to the PAGE
    // (not the margins) at roughly (0,0) with no text wrap, so it spans the
    // entire page width edge-to-edge - matching that exactly rather than
    // embedding it within the content margins like a normal inline image.
    const targetWidthPx = Math.round((PAGE_WIDTH / 1440) * 96);
    const targetHeightPx = Math.round(targetWidthPx * (NATIVE_H / NATIVE_W));
    return new Header({
      children: [
        new Paragraph({
          spacing: { after: 0 },
          children: [
            new ImageRun({
              data: buf,
              transformation: { width: targetWidthPx, height: targetHeightPx },
              floating: {
                horizontalPosition: { relative: HorizontalPositionRelativeFrom.PAGE, offset: 0 },
                verticalPosition: { relative: VerticalPositionRelativeFrom.PAGE, offset: 0 },
                wrap: { type: TextWrappingType.NONE },
                behindDocument: false,
                allowOverlap: true,
              },
            }),
          ],
        }),
      ],
    });
  }

  function makeFooter() {
    return new Footer({
      children: [
        new Paragraph({
          alignment: AlignmentType.RIGHT,
          children: [new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: BODY_SIZE })],
        }),
      ],
    });
  }

  async function generateDocx(brief) {
    const children = [];
    const { numberedLines, newListRef, numberingConfigs } = makeNumberedLinesFactory();
    const assets = await loadAssets();

    const halfW = dxa(50);
    children.push(table([halfW, CONTENT_WIDTH - halfW], [
      new TableRow({
        children: [
          cell(p(`Date: ${brief.date || ""}`, { bold: true }), halfW, { noBorder: true }),
          cell(p(`PO Number: ${brief.poNumber || ""}`, { bold: true, align: AlignmentType.RIGHT }), CONTENT_WIDTH - halfW, { noBorder: true }),
        ],
      }),
    ]));
    children.push(p("", { after: 100 }));

    // Skip optional lines that have nothing in them (e.g. Step 3 wasn't
    // filled in yet) rather than rendering dangling labels like "Tel/Fax:"
    // with nothing after them, or blank gap rows.
    const toRows = [["To:", [p(brief.supplierName || "", { bold: true })]]];
    if (brief.supplierAddress) toRows.push(["", brief.supplierAddress]);
    if (brief.supplierPhone) toRows.push(["", `Tel/Fax: ${brief.supplierPhone}`]);
    if (brief.supplierEmail) toRows.push(["", `Email: ${brief.supplierEmail}`]);
    if (brief.supplierContact) toRows.push(["", `Contact person: ${brief.supplierContact}`]);
    children.push(labelValueTable(toRows, 15, true));

    children.push(p("", { after: 200 }));

    children.push(labelValueTable([
      ["Article No.", brief.articleNo],
      ["Item", brief.item],
      ["Description", brief.description],
      ["Material", brief.material],
      ["Color", brief.color],
    ], 22));

    children.push(p("", { after: 200 }));

    if (brief.includeImages && brief.images && brief.images.hero && brief.images.hero.length) {
      children.push(p("Product Image(s)", { bold: true, after: 100, keepNext: true }));
      children.push(...imageParagraphs(brief.images.hero, 220));
      children.push(p("", { after: 100 }));
    }

    children.push(p("I. PRODUCT INDIVIDUALIZATION", { bold: true, after: 150, keepNext: true }));
    children.push(productIndividualizationTable(brief, numberedLines, newListRef));
    children.push(p("", { after: 200 }));

    children.push(p("II. COMPLIANCE", { bold: true, after: 100, keepNext: true }));
    children.push(...complianceParagraphs(brief.compliance));
    children.push(p("", { after: 120 }));

    children.push(p("III. TYPICAL PRODUCTION MISTAKES/QUALITY CHECK ACCEPTABLE LIMITS", { bold: true, after: 150, keepNext: true }));
    children.push(qcTable(brief.qcRows, assets));
    children.push(p("", { after: 200 }));

    // One paragraph, as in the approved brief: the sentence followed directly
    // by the contact names (typed in Step 3, separate lines are joined).
    const contacts = String(brief.approvalContacts || "").split("\n").map((l) => l.trim()).filter(Boolean).join(" ");
    children.push(new Paragraph({
      spacing: { after: 200 },
      children: [run("Alterations and additions (in terms of Quantity, Quality, Product Packing or Carton Packing) to this contract must be made in writing directly to" + (contacts ? " " + contacts.replace(/^:\s*/, "") : ":"))],
    }));

    // "page 1-N" is the document's own total page count (a Word field), so it
    // is right whatever length the brief turns out to be.
    children.push(new Paragraph({
      spacing: { after: 300 },
      children: [
        run("Herewith I confirm that the product sample will be produced exactly as stated in the Product Briefing Agreement. I confirm to use the exact same materials as shown on page 1-"),
        new TextRun({ children: [PageNumber.TOTAL_PAGES], font: FONT, size: BODY_SIZE }),
        run(" as agreed with the customer. As well I confirm that the order details above will be met."),
      ],
    }));

    // A "\n" inside a single text run is not a line break in Word - it's
    // silently collapsed, which is why the line and the label used to run
    // together on one row. Two separate (centered) paragraphs instead, so
    // the label actually sits on its own line below the underscores.
    function signatureLine(label) {
      return [
        p("________________________", { align: AlignmentType.CENTER, after: 40 }),
        p(label, { align: AlignmentType.CENTER }),
      ];
    }
    children.push(table([halfW, CONTENT_WIDTH - halfW], [
      new TableRow({
        children: [
          cell(signatureLine("Supplier"), halfW, { noBorder: true }),
          cell(signatureLine("Customer - CD Commerce GmbH"), CONTENT_WIDTH - halfW, { noBorder: true }),
        ],
      }),
    ]));

    const doc = new Document({
      styles: { default: { document: { run: { font: FONT, size: BODY_SIZE } } } },
      numbering: { config: numberingConfigs },
      sections: [{
        properties: {
          page: {
            // docx.js swaps width/height itself when orientation is set to
            // LANDSCAPE (it expects portrait-equivalent values and flips
            // them) - PAGE_WIDTH/PAGE_HEIGHT here are already the landscape
            // values used everywhere else (table widths etc.), so they must
            // be passed swapped to end up correct after that internal flip.
            size: { width: PAGE_HEIGHT, height: PAGE_WIDTH, orientation: PageOrientation.LANDSCAPE },
            margin: { top: MARGIN, bottom: MARGIN, left: MARGIN, right: MARGIN, header: 720, footer: 720 },
          },
        },
        headers: { default: await makeHeader() },
        footers: { default: makeFooter() },
        children,
      }],
    });

    return Packer.toBlob(doc);
  }

  function downloadBlob(blob, filename) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
  }

  global.PBDocx = { generateDocx, downloadBlob };
})(window);
