/* docx-generator.js
 * Builds a .docx matching CD Commerce's updated Product Brief template
 * (CD_Commerce_-_Product_Brief_Template.docx): landscape page, a full-width
 * orange branded header banner instead of a text title, "Date / PO Number"
 * as the first line, a "To:" supplier block, the Article info table, a
 * unified Product Individualization table (with per-component sub-rows),
 * Compliance before the QC mistakes table, then the QC table and sign-off.
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
    AlignmentType, ImageRun, PageNumber,
  } = docx;

  // Landscape A4, 1in (1440 twip) margins, matching the reference template.
  const PAGE_WIDTH = 16834;
  const PAGE_HEIGHT = 11909;
  const MARGIN = 1440;
  const CONTENT_WIDTH = PAGE_WIDTH - MARGIN * 2; // 13954 DXA

  function dxa(pct) {
    return Math.round((CONTENT_WIDTH * pct) / 100);
  }

  const BORDER = {
    top: { style: BorderStyle.SINGLE, size: 4, color: "AAAAAA" },
    bottom: { style: BorderStyle.SINGLE, size: 4, color: "AAAAAA" },
    left: { style: BorderStyle.SINGLE, size: 4, color: "AAAAAA" },
    right: { style: BorderStyle.SINGLE, size: 4, color: "AAAAAA" },
  };
  const NO_BORDER = {
    top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
    bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
    left: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
    right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  };

  function p(text, opts = {}) {
    return new Paragraph({
      children: [new TextRun({ text: text || "", bold: !!opts.bold, size: opts.size || 22 })],
      spacing: { after: opts.after ?? 120 },
      heading: opts.heading,
      alignment: opts.align,
    });
  }

  function multiline(text, opts = {}) {
    const lines = (text || "").split("\n");
    return lines.map((line) => p(line, opts));
  }

  function cell(children, widthDxa, opts = {}) {
    return new TableCell({
      borders: opts.noBorder ? NO_BORDER : BORDER,
      width: { size: widthDxa, type: WidthType.DXA },
      columnSpan: opts.columnSpan,
      verticalAlign: VerticalAlign.TOP,
      children: Array.isArray(children) ? children : [children],
      shading: opts.shaded ? { fill: "F2F2F2" } : undefined,
    });
  }

  function table(columnWidthsDxa, rows) {
    return new Table({
      width: { size: CONTENT_WIDTH, type: WidthType.DXA },
      columnWidths: columnWidthsDxa,
      rows,
    });
  }

  // Two-column "label | value" table (header block, To:, Article info).
  function labelValueTable(pairs, labelPct = 25) {
    const labelW = dxa(labelPct);
    const valueW = CONTENT_WIDTH - labelW;
    const rows = pairs.map(
      ([label, value, opts = {}]) => new TableRow({
        children: [
          cell(p(label, { bold: true }), labelW, { shaded: true }),
          cell(Array.isArray(value) ? value : multiline(value), valueW, opts),
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
      const looksLikeLabel = first.length > 0 && first.length <= 60 && !rest ? false : first.length <= 60;
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
      children: [cell(p(label, { bold: true }), CONTENT_WIDTH, { shaded: true, columnSpan: 2 })],
    });
  }

  function productIndividualizationTable(brief) {
    const labelW = dxa(25);
    const valueW = CONTENT_WIDTH - labelW;
    const rows = [];

    rows.push(new TableRow({
      children: [
        cell(p("Product Inclusions", { bold: true }), labelW, { shaded: true }),
        cell(multiline(brief.inclusions), valueW),
      ],
    }));

    rows.push(dividerRow("Material and Workmanship Instructions"));
    splitIntoSubRows(brief.materialInstructions).forEach(([label, content]) => {
      rows.push(new TableRow({
        children: [
          cell(p(label, { bold: true }), labelW, { shaded: true }),
          cell(multiline(content), valueW),
        ],
      }));
    });
    if (brief.techSpecs && brief.techSpecs.length) {
      rows.push(new TableRow({
        children: [
          cell(p("Technical Specifications", { bold: true }), labelW, { shaded: true }),
          cell(brief.techSpecs.map((r) => p(`${r.label}: ${r.value}`, { after: 40 })), valueW),
        ],
      }));
    }

    rows.push(new TableRow({
      children: [
        cell(p("Packaging", { bold: true }), labelW, { shaded: true }),
        cell(multiline(brief.packaging), valueW),
      ],
    }));

    return table([labelW, valueW], rows);
  }

  function qcTable(rows) {
    const w1 = dxa(25), w2 = dxa(37.5), w3 = CONTENT_WIDTH - w1 - w2;
    const header = new TableRow({
      children: [
        cell(p("Parameter", { bold: true }), w1, { shaded: true }),
        cell(p("Acceptable", { bold: true }), w2, { shaded: true }),
        cell(p("Not Acceptable", { bold: true }), w3, { shaded: true }),
      ],
    });
    const body = (rows || []).map(
      (r) => new TableRow({
        children: [
          cell(p(r.parameter), w1),
          cell(multiline(r.acceptable), w2),
          cell(multiline(r.notAcceptable), w3),
        ],
      })
    );
    return table([w1, w2, w3], [header, ...body]);
  }

  // Full-width branded banner (logo lockup on the CD Commerce brand
  // terracotta) fetched from the app's own assets and embedded in the page
  // header, exactly as in the reference template - replaces a text title.
  async function makeHeader() {
    const resp = await fetch("img/header-banner.jpg");
    const buf = new Uint8Array(await resp.arrayBuffer());
    const NATIVE_W = 2048, NATIVE_H = 146;
    const targetWidthPx = Math.round((CONTENT_WIDTH / 1440) * 96);
    const targetHeightPx = Math.round(targetWidthPx * (NATIVE_H / NATIVE_W));
    return new Header({
      children: [
        new Paragraph({
          spacing: { after: 0 },
          children: [
            new ImageRun({
              data: buf,
              transformation: { width: targetWidthPx, height: targetHeightPx },
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
          children: [new TextRun({ children: [PageNumber.CURRENT] })],
        }),
      ],
    });
  }

  async function generateDocx(brief) {
    const children = [];

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

    children.push(labelValueTable([
      ["To:", [p(brief.supplierName || "", { bold: true })]],
      ["", brief.supplierAddress || ""],
      ["", `Tel/Fax: ${brief.supplierPhone || ""}`],
      ["", `Contact person: ${brief.supplierContact || ""}`],
    ], 15));

    children.push(p("", { after: 200 }));

    children.push(labelValueTable([
      ["Article No.", brief.articleNo],
      ["Item", brief.item],
      ["Description", brief.description],
      ["Material", brief.material],
      ["Color", brief.color],
    ]));

    children.push(p("", { after: 200 }));
    children.push(p("PRODUCT INDIVIDUALIZATION", { bold: true, after: 150 }));
    children.push(productIndividualizationTable(brief));
    children.push(p("", { after: 200 }));

    if (brief.knownIssues) {
      children.push(p("Known Market Complaints & Preventive Actions", { bold: true, after: 100 }));
      children.push(...multiline(brief.knownIssues, { after: 200 }));
    }

    if (brief.qualityInspectionNotes) {
      children.push(p("Quality Inspection Notes", { bold: true, after: 100 }));
      children.push(...multiline(brief.qualityInspectionNotes, { after: 200 }));
    }

    children.push(p("II. COMPLIANCE", { bold: true, after: 100 }));
    children.push(...multiline(brief.compliance, { after: 200 }));

    children.push(p("III. TYPICAL PRODUCTION MISTAKES/QUALITY CHECK ACCEPTABLE LIMITS", { bold: true, after: 150 }));
    children.push(qcTable(brief.qcRows));
    children.push(p("", { after: 200 }));

    children.push(p("Alterations and additions (in terms of Quantity, Quality, Product Packing or Carton Packing) to this contract must be made in writing directly to:", { after: 80 }));
    children.push(...multiline(brief.approvalContacts, { after: 200 }));

    children.push(p("Herewith I confirm that the product sample will be produced exactly as stated in the Product Briefing Agreement. I confirm to use the exact same materials as shown on page 1-22 as agreed with the customer. As well I confirm that the order details above will be met.", { after: 300 }));

    children.push(table([halfW, CONTENT_WIDTH - halfW], [
      new TableRow({
        children: [
          cell(p("________________________\nSupplier"), halfW, { noBorder: true }),
          cell(p("________________________\nCustomer - CD Commerce GmbH"), CONTENT_WIDTH - halfW, { noBorder: true }),
        ],
      }),
    ]));

    const doc = new Document({
      sections: [{
        properties: {
          page: {
            size: { width: PAGE_WIDTH, height: PAGE_HEIGHT },
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
