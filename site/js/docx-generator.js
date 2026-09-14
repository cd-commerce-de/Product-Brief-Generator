/* docx-generator.js
 * Builds a .docx matching the standard CD Commerce Product Brief layout.
 * Uses the `docx` library (loaded globally via CDN). Plain .docx opens
 * cleanly in Google Docs (File > Open > Upload, or right-click > Open with
 * Google Docs in Drive) - no special export settings needed.
 *
 * IMPORTANT width note: docx.js's Table defaults `columnWidths` to 100 DXA
 * per column (about 1.7mm) if it isn't set explicitly - individual cell
 * `width` values are NOT enough on their own to size columns; the table's
 * own `columnWidths` grid is what most renderers (incl. Google Docs)
 * actually use. Every table below sets `columnWidths` in DXA (twentieths of
 * a point) explicitly for this reason - leaving it out is what previously
 * caused every column to collapse to near-zero width (text wrapping one
 * character per line).
 */
(function (global) {
  const {
    Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
    HeadingLevel, WidthType, BorderStyle, VerticalAlign, PageOrientation,
  } = docx;

  // A4 page, 1in (1440 twip) margins on all sides -> usable width in DXA.
  const PAGE_WIDTH = 11906;
  const PAGE_HEIGHT = 16838;
  const MARGIN = 1440;
  const CONTENT_WIDTH = PAGE_WIDTH - MARGIN * 2; // 9026 DXA

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

  // Two-column "label | value" table (used for the header block and for
  // PRODUCT INDIVIDUALIZATION, mirroring the reference brief's layout).
  function labelValueTable(pairs, labelPct = 25) {
    const labelW = dxa(labelPct);
    const valueW = CONTENT_WIDTH - labelW;
    const rows = pairs.map(
      ([label, value]) => new TableRow({
        children: [
          cell(p(label, { bold: true }), labelW, { shaded: true }),
          cell(Array.isArray(value) ? value : multiline(value), valueW),
        ],
      })
    );
    return table([labelW, valueW], rows);
  }

  function techSpecBlock(rows) {
    if (!rows || !rows.length) return [];
    return [
      p("", { after: 100 }),
      p("TECHNICAL SPECIFICATIONS", { bold: true, after: 80 }),
      ...rows.map((r) => p(`${r.label}: ${r.value}`, { after: 40 })),
    ];
  }

  function qcTable(rows) {
    const w1 = dxa(30), w2 = dxa(35), w3 = CONTENT_WIDTH - w1 - w2;
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

  async function generateDocx(brief) {
    const children = [];

    children.push(p(`Date: ${brief.date || ""}`, { bold: true, after: 200 }));
    children.push(p("PRODUCT BRIEFING", { heading: HeadingLevel.HEADING_1, after: 200 }));

    children.push(p("CD Commerce GmbH", { bold: true }));
    children.push(p("Ernst-Abbe-Str. 2"));
    children.push(p("136179 Bebra"));
    children.push(p("Deutschland", { after: 200 }));

    children.push(p("To:", { bold: true }));
    children.push(p(brief.supplierName || "", { bold: true }));
    children.push(p(brief.supplierAddress || ""));
    children.push(p(`Contact: ${brief.supplierContact || ""}`));
    children.push(p(`Tel/Fax: ${brief.supplierPhone || ""}`));
    children.push(p(`PO Number: ${brief.poNumber || ""}`, { bold: true, after: 300 }));

    children.push(labelValueTable([
      ["Article No.", brief.articleNo],
      ["Item", brief.item],
      ["Description", brief.description],
      ["Material", brief.material],
      ["Color", brief.color],
    ]));

    children.push(p("", { after: 200 }));
    children.push(p("PRODUCT INDIVIDUALIZATION", { heading: HeadingLevel.HEADING_2, after: 150 }));

    // One unified table for inclusions / material & workmanship (with tech
    // specs folded into the same cell) / packaging - matching the reference
    // brief's single-table section rather than splitting into separate
    // floating tables.
    children.push(labelValueTable([
      ["Product Inclusions", multiline(brief.inclusions)],
      [
        "Material and Workmanship Instructions",
        [...multiline(brief.materialInstructions), ...techSpecBlock(brief.techSpecs)],
      ],
      ["Packaging", multiline(brief.packaging)],
    ], 30));

    children.push(p("", { after: 200 }));
    children.push(p("TYPICAL PRODUCTION MISTAKES / QUALITY CHECK ACCEPTABLE LIMITS", {
      heading: HeadingLevel.HEADING_2, after: 150,
    }));
    children.push(qcTable(brief.qcRows));
    children.push(p("", { after: 200 }));

    if (brief.knownIssues) {
      children.push(p("Known Market Complaints & Preventive Actions", { heading: HeadingLevel.HEADING_2, after: 100 }));
      children.push(...multiline(brief.knownIssues, { after: 200 }));
    }

    if (brief.qualityInspectionNotes) {
      children.push(p("Quality Inspection Notes", { heading: HeadingLevel.HEADING_2, after: 100 }));
      children.push(...multiline(brief.qualityInspectionNotes, { after: 200 }));
    }

    children.push(p("Compliance", { heading: HeadingLevel.HEADING_2, after: 100 }));
    children.push(...multiline(brief.compliance, { after: 200 }));

    if (brief.approvalContacts) {
      children.push(p("If any alterations and additions to this document will be made, the Supplier must contact:", { after: 80 }));
      children.push(...multiline(brief.approvalContacts, { after: 200 }));
    }

    children.push(p("Herewith I confirm that the product sample will be produced exactly as stated in the Purchase Order Agreement.", { after: 300 }));

    const halfW = dxa(50);
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
            margin: { top: MARGIN, bottom: MARGIN, left: MARGIN, right: MARGIN },
          },
        },
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
