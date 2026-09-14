/* docx-generator.js
 * Builds a .docx matching the standard CD Commerce Product Brief layout.
 * Uses the `docx` library (loaded globally via CDN). Plain .docx opens
 * cleanly in Google Docs (File > Open > Upload, or right-click > Open with
 * Google Docs in Drive) — no special export settings needed.
 */
(function (global) {
  const {
    Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
    HeadingLevel, WidthType, BorderStyle, AlignmentType, VerticalAlign,
  } = docx;

  const NOBORDER = {
    top: { style: BorderStyle.SINGLE, size: 2, color: "CCCCCC" },
    bottom: { style: BorderStyle.SINGLE, size: 2, color: "CCCCCC" },
    left: { style: BorderStyle.SINGLE, size: 2, color: "CCCCCC" },
    right: { style: BorderStyle.SINGLE, size: 2, color: "CCCCCC" },
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

  function cell(children, opts = {}) {
    return new TableCell({
      borders: NOBORDER,
      width: opts.width ? { size: opts.width, type: WidthType.PERCENTAGE } : undefined,
      verticalAlign: VerticalAlign.TOP,
      children: Array.isArray(children) ? children : [children],
      shading: opts.shaded ? { fill: "F2F2F2" } : undefined,
    });
  }

  function labeledRow(label, value) {
    return new TableRow({
      children: [
        cell(p(label, { bold: true }), { width: 25, shaded: true }),
        cell(multiline(value), { width: 75 }),
      ],
    });
  }

  function fullTable(rows) {
    return new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows });
  }

  function techSpecTable(rows) {
    const header = new TableRow({
      children: [
        cell(p("Parameter", { bold: true }), { width: 40, shaded: true }),
        cell(p("Value", { bold: true }), { width: 60, shaded: true }),
      ],
    });
    const body = (rows || []).map(
      (r) => new TableRow({ children: [cell(p(r.label)), cell(p(r.value))] })
    );
    return fullTable([header, ...body]);
  }

  function qcTable(rows) {
    const header = new TableRow({
      children: [
        cell(p("Parameter", { bold: true }), { width: 30, shaded: true }),
        cell(p("Acceptable", { bold: true }), { width: 35, shaded: true }),
        cell(p("Not Acceptable", { bold: true }), { width: 35, shaded: true }),
      ],
    });
    const body = (rows || []).map(
      (r) => new TableRow({
        children: [
          cell(p(r.parameter)),
          cell(multiline(r.acceptable)),
          cell(multiline(r.notAcceptable)),
        ],
      })
    );
    return fullTable([header, ...body]);
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

    children.push(fullTable([
      labeledRow("Article No.", brief.articleNo),
      labeledRow("Item", brief.item),
      labeledRow("Description", brief.description),
      labeledRow("Material", brief.material),
      labeledRow("Color", brief.color),
    ]));

    children.push(p("", { after: 200 }));
    children.push(p("PRODUCT INDIVIDUALIZATION", { heading: HeadingLevel.HEADING_2, after: 150 }));

    children.push(p("Product Inclusions", { bold: true }));
    children.push(...multiline(brief.inclusions, { after: 150 }));

    children.push(p("Material and Workmanship Instructions", { bold: true }));
    children.push(...multiline(brief.materialInstructions, { after: 150 }));

    children.push(p("Technical Specifications", { bold: true, after: 100 }));
    children.push(techSpecTable(brief.techSpecs));
    children.push(p("", { after: 200 }));

    children.push(p("Packaging", { bold: true }));
    children.push(...multiline(brief.packaging, { after: 200 }));

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

    children.push(new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: [new TableRow({
        children: [
          cell(p("________________________\nSupplier"), { width: 50 }),
          cell(p("________________________\nCustomer - CD Commerce GmbH"), { width: 50 }),
        ],
      })],
    }));

    const doc = new Document({
      sections: [{ properties: {}, children }],
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
