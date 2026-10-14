/* xlsx-generators.js
 * Builds the three downstream spreadsheets that reuse Product Brief data:
 * Pre-inspection Briefing Form, Pre-QC (Quality Control) Check, and
 * Marketing Guide Sheet. Plain .xlsx, opens natively in Google Sheets
 * (Drive -> right-click -> Open with -> Google Sheets).
 *
 * These intentionally don't try to invent product-specific test steps or
 * marketing copy out of thin air - they carry over what the brief actually
 * knows (product/PO/supplier identity, dimensions, QC notes, USPs/ESPs) and
 * leave clearly-labeled prompts where a human still needs to write original
 * content (benefit copy, target group, photos).
 *
 * FORMATTING: matches the color/style conventions of CD Commerce's actual
 * reference documents exactly (checked against real SUP examples) - orange
 * #FF9900 title/field bars, green #00FF00 table header rows, bold labels,
 * wrapped text, thin borders throughout, and the same column widths. This
 * requires writing cell styles, which plain SheetJS (xlsx.full.min.js)
 * deliberately does not support at all - see the xlsx-js-style CDN script
 * in index.html, a maintained fork with the identical read/parse API (so
 * nothing else in the app changes) that adds style-writing on top.
 */
(function (global) {
  const BORDER_THIN = { style: "thin", color: { rgb: "FF000000" } };
  const ALL_BORDERS = { top: BORDER_THIN, bottom: BORDER_THIN, left: BORDER_THIN, right: BORDER_THIN };
  const ORANGE = "FFFF9900";
  const LIGHT_ORANGE = "FFF6B26B";
  const GREEN = "FF00FF00";

  const TITLE_STYLE = {
    font: { bold: true, sz: 13 },
    fill: { patternType: "solid", fgColor: { rgb: ORANGE } },
    border: ALL_BORDERS,
  };
  const FIELD_LABEL_STYLE = {
    font: { bold: true },
    border: ALL_BORDERS,
    alignment: { vertical: "top", wrapText: true },
  };
  const FIELD_VALUE_STYLE = {
    border: ALL_BORDERS,
    alignment: { vertical: "top", wrapText: true },
  };
  const TABLE_HEADER_STYLE = {
    font: { bold: true },
    fill: { patternType: "solid", fgColor: { rgb: GREEN } },
    border: ALL_BORDERS,
    alignment: { vertical: "center", horizontal: "center", wrapText: true },
  };
  const ROW_LABEL_STYLE = {
    font: { bold: true },
    border: ALL_BORDERS,
    alignment: { vertical: "top", wrapText: true },
  };
  const ROW_VALUE_STYLE = {
    border: ALL_BORDERS,
    alignment: { vertical: "top", wrapText: true },
  };
  const SUBSECTION_STYLE = {
    font: { bold: true },
    fill: { patternType: "solid", fgColor: { rgb: LIGHT_ORANGE } },
    border: ALL_BORDERS,
    alignment: { vertical: "top", wrapText: true },
  };

  // Applies a style to every cell in a rectangular range (inclusive),
  // creating empty cells where needed so borders/fills still show even if
  // a cell in the range has no value (e.g. a title merged across columns
  // with nothing in the later cells).
  function styleRange(ws, r0, c0, r1, c1, style) {
    for (let r = r0; r <= r1; r++) {
      for (let c = c0; c <= c1; c++) {
        const addr = XLSX.utils.encode_cell({ r, c });
        if (!ws[addr]) ws[addr] = { t: "z", v: undefined };
        ws[addr].s = style;
      }
    }
  }

  function setColWidths(ws, widths) {
    ws["!cols"] = widths.map((wch) => ({ wch }));
  }

  function addMerge(ws, r0, c0, r1, c1) {
    ws["!merges"] = ws["!merges"] || [];
    ws["!merges"].push({ s: { r: r0, c: c0 }, e: { r: r1, c: c1 } });
  }

  function aoaToStyledBlob(rows, sheetName, colWidths) {
    const ws = XLSX.utils.aoa_to_sheet(rows);
    if (colWidths) setColWidths(ws, colWidths);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, sheetName);
    const out = XLSX.write(wb, { bookType: "xlsx", type: "array", bookSST: true });
    return { ws, blob: new Blob([out], { type: "application/octet-stream" }) };
  }

  function finalizeBlob(wb) {
    const out = XLSX.write(wb, { bookType: "xlsx", type: "array", bookSST: true });
    return new Blob([out], { type: "application/octet-stream" });
  }

  function findRawSection(brief, keywords) {
    const sections = (brief.raw && brief.raw.sections) || [];
    return sections.find((s) =>
      keywords.some((k) => s.title.toLowerCase().includes(k))
    );
  }

  // Dimensions live in brief.techSpecs only when the PD sheet had an
  // explicit "Technical Specifications" sub-table (legacy LTF/SUP-style
  // sheets). Post-optimized sheets (KPM/AKP/SSC-style) fold dimensions into
  // the Material & Workmanship Instructions text as a "DIMENSIONS" block
  // instead - fall back to pulling that out when techSpecs doesn't have it,
  // rather than leaving the field blank.
  function findDimensionSpec(brief) {
    const spec = (brief.techSpecs || []).find((r) => /dimension|size/i.test(r.label));
    if (spec) return spec.value;
    const m = /DIMENSIONS\n([\s\S]*?)(?:\n\n|$)/.exec(brief.materialInstructions || "");
    return m ? m[1].trim() : "";
  }

  // "Packaging Size" should be the short carton dimension (e.g. "68 x 52 x
  // 12 cm"), not the full packaging write-up (material, weight, shipping
  // fees, etc.) that brief.packaging holds in full elsewhere in this doc
  // and in the Product Brief itself. Pull just the "Packaging Dimensions"
  // line out of that text; fall back to the whole thing only if that
  // specific line isn't present, so the field is never silently empty.
  function findPackagingSize(brief) {
    const text = brief.packaging || "";
    const m = /Packaging Dimensions\s*[—-]\s*(.+)/i.exec(text);
    if (m) return m[1].trim();
    return text;
  }

  // ---------- one source of truth for sizes, cartons and compliance ----------
  // When the PD sheet has a Product Brief Input tab, the brief's own
  // components, SKU table and QC rows are the source for every figure in
  // these forms - NOT the older spec tabs, which can disagree with it (the
  // SSC sheet's old tab says 46 cm wide, its Brief Input says 50). The older
  // lookups above are only the fallback for sheets without that tab.
  function fmtNum(v) {
    const t = String(v == null ? "" : v).trim();
    return /^-?\d+(\.\d+)?$/.test(t) ? String(Number(t)) : t;
  }

  // The brief's compliance box can carry **bold** markers; spreadsheets
  // can't show them, so they must never reach a cell.
  function stripMarkup(text) {
    return String(text || "").replace(/\*\*/g, "");
  }

  const MEASURE_RE = /\d+(?:\.\d+)?\s*[x×]\s*\d+|\d+(?:[.,]\d+)?\s*(?:mm|cm|kg|g)\b/i;
  const DIM_RE = /\d+(?:\.\d+)?(?:\s*[x×]\s*\d+(?:\.\d+)?)+\s*(?:mm|cm|m)?/i;

  function mainItems(brief) {
    const out = [];
    (brief.components || []).filter((c) => c.kind !== "packaging").forEach((c) => (c.items || []).forEach((it) => out.push(it)));
    return out;
  }

  // Short spec statements from the brief's instructions ("Wheels: 110 x 24 mm
  // PU ..."). Long sentences are rules, not measurements, so they're skipped.
  function measurableSpecs(brief) {
    if (!(brief.components || []).length) return null;
    return mainItems(brief)
      .map((it) => String(it.text || "").trim())
      .filter((t) => t && t.length <= 140 && MEASURE_RE.test(t))
      .map((t) => t.replace(/\.$/, ""));
  }

  // "The overall assembled size is 69 x 50 x84 cm." -> "69 x 50 x84 cm";
  // "The deck size is 52 x 12 cm." -> "Deck: 52 x 12 cm".
  function sizeLines(brief) {
    if (!(brief.components || []).length) return null;
    const out = [];
    mainItems(brief).forEach((it) => {
      const t = String(it.text || "");
      if (!/\b(size|dimensions?)\b/i.test(t)) return;
      const m = DIM_RE.exec(t);
      if (!m) return;
      const lab = /^the\s+(.+?)\s+(?:size|dimensions?)\b/i.exec(t.trim());
      let label = lab ? lab[1].trim() : "";
      if (/overall|assembled|product|total/i.test(label)) label = "";
      label = label.charAt(0).toUpperCase() + label.slice(1);
      out.push(label ? `${label}: ${m[0].trim()}` : m[0].trim());
    });
    return out;
  }

  function dimensionText(brief) {
    const sz = sizeLines(brief);
    if (sz && sz.length) return sz.join("\n");
    return findDimensionSpec(brief);
  }

  function specsText(brief) {
    const ms = measurableSpecs(brief);
    if (ms && ms.length) return ms.slice(0, 10).join("\n");
    return findDimensionSpec(brief);
  }

  // Carton sizes/weights straight from the SKU table. Returns lines, or
  // null when the brief has no SKU table (older sheets).
  function cartonLines(brief) {
    const skus = (brief.skuTable || []).filter((s) => s.itemCarton || s.masterCarton);
    if (!skus.length) return null;
    const uniq = (key) => {
      const vals = skus.map((s) => fmtNum(s[key])).filter(Boolean);
      return vals.length && vals.every((v) => v === vals[0]) ? vals[0] : null;
    };
    const fmt = (dims, wt, pcs) => [dims && `${dims} cm`, wt && `${wt} kg`, pcs && `${pcs} pcs`].filter(Boolean).join(", ");
    const out = [];
    const itemDims = uniq("itemCarton"), itemWt = uniq("itemCartonWeight");
    if (skus.some((s) => s.itemCarton)) {
      if (itemDims) out.push("Item carton: " + fmt(itemDims, itemWt));
      else skus.forEach((s) => out.push(`Item carton ${s.sku}: ` + fmt(fmtNum(s.itemCarton), fmtNum(s.itemCartonWeight))));
    }
    const mDims = uniq("masterCarton"), mWt = uniq("masterCartonWeight"), pcs = uniq("pcsPerMaster");
    if (skus.some((s) => s.masterCarton)) {
      if (mDims) out.push("Master carton: " + fmt(mDims, mWt, pcs));
      else skus.forEach((s) => out.push(`Master carton ${s.sku}: ` + fmt(fmtNum(s.masterCarton), fmtNum(s.masterCartonWeight), fmtNum(s.pcsPerMaster))));
    }
    return out;
  }

  // Marketing "Packaging size": just the single-item carton size.
  function packagingSizeText(brief) {
    const skus = (brief.skuTable || []).filter((s) => s.itemCarton);
    if (skus.length) {
      const dims = [...new Set(skus.map((s) => fmtNum(s.itemCarton)))];
      return dims.length === 1 ? `${dims[0]} cm` : skus.map((s) => `${s.sku}: ${fmtNum(s.itemCarton)} cm`).join("\n");
    }
    return findPackagingSize(brief);
  }

  // The brief's first packaging instruction (what the packing is made of).
  function packagingMaterialLine(brief) {
    const pk = (brief.components || []).find((c) => c.kind === "packaging");
    return pk && pk.items && pk.items[0] ? String(pk.items[0].text || "").trim() : "";
  }

  // QC rows the team actually typed in Part C (not the tool's placeholders).
  function sheetQcRows(brief) {
    return (brief.qcRows || []).filter((r) => r.fromSheet && r.parameter);
  }

  function qcCriteria(r) {
    return [
      r.acceptable && `Acceptable: ${r.acceptable}`,
      r.notAcceptable && `Not acceptable: ${r.notAcceptable}`,
      r.method && `Method: ${r.method}`,
    ].filter(Boolean).join("\n");
  }

  // Excel and Google Sheets don't grow a row to fit wrapped text inside
  // MERGED cells (and may not on open at all), so long text got cut off.
  // Set each row's height explicitly from how many lines its text needs at
  // the cell's (merged) width.
  function fitRows(ws, colWidths) {
    const ref = XLSX.utils.decode_range(ws["!ref"]);
    const origins = {};
    (ws["!merges"] || []).forEach((m) => { origins[`${m.s.r},${m.s.c}`] = m; });
    const rowsOut = [];
    for (let r = ref.s.r; r <= ref.e.r; r++) {
      let maxLines = 1;
      for (let c = ref.s.c; c <= ref.e.c; c++) {
        const cell = ws[XLSX.utils.encode_cell({ r, c })];
        if (!cell || typeof cell.v !== "string" || !cell.v) continue;
        const m = origins[`${r},${c}`];
        let width = 0;
        for (let k = c; k <= (m ? m.e.c : c); k++) width += colWidths[k] || 9;
        const bold = cell.s && cell.s.font && cell.s.font.bold;
        const perLine = Math.max(4, Math.floor(width * (bold ? 0.95 : 1.1) - 1));
        let lines = 0;
        cell.v.split("\n").forEach((seg) => { lines += Math.max(1, Math.ceil(seg.length / perLine)); });
        if (lines > maxLines) maxLines = lines;
      }
      rowsOut[r] = maxLines > 1 ? { hpt: Math.min(409, maxLines * 15 + 3) } : {};
    }
    ws["!rows"] = rowsOut;
  }

  // ---------- Pre-inspection Briefing Form ----------
  function generatePreInspectionXlsx(brief) {
    const fieldRows = [
      ["Supplier Company Name", brief.supplierName || ""],
      ["Contact Person", brief.supplierContact || ""],
      ["Contact Email", brief.supplierEmail || ""],
      ["Product Name", brief.item || ""],
      ["Purchase Order (PO) Number", brief.poNumber || ""],
      ["Pre-QC check (in-house) DATE", brief.preQcDate || ""],
      ["Planned 3rd Party Inspection Date", brief.inspectionDate || ""],
      ["# of samples to be inspected", brief.sampleCount || ""],
    ];

    const testRows = [
      [
        "Package completeness",
        `${brief.item || "Product"} complete set. Must match inclusions list exactly per the approved Product Brief.\n\nInclusions to verify:\n${brief.inclusions || "(see Product Brief)"}`,
        "Available/Not Available",
      ],
      [
        "Measurement",
        `~Measuring tape\n~Vernier caliper / weighing scale\n\nAgainst spec:\n${specsText(brief) || "(see Technical Specifications in Product Brief)"}`,
        "Available/Not Available",
      ],
      [
        "Visual inspection",
        "~Good lighting setup\n~Clean inspection table\n~All reference materials:\n     - Approved design artwork\n     - Approved color references\n     - Product briefing document\n     - Golden sample or retention sample",
        "Available/Not Available",
      ],
      [
        "Packaging",
        "~Complete carton with all packaging materials\n~Tape and sealing materials\n\nSpec:\n" + (
          cartonLines(brief)
            ? [packagingMaterialLine(brief), ...cartonLines(brief)].filter(Boolean).join("\n")
            : (brief.packaging || "(see Product Brief)")
        ),
        "Available/Not Available",
      ],
    ];

    if (brief.qualityInspectionNotes) {
      testRows.push(["Product-specific test (from PD sheet)", brief.qualityInspectionNotes, "Available/Not Available"]);
    }
    // The QC limits the team defined in the brief's Part C: each becomes a
    // check, with its acceptance limits and test method.
    sheetQcRows(brief).forEach((r) => {
      testRows.push([r.parameter, "~Check against the brief's limits:\n" + qcCriteria(r), "Available/Not Available"]);
    });
    (brief.extraPreInspectionTests || []).forEach((t) => {
      if (!t.test && !t.items) return;
      testRows.push([t.test || "", t.items || "", "Available/Not Available"]);
    });

    const rows = [
      ["Pre-inspection Briefing Form", "", ""],
      ...fieldRows.map(([l, v]) => [l, v, ""]),
      ["", "", ""],
      ["Test(s) to be Conducted", "Items to be Prepared for Testing", "Availability"],
      ...testRows,
    ];

    let r = rows.length;
    if (brief.productionStatus) {
      rows.push([], ["Production status at time of inspection", brief.productionStatus]);
      r = rows.length;
    }
    // Without Part C limits, fall back to the PD sheet's market-complaint
    // analysis so inspectors still know what to watch for.
    if (brief.knownIssues && !sheetQcRows(brief).length) {
      rows.push([], ["Known Market Complaints to Watch For (from PD review analysis)"]);
      brief.knownIssues.split("\n\n").forEach((block) => rows.push([block]));
    }

    const ws = XLSX.utils.aoa_to_sheet(rows);
    setColWidths(ws, [34, 58, 20]);

    addMerge(ws, 0, 0, 0, 2);
    styleRange(ws, 0, 0, 0, 2, TITLE_STYLE);
    fieldRows.forEach((_, i) => {
      const rr = 1 + i;
      addMerge(ws, rr, 1, rr, 2);
      styleRange(ws, rr, 0, rr, 0, FIELD_LABEL_STYLE);
      styleRange(ws, rr, 1, rr, 2, FIELD_VALUE_STYLE);
    });

    const headerRowIdx = 1 + fieldRows.length + 1;
    styleRange(ws, headerRowIdx, 0, headerRowIdx, 2, TABLE_HEADER_STYLE);
    testRows.forEach((_, i) => {
      const rr = headerRowIdx + 1 + i;
      styleRange(ws, rr, 0, rr, 0, ROW_LABEL_STYLE);
      styleRange(ws, rr, 1, rr, 2, ROW_VALUE_STYLE);
    });

    fitRows(ws, [34, 58, 20]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Pre-inspection");
    return finalizeBlob(wb);
  }

  // ---------- Pre-QC (Quality Control) Check ----------
  function generatePreQCXlsx(brief) {
    const fieldRows = [
      ["Supplier Company Name", brief.supplierName || ""],
      ["Contact Person", brief.supplierContact || ""],
      ["Contact Email", brief.supplierEmail || ""],
      ["Product Name", brief.item || ""],
      ["Purchase Order (PO) Number", brief.poNumber || ""],
      ["Target 3rd party inspection date", brief.inspectionDate || ""],
      ["# of samples to be inspected", brief.sampleCount || ""],
    ];

    const testRows = [
      [
        "Package completeness check",
        "~Verify that each set contains all required items and matches the approved Product Brief.\n~Quantity, specification, and general condition of each item are aligned.\n~Confirm nothing is substituted or missing.\n\nInclusions:\n" + (brief.inclusions || ""),
        "", "",
      ],
      ["Material / surface check", "~Inspect visible quality of materials against the Product Brief.\n\nSpec:\n" + (brief.material || ""), "", ""],
      ["Dimensions check", "~Confirm product dimensions match the approved specification.\n\nSpec:\n" + (dimensionText(brief) || specsText(brief) || "(see Technical Specifications)"), "", ""],
      ["Odor check", "~Smell by nose within 10cm.", "", ""],
      ["Printing, artwork, colors check", "~Compare product print with approved artwork.\n~Check logo orientation, position, sharpness, and color.\n\nColor spec: " + (brief.color || ""), "", ""],
    ];

    if (brief.qualityInspectionNotes) {
      testRows.push(["Product-specific test (from PD sheet)", brief.qualityInspectionNotes, "", ""]);
    }
    // One test per QC limit the team wrote in the brief's Part C, with its
    // acceptance limits and test method.
    sheetQcRows(brief).forEach((r) => testRows.push([r.parameter, qcCriteria(r), "", ""]));
    (brief.extraPreQcTests || []).forEach((t) => {
      if (!t.test && !t.methods) return;
      testRows.push([t.test || "", t.methods || "", "", ""]);
    });

    const rows = [
      ["Pre-QC (Quality Control) Check", "", "", ""],
      ...fieldRows.map(([l, v]) => [l, v, "", ""]),
      ["", "", "", ""],
      ["Test/s", "Methods", "Result (PASSED/FAILED)", "Photos/Videos with Remarks"],
      ...testRows,
    ];

    const packingStart = rows.length + 1; // +1 for the blank spacer pushed below
    rows.push([], ["Packing List Overview", "Requirement", "ACTUAL", "Photo Proof"]);
    const packingRows = [];
    if (brief.productionStatus) packingRows.push(["Production status", brief.productionStatus, "", ""]);
    packingRows.push(["Carton dimensions / weight", cartonLines(brief) ? cartonLines(brief).join("\n") : (brief.packaging || ""), "", ""]);
    packingRows.forEach((row) => rows.push(row));

    const complianceHeaderIdx = rows.length + 1;
    rows.push([], ["Required Certification / Compliance Declarations", "", "AVAILABLE/NOT AVAILABLE", "Remarks"]);
    rows.push([stripMarkup(brief.compliance), "", "", ""]);

    const signOffIdx = rows.length + 1;
    rows.push([], ["Supplier Signature & Date"], ["Name", "Signature", "", "Date"]);

    const ws = XLSX.utils.aoa_to_sheet(rows);
    setColWidths(ws, [34, 56, 24, 30]);

    addMerge(ws, 0, 0, 0, 3);
    styleRange(ws, 0, 0, 0, 3, TITLE_STYLE);
    fieldRows.forEach((_, i) => {
      const rr = 1 + i;
      addMerge(ws, rr, 1, rr, 3);
      styleRange(ws, rr, 0, rr, 0, FIELD_LABEL_STYLE);
      styleRange(ws, rr, 1, rr, 3, FIELD_VALUE_STYLE);
    });

    const headerRowIdx = 1 + fieldRows.length + 1;
    styleRange(ws, headerRowIdx, 0, headerRowIdx, 3, TABLE_HEADER_STYLE);
    testRows.forEach((_, i) => {
      const rr = headerRowIdx + 1 + i;
      styleRange(ws, rr, 0, rr, 0, ROW_LABEL_STYLE);
      styleRange(ws, rr, 1, rr, 3, ROW_VALUE_STYLE);
    });

    styleRange(ws, packingStart, 0, packingStart, 3, TABLE_HEADER_STYLE);
    packingRows.forEach((_, i) => {
      const rr = packingStart + 1 + i;
      styleRange(ws, rr, 0, rr, 0, ROW_LABEL_STYLE);
      styleRange(ws, rr, 1, rr, 3, ROW_VALUE_STYLE);
    });

    styleRange(ws, complianceHeaderIdx, 0, complianceHeaderIdx, 3, TABLE_HEADER_STYLE);
    // The compliance text spans the full width instead of one narrow column.
    addMerge(ws, complianceHeaderIdx + 1, 0, complianceHeaderIdx + 1, 3);
    styleRange(ws, complianceHeaderIdx + 1, 0, complianceHeaderIdx + 1, 3, ROW_VALUE_STYLE);

    styleRange(ws, signOffIdx, 0, signOffIdx, 1, SUBSECTION_STYLE);
    styleRange(ws, signOffIdx + 1, 0, signOffIdx + 1, 3, FIELD_LABEL_STYLE);

    fitRows(ws, [34, 56, 24, 30]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "QC Results");
    return finalizeBlob(wb);
  }

  // ---------- Marketing Guide Sheet ----------
  // Structure matches three real reference documents checked directly (SUP,
  // FKT, BRS - an inflatable sports product, a pool chemical, and an office
  // chair) - all three share this exact section order, label wording, and
  // column layout, confirming it's the standard template rather than a
  // one-off: ITEM/BRAND NAME/DESCRIPTION -> INCLUSIONS table ->
  // CERTIFICATIONS -> PRODUCT SPECIFICATIONS (a second, more detailed
  // restatement of article/material/color/size plus any extra
  // category-specific specs) -> IMPORTANT FEATURES TO HIGHLIGHT, split into
  // two sub-groups (features competitors also have vs. underused
  // differentiators - mapped from the PD sheet's ESP vs. USP sections,
  // which this template's two sub-groups directly correspond to) ->
  // Weaknesses in our product -> Target group -> Sample Picture.
  //
  // Color convention, confirmed identical across all three reference files:
  // strong orange (#FF9900) marks top-level section dividers (ITEM, BRAND
  // NAME, PRODUCT SPECIFICATIONS, IMPORTANT FEATURES TO HIGHLIGHT,
  // Weaknesses in our product., Target group, Sample Picture); a lighter
  // orange (#F6B26B) marks sub-field labels within a section. Green is
  // never used anywhere in this document - that's specific to the
  // Pre-inspection/Pre-QC table headers.
  //
  // COLUMN LAYOUT - also confirmed identical in all three references, and
  // easy to get wrong (an earlier version did): everything from ITEM
  // through the end of PRODUCT SPECIFICATIONS has its label in column B,
  // with content/sub-columns starting at C. But starting at "IMPORTANT
  // FEATURES TO HIGHLIGHT" and continuing through Weaknesses/Target
  // group/Sample Picture, the WHOLE section shifts one column right -
  // labels and content both start at column C instead of B, with column B
  // left empty throughout. This is consistent across all three documents,
  // not a one-off formatting choice.
  function generateMarketingXlsx(brief) {
    const uspSection = findRawSection(brief, ["usp"]);
    const espSection = findRawSection(brief, ["esp"]);
    const uspLines = uspSection ? uspSection.text.split("\n").filter(Boolean) : [];
    const espLines = espSection ? espSection.text.split("\n").filter(Boolean) : [];

    const rows = [[]];
    // Every row is tagged with how it is laid out (matching the reference
    // sheets): a label cell plus a value merged across to column F, a full-
    // width divider, or one of the two small header rows with a Benefit
    // column. labelCol is 1 (B) at the top and 2 (C) from the features
    // section down.
    const meta = [];
    function push(row, kind, labelCol) { rows.push(row); meta.push({ row: rows.length - 1, kind, labelCol }); return rows.length - 1; }

    push(["", "ITEM", brief.item || ""], "strongField", 1);
    push(["", "BRAND NAME", brief.brandName || ""], "strongField", 1);
    push(["", "DESCRIPTION", brief.description || ""], "lightField", 1);

    rows.push([]);
    const inclusionsHeaderIdx = push(["", "INCLUSIONS", "Item/Part", "", "Benefit", "Expectation Setting"], "incHeader", 1);
    const inclusionLines = (brief.inclusions || "").split("\n").filter(Boolean);
    inclusionLines.forEach((line) => rows.push(["", "", line, "", "", ""]));

    rows.push([]);
    push(["", "CERTIFICATIONS\n(Link to the certificates)", ""], "lightField", 1);
    push(["", "Other Compliance related informations", ""], "lightField", 1);

    rows.push([]);
    push(["", "PRODUCT SPECIFICATIONS"], "divider", 1);
    rows.push([]);
    push(["", "Link POE Table", ""], "lightField", 1);
    push(["", "Article No.", brief.articleNo || ""], "lightField", 1);
    push(["", "Artworks/Image/Color", ""], "lightField", 1);
    push(["", "Material", brief.material || ""], "lightField", 1);
    push(["", "Color", brief.color || ""], "lightField", 1);
    push(["", "Size", dimensionText(brief) || ""], "lightField", 1);
    // Any tech specs beyond the one used for "Size" (e.g. operating
    // pressure, load capacity, coverage) - the category-specific extras
    // all three reference documents include here.
    const dimLabelMatch = /dimension|size/i;
    (brief.techSpecs || []).filter((s) => !dimLabelMatch.test(s.label)).forEach((spec) => {
      push(["", spec.label, spec.value], "lightField", 1);
    });
    push(["", "Packaging size", packagingSizeText(brief)], "lightField", 1);

    rows.push([]);
    push(["", "", "IMPORTANT FEATURES TO HIGHLIGHT\nMention all the product's features, no matter how small - better too many than too few. All features must be explained in a comprehensible way."], "divider", 2);
    const espIntroIdx = push(["", "", "These features are also offered by other competitors in the market.", "", "Benefit"], "intro", 2);
    espLines.forEach((line) => rows.push(["", "", line, "", ""]));
    const espRowCount = espLines.length;

    rows.push([]);
    const uspIntroIdx = push(["", "", "These features are underutilized. Can be positioned as USPs", "", "Benefit"], "intro", 2);
    uspLines.forEach((line) => rows.push(["", "", line, "", ""]));
    const uspRowCount = uspLines.length;

    rows.push([]);
    push(["", "", "Weaknesses in our product."], "divider", 2);
    // Genuinely not PD-sheet-derivable: the PD sheet's review-analysis data
    // is about competitor products, not an honest self-assessment of this
    // specific design's own limitations - left blank rather than reused
    // from a different (related but not equivalent) source.
    rows.push(["", ""]);

    rows.push([]);
    push(["", "", "Target group"], "divider", 2);
    rows.push(["", ""]);

    rows.push([]);
    push(["", "", "Sample Picture"], "divider", 2);
    rows.push(["", "", brief.includeImages && brief.images && (brief.images.hero.length || brief.images.feature.length)
      ? "See the Product Brief's Product Image(s) and Reference Images for source photos to use here."
      : ""]);

    const ws = XLSX.utils.aoa_to_sheet(rows);
    const widths = [4, 22, 32, 20, 34, 30];
    setColWidths(ws, widths);

    // In the reference sheets ITEM / BRAND NAME carry an orange value that
    // is NOT bold, and every other value is plain white - only the labels
    // are filled.
    const STRONG_VALUE_STYLE = { fill: TITLE_STYLE.fill, border: ALL_BORDERS, alignment: { vertical: "top", wrapText: true } };
    meta.forEach(({ row: rr, kind, labelCol }) => {
      if (kind === "strongField") {
        styleRange(ws, rr, labelCol, rr, labelCol, TITLE_STYLE);
        addMerge(ws, rr, labelCol + 1, rr, 5);
        styleRange(ws, rr, labelCol + 1, rr, 5, STRONG_VALUE_STYLE);
      } else if (kind === "lightField") {
        styleRange(ws, rr, labelCol, rr, labelCol, SUBSECTION_STYLE);
        addMerge(ws, rr, labelCol + 1, rr, 5);
        styleRange(ws, rr, labelCol + 1, rr, 5, ROW_VALUE_STYLE);
      } else if (kind === "divider") {
        // The label itself spans the whole row (otherwise it is clipped to
        // its own narrow column).
        addMerge(ws, rr, labelCol, rr, 5);
        styleRange(ws, rr, labelCol, rr, 5, TITLE_STYLE);
      } else if (kind === "incHeader") {
        addMerge(ws, rr, 2, rr, 3);
        styleRange(ws, rr, labelCol, rr, 5, SUBSECTION_STYLE);
      } else if (kind === "intro") {
        addMerge(ws, rr, 2, rr, 3);
        styleRange(ws, rr, 2, rr, 5, SUBSECTION_STYLE);
      }
    });

    inclusionLines.forEach((_, i) => {
      const rr = inclusionsHeaderIdx + 1 + i;
      addMerge(ws, rr, 2, rr, 3);
      styleRange(ws, rr, 2, rr, 5, ROW_VALUE_STYLE);
    });
    for (let i = 0; i < espRowCount; i++) {
      const rr = espIntroIdx + 1 + i;
      addMerge(ws, rr, 2, rr, 3);
      styleRange(ws, rr, 2, rr, 5, ROW_VALUE_STYLE);
    }
    for (let i = 0; i < uspRowCount; i++) {
      const rr = uspIntroIdx + 1 + i;
      addMerge(ws, rr, 2, rr, 3);
      styleRange(ws, rr, 2, rr, 5, ROW_VALUE_STYLE);
    }

    // The three freeform blank-or-filled value rows (Weaknesses, Target
    // group, Sample Picture) sit right under their dividers.
    const dividers = meta.filter((m) => m.kind === "divider");
    dividers.slice(-3).forEach(({ row }) => {
      const rr = row + 1;
      addMerge(ws, rr, 2, rr, 5);
      styleRange(ws, rr, 2, rr, 5, ROW_VALUE_STYLE);
    });

    fitRows(ws, widths);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Marketing Guide");
    return finalizeBlob(wb);
  }

  global.PBXlsxGen = { generatePreInspectionXlsx, generatePreQCXlsx, generateMarketingXlsx };
})(window);
