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
    const out = XLSX.write(wb, { bookType: "xlsx", type: "array" });
    return { ws, blob: new Blob([out], { type: "application/octet-stream" }) };
  }

  function finalizeBlob(wb) {
    const out = XLSX.write(wb, { bookType: "xlsx", type: "array" });
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
        `~Measuring tape\n~Vernier caliper / weighing scale\n\nAgainst spec: ${findDimensionSpec(brief) || "(see Technical Specifications in Product Brief)"}`,
        "Available/Not Available",
      ],
      [
        "Visual inspection",
        "~Good lighting setup\n~Clean inspection table\n~All reference materials:\n     - Approved design artwork\n     - Approved color references\n     - Product briefing document\n     - Golden sample or retention sample",
        "Available/Not Available",
      ],
      [
        "Packaging",
        "~Complete carton with all packaging materials\n~Tape and sealing materials\n\nSpec: " + (brief.packaging || "(see Product Brief)"),
        "Available/Not Available",
      ],
    ];

    if (brief.qualityInspectionNotes) {
      testRows.push(["Product-specific test (from PD sheet)", brief.qualityInspectionNotes, "Available/Not Available"]);
    }
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
    if (brief.knownIssues) {
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
      ["Dimensions check", "~Confirm product dimensions match the approved specification.\n\nSpec: " + (findDimensionSpec(brief) || "(see Technical Specifications)"), "", ""],
      ["Odor check", "~Smell by nose within 10cm.", "", ""],
      ["Printing, artwork, colors check", "~Compare product print with approved artwork.\n~Check logo orientation, position, sharpness, and color.\n\nColor spec: " + (brief.color || ""), "", ""],
    ];

    if (brief.qualityInspectionNotes) {
      testRows.push(["Product-specific test (from PD sheet)", brief.qualityInspectionNotes, "", ""]);
    }
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
    packingRows.push(["Carton dimensions / weight", brief.packaging || "", "", ""]);
    packingRows.forEach((row) => rows.push(row));

    const complianceHeaderIdx = rows.length + 1;
    rows.push([], ["Required Certification / Compliance Declarations", "", "AVAILABLE/NOT AVAILABLE", "Remarks"]);
    rows.push([brief.compliance || "", "", "", ""]);

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
    styleRange(ws, complianceHeaderIdx + 1, 0, complianceHeaderIdx + 1, 3, ROW_VALUE_STYLE);

    styleRange(ws, signOffIdx, 0, signOffIdx, 1, SUBSECTION_STYLE);
    styleRange(ws, signOffIdx + 1, 0, signOffIdx + 1, 3, FIELD_LABEL_STYLE);

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "QC Results");
    return finalizeBlob(wb);
  }

  // ---------- Marketing Guide Sheet ----------
  function generateMarketingXlsx(brief) {
    const uspSection = findRawSection(brief, ["usp"]);
    const espSection = findRawSection(brief, ["esp"]);
    const highlightLines = []
      .concat(uspSection ? uspSection.text.split("\n") : [])
      .concat(espSection ? espSection.text.split("\n") : [])
      .filter(Boolean);

    const topFields = [
      ["ITEM", brief.item || ""],
      ["BRAND NAME", brief.brandName || ""],
    ];
    const rows = [
      [],
      ["", "ITEM", brief.item || ""],
      ["", "BRAND NAME", brief.brandName || ""],
      ["", "DESCRIPTION", brief.description || ""],
      ["", "ARTICLE NO.", brief.articleNo || ""],
      ["", "MATERIAL", brief.material || ""],
      ["", "SIZE", findDimensionSpec(brief) || ""],
      ["", "COLOR", brief.color || ""],
      ["", "PACKAGING SIZE", findPackagingSize(brief)],
    ];
    const topFieldEnd = rows.length - 1;

    let techStart = null;
    if (brief.techSpecs && brief.techSpecs.length) {
      rows.push([], ["", "TECHNICAL SPECIFICATIONS"]);
      techStart = rows.length - 1;
      brief.techSpecs.forEach((spec) => rows.push(["", "", `${spec.label}: ${spec.value}`]));
    }

    rows.push([], ["", "INCLUSIONS", "Item/Part", "Benefit", "Expectation Setting"]);
    const inclusionsHeaderIdx = rows.length - 1;
    const inclusionLines = (brief.inclusions || "").split("\n").filter(Boolean);
    inclusionLines.forEach((line) => {
      rows.push(["", "", line, "TODO: benefit copy", "TODO: expectation-setting note"]);
    });

    rows.push([], ["", "IMPORTANT FEATURES TO HIGHLIGHT", "", "Benefit", ""]);
    const featuresHeaderIdx = rows.length - 1;
    rows.push(["", "Mention all the product's features, no matter how small - better too many than too few. All features must be explained in a comprehensible way."]);
    const featureLines = highlightLines.length ? highlightLines : ["TODO: pull feature list from Product Brief"];
    featureLines.forEach((line) => rows.push(["", "", line, "TODO: benefit copy", ""]));

    rows.push([], ["", "Known weaknesses / trade-offs to acknowledge"]);
    const weaknessesHeaderIdx = rows.length - 1;
    if (brief.knownIssues) {
      brief.knownIssues.split("\n\n").forEach((block) => rows.push(["", block]));
    } else {
      rows.push(["", "TODO - what should buyers not expect this product to do?"]);
    }

    rows.push([], ["", "Target group"]);
    const targetGroupIdx = rows.length - 1;
    rows.push(["", "TODO: who is this for? (age, use-case, lifestyle)"]);

    rows.push([], ["", "Sample Pictures"]);
    const samplePicsIdx = rows.length - 1;
    rows.push(["", brief.includeImages && brief.images && (brief.images.hero.length || brief.images.feature.length)
      ? "See the Product Brief's Product Image(s) and Reference Images for source photos to use here."
      : "TODO: attach/link product photos"]);

    const ws = XLSX.utils.aoa_to_sheet(rows);
    setColWidths(ws, [3, 24, 38, 44, 36]);

    // ITEM/BRAND NAME rows get the strong orange (matches the reference
    // exactly - those two specifically, not the rest of the field list).
    [1, 2].forEach((rr) => {
      addMerge(ws, rr, 2, rr, 4);
      styleRange(ws, rr, 1, rr, 1, TITLE_STYLE);
      styleRange(ws, rr, 2, rr, 4, TITLE_STYLE);
    });
    for (let rr = 3; rr <= topFieldEnd; rr++) {
      addMerge(ws, rr, 2, rr, 4);
      styleRange(ws, rr, 1, rr, 1, SUBSECTION_STYLE);
      styleRange(ws, rr, 2, rr, 4, FIELD_VALUE_STYLE);
    }

    if (techStart !== null) {
      styleRange(ws, techStart, 1, techStart, 4, SUBSECTION_STYLE);
      brief.techSpecs.forEach((_, i) => {
        const rr = techStart + 1 + i;
        styleRange(ws, rr, 2, rr, 4, ROW_VALUE_STYLE);
      });
    }

    styleRange(ws, inclusionsHeaderIdx, 1, inclusionsHeaderIdx, 4, TABLE_HEADER_STYLE);
    inclusionLines.forEach((_, i) => {
      const rr = inclusionsHeaderIdx + 1 + i;
      styleRange(ws, rr, 2, rr, 4, ROW_VALUE_STYLE);
    });

    styleRange(ws, featuresHeaderIdx, 1, featuresHeaderIdx, 4, TABLE_HEADER_STYLE);
    styleRange(ws, featuresHeaderIdx + 1, 1, featuresHeaderIdx + 1, 4, ROW_VALUE_STYLE);
    featureLines.forEach((_, i) => {
      const rr = featuresHeaderIdx + 2 + i;
      styleRange(ws, rr, 2, rr, 4, ROW_VALUE_STYLE);
    });

    styleRange(ws, weaknessesHeaderIdx, 1, weaknessesHeaderIdx, 1, SUBSECTION_STYLE);
    styleRange(ws, targetGroupIdx, 1, targetGroupIdx, 1, SUBSECTION_STYLE);
    styleRange(ws, targetGroupIdx + 1, 1, targetGroupIdx + 1, 4, ROW_VALUE_STYLE);
    styleRange(ws, samplePicsIdx, 1, samplePicsIdx, 1, SUBSECTION_STYLE);
    styleRange(ws, samplePicsIdx + 1, 1, samplePicsIdx + 1, 4, ROW_VALUE_STYLE);

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Marketing Guide");
    return finalizeBlob(wb);
  }

  global.PBXlsxGen = { generatePreInspectionXlsx, generatePreQCXlsx, generateMarketingXlsx };
})(window);
