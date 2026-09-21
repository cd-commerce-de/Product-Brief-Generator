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
 */
(function (global) {
  function aoaToBlob(rows, sheetName) {
    const ws = XLSX.utils.aoa_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, sheetName);
    const out = XLSX.write(wb, { bookType: "xlsx", type: "array" });
    return new Blob([out], { type: "application/octet-stream" });
  }

  function findRawSection(brief, keywords) {
    const sections = (brief.raw && brief.raw.sections) || [];
    return sections.find((s) =>
      keywords.some((k) => s.title.toLowerCase().includes(k))
    );
  }

  function findDimensionSpec(brief) {
    const spec = (brief.techSpecs || []).find((r) => /dimension|size/i.test(r.label));
    return spec ? spec.value : "";
  }

  // ---------- Pre-inspection Briefing Form ----------
  function generatePreInspectionXlsx(brief) {
    const rows = [
      ["Pre-inspection Briefing Form"],
      ["Supplier Company Name", brief.supplierName || ""],
      ["Contact Person", brief.supplierContact || ""],
      ["Contact Email", brief.supplierEmail || ""],
      ["Product Name", brief.item || ""],
      ["Purchase Order (PO) Number", brief.poNumber || ""],
      ["Pre-QC check (in-house) DATE", brief.preQcDate || ""],
      ["Planned 3rd Party Inspection Date", brief.inspectionDate || ""],
      ["# of samples to be inspected", brief.sampleCount || ""],
      [],
      ["Test(s) to be Conducted", "Items to be Prepared for Testing", "Availability"],
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
      rows.push([
        "Product-specific test (from PD sheet)",
        brief.qualityInspectionNotes,
        "Available/Not Available",
      ]);
    }

    if (brief.knownIssues) {
      rows.push([], ["Known Market Complaints to Watch For (from PD review analysis)"]);
      brief.knownIssues.split("\n\n").forEach((block) => rows.push([block]));
    }

    return aoaToBlob(rows, "Pre-inspection");
  }

  // ---------- Pre-QC (Quality Control) Check ----------
  function generatePreQCXlsx(brief) {
    const rows = [
      ["Pre-QC (Quality Control) Check"],
      ["Supplier Company Name", brief.supplierName || ""],
      ["Contact Person", brief.supplierContact || ""],
      ["Contact Email", brief.supplierEmail || ""],
      ["Product Name", brief.item || ""],
      ["Purchase Order (PO) Number", brief.poNumber || ""],
      ["Target 3rd party inspection date", brief.inspectionDate || ""],
      ["# of samples to be inspected", brief.sampleCount || ""],
      [],
      ["Test/s", "Methods", "Result (PASSED/FAILED)", "Photos/Videos with Remarks"],
      [
        "Package completeness check",
        "~Verify that each set contains all required items and matches the approved Product Brief.\n~Quantity, specification, and general condition of each item are aligned.\n~Confirm nothing is substituted or missing.\n\nInclusions:\n" + (brief.inclusions || ""),
        "",
        "",
      ],
      [
        "Material / surface check",
        "~Inspect visible quality of materials against the Product Brief.\n\nSpec:\n" + (brief.material || ""),
        "",
        "",
      ],
      [
        "Dimensions check",
        "~Confirm product dimensions match the approved specification.\n\nSpec: " + (findDimensionSpec(brief) || "(see Technical Specifications)"),
        "",
        "",
      ],
      ["Odor check", "~Smell by nose within 10cm.", "", ""],
      [
        "Printing, artwork, colors check",
        "~Compare product print with approved artwork.\n~Check logo orientation, position, sharpness, and color.\n\nColor spec: " + (brief.color || ""),
        "",
        "",
      ],
    ];

    if (brief.qualityInspectionNotes) {
      rows.push(["Product-specific test (from PD sheet)", brief.qualityInspectionNotes, "", ""]);
    }

    rows.push(
      [],
      ["Packing List Overview", "Requirement", "ACTUAL", "Photo Proof"],
      ["Carton dimensions / weight", brief.packaging || "", "", ""],
      [],
      ["Required Certification / Compliance Declarations", "", "AVAILABLE/NOT AVAILABLE", "Remarks"],
      [brief.compliance || "", "", "", ""],
      [],
      ["Supplier Signature & Date"],
      ["Name", "Signature", "", "Date"]
    );

    return aoaToBlob(rows, "QC Results");
  }

  // ---------- Marketing Guide Sheet ----------
  function generateMarketingXlsx(brief) {
    const uspSection = findRawSection(brief, ["usp"]);
    const espSection = findRawSection(brief, ["esp"]);
    const highlightLines = []
      .concat(uspSection ? uspSection.text.split("\n") : [])
      .concat(espSection ? espSection.text.split("\n") : [])
      .filter(Boolean);

    const rows = [
      ["ITEM", brief.item || ""],
      ["BRAND NAME", brief.brandName || ""],
      ["DESCRIPTION", brief.description || ""],
      ["ARTICLE NO.", brief.articleNo || ""],
      ["MATERIAL", brief.material || ""],
      ["SIZE", findDimensionSpec(brief) || ""],
      ["COLOR", brief.color || ""],
      ["PACKAGING SIZE", brief.packaging || ""],
      [],
      ["INCLUSIONS", "Item/Part", "Benefit", "Expectation Setting"],
    ];

    (brief.inclusions || "").split("\n").filter(Boolean).forEach((line) => {
      rows.push(["", line, "TODO: benefit copy", "TODO: expectation-setting note"]);
    });

    rows.push(
      [],
      ["IMPORTANT FEATURES TO HIGHLIGHT", "", "Benefit", ""],
      ["Mention all the product's features, no matter how small - better too many than too few. All features must be explained in a comprehensible way."]
    );

    if (highlightLines.length) {
      highlightLines.forEach((line) => rows.push(["", line, "TODO: benefit copy", ""]));
    } else {
      rows.push(["", "TODO: pull feature list from Product Brief", "TODO: benefit copy", ""]);
    }

    rows.push(
      [],
      ["Known weaknesses / trade-offs to acknowledge", "TODO - see Known Market Complaints in the Product Brief for what NOT to overclaim"],
      [],
      ["Target group", "TODO: who is this for? (age, use-case, lifestyle)"],
      [],
      ["Sample Pictures", "TODO: attach/link product photos"]
    );

    return aoaToBlob(rows, "Marketing Guide");
  }

  global.PBXlsxGen = { generatePreInspectionXlsx, generatePreQCXlsx, generateMarketingXlsx };
})(window);
