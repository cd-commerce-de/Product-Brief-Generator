/* xlsx-parser.js
 * Parser for CD Commerce "Product Development Sheet" workbooks.
 *
 * Primary target: the POST-OPTIMIZED layout (KPM/AKP-style) - a single
 * "Final Product Specifications" tab that already contains our target specs
 * from the start, laid out feature-by-feature with an inline comparison
 * column against competitor specs (e.g. a "Material" row followed by
 * sub-rows like "Top layer | 4cm memory foam (ours) | 25D PU foam (Kesser)").
 * This is the standard going forward, per team direction.
 *
 * Legacy fallback: the PRE-OPTIMIZED layout (LTF/SUP-style) - specs were
 * split across a separate "Initial Product Specifications" tab (no
 * comparison columns) and a "Final Product Specifications" tab that was
 * often left as an unfilled template. Still supported so old SKUs remain
 * usable, but not the sheet to design new PD sheets around.
 */
(function (global) {
  function cellText(v) {
    if (v === null || v === undefined) return "";
    if (v instanceof Date) return v.toISOString().slice(0, 10);
    return String(v).trim();
  }

  function rowToArray(row, colCount) {
    const out = [];
    for (let c = 0; c < colCount; c++) out.push(cellText(row[c]));
    return out;
  }

  function sheetToRows(ws) {
    return XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: null });
  }

  function isSectionHeaderRow(cells) {
    const nonEmpty = cells.filter((c) => c !== "");
    if (nonEmpty.length !== 1) return false;
    const text = nonEmpty[0];
    if (text.length < 4) return false;
    // Header rows are (mostly) ALL CAPS labels like "COMPONENT INFORMATION"
    // or "USPs TO HIGHLIGHT" (note the trailing lowercase 's' - ratio-based
    // rather than strict equality so plurals like that still count).
    const letters = text.replace(/[^A-Za-z]/g, "");
    if (!letters) return false;
    const upperCount = (letters.match(/[A-Z]/g) || []).length;
    return upperCount / letters.length >= 0.75;
  }

  function findSheet(workbook, predicate) {
    const name = workbook.SheetNames.find(predicate);
    return name ? workbook.Sheets[name] : null;
  }

  function parseOverviewSheet(ws) {
    if (!ws) return {};
    const rows = sheetToRows(ws);
    const colCount = Math.max(...rows.map((r) => r.length), 2);
    const result = {};
    let currentLabel = null;
    rows.forEach((rawRow) => {
      const cells = rowToArray(rawRow, colCount);
      const label = cells[0];
      const rest = cells.slice(1).filter((c) => c !== "");
      if (label) {
        currentLabel = label.replace(/:\s*$/, "");
        if (!result[currentLabel]) result[currentLabel] = [];
        result[currentLabel].push(...rest);
      } else if (currentLabel && rest.length) {
        result[currentLabel].push(...rest);
      }
    });
    return result;
  }

  // Splits a spec sheet into { title, lines: [ [cell,cell,...], ... ] } blocks
  function parseSpecSheetIntoSections(ws) {
    if (!ws) return [];
    const rows = sheetToRows(ws);
    const colCount = Math.max(...rows.map((r) => r.length), 2);
    const sections = [];
    let current = null;
    rows.forEach((rawRow) => {
      const cells = rowToArray(rawRow, colCount);
      if (cells.every((c) => c === "")) return; // skip blank rows
      if (isSectionHeaderRow(cells)) {
        current = { title: cells.find((c) => c !== ""), lines: [] };
        sections.push(current);
      } else {
        if (!current) { current = { title: "General", lines: [] }; sections.push(current); }
        current.lines.push(cells);
      }
    });
    return sections;
  }

  function sectionToPlainText(section) {
    return section.lines
      .map((cells) => cells.filter((c) => c !== "").join(" — "))
      .filter((l) => l.length)
      .join("\n");
  }

  function textVolume(ws) {
    if (!ws) return 0;
    const rows = sheetToRows(ws);
    let n = 0;
    rows.forEach((r) => r.forEach((c) => { n += cellText(c).length; }));
    return n;
  }

  function findSectionsMatching(sections, keywords) {
    return sections.filter((s) =>
      keywords.some((k) => s.title.toLowerCase().includes(k.toLowerCase()))
    );
  }

  // --- Component Information parsing -------------------------------------
  // Row shapes seen in practice (0-indexed columns):
  //   [component/SKU, FEATURE label, value, ...]                     <- leaf
  //   [component/SKU, FEATURE label, "", ourSpec, competitorSpec...] <- group header
  //   ["", "", sub-feature name, ourValue, competitorValue...]       <- group sub-row
  //   ["", "", bullet text, "", ""]                                  <- plain continuation
  // "Group header" is distinguished from "leaf" by: description column (2)
  // empty AND columns 3+ hold short (<=30 char) strings - those are read as
  // competitor/brand column labels, not data, and the *first* of them is
  // where "our" value lives on the sub-rows that follow.
  function parseComponentInformation(section) {
    const featureMap = {}; // label -> { isGroup, value, subFeatures:[{name,value}] }
    const order = [];
    let currentLabel = null;
    let currentGroup = null;

    (section ? section.lines : []).forEach((cells) => {
      const c1 = cells[1] || "";
      const c2 = cells[2] || "";
      const extra = cells.slice(3).filter((c) => c !== "");

      if (c1) {
        currentLabel = c1;
        if (!featureMap[currentLabel]) order.push(currentLabel);
        const looksLikeGroupHeader = !c2 && extra.length > 0 && extra.every((e) => e.length <= 30);
        if (looksLikeGroupHeader) {
          featureMap[currentLabel] = { isGroup: true, subFeatures: [] };
          currentGroup = featureMap[currentLabel];
        } else {
          const val = [c2, ...extra].filter(Boolean).join(" ");
          featureMap[currentLabel] = { isGroup: false, value: val };
          currentGroup = null;
        }
      } else if (c2 || extra.length) {
        if (currentGroup) {
          const ourValue = extra.length ? extra[0] : "";
          if (c2 || ourValue) currentGroup.subFeatures.push({ name: c2, value: ourValue });
        } else if (currentLabel && featureMap[currentLabel]) {
          const text = [c2, ...extra].filter(Boolean).join(" ");
          featureMap[currentLabel].value = featureMap[currentLabel].value
            ? `${featureMap[currentLabel].value}\n${text}`
            : text;
        }
      }
    });

    return { featureMap, order };
  }

  function renderFeature(entry) {
    if (!entry) return "";
    if (entry.isGroup) {
      return entry.subFeatures
        .map((sf) => (sf.name ? `${sf.name}: ${sf.value}` : sf.value))
        .filter(Boolean)
        .join("\n");
    }
    return entry.value || "";
  }

  function getFeature(featureMap, keywords) {
    const key = Object.keys(featureMap).find((k) =>
      keywords.some((kw) => k.toLowerCase().includes(kw))
    );
    return key ? featureMap[key] : null;
  }

  // Pulls a "Technical Specifications" mini-table, e.g.:
  // (null,'Technical Specifications',null,'LTF01','LTF02')
  // (null,null,'Dehumidification capacity','18L/24H','30L/24H')
  function extractTechSpecRows(sections) {
    const out = [];
    sections.forEach((section) => {
      let inTechBlock = false;
      section.lines.forEach((cells) => {
        const joined = cells.join(" ").toLowerCase();
        if (joined.includes("technical specification")) { inTechBlock = true; return; }
        if (inTechBlock) {
          const nonEmpty = cells.filter((c) => c !== "");
          if (nonEmpty.length === 0) { inTechBlock = false; return; }
          const label = nonEmpty[0];
          const values = nonEmpty.slice(1);
          if (values.length === 0) { inTechBlock = false; return; }
          out.push({ label, value: values.join(" / ") });
        }
      });
    });
    return out;
  }

  // Pulls the "Review Analyses" / known-field-complaints table that's built
  // into post-optimized sheets under "CONCERNS TO BE DISCUSSED WITH SUPPLIER"
  // - rows shaped [ISSUE, "", RECOMMENDED SOLUTION, "", # OF MENTIONS].
  function extractKnownIssues(sections) {
    const concernSections = findSectionsMatching(sections, ["concerns to be discussed", "review analys"]);
    const out = [];
    concernSections.forEach((section) => {
      section.lines.forEach((cells) => {
        const nonEmpty = cells.filter((c) => c !== "");
        if (nonEmpty.length < 2) return;
        if (/^issue$/i.test(nonEmpty[0])) return; // header row
        const mentionsRaw = nonEmpty[nonEmpty.length - 1];
        const mentions = /^\d+(\.\d+)?$/.test(mentionsRaw) ? Number(mentionsRaw) : null;
        const issue = cells[0] || "";
        const solution = cells[2] || nonEmpty[1] || "";
        if (!issue || issue.length < 15) return; // skip stray short/header fragments
        out.push({ issue, solution, mentions });
      });
    });
    out.sort((a, b) => (b.mentions || 0) - (a.mentions || 0));
    return out;
  }

  function guessField(overview, keys) {
    for (const k of keys) {
      const match = Object.keys(overview).find((label) => label.toLowerCase().includes(k));
      if (match) return overview[match].join(", ");
    }
    return "";
  }

  // "Quality InspectionPO Informatio[n]" tab: numbered free-text notes
  // specific to this product that QC must observe (e.g. a load/rope test).
  function parseQualityInspectionSheet(ws) {
    if (!ws) return [];
    const rows = sheetToRows(ws);
    const colCount = Math.max(...rows.map((r) => r.length), 2);
    const notes = [];
    rows.forEach((rawRow) => {
      const cells = rowToArray(rawRow, colCount);
      const first = cells[0];
      if (/^\d+\.?$/.test(first)) {
        const text = cells.slice(1).filter((c) => c !== "").join(" ");
        if (text) notes.push(text);
      }
    });
    return notes;
  }

  function parseWorkbookFile(arrayBuffer) {
    const workbook = XLSX.read(arrayBuffer, { type: "array", cellDates: true });

    const overviewWs = findSheet(workbook, (n) => /overview/i.test(n));
    const overview = parseOverviewSheet(overviewWs);

    const finalWs = findSheet(workbook, (n) => /final product spec/i.test(n));
    const initialWs = findSheet(workbook, (n) => /initial product spec/i.test(n));
    const qualityWs = findSheet(workbook, (n) => /quality inspection/i.test(n));

    let specWs = finalWs;
    let specSource = "Final Product Specifications";
    const finalVol = textVolume(finalWs);
    const initialVol = textVolume(initialWs);
    // Post-optimized sheets only have a "Final" tab and it's fully populated
    // from the start, so this branch only fires for legacy pre-optimized
    // sheets where "Final" was left as an unfilled template.
    if (initialWs && initialVol > finalVol * 1.3) {
      specWs = initialWs;
      specSource = "Initial Product Specifications (legacy layout - Final tab was empty)";
    } else if (!finalWs && initialWs) {
      specWs = initialWs;
      specSource = "Initial Product Specifications (legacy layout)";
    }

    const sections = specWs ? parseSpecSheetIntoSections(specWs) : [];
    const techSpecRows = extractTechSpecRows(sections);
    const knownIssues = extractKnownIssues(sections);
    const qualityInspectionNotes = parseQualityInspectionSheet(qualityWs);

    const componentSection = findSectionsMatching(sections, ["component"])[0];
    const { featureMap, order } = parseComponentInformation(componentSection);

    const materialEntry = getFeature(featureMap, ["material"]);
    const colorEntry = getFeature(featureMap, ["color"]);
    const inclusionsEntry = getFeature(featureMap, ["inclusion"]);

    const packagingSection = findSectionsMatching(sections, ["packaging"]);
    const packagingText = packagingSection.map(sectionToPlainText).join("\n\n");
    const uspSections = findSectionsMatching(sections, ["usp", "esp"]);
    const uspText = uspSections.map(sectionToPlainText).join("\n\n");

    // Material & Workmanship Instructions = every component-info feature
    // (Material, Dimensions, Color, Features, etc.) rendered in full, plus
    // USPs/ESPs - mirrors how the reference LTF brief combines these into
    // one narrative block, minus the raw competitor-comparison noise.
    const materialInstructions = [
      ...order
        .filter((k) => !/inclusion/i.test(k) && !/technical spec/i.test(k))
        .map((k) => `${k.toUpperCase()}\n${renderFeature(featureMap[k])}`),
      uspText ? `USPs / ESPs\n${uspText}` : "",
    ].filter(Boolean).join("\n\n");

    const knownIssuesText = knownIssues
      .slice(0, 8)
      .map((k, i) => `${i + 1}. ${k.issue}${k.mentions ? ` (${k.mentions} mentions)` : ""}\n   -> Preventive spec/action: ${k.solution}`)
      .join("\n\n");

    return {
      meta: { specSource, sheetNames: workbook.SheetNames },
      overview,
      sections: sections.map((s) => ({ title: s.title, text: sectionToPlainText(s) })),
      guessed: {
        articleNo: guessField(overview, ["product ident"]),
        item: guessField(overview, ["product name"]),
        description: guessField(overview, ["product description"]),
        material: renderFeature(materialEntry).slice(0, 400),
        color: renderFeature(colorEntry),
        inclusions: renderFeature(inclusionsEntry),
        packaging: packagingText,
        materialInstructions,
        techSpecRows,
        knownIssues,
        knownIssuesText,
        qualityInspectionNotes,
      },
    };
  }

  global.PBParser = { parseWorkbookFile };
})(window);
