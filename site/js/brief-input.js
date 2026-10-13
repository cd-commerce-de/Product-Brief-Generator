/* brief-input.js
 * Turns the parsed "Product Brief Input" tab (see parseBriefInputSheet in
 * xlsx-parser.js) into the fields of a Product Brief. Kept separate from the
 * parser (which only reads cells) and from the Word generator (which only
 * draws) so the mapping rules live in one place and can be tested alone.
 *
 * Rules, each taken from how the team's approved SSC brief is built:
 *  - Article No. = the SKU column; Color lines = "<SKU suffix> - <colour>"
 *    (SSC01-SW + Black => "SW - Black").
 *  - Inclusions drop a leading "1 x " ("1 x Allen key" => "Allen key").
 *  - Part B rows become numbered instructions per component; text after
 *    "SKU-specific detail" is printed under its instruction (lettered a./b./c.
 *    when the instruction ends with ":", plain lines otherwise).
 *  - The SKU table supplies the carton dimensions/weights that the approved
 *    brief states in Packaging: an "individual carton" statement is added,
 *    and "(dimensions and weight per SKU table)" in the master-carton row is
 *    replaced with the real numbers.
 *  - Part C rows are merged with default QC rows (Odor, Finish/Paint, Bad
 *    printing, Packaging): a sheet row overrides the default of the same
 *    kind and keeps its own position; defaults not covered are added.
 *  - Compliance text = test-cost clause + requirement list + test report
 *    number, all taken from the sheet (nothing else is added).
 */
(function (global) {
  // ---------- small helpers ----------
  function fmtNum(v) {
    const s = String(v == null ? "" : v).trim();
    if (!/^-?\d+(\.\d+)?$/.test(s)) return s;
    return String(Number(s)); // "4.50" -> "4.5", "19.50" -> "19.5", "4.0" -> "4"
  }

  function lines(text) {
    return String(text || "").split("\n").map((l) => l.trim()).filter(Boolean);
  }

  // "SSC01-SW" with product ident "SSC01" -> "SW"; falls back to the part
  // after the first dash, or the whole SKU.
  function skuSuffix(sku, ident) {
    const s = String(sku || "").trim();
    if (ident && s.toLowerCase().startsWith(ident.toLowerCase() + "-")) return s.slice(ident.length + 1);
    const m = /^[A-Za-z]+\d+-(.+)$/.exec(s);
    return m ? m[1] : s;
  }

  function uniform(skus, key) {
    const vals = skus.map((s) => fmtNum(s[key])).filter(Boolean);
    return vals.length > 0 && vals.every((v) => v === vals[0]) ? vals[0] : null;
  }

  // ---------- default QC rows ----------
  // The sheet says these four rows are added automatically but doesn't give
  // their wording, so these are neutral placeholders - replace them by adding
  // a row of the same kind to Part C (it overrides the default), or edit them
  // in Step 2. They are not taken from any existing brief.
  const DEFAULT_QC = [
    {
      kind: "odor", match: /odou?r|smell/i, first: true,
      parameter: "Odor / smell",
      acceptable: "No noticeable strong or unusual odor",
      notAcceptable: "Persistent strong or unusual odor",
    },
    {
      kind: "finish", match: /finish|paint|coat/i,
      parameter: "Finish / Paint",
      acceptable: "Even finish, matches the approved sample; no chips, scratches or bare spots",
      notAcceptable: "Uneven finish; chips, scratches, discoloration or bare spots",
    },
    {
      kind: "print", match: /print|artwork/i,
      parameter: "Printing",
      acceptable: "Print is sharp, complete and positioned as in the approved artwork",
      notAcceptable: "Smeared, incomplete or misplaced print",
    },
    {
      kind: "packaging", match: /packag|carton/i,
      parameter: "Packaging",
      acceptable: "Packaging intact, no deformation or damage",
      notAcceptable: "Deformed or damaged packaging",
    },
  ];

  function mergeQcRows(partCRows) {
    const rows = (partCRows || []).map((r) => ({
      parameter: r.parameter, acceptable: r.acceptable || "", notAcceptable: r.notAcceptable || "",
      method: r.method || "", images: r.images || [],
    }));
    const covered = new Set();
    rows.forEach((r) => {
      const d = DEFAULT_QC.find((x) => x.match.test(r.parameter));
      if (d) covered.add(d.kind);
    });
    const out = [];
    DEFAULT_QC.filter((d) => d.first && !covered.has(d.kind)).forEach((d) => {
      out.push({ parameter: d.parameter, acceptable: d.acceptable, notAcceptable: d.notAcceptable, method: "", images: [] });
    });
    rows.forEach((r) => out.push(r));
    DEFAULT_QC.filter((d) => !d.first && !covered.has(d.kind)).forEach((d) => {
      out.push({ parameter: d.parameter, acceptable: d.acceptable, notAcceptable: d.notAcceptable, method: "", images: [] });
    });
    return out;
  }

  // ---------- components ----------
  function normItem(it) {
    return {
      text: it.text, skuDetail: it.skuDetail || "", caption: it.caption || "",
      status: it.status || "", appliesTo: it.appliesTo || "", images: it.images || [],
    };
  }

  function cartonStatements(skus, items) {
    const mentions = (re) => items.some((it) => re.test(it.text));
    const result = { items: items.slice() };
    const withItem = skus.filter((s) => s.itemCarton);
    const idxMaster = result.items.findIndex((it) => /master carton/i.test(it.text));

    // 1) individual carton dimensions / weight (only if not already written)
    if (withItem.length && !mentions(/individual carton dimension|carton dimensions/i)) {
      const dims = uniform(withItem, "itemCarton");
      const wt = uniform(withItem, "itemCartonWeight");
      let item;
      if (dims) {
        item = {
          text: `The individual carton dimensions should be ${dims} cm` + (wt ? ` and packaging weight is ${wt} kg.` : "."),
          skuDetail: "",
        };
      } else {
        item = {
          text: "The individual carton dimensions and packaging weight per SKU:",
          skuDetail: withItem.map((s) => `${s.sku}: ${fmtNum(s.itemCarton)} cm` + (s.itemCartonWeight ? `, ${fmtNum(s.itemCartonWeight)} kg` : "")).join("\n"),
        };
      }
      item = { caption: "", status: "Final", appliesTo: "All SKUs", images: [], generated: true, ...item };
      if (idxMaster >= 0) result.items.splice(idxMaster, 0, item); else result.items.push(item);
    }

    // 2) master carton row that says "per SKU table": put the numbers in
    const mi = result.items.findIndex((it) => /master carton/i.test(it.text) && /sku table/i.test(it.text));
    const withMaster = skus.filter((s) => s.masterCarton);
    if (mi >= 0) {
      const it = result.items[mi];
      const pcsMatch = /(\d+)\s*(?:pieces|pcs)/i.exec(it.text);
      const dims = uniform(withMaster, "masterCarton");
      const wt = uniform(withMaster, "masterCartonWeight");
      const pcs = uniform(skus.filter((s) => s.pcsPerMaster), "pcsPerMaster") || (pcsMatch && pcsMatch[1]) || "";
      if (withMaster.length && dims) {
        result.items[mi] = {
          ...it,
          text: `Master carton: ${dims} cm` + (wt ? `, ${wt} kg` : "") + (pcs ? `, ${pcs} pieces per master carton.` : "."),
        };
      } else if (withMaster.length) {
        result.items[mi] = {
          ...it,
          text: "Master carton" + (pcs ? `: ${pcs} pieces per master carton.` : ":") + " Dimensions and weight per SKU:",
          skuDetail: withMaster.map((s) => `${s.sku}: ${fmtNum(s.masterCarton)} cm` + (s.masterCartonWeight ? `, ${fmtNum(s.masterCartonWeight)} kg` : "")).join("\n"),
        };
      } else {
        // No SKU-table numbers to insert: drop the dangling pointer.
        result.items[mi] = { ...it, text: it.text.replace(/\s*\([^)]*SKU table[^)]*\)/i, "") };
      }
    }
    return result.items;
  }

  function buildComponents(bi) {
    return (bi.components || []).map((c) => {
      let items = c.items.map(normItem);
      if (c.kind === "packaging") items = cartonStatements(bi.skus || [], items);
      return { name: c.name, kind: c.kind, items };
    });
  }

  function packagingText(components) {
    const pk = (components || []).filter((c) => c.kind === "packaging");
    const out = [];
    pk.forEach((c) => c.items.forEach((it) => {
      out.push(it.text);
      lines(it.skuDetail).forEach((l) => out.push(l));
    }));
    return out.join("\n");
  }

  // ---------- highlighting ----------
  // The approved briefs mark "... will be sent separately" sentences in
  // yellow (files the supplier still has to receive). A row whose status is
  // Pending is highlighted whole. Returns [{ text, hl }] segments.
  function splitHighlights(text, status) {
    const t = String(text || "");
    if (/^pending$/i.test(String(status || "").trim())) return [{ text: t, hl: true }];
    const sentences = t.match(/[^.!?]+[.!?]+["')\]]*\s*|[^.!?]+$/g) || [t];
    const segs = [];
    const push = (txt, hl) => {
      if (!txt) return;
      const last = segs[segs.length - 1];
      if (last && last.hl === hl) last.text += txt; else segs.push({ text: txt, hl });
    };
    sentences.forEach((s) => {
      if (/\bsent separately\b/i.test(s)) {
        const core = s.replace(/\s+$/, "");
        push(core, true);
        push(s.slice(core.length), false);
      } else {
        push(s, false);
      }
    });
    return segs;
  }

  // ---------- compliance ----------
  // Light markup understood by the Word generator: **bold**; lines starting
  // with "-" stay as typed.
  function boldAmounts(text) {
    return text.replace(/(\d[\d,.]*\s?(?:USD|EUR|€|\$)|\d[\d,.]*\s+pieces)/g, "**$1**");
  }

  // Built only from what the sheet says: the test-cost clause, the list of
  // requirements and the test report number. Nothing else is added - any
  // standard wording a brief needs must be in the sheet (or typed in Step 3).
  function composeCompliance(bi) {
    const out = [];
    if (bi.testCostClause) out.push(boldAmounts(bi.testCostClause), "");
    const reqs = lines(bi.complianceRequirements);
    reqs.forEach((r) => out.push(`**${r.startsWith("-") ? r : "-" + r}**`));
    if (reqs.length) out.push("");
    if (bi.testReportNo) out.push(`Test Report No.: ${bi.testReportNo}`);
    return out.join("\n").replace(/\n{3,}/g, "\n\n").trim();
  }

  // ---------- main ----------
  function compose(bi) {
    const skus = bi.skus || [];
    const components = buildComponents(bi);
    const colorLines = skus
      .filter((s) => s.color)
      .map((s) => `${skuSuffix(s.sku, bi.productIdent)} - ${s.color}`);
    const inclusions = lines(bi.inclusions).map((l) => l.replace(/^\s*1\s*[x×]\s*/i, "")).join("\n");
    return {
      articleNo: skus.length ? skus.map((s) => s.sku).join("\n") : bi.productIdent || "",
      item: bi.item || "",
      description: bi.description || "",
      material: bi.material || "",
      color: colorLines.join("\n"),
      inclusions,
      components,
      skuTable: skus.map((s) => ({ ...s })),
      category: bi.category || "",
      qcRows: mergeQcRows(bi.qcRows),
      compliance: composeCompliance(bi),
      packaging: packagingText(components),
    };
  }

  // Writes the composed fields onto a brief. Fields the sheet left empty
  // keep whatever the PD-sheet parser or the person already put there.
  function apply(b, bi) {
    const c = compose(bi);
    ["articleNo", "item", "description", "material", "color", "inclusions"].forEach((k) => {
      if (c[k]) b[k] = c[k];
    });
    b.components = c.components;
    b.skuTable = c.skuTable;
    b.category = c.category;
    b.qcRows = c.qcRows;
    if (c.packaging) b.packaging = c.packaging;
    // Compliance text from the sheet replaces the box only when the person
    // hasn't edited it since the last import (or it's empty).
    if (!(b.compliance || "").trim() || b.compliance === b.complianceAuto) b.compliance = c.compliance;
    b.complianceAuto = c.compliance;
    return c;
  }

  global.PBBriefInput = {
    compose, apply, packagingText, splitHighlights, mergeQcRows, composeCompliance,
    fmtNum, skuSuffix, DEFAULT_QC,
  };
})(window);
