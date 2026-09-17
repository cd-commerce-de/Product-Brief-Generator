/* app.js - UI glue */
(function () {
  const $ = (sel) => document.querySelector(sel);
  const $$ = (sel) => Array.from(document.querySelectorAll(sel));

  const B = () => PBState.brief;

  // ---------- Bullet-list helper buttons ----------
  // Generic: any button with data-bullet-for="<textareaId>" appends a new
  // bullet line to that textarea (or starts one if the field is empty),
  // then focuses it so the person can type immediately.
  $$("[data-bullet-for]").forEach((btn) => {
    btn.addEventListener("click", () => {
      const ta = document.getElementById(btn.dataset.bulletFor);
      if (!ta) return;
      const needsNewline = ta.value.length > 0 && !ta.value.endsWith("\n");
      ta.value += (needsNewline ? "\n" : "") + "• ";
      ta.focus();
      ta.selectionStart = ta.selectionEnd = ta.value.length;
    });
  });

  // ---------- Step navigation ----------
  function goToStep(n) {
    $$(".panel").forEach((p) => p.classList.add("hidden"));
    $(`#panel-${n}`).classList.remove("hidden");
    $$(".step").forEach((s) => {
      const step = Number(s.dataset.step);
      s.classList.toggle("active", step === n);
      s.classList.toggle("done", step < n);
    });
    if (n === 2) fillReviewForm();
    if (n === 4) renderVersionList("#versionList");
  }

  $$("[data-next]").forEach((btn) =>
    btn.addEventListener("click", () => { syncFormToState(); goToStep(Number(btn.dataset.next)); })
  );
  $$("[data-back]").forEach((btn) =>
    btn.addEventListener("click", () => { syncFormToState(); goToStep(Number(btn.dataset.back)); })
  );
  $$(".step").forEach((s) =>
    s.addEventListener("click", () => { syncFormToState(); goToStep(Number(s.dataset.step)); })
  );

  // ---------- Step 1: Import ----------
  const dropzone = $("#dropzone");
  const fileInput = $("#fileInput");

  dropzone.addEventListener("click", () => fileInput.click());
  dropzone.addEventListener("dragover", (e) => { e.preventDefault(); dropzone.classList.add("dragover"); });
  dropzone.addEventListener("dragleave", () => dropzone.classList.remove("dragover"));
  dropzone.addEventListener("drop", (e) => {
    e.preventDefault();
    dropzone.classList.remove("dragover");
    if (e.dataTransfer.files.length) handleFile(e.dataTransfer.files[0]);
  });
  fileInput.addEventListener("change", (e) => {
    if (e.target.files.length) handleFile(e.target.files[0]);
  });

  function setStatus(sel, msg, cls) {
    const el = $(sel);
    el.textContent = msg;
    el.className = "status" + (cls ? ` ${cls}` : "");
  }

  function handleFile(file) {
    if (!/\.xlsx$/i.test(file.name)) {
      setStatus("#importStatus", "Please upload an .xlsx file.", "error");
      return;
    }
    setStatus("#importStatus", "Parsing…");
    const reader = new FileReader();
    reader.onload = async (e) => {
      const buffer = e.target.result;
      try {
        const parsed = PBParser.parseWorkbookFile(buffer);
        applyParsedToState(parsed);
        renderImportPreview(parsed);
        setStatus("#importStatus", `Parsed successfully (source tab: ${parsed.meta.specSource}). Looking for images…`, "ok");
        $("#importPreview").classList.remove("hidden");

        // Image extraction is a separate pass (SheetJS doesn't expose
        // embedded drawings) and runs after the text parse so a slow/failed
        // image extraction never blocks the fields the person actually needs.
        try {
          const images = await PBImages.extractImages(buffer);
          B().images = images;
          renderImagesPreview(images);
          const total = images.hero.length + images.feature.length;
          setStatus("#importStatus", `Parsed successfully (source tab: ${parsed.meta.specSource}). Found ${total} image${total === 1 ? "" : "s"}.`, "ok");
        } catch (imgErr) {
          console.error(imgErr);
          setStatus("#importStatus", `Parsed successfully (source tab: ${parsed.meta.specSource}). Couldn't read images from this file.`, "ok");
        }
      } catch (err) {
        console.error(err);
        setStatus("#importStatus", "Couldn't parse this file — is it a standard PD sheet .xlsx? " + err.message, "error");
      }
    };
    reader.readAsArrayBuffer(file);
  }

  function renderImagesPreview(images) {
    const wrap = $("#imagesPreviewWrap");
    const gallery = $("#imagesPreview");
    gallery.innerHTML = "";
    const all = [...images.hero.map((im) => ({ ...im, label: "Product image" })),
                 ...images.feature.map((im) => ({ ...im, label: "Reference image" }))];
    if (!all.length) {
      wrap.classList.add("hidden");
      return;
    }
    wrap.classList.remove("hidden");
    all.forEach((im) => {
      const fig = document.createElement("figure");
      const img = document.createElement("img");
      img.src = im.dataUrl;
      const cap = document.createElement("figcaption");
      cap.textContent = im.label;
      fig.appendChild(img);
      fig.appendChild(cap);
      gallery.appendChild(fig);
    });
  }

  $("#f_includeImages").addEventListener("change", (e) => {
    B().includeImages = e.target.checked;
  });

  function applyParsedToState(parsed) {
    const b = B();
    b.raw = parsed;
    b.articleNo = parsed.guessed.articleNo || b.articleNo;
    b.item = parsed.guessed.item || b.item;
    b.description = parsed.guessed.description || b.description;
    b.material = parsed.guessed.material || b.material;
    b.color = parsed.guessed.color || b.color;
    b.inclusions = parsed.guessed.inclusions || b.inclusions;
    b.materialInstructions = parsed.guessed.materialInstructions || b.materialInstructions;
    b.packaging = parsed.guessed.packaging || b.packaging;
    if (parsed.guessed.techSpecRows && parsed.guessed.techSpecRows.length) {
      b.techSpecs = parsed.guessed.techSpecRows.map((r) => ({ label: r.label, value: r.value }));
    }
    b.knownIssues = parsed.guessed.knownIssuesText || b.knownIssues;
    b.qualityInspectionNotes = (parsed.guessed.qualityInspectionNotes || []).join("\n\n") || b.qualityInspectionNotes;
  }

  function renderImportPreview(parsed) {
    const ov = $("#overviewPreview");
    ov.innerHTML = "";
    Object.entries(parsed.overview).forEach(([k, values]) => {
      const kEl = document.createElement("div");
      kEl.className = "k";
      kEl.textContent = k;
      const vEl = document.createElement("div");
      vEl.className = "v";
      vEl.textContent = values.join("\n");
      ov.appendChild(kEl);
      ov.appendChild(vEl);
    });

    const sec = $("#sectionsPreview");
    sec.innerHTML = "";
    parsed.sections.forEach((s) => {
      const details = document.createElement("details");
      details.className = "section-item";
      const summary = document.createElement("summary");
      summary.textContent = s.title;
      const pre = document.createElement("pre");
      pre.textContent = s.text || "(no content)";
      details.appendChild(summary);
      details.appendChild(pre);
      sec.appendChild(details);
    });
  }

  $("#btnContinueToReview").addEventListener("click", () => goToStep(2));

  // ---------- Step 2: Review & Edit ----------
  function fillReviewForm() {
    const b = B();
    $("#f_articleNo").value = b.articleNo;
    $("#f_item").value = b.item;
    $("#f_material").value = b.material;
    $("#f_color").value = b.color;
    $("#f_description").value = b.description;
    $("#f_inclusions").value = b.inclusions;
    $("#f_materialInstructions").value = b.materialInstructions;
    $("#f_packaging").value = b.packaging;
    $("#f_knownIssues").value = b.knownIssues;
    renderTechSpecTable();
    renderQcTable();
  }

  function renderTechSpecTable() {
    const wrap = $("#techSpecTable");
    wrap.innerHTML = "";
    B().techSpecs.forEach((row, idx) => {
      wrap.appendChild(buildRow(
        [
          { key: "label", val: row.label, placeholder: "Parameter (e.g. Voltage)", cls: "col-narrow" },
          { key: "value", val: row.value, placeholder: "Value (e.g. 230V AC / 50Hz)" },
        ],
        () => { B().techSpecs.splice(idx, 1); renderTechSpecTable(); },
        (key, val) => { B().techSpecs[idx][key] = val; }
      ));
    });
  }

  $("#btnAddTechSpec").addEventListener("click", () => {
    B().techSpecs.push({ label: "", value: "" });
    renderTechSpecTable();
  });

  function renderQcTable() {
    const wrap = $("#qcTable");
    wrap.innerHTML = "";
    B().qcRows.forEach((row, idx) => {
      wrap.appendChild(buildRow(
        [
          { key: "parameter", val: row.parameter, placeholder: "Parameter", cls: "col-narrow" },
          { key: "acceptable", val: row.acceptable, placeholder: "Acceptable" },
          { key: "notAcceptable", val: row.notAcceptable, placeholder: "Not acceptable" },
        ],
        () => { B().qcRows.splice(idx, 1); renderQcTable(); },
        (key, val) => { B().qcRows[idx][key] = val; }
      ));
    });
  }

  $("#btnAddQcRow").addEventListener("click", () => {
    B().qcRows.push({ parameter: "", acceptable: "", notAcceptable: "" });
    renderQcTable();
  });

  function buildRow(fields, onRemove, onChange) {
    const row = document.createElement("div");
    row.className = "editable-row";
    fields.forEach((f) => {
      const input = document.createElement("textarea");
      input.rows = 1;
      input.className = f.cls || "";
      input.placeholder = f.placeholder || "";
      input.value = f.val || "";
      input.addEventListener("input", () => onChange(f.key, input.value));
      row.appendChild(input);
    });
    const rm = document.createElement("button");
    rm.className = "row-remove";
    rm.textContent = "✕";
    rm.addEventListener("click", onRemove);
    row.appendChild(rm);
    return row;
  }

  function syncFormToState() {
    const b = B();
    if (!$("#panel-2").classList.contains("hidden")) {
      b.articleNo = $("#f_articleNo").value;
      b.item = $("#f_item").value;
      b.material = $("#f_material").value;
      b.color = $("#f_color").value;
      b.description = $("#f_description").value;
      b.inclusions = $("#f_inclusions").value;
      b.materialInstructions = $("#f_materialInstructions").value;
      b.packaging = $("#f_packaging").value;
      b.knownIssues = $("#f_knownIssues").value;
    }
    if (!$("#panel-3").classList.contains("hidden")) {
      b.date = $("#f_date").value;
      b.poNumber = $("#f_poNumber").value;
      b.supplierName = $("#f_supplierName").value;
      b.supplierContact = $("#f_supplierContact").value;
      b.supplierPhone = $("#f_supplierPhone").value;
      b.supplierAddress = $("#f_supplierAddress").value;
      b.compliance = $("#f_compliance").value;
      b.approvalContacts = $("#f_approvalContacts").value;
      b.brandName = $("#f_brandName").value;
      b.sampleCount = $("#f_sampleCount").value;
      b.preQcDate = $("#f_preQcDate").value;
      b.inspectionDate = $("#f_inspectionDate").value;
      b.qualityInspectionNotes = $("#f_qualityInspectionNotes").value;
    }
    if (!$("#panel-4").classList.contains("hidden")) {
      b.version = $("#f_version").value || b.version;
      b.versionLabel = $("#f_versionLabel").value;
    }
  }

  // ---------- Step 3 prefill ----------
  const originalGoToStep = goToStep;
  function fillPoForm() {
    const b = B();
    $("#f_date").value = b.date || new Date().toISOString().slice(0, 10);
    $("#f_poNumber").value = b.poNumber;
    $("#f_supplierName").value = b.supplierName;
    $("#f_supplierContact").value = b.supplierContact;
    $("#f_supplierPhone").value = b.supplierPhone;
    $("#f_supplierAddress").value = b.supplierAddress;
    $("#f_compliance").value = b.compliance;
    $("#f_approvalContacts").value = b.approvalContacts;
    $("#f_brandName").value = b.brandName;
    $("#f_sampleCount").value = b.sampleCount;
    $("#f_preQcDate").value = b.preQcDate;
    $("#f_inspectionDate").value = b.inspectionDate;
    $("#f_qualityInspectionNotes").value = b.qualityInspectionNotes;
  }

  // hook into step 3 render
  const panel3Observer = () => {
    if (!$("#panel-3").classList.contains("hidden")) fillPoForm();
  };
  new MutationObserver(panel3Observer).observe($("#panel-3"), { attributes: true, attributeFilter: ["class"] });

  // ---------- Step 4: Export ----------
  $("#btnExportDocx").addEventListener("click", async () => {
    syncFormToState();
    setStatus("#exportStatus", "Generating .docx…");
    try {
      const blob = await PBDocx.generateDocx(B());
      const filename = `Product Brief_${B().articleNo || "SKU"}_v${B().version || "1.0"}.docx`;
      PBDocx.downloadBlob(blob, filename);
      setStatus("#exportStatus", `Downloaded ${filename}. Open it in Google Drive (right-click > Open with > Google Docs) to continue editing there.`, "ok");
    } catch (err) {
      console.error(err);
      setStatus("#exportStatus", "Export failed: " + err.message, "error");
    }
  });

  $("#btnExportJson").addEventListener("click", () => {
    syncFormToState();
    const blob = new Blob([JSON.stringify(B(), null, 2)], { type: "application/json" });
    PBDocx.downloadBlob(blob, `ProductBrief_${B().articleNo || "SKU"}_v${B().version || "1.0"}.json`);
  });

  $("#btnSaveVersion").addEventListener("click", () => {
    syncFormToState();
    const entry = PBStorage.save(B());
    setStatus("#exportStatus", `Saved as ${entry.sku} v${entry.version}.`, "ok");
    renderVersionList("#versionList");
  });

  function renderVersionList(sel) {
    const wrap = $(sel);
    wrap.innerHTML = "";
    const list = PBStorage.listAll();
    if (!list.length) {
      wrap.innerHTML = '<p class="hint">No saved versions yet.</p>';
      return;
    }
    list.forEach((entry) => {
      const card = document.createElement("div");
      card.className = "version-card";
      const date = new Date(entry.savedAt).toLocaleString();
      card.innerHTML = `
        <div>
          <strong>${entry.sku}</strong> — v${entry.version} ${entry.label ? "· " + entry.label : ""}
          <div class="meta">${date}</div>
        </div>
        <div class="actions">
          <button class="btn btn-small" data-load="${entry.id}">Open</button>
          <button class="btn btn-small" data-del="${entry.id}">Delete</button>
        </div>`;
      wrap.appendChild(card);
    });
    wrap.querySelectorAll("[data-load]").forEach((btn) =>
      btn.addEventListener("click", () => {
        const entry = PBStorage.load(btn.dataset.load);
        if (entry) {
          PBState.brief = entry.brief;
          $("#versionModal").classList.add("hidden");
          goToStep(2);
        }
      })
    );
    wrap.querySelectorAll("[data-del]").forEach((btn) =>
      btn.addEventListener("click", () => {
        PBStorage.remove(btn.dataset.del);
        renderVersionList(sel);
        renderVersionList("#versionModalList");
      })
    );
  }

  // ---------- Step 5: Downstream Documents ----------
  function exportDownstream(genFn, filenamePrefix) {
    syncFormToState();
    try {
      const blob = genFn(B());
      const filename = `${filenamePrefix}_${B().articleNo || "SKU"}_${B().poNumber || ""}.xlsx`.replace(/_+/g, "_");
      PBDocx.downloadBlob(blob, filename);
      setStatus("#downstreamStatus", `Downloaded ${filename}.`, "ok");
    } catch (err) {
      console.error(err);
      setStatus("#downstreamStatus", "Export failed: " + err.message, "error");
    }
  }

  $("#btnExportPreInspection").addEventListener("click", () =>
    exportDownstream(PBXlsxGen.generatePreInspectionXlsx, "Pre-inspection Briefing Form")
  );
  $("#btnExportPreQC").addEventListener("click", () =>
    exportDownstream(PBXlsxGen.generatePreQCXlsx, "Pre-QC Check")
  );
  $("#btnExportMarketing").addEventListener("click", () =>
    exportDownstream(PBXlsxGen.generateMarketingXlsx, "Marketing Guide Sheet")
  );

  // ---------- Top bar ----------
  $("#btnNew").addEventListener("click", () => {
    if (confirm("Start a new brief? Unsaved changes will be lost.")) {
      PBState.reset();
      $("#importPreview").classList.add("hidden");
      fileInput.value = "";
      setStatus("#importStatus", "");
      goToStep(1);
    }
  });

  $("#btnLoadVersion").addEventListener("click", () => {
    renderVersionList("#versionModalList");
    $("#versionModal").classList.remove("hidden");
  });
  $("#btnCloseModal").addEventListener("click", () => $("#versionModal").classList.add("hidden"));

  // init
  goToStep(1);
})();
