/* state.js - single source of truth for the current brief being edited */
(function (global) {
  function emptyBrief() {
    return {
      articleNo: "",
      item: "",
      description: "",
      material: "",
      color: "",
      inclusions: "",
      materialInstructions: "",
      techSpecs: [], // [{label, value}]
      packaging: "",
      qcRows: [], // [{parameter, acceptable, notAcceptable}]
      knownIssues: "", // known field complaints + preventive specs, from PD review analysis
      qualityInspectionNotes: "", // per-SKU QC test notes from the PD sheet
      date: "",
      poNumber: "",
      supplierName: "",
      supplierContact: "",
      supplierPhone: "",
      supplierAddress: "",
      compliance: "",
      approvalContacts: "",
      brandName: "",
      preQcDate: "",
      inspectionDate: "",
      sampleCount: "",
      version: "1.0",
      versionLabel: "",
      raw: null, // raw parser output, for reference/debug
    };
  }

  const State = {
    brief: emptyBrief(),
    reset() { this.brief = emptyBrief(); },
  };

  global.PBState = State;
})(window);
