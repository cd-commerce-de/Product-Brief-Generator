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
      knownIssues: "", // known field complaints + preventive specs, from PD review analysis (reference only - used to seed qcRows, not exported as its own section)
      qualityInspectionNotes: "", // per-SKU QC test notes from the PD sheet (feeds the Pre-QC Check downstream doc)
      images: { hero: [], feature: [] }, // extracted from the PD sheet, see xlsx-images.js
      includeImages: true,
      date: "",
      poNumber: "",
      supplierName: "",
      supplierContact: "",
      supplierPhone: "",
      supplierEmail: "",
      supplierAddress: "",
      compliance: "",
      approvalContacts: "",
      brandName: "",
      preQcDate: "",
      inspectionDate: "",
      sampleCount: "",
      version: "1.0",
      versionLabel: "",
      productKey: "", // identifies "this product" for the one-per-product history (see storage.js)
      raw: null, // raw parser output, for reference/debug
    };
  }

  const State = {
    brief: emptyBrief(),
    reset() { this.brief = emptyBrief(); },
  };

  global.PBState = State;
})(window);
