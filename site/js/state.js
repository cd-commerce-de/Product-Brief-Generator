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
      // From the PD sheet's "Product Brief Input" tab (see brief-input.js).
      // When components is non-empty it replaces the free-text Material &
      // Workmanship / Packaging boxes: per-component numbered instructions
      // with their own pictures, status and SKU-specific detail.
      components: [], // [{name, kind:"main"|"packaging", items:[{text, skuDetail, caption, status, appliesTo, images, generated}]}]
      skuTable: [], // [{sku, color, itemCarton, itemCartonWeight, masterCarton, masterCartonWeight, pcsPerMaster}]
      category: "",
      showCaptions: true, // print each picture's caption under it (Step 2 switch)
      complianceAuto: "", // the compliance text last generated from the sheet (to tell if the person has edited it since)
      qcRows: [], // [{parameter, acceptable, notAcceptable, method?, images?}]
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
      productionStatus: "", // e.g. "70% finished, 20% packed" - inspection-day status, feeds Pre-inspection/Pre-QC
      // Downstream docs always include a fixed generic baseline (package
      // completeness, measurement, visual, packaging/printing checks).
      // These two arrays are PRODUCT-SPECIFIC tests added on top of that
      // baseline (e.g. an inflatable product needs inflation/leakage/
      // deflation tests many products never need) - empty by default,
      // fully editable in Step 3.
      extraPreInspectionTests: [], // [{test, items}]
      extraPreQcTests: [], // [{test, methods}]
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
