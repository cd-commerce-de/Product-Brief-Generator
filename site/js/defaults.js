/* defaults.js - speeds up repetitive data entry across briefs:
 *
 * 1. A supplier directory, built automatically every time a brief is saved
 *    with a non-empty supplier name - no separate "add a supplier" step.
 *    Step 3 offers a dropdown to instantly re-fill name/address/phone/
 *    email/contact for a supplier you've used before, instead of retyping
 *    it for every new product.
 * 2. Sticky defaults for fields that are nearly identical across every
 *    brief (Approval Contacts has been the exact same 2-3 names in every
 *    PD sheet seen so far) - the last value used pre-fills the field on the
 *    next brief, without forcing it (still fully editable, never silently
 *    overwrites something already typed).
 *
 * Both are plain localStorage, same as storage.js - per-browser, not a
 * shared team resource (see README "Known limitations").
 */
(function (global) {
  const SUPPLIERS_KEY = "pbg_suppliers_v1";
  const DEFAULTS_KEY = "pbg_defaults_v1";

  function readJson(key, fallback) {
    try {
      return JSON.parse(localStorage.getItem(key) || JSON.stringify(fallback));
    } catch (e) {
      return fallback;
    }
  }
  function writeJson(key, value) {
    localStorage.setItem(key, JSON.stringify(value));
  }

  const Defaults = {
    // ---- Supplier directory ----
    listSuppliers() {
      const map = readJson(SUPPLIERS_KEY, {});
      return Object.values(map).sort((a, b) => a.name.localeCompare(b.name));
    },

    // Called automatically whenever a brief with a supplier name is saved.
    // Last-used values win for a given supplier name, since contact details
    // do sometimes change (new phone, new contact person).
    rememberSupplier(brief) {
      const name = (brief.supplierName || "").trim();
      if (!name) return;
      const map = readJson(SUPPLIERS_KEY, {});
      map[name] = {
        name,
        address: brief.supplierAddress || "",
        phone: brief.supplierPhone || "",
        email: brief.supplierEmail || "",
        contact: brief.supplierContact || "",
        updatedAt: Date.now(),
      };
      writeJson(SUPPLIERS_KEY, map);
    },

    getSupplier(name) {
      const map = readJson(SUPPLIERS_KEY, {});
      return map[name] || null;
    },

    // ---- Sticky field defaults ----
    getFieldDefaults() {
      return readJson(DEFAULTS_KEY, { approvalContacts: "" });
    },

    // Called automatically whenever Step 3 is saved - remembers whichever
    // fields are non-empty so the *next* brief starts pre-filled.
    rememberFieldDefaults(brief) {
      const current = readJson(DEFAULTS_KEY, {});
      if (brief.approvalContacts) current.approvalContacts = brief.approvalContacts;
      writeJson(DEFAULTS_KEY, current);
    },
  };

  global.PBDefaults = Defaults;
})(window);
