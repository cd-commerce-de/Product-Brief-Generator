/* storage.js - localStorage-based, ONE RECORD PER PRODUCT.
 *
 * Earlier versions of this file kept a flat list where every "Save version"
 * click created a brand-new top-level entry - so a product revised five
 * times showed up as five separate cards. This version keys everything by
 * a "product key" (see defaultProductKey) so re-saving the same product
 * always updates its ONE record; each save just appends a lightweight
 * snapshot to that record's own `history` array.
 *
 * This is still a client-only MVP store (browser localStorage), not a
 * shared database - see README "Phase 3" for what a real shared version
 * history would need.
 */
(function (global) {
  const KEY = "pbg_products_v2";
  const MAX_HISTORY_PER_PRODUCT = 25;

  function readAll() {
    try {
      return JSON.parse(localStorage.getItem(KEY) || "{}");
    } catch (e) {
      return {};
    }
  }

  function writeAll(map) {
    localStorage.setItem(KEY, JSON.stringify(map));
  }

  // Best-effort default grouping key for "this product": the first SKU in
  // Article No. (which already aggregates every color/size variant of one
  // product into a single multi-line value upstream in the parser), falling
  // back to the Item name. This is only ever a *suggestion* - the person can
  // always override it in Step 4, since no automatic heuristic is fully
  // reliable across every PD sheet naming convention.
  function defaultProductKey(brief) {
    const firstSku = (brief.articleNo || "").split("\n")[0].trim();
    const key = firstSku || (brief.item || "").trim() || "Untitled product";
    return key;
  }

  function snapshot(brief) {
    // Deep-clone via JSON round-trip so later edits to the live brief object
    // can never mutate a saved history entry.
    return JSON.parse(JSON.stringify(brief));
  }

  const Storage = {
    defaultProductKey,

    listProducts() {
      return Object.values(readAll()).sort((a, b) => b.updatedAt - a.updatedAt);
    },

    // Saves/updates the ONE record for this product. Always the same record
    // for the same key - re-saving never creates a second card, it just
    // updates `brief` (the current/latest state) and appends to `history`.
    save(brief, productKeyOverride) {
      const map = readAll();
      const key = (productKeyOverride || brief.productKey || defaultProductKey(brief)).trim() || "Untitled product";
      const now = Date.now();
      const historyEntry = {
        savedAt: now,
        version: brief.version || "1.0",
        label: brief.versionLabel || "",
        brief: snapshot(brief),
      };

      const existing = map[key];
      if (existing) {
        existing.brief = snapshot(brief);
        existing.item = brief.item || existing.item;
        existing.articleNo = brief.articleNo || existing.articleNo;
        existing.updatedAt = now;
        existing.history.push(historyEntry);
        if (existing.history.length > MAX_HISTORY_PER_PRODUCT) {
          existing.history = existing.history.slice(-MAX_HISTORY_PER_PRODUCT);
        }
      } else {
        map[key] = {
          key,
          item: brief.item || "",
          articleNo: brief.articleNo || "",
          createdAt: now,
          updatedAt: now,
          brief: snapshot(brief),
          history: [historyEntry],
        };
      }

      writeAll(map);
      return map[key];
    },

    load(key) {
      return readAll()[key] || null;
    },

    remove(key) {
      const map = readAll();
      delete map[key];
      writeAll(map);
    },

    // Rolls a product's *current* brief back to one specific past save,
    // identified by that history entry's timestamp. The history itself is
    // left untouched (restoring is its own new fact, not an erasure).
    restoreHistoryEntry(key, savedAt) {
      const map = readAll();
      const product = map[key];
      if (!product) return null;
      const entry = product.history.find((h) => h.savedAt === savedAt);
      if (!entry) return null;
      product.brief = snapshot(entry.brief);
      product.updatedAt = Date.now();
      writeAll(map);
      return product;
    },

    // "1.0" -> "1.1". Falls back to "1.0" for anything that doesn't parse.
    suggestNextVersion(currentVersion) {
      const m = /^(\d+)\.(\d+)$/.exec((currentVersion || "").trim());
      if (!m) return "1.0";
      return `${m[1]}.${Number(m[2]) + 1}`;
    },
  };

  global.PBStorage = Storage;
})(window);
