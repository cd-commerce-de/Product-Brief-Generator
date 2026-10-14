/* storage.js - localStorage-based, ONE RECORD PER PRODUCT.
 *
 * Earlier versions of this file kept a flat list where every "Save version"
 * click created a brand-new top-level entry - so a product revised five
 * times showed up as five separate cards. This version keys everything by
 * a "product key" (see defaultProductKey) so re-saving the same product
 * always updates its ONE record; each save just appends a lightweight
 * snapshot to that record's own `history` array.
 *
 * This file is the LOCAL copy: it keeps screens instant and works offline.
 * Sharing across people is layered on top by shared.js, which syncs this
 * copy with the team database (api/store.js). Each history entry carries:
 *   id       - unique, so the same save is never duplicated when synced
 *   savedBy  - the name of whoever saved it
 *   synced   - true once it has reached the shared database
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

  // Pictures can sit in several places in a brief: the PD sheet gallery,
  // every instruction in a component, and the QC rows' defect photos.
  function stripImages(brief) {
    const c = JSON.parse(JSON.stringify(brief || {}));
    c.images = { hero: [], feature: [] };
    (c.components || []).forEach((co) => (co.items || []).forEach((it) => { it.images = []; }));
    (c.qcRows || []).forEach((r) => { r.images = []; });
    return c;
  }

  // Copies pictures from a previous version of the same brief onto this one
  // wherever the instruction / QC row it belongs to is unchanged.
  function carryImages(brief, prev) {
    if (!prev) return;
    const has = (im) => im && ((im.hero || []).length || (im.feature || []).length);
    if (!has(brief.images) && has(prev.images)) brief.images = prev.images;
    (brief.components || []).forEach((co) => {
      const pc = (prev.components || []).find((x) => x.name === co.name);
      if (!pc) return;
      (co.items || []).forEach((it) => {
        if ((it.images || []).length) return;
        const pi = pc.items.find((x) => x.text === it.text && (x.images || []).length);
        if (pi) it.images = pi.images;
      });
    });
    (brief.qcRows || []).forEach((r) => {
      if ((r.images || []).length) return;
      const pr = (prev.qcRows || []).find((x) => x.parameter === r.parameter && (x.images || []).length);
      if (pr) r.images = pr.images;
    });
  }

  // A browser only gives a page about 5 MB of localStorage and pictures are
  // the bulk of a brief. If it's full, drop the pictures from every saved
  // snapshot (the current brief of each product keeps its own) and retry;
  // if that still doesn't fit, say so plainly rather than failing silently.
  function writeAll(map) {
    try {
      localStorage.setItem(KEY, JSON.stringify(map));
    } catch (e) {
      Object.values(map).forEach((p) => p.history.forEach((h) => { h.brief = stripImages(h.brief); }));
      try {
        localStorage.setItem(KEY, JSON.stringify(map));
      } catch (e2) {
        throw new Error("This browser's storage is full. Delete products you no longer need in Product History (share them first if the team needs them), then save again.");
      }
    }
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
    stripImages,
    carryImages,

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
        id: `${now}-${Math.random().toString(36).slice(2, 8)}`,
        savedAt: now,
        version: brief.version || "1.0",
        label: brief.versionLabel || "",
        savedBy: brief.savedBy || "",
        synced: false,
        brief: stripImages(brief),
      };

      const existing = map[key];
      if (existing) {
        existing.history.forEach((h) => { h.brief = stripImages(h.brief); });
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

    // Used when a document is generated: if the brief is unchanged since the
    // newest saved version, that version is reused (so downloading the same
    // brief twice doesn't fill the history with duplicates); otherwise a new
    // version is saved. Returns { product, entry, created }.
    saveIfChanged(brief, productKeyOverride) {
      const key = (productKeyOverride || brief.productKey || defaultProductKey(brief)).trim() || "Untitled product";
      const existing = readAll()[key];
      const latest = existing && existing.history[existing.history.length - 1];
      if (latest && JSON.stringify(stripImages(brief)) === JSON.stringify(stripImages(latest.brief))) {
        const map = readAll();
        map[key].brief = snapshot(brief); // keep this computer's pictures on the current brief
        writeAll(map);
        return { product: map[key], entry: map[key].history[map[key].history.length - 1], created: false };
      }
      const product = this.save(brief, key);
      return { product, entry: product.history[product.history.length - 1], created: true };
    },

    load(key) {
      return readAll()[key] || null;
    },

    // Low-level access for shared.js (the sync layer) - not for UI code.
    getAll() {
      return readAll();
    },

    putRecord(record) {
      const map = readAll();
      map[record.key] = record;
      writeAll(map);
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
      const restored = snapshot(entry.brief);
      carryImages(restored, product.brief); // snapshots hold no pictures; reuse the current ones where unchanged
      product.brief = restored;
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
