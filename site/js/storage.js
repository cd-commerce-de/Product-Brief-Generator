/* storage.js - simple localStorage version tracking.
 * This is a client-only MVP store. For real multi-user version control,
 * swap listAll/save/load for calls to Google Drive/Sheets API or a small
 * Vercel serverless function backed by a database — the call sites below
 * are the only places that would need to change. See README "Phase 2".
 */
(function (global) {
  const KEY = "pbg_versions_v1";

  function readAll() {
    try {
      return JSON.parse(localStorage.getItem(KEY) || "[]");
    } catch (e) {
      return [];
    }
  }

  function writeAll(list) {
    localStorage.setItem(KEY, JSON.stringify(list));
  }

  function makeId(brief) {
    const sku = (brief.articleNo || "SKU").replace(/\s+/g, "");
    return `${sku}_v${brief.version || "1.0"}_${Date.now()}`;
  }

  const Storage = {
    listAll() {
      return readAll().sort((a, b) => b.savedAt - a.savedAt);
    },
    save(brief) {
      const list = readAll();
      const entry = {
        id: makeId(brief),
        sku: brief.articleNo || "(no SKU)",
        item: brief.item || "",
        version: brief.version || "1.0",
        label: brief.versionLabel || "",
        savedAt: Date.now(),
        brief,
      };
      list.push(entry);
      writeAll(list);
      return entry;
    },
    load(id) {
      const list = readAll();
      return list.find((e) => e.id === id);
    },
    remove(id) {
      const list = readAll().filter((e) => e.id !== id);
      writeAll(list);
    },
  };

  global.PBStorage = Storage;
})(window);
