/* api/store.js
 * Shared product history, stored in an Upstash Redis database connected
 * through Vercel's Storage tab. Same Basic Auth password as the rest of the
 * site - nobody can read or write this without logging in.
 *
 * Talks to Redis over plain HTTPS (Upstash's REST API), so there are no npm
 * packages and the project keeps its zero-build setup.
 *
 *   GET    /api/store           -> [{ key, updatedAt }] for every product
 *   GET    /api/store?key=K     -> one full product record (with history)
 *   POST   /api/store           -> add history entries to a product
 *                                  body: { key, entries: [ {id, savedAt, version,
 *                                  label, savedBy, brief}, ... ] }
 *                                  or { key, attach: { entryId, file: {id, name,
 *                                  size, chunks} } } to attach a generated file
 *                                  (uploaded via api/files.js) to a saved version
 *   DELETE /api/store?key=K     -> delete a product for everyone (and its files)
 *
 * WHY EVERY SAVE IS ITS OWN DATABASE ENTRY: if the whole product were stored
 * as one blob, two people saving at the same moment would each read the old
 * blob and the second write would silently erase the first person's save.
 * Instead each saved version is its own field in a Redis hash (added with
 * HSETNX, which can't overwrite anything), so simultaneous saves never lose
 * each other. The product's "last changed" time lives in a sorted set updated
 * with ZADD GT ("only if newer"), which is also safe to do simultaneously.
 *
 * Redis layout:
 *   pbg:p:<key>   hash   entry id -> JSON of that saved version
 *   pbg:o:<key>   zset   entry id scored by savedAt (to find the oldest to trim)
 *   pbg:updated   zset   product key scored by when it last gained a save
 *
 * Pictures are NOT stored in the saved text of a version (they're large).
 * The generated .docx itself - pictures included - is what keeps them: it is
 * stored as chunks (api/files.js) and referenced from the version's `file`
 * field. Only the newest MAX_FILES_PER_PRODUCT files per product are kept;
 * older versions keep their text but lose the stored file.
 */
const { checkAuth, deny } = require("./_auth");
const { redisConfig, redis, send, readBody, FILE_ID_RE, MAX_FILE_BYTES, MAX_CHUNKS, mimeFor, fileKeys } = require("./_redis");

const MAX_HISTORY = 25; // per product, same as the browser-only version
const MAX_ENTRY_BYTES = 400 * 1024; // one saved version; keeps each request well under Upstash's size limit
const MAX_ENTRIES_PER_REQUEST = 30;
const hashKey = (k) => `pbg:p:${k}`;
const orderKey = (k) => `pbg:o:${k}`;
const UPDATED_KEY = "pbg:updated";

const MAX_FILES_PER_PRODUCT = 10;

function stripImages(brief) {
  const copy = JSON.parse(JSON.stringify(brief || {}));
  copy.images = { hero: [], feature: [] };
  (copy.components || []).forEach((c) => (c.items || []).forEach((it) => { it.images = []; }));
  (copy.qcRows || []).forEach((r) => { r.images = []; });
  return copy;
}

function entryId(e) {
  return e.id || `L${e.savedAt}`;
}

function newestEntry(history) {
  return history.reduce((a, b) => (b.savedAt >= a.savedAt ? b : a), history[0]);
}

// Builds the record the browser sees from the raw hash contents. The
// "current brief" is simply the newest saved version.
function buildRecord(key, flatHash, updatedAtScore) {
  const history = [];
  for (let i = 0; i < (flatHash || []).length; i += 2) {
    try { history.push(JSON.parse(flatHash[i + 1])); } catch (e) { /* skip a corrupt entry */ }
  }
  if (!history.length) return null;
  history.sort((a, b) => a.savedAt - b.savedAt);
  const newest = newestEntry(history);
  return {
    key,
    item: (newest.brief && newest.brief.item) || "",
    articleNo: (newest.brief && newest.brief.articleNo) || "",
    createdAt: history[0].savedAt,
    updatedAt: Number(updatedAtScore) || newest.savedAt,
    brief: newest.brief,
    history,
  };
}

// Links an uploaded file to a saved version, makes its chunks permanent, and
// keeps only the newest MAX_FILES_PER_PRODUCT files for the product.
async function attachFile(cfg, res, productKey, attach) {
  const entryId = String(attach.entryId || "");
  const f = attach.file || {};
  const file = {
    id: String(f.id || ""), name: String(f.name || "").slice(0, 200),
    size: Number(f.size), chunks: Number(f.chunks),
  };
  if (!productKey || !entryId || !FILE_ID_RE.test(file.id) || !mimeFor(file.name) ||
      !Number.isInteger(file.chunks) || file.chunks < 1 || file.chunks > MAX_CHUNKS || !(file.size > 0 && file.size <= MAX_FILE_BYTES)) {
    return send(res, 400, { configured: true, error: "Bad file details." });
  }
  const [raw, metaCheck] = await redis(cfg, [["HGET", hashKey(productKey), entryId], ["GET", `pbg:f:${file.id}:m`]]);
  if (!raw) return send(res, 404, { configured: true, error: "That saved version no longer exists." });
  if (!metaCheck) return send(res, 400, { configured: true, error: "The file upload is missing or has expired." });

  const entry = JSON.parse(raw);
  const cmds = [];
  if (entry.file && entry.file.id !== file.id) fileKeys(entry.file).forEach((k) => cmds.push(["DEL", k]));
  entry.file = { ...file, mime: mimeFor(file.name), generatedAt: Date.now() };
  cmds.push(["HSET", hashKey(productKey), entryId, JSON.stringify(entry)]);
  fileKeys(file).forEach((k) => cmds.push(["PERSIST", k]));
  cmds.push(["ZADD", UPDATED_KEY, "GT", Date.now(), productKey]);
  await redis(cfg, cmds);

  // Retention: keep the newest few files; older versions keep their text only.
  const [flat] = await redis(cfg, [["HGETALL", hashKey(productKey)]]);
  const entries = [];
  for (let i = 0; i < (flat || []).length; i += 2) {
    try { entries.push(JSON.parse(flat[i + 1])); } catch (e) { /* skip */ }
  }
  const withFile = entries.filter((e) => e.file).sort((a, b) => a.savedAt - b.savedAt);
  const drops = [];
  while (withFile.length > MAX_FILES_PER_PRODUCT) {
    const old = withFile.shift();
    fileKeys(old.file).forEach((k) => drops.push(["DEL", k]));
    const stripped = { ...old }; delete stripped.file;
    drops.push(["HSET", hashKey(productKey), old.id, JSON.stringify(stripped)]);
  }
  if (drops.length) await redis(cfg, drops);

  const [flat2, score] = await redis(cfg, [["HGETALL", hashKey(productKey)], ["ZSCORE", UPDATED_KEY, productKey]]);
  return send(res, 200, { configured: true, product: buildRecord(productKey, flat2, score) });
}

module.exports = async (req, res) => {
  if (!checkAuth(req)) return deny(res);

  const cfg = redisConfig();
  if (!cfg) {
    return send(res, 503, { configured: false, error: "Shared storage isn't set up yet." });
  }

  try {
    const url = new URL(req.url, "http://localhost");
    const key = url.searchParams.get("key");

    if (req.method === "GET" && !key) {
      const [flat] = await redis(cfg, [["ZRANGE", UPDATED_KEY, 0, -1, "WITHSCORES"]]);
      const products = [];
      for (let i = 0; i < (flat || []).length; i += 2) {
        products.push({ key: flat[i], updatedAt: Number(flat[i + 1]) });
      }
      products.sort((a, b) => b.updatedAt - a.updatedAt);
      return send(res, 200, { configured: true, products });
    }

    if (req.method === "GET") {
      const [flat, score] = await redis(cfg, [["HGETALL", hashKey(key)], ["ZSCORE", UPDATED_KEY, key]]);
      const record = buildRecord(key, flat, score);
      if (!record) return send(res, 404, { configured: true, error: "Not found" });
      return send(res, 200, { configured: true, product: record });
    }

    if (req.method === "DELETE") {
      if (!key) return send(res, 400, { configured: true, error: "Missing key" });
      const [flat] = await redis(cfg, [["HGETALL", hashKey(key)]]);
      const fileDels = [];
      for (let i = 1; i < (flat || []).length; i += 2) {
        try { const e = JSON.parse(flat[i]); if (e.file) fileKeys(e.file).forEach((k) => fileDels.push(["DEL", k])); } catch (e) { /* skip */ }
      }
      await redis(cfg, [...fileDels, ["DEL", hashKey(key)], ["DEL", orderKey(key)], ["ZREM", UPDATED_KEY, key]]);
      return send(res, 200, { configured: true, deleted: key });
    }

    if (req.method === "POST") {
      const body = await readBody(req);
      const productKey = typeof body.key === "string" ? body.key.trim() : "";
      if (body.attach) return attachFile(cfg, res, productKey, body.attach);
      const incoming = Array.isArray(body.entries) ? body.entries : [];
      if (!productKey || productKey.length > 200) {
        return send(res, 400, { configured: true, error: "A product key (max 200 characters) is required." });
      }
      if (!incoming.length || incoming.length > MAX_ENTRIES_PER_REQUEST) {
        return send(res, 400, { configured: true, error: `Send between 1 and ${MAX_ENTRIES_PER_REQUEST} entries.` });
      }

      const entries = [];
      let tooBig = 0;
      for (const e of incoming) {
        if (!e || typeof e.savedAt !== "number" || !e.brief || typeof e.brief !== "object") continue;
        const entry = {
          id: entryId(e),
          savedAt: e.savedAt,
          version: String(e.version || "1.0").slice(0, 40),
          label: String(e.label || "").slice(0, 200),
          savedBy: String(e.savedBy || "").slice(0, 80),
          brief: stripImages(e.brief),
          // `file` is only ever set through the attach step above
        };
        const json = JSON.stringify(entry);
        if (json.length > MAX_ENTRY_BYTES) { tooBig++; continue; }
        entries.push({ entry, json });
      }
      if (!entries.length) {
        return send(res, 400, { configured: true, error: tooBig ? "That save is too large to share." : "No valid entries." });
      }

      // 1) Add each save as its own entry. HSETNX never overwrites, so a
      //    retry (or two people at once) can't clobber anything.
      const addCmds = [];
      entries.forEach(({ entry, json }) => {
        addCmds.push(["HSETNX", hashKey(productKey), entry.id, json]);
        addCmds.push(["ZADD", orderKey(productKey), entry.savedAt, entry.id]);
      });
      const addResults = await redis(cfg, addCmds);
      let added = 0;
      entries.forEach((_, i) => { added += Number(addResults[i * 2]) || 0; });

      // 2) Mark the product as changed (only moves forward), and find any
      //    saves beyond the history limit.
      const step2 = [];
      if (added) step2.push(["ZADD", UPDATED_KEY, "GT", Date.now(), productKey]);
      step2.push(["ZRANGE", orderKey(productKey), 0, -(MAX_HISTORY + 1)]);
      const step2Results = await redis(cfg, step2);
      const trimIds = step2Results[step2Results.length - 1] || [];

      // 3) Drop the oldest beyond the limit, then read the result back.
      const step3 = [];
      if (trimIds.length) {
        const [trimmed] = await redis(cfg, [["HMGET", hashKey(productKey), ...trimIds]]);
        (trimmed || []).forEach((raw) => {
          try { const e = JSON.parse(raw); if (e && e.file) fileKeys(e.file).forEach((k) => step3.push(["DEL", k])); } catch (e) { /* skip */ }
        });
        step3.push(["ZREM", orderKey(productKey), ...trimIds]);
        step3.push(["HDEL", hashKey(productKey), ...trimIds]);
      }
      step3.push(["HGETALL", hashKey(productKey)], ["ZSCORE", UPDATED_KEY, productKey]);
      const step3Results = await redis(cfg, step3);
      const flat = step3Results[step3Results.length - 2];
      const score = step3Results[step3Results.length - 1];
      const record = buildRecord(productKey, flat, score);
      return send(res, 200, { configured: true, added, rejected: tooBig, product: record });
    }

    res.setHeader("Allow", "GET, POST, DELETE");
    return send(res, 405, { configured: true, error: "Method not allowed" });
  } catch (err) {
    return send(res, 500, { configured: true, error: err.message || "Server error" });
  }
};
