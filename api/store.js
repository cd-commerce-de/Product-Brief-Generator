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
 *   DELETE /api/store?key=K     -> delete a product for everyone
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
 * Images are deliberately NOT stored: they're large, and they come straight
 * from the PD sheet, so re-importing the sheet brings them back.
 */
const { checkAuth, deny } = require("./_auth");

const MAX_HISTORY = 25; // per product, same as the browser-only version
const MAX_ENTRY_BYTES = 400 * 1024; // one saved version; keeps each request well under Upstash's size limit
const MAX_ENTRIES_PER_REQUEST = 30;
const hashKey = (k) => `pbg:p:${k}`;
const orderKey = (k) => `pbg:o:${k}`;
const UPDATED_KEY = "pbg:updated";

function redisConfig() {
  // Vercel's Upstash integration has used both naming schemes.
  const url = process.env.KV_REST_API_URL || process.env.UPSTASH_REDIS_REST_URL || "";
  const token = process.env.KV_REST_API_TOKEN || process.env.UPSTASH_REDIS_REST_TOKEN || "";
  return url && token ? { url: url.replace(/\/+$/, ""), token } : null;
}

async function redis(cfg, commands) {
  const resp = await fetch(cfg.url + "/pipeline", {
    method: "POST",
    headers: { Authorization: `Bearer ${cfg.token}`, "Content-Type": "application/json" },
    body: JSON.stringify(commands),
  });
  if (!resp.ok) throw new Error(`Database request failed (HTTP ${resp.status})`);
  const out = await resp.json();
  return out.map((o) => {
    if (o && o.error) throw new Error(o.error);
    return o ? o.result : null;
  });
}

function send(res, status, body) {
  res.statusCode = status;
  res.setHeader("Content-Type", "application/json; charset=utf-8");
  res.setHeader("Cache-Control", "no-store");
  res.end(JSON.stringify(body));
}

async function readBody(req) {
  if (req.body && typeof req.body === "object") return req.body;
  if (typeof req.body === "string") return JSON.parse(req.body || "{}");
  const chunks = [];
  for await (const c of req) chunks.push(c);
  const text = Buffer.concat(chunks).toString("utf8");
  return text ? JSON.parse(text) : {};
}

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
      await redis(cfg, [["DEL", hashKey(key)], ["DEL", orderKey(key)], ["ZREM", UPDATED_KEY, key]]);
      return send(res, 200, { configured: true, deleted: key });
    }

    if (req.method === "POST") {
      const body = await readBody(req);
      const productKey = typeof body.key === "string" ? body.key.trim() : "";
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
