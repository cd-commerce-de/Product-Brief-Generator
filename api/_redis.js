/* api/_redis.js
 * Shared helpers for the functions that talk to the team database (Upstash
 * Redis, connected through Vercel's Storage tab) over plain HTTPS - no npm
 * packages, so the project keeps its zero-build setup. The leading underscore
 * keeps Vercel from treating this file as a function of its own.
 */
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

// ---- generated files ----
// A generated brief is stored as the actual .docx (so its pictures are kept),
// split into base64 chunks that each fit under the database's per-request
// size limit. Keys: pbg:f:<id>:m (small metadata) and pbg:f:<id>:0..n-1.
const FILE_ID_RE = /^[a-z0-9-]{6,48}$/;
const MAX_FILE_BYTES = 3 * 1024 * 1024;
const MAX_CHUNKS = 8;
const MAX_CHUNK_CHARS = 900 * 1024;
const MIME_BY_EXT = {
  docx: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  zip: "application/zip",
};

function mimeFor(name) {
  const m = /\.([a-z0-9]+)$/i.exec(String(name || ""));
  return m ? MIME_BY_EXT[m[1].toLowerCase()] || null : null;
}

function fileKeys(file) {
  const keys = [`pbg:f:${file.id}:m`];
  for (let i = 0; i < (file.chunks || 0); i++) keys.push(`pbg:f:${file.id}:${i}`);
  return keys;
}

module.exports = {
  redisConfig, redis, send, readBody,
  FILE_ID_RE, MAX_FILE_BYTES, MAX_CHUNKS, MAX_CHUNK_CHARS, mimeFor, fileKeys,
};
