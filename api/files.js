/* api/files.js
 * Generated documents (the actual .docx) kept in the shared database, so a
 * previous brief can be downloaded exactly as it was generated - pictures
 * included. Same Basic Auth login as the rest of the site.
 *
 *   POST /api/files   one chunk of a file:
 *                     { id, index, total, name, size, data(base64) }
 *                     Chunks expire after a day unless the upload is
 *                     "attached" to a saved version (see api/store.js), so an
 *                     abandoned upload cleans itself up.
 *   GET  /api/files?id=ID   the file, as a download
 */
const { checkAuth, deny } = require("./_auth");
const {
  redisConfig, redis, send, readBody,
  FILE_ID_RE, MAX_FILE_BYTES, MAX_CHUNKS, MAX_CHUNK_CHARS, mimeFor,
} = require("./_redis");

const ONE_DAY = 24 * 60 * 60;

module.exports = async (req, res) => {
  if (!checkAuth(req)) return deny(res);
  const cfg = redisConfig();
  if (!cfg) return send(res, 503, { configured: false, error: "Shared storage isn't set up yet." });

  try {
    const url = new URL(req.url, "http://localhost");

    if (req.method === "POST") {
      const b = await readBody(req);
      const id = String(b.id || "");
      const index = Number(b.index), total = Number(b.total), size = Number(b.size);
      const name = String(b.name || "").slice(0, 200);
      const data = typeof b.data === "string" ? b.data : "";
      if (!FILE_ID_RE.test(id)) return send(res, 400, { configured: true, error: "Bad file id." });
      if (!Number.isInteger(index) || !Number.isInteger(total) || total < 1 || total > MAX_CHUNKS || index < 0 || index >= total) {
        return send(res, 400, { configured: true, error: "Bad chunk numbers." });
      }
      if (!mimeFor(name)) return send(res, 400, { configured: true, error: "Only .docx, .xlsx and .zip files can be stored." });
      if (!data || data.length > MAX_CHUNK_CHARS || !/^[A-Za-z0-9+/=]+$/.test(data)) {
        return send(res, 400, { configured: true, error: "Bad chunk data." });
      }
      if (!(size > 0 && size <= MAX_FILE_BYTES)) {
        return send(res, 413, { configured: true, error: `That file is too large to store (limit ${Math.round(MAX_FILE_BYTES / 1024 / 1024)} MB).` });
      }
      await redis(cfg, [
        ["SET", `pbg:f:${id}:${index}`, data, "EX", ONE_DAY],
        ["SET", `pbg:f:${id}:m`, JSON.stringify({ total, name, size, mime: mimeFor(name) }), "EX", ONE_DAY],
      ]);
      return send(res, 200, { configured: true, stored: index });
    }

    if (req.method === "GET") {
      const id = url.searchParams.get("id") || "";
      if (!FILE_ID_RE.test(id)) return send(res, 400, { configured: true, error: "Bad file id." });
      const [metaRaw] = await redis(cfg, [["GET", `pbg:f:${id}:m`]]);
      if (!metaRaw) return send(res, 404, { configured: true, error: "That file is no longer stored." });
      const meta = JSON.parse(metaRaw);
      const parts = await redis(cfg, Array.from({ length: meta.total }, (_, i) => ["GET", `pbg:f:${id}:${i}`]));
      if (parts.some((p) => !p)) return send(res, 404, { configured: true, error: "That file is incomplete or has expired." });
      const buf = Buffer.concat(parts.map((p) => Buffer.from(p, "base64")));
      res.statusCode = 200;
      res.setHeader("Content-Type", meta.mime || "application/octet-stream");
      res.setHeader("Content-Length", String(buf.length));
      res.setHeader("Cache-Control", "no-store");
      res.setHeader("Content-Disposition", `attachment; filename="${meta.name.replace(/[^\x20-\x7e]/g, "_").replace(/"/g, "")}"; filename*=UTF-8''${encodeURIComponent(meta.name)}`);
      res.end(buf);
      return;
    }

    res.setHeader("Allow", "GET, POST");
    return send(res, 405, { configured: true, error: "Method not allowed" });
  } catch (err) {
    return send(res, 500, { configured: true, error: err.message || "Server error" });
  }
};
