/* api/gate.js
 * A minimal password gate in front of the static app in /site.
 *
 * Every request is rewritten to this function (see vercel.json). It checks
 * HTTP Basic Auth against SITE_USER / SITE_PASSWORD (set as Vercel
 * Environment Variables — never hard-code them here). If the credentials
 * are missing or wrong, the browser gets a 401 with WWW-Authenticate, which
 * makes it pop up its native username/password dialog. If they're correct,
 * this reads and serves the requested file straight out of /site.
 *
 * This is real, server-side protection (unlike a client-side JS password
 * prompt, which can be bypassed by reading the page source) and it works on
 * Vercel's free Hobby plan — no Pro-plan "Password Protection" add-on
 * needed. Trade-off vs. that add-on: no "log out" button (closing the
 * browser / using a private window is the equivalent), and changing the
 * password requires updating the env vars + redeploying.
 */
const fs = require("fs");
const path = require("path");

const SITE_ROOT = path.join(process.cwd(), "site");

const MIME = {
  ".html": "text/html; charset=utf-8",
  ".js": "application/javascript; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".svg": "image/svg+xml",
  ".png": "image/png",
  ".jpg": "image/jpeg",
  ".ico": "image/x-icon",
  ".map": "application/json; charset=utf-8",
};

function checkAuth(req) {
  const header = req.headers.authorization || "";
  const [scheme, encoded] = header.split(" ");
  if (scheme !== "Basic" || !encoded) return false;
  const decoded = Buffer.from(encoded, "base64").toString("utf8");
  const sep = decoded.indexOf(":");
  if (sep === -1) return false;
  const user = decoded.slice(0, sep);
  const pass = decoded.slice(sep + 1);
  const expectedUser = process.env.SITE_USER || "";
  const expectedPass = process.env.SITE_PASSWORD || "";
  return user === expectedUser && pass === expectedPass && expectedPass.length > 0;
}

module.exports = (req, res) => {
  if (!checkAuth(req)) {
    res.setHeader("WWW-Authenticate", 'Basic realm="Product Brief Generator"');
    res.statusCode = 401;
    res.end("Authentication required.");
    return;
  }

  let urlPath = (req.url || "/").split("?")[0];
  if (urlPath === "/" || urlPath === "") urlPath = "/index.html";

  const filePath = path.normalize(path.join(SITE_ROOT, urlPath));
  // Guard against path traversal escaping /site
  if (!filePath.startsWith(SITE_ROOT)) {
    res.statusCode = 400;
    res.end("Bad request");
    return;
  }

  fs.readFile(filePath, (err, data) => {
    if (err) {
      res.statusCode = 404;
      res.end("Not found");
      return;
    }
    const ext = path.extname(filePath);
    res.setHeader("Content-Type", MIME[ext] || "application/octet-stream");
    res.statusCode = 200;
    res.end(data);
  });
};
