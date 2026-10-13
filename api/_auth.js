/* api/_auth.js
 * Shared HTTP Basic Auth check used by every function in /api (the leading
 * underscore tells Vercel this file is a helper, not a function of its own).
 * Credentials come from the SITE_USER / SITE_PASSWORD environment variables.
 */
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

function deny(res) {
  res.setHeader("WWW-Authenticate", 'Basic realm="Product Brief Generator"');
  res.statusCode = 401;
  res.end("Authentication required.");
}

module.exports = { checkAuth, deny };
