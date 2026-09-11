// Thin Microsoft Graph client: app-only (client-credentials) auth + a small
// fetch wrapper. This is the ONLY module in the whole backend that talks to
// Microsoft. Everything else goes through sheet.js, which calls this.

var GRAPH_ROOT = "https://graph.microsoft.com/v1.0";

var tokenCache = { value: null, expiresAt: 0 };

function requireEnv(name) {
  var v = process.env[name];
  if (!v) throw new Error("Missing required env var: " + name);
  return v;
}

// Client-credentials flow: an app-only token, no end user involved.
// https://learn.microsoft.com/graph/auth-v2-service
async function getAppToken() {
  var now = Date.now();
  if (tokenCache.value && now < tokenCache.expiresAt - 60000) return tokenCache.value;

  var tenantId = requireEnv("GRAPH_TENANT_ID");
  var clientId = requireEnv("GRAPH_CLIENT_ID");
  var clientSecret = requireEnv("GRAPH_CLIENT_SECRET");

  var body = new URLSearchParams({
    grant_type: "client_credentials",
    client_id: clientId,
    client_secret: clientSecret,
    scope: "https://graph.microsoft.com/.default"
  });

  var res = await fetch("https://login.microsoftonline.com/" + tenantId + "/oauth2/v2.0/token", {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: body.toString()
  });
  var json = await res.json();
  if (!res.ok) {
    throw new Error("Could not get a Graph token: " + (json.error_description || json.error || res.status));
  }
  tokenCache.value = json.access_token;
  tokenCache.expiresAt = now + (json.expires_in || 3600) * 1000;
  return tokenCache.value;
}

// Encode a SharePoint sharing URL the way Graph's /shares endpoint expects.
// https://learn.microsoft.com/graph/api/shares-get
function encodeShareUrl(url) {
  var base64 = Buffer.from(url, "utf8").toString("base64");
  var base64url = base64.replace(/=+$/, "").replace(/\//g, "_").replace(/\+/g, "-");
  return "u!" + base64url;
}

async function graphFetch(path, options) {
  options = options || {};
  var token = await getAppToken();
  var url = /^https?:\/\//.test(path) ? path : GRAPH_ROOT + path;
  var headers = Object.assign({ Authorization: "Bearer " + token }, options.headers || {});
  if (options.body != null && !headers["Content-Type"]) headers["Content-Type"] = "application/json";
  var res = await fetch(url, {
    method: options.method || "GET",
    headers: headers,
    body: options.body != null ? (typeof options.body === "string" ? options.body : JSON.stringify(options.body)) : undefined
  });
  var text = await res.text();
  var json = null;
  try { json = text ? JSON.parse(text) : null; } catch (e) { /* non-JSON response, e.g. 204 */ }
  if (!res.ok) {
    var msg = (json && json.error && json.error.message) || res.statusText || ("HTTP " + res.status);
    var err = new Error("Graph " + (options.method || "GET") + " " + path + " failed: " + msg);
    err.status = res.status;
    err.body = json;
    throw err;
  }
  return json;
}

module.exports = { graphFetch: graphFetch, encodeShareUrl: encodeShareUrl, getAppToken: getAppToken };
