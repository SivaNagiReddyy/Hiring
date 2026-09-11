// Thin Google Drive client. A service account holds this app's only
// credential, and only reaches the one file you share with it. Everything
// else in the backend goes through sheet.js, which calls this.
//
// No Google Workspace admin is needed: sharing a file with a service
// account's email is exactly the same action as sharing it with a
// colleague, and any file owner/editor can do it themselves.

var stream = require("stream");
var google = require("googleapis").google;

var XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

function requireEnv(name) {
  var v = process.env[name];
  if (!v) throw new Error("Missing required env var: " + name);
  return v;
}

// Accepts a plain file id, or any normal Google share URL, e.g.
//   https://drive.google.com/file/d/FILE_ID/view?usp=sharing
//   https://docs.google.com/spreadsheets/d/FILE_ID/edit
function extractFileId(urlOrId) {
  if (!urlOrId) return null;
  var m = String(urlOrId).match(/\/d\/([a-zA-Z0-9_-]+)/);
  if (m) return m[1];
  var m2 = String(urlOrId).match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (m2) return m2[1];
  if (/^[a-zA-Z0-9_-]{15,}$/.test(urlOrId)) return urlOrId;
  return null;
}

var authClient = null;
function getAuth() {
  if (authClient) return authClient;
  var email = requireEnv("GOOGLE_CLIENT_EMAIL");
  var key = requireEnv("GOOGLE_PRIVATE_KEY").replace(/\\n/g, "\n");
  authClient = new google.auth.GoogleAuth({
    credentials: { client_email: email, private_key: key },
    scopes: ["https://www.googleapis.com/auth/drive"]
  });
  return authClient;
}

var driveClient = null;
function getDrive() {
  if (!driveClient) driveClient = google.drive({ version: "v3", auth: getAuth() });
  return driveClient;
}

function bufferToStream(buffer) {
  var s = new stream.Readable();
  s.push(buffer);
  s.push(null);
  return s;
}

async function downloadFile(fileId) {
  var res = await getDrive().files.get({ fileId: fileId, alt: "media" }, { responseType: "arraybuffer" });
  return Buffer.from(res.data);
}

async function uploadFile(fileId, buffer) {
  await getDrive().files.update({
    fileId: fileId,
    media: { mimeType: XLSX_MIME, body: bufferToStream(buffer) }
  });
}

async function getFileName(fileId) {
  var res = await getDrive().files.get({ fileId: fileId, fields: "name" });
  return res.data.name;
}

module.exports = {
  extractFileId: extractFileId,
  downloadFile: downloadFile,
  uploadFile: uploadFile,
  getFileName: getFileName
};
