// Login accounts. Two sources, in priority order:
//
//   1. USERS_JSON env var: a JSON array of {username, passwordHash}. This is
//      the PRODUCTION source: most hosting platforms (Render, Fly.io, Azure
//      App Service, etc.) wipe the local filesystem on every redeploy/restart,
//      so a plain file on disk is not a safe place to keep accounts once this
//      is actually deployed. Set this once in your host's environment-variable
//      dashboard and it survives every redeploy.
//   2. server/users.local.json: a local file, used only for local dev when
//      USERS_JSON isn't set. Gitignored; never deployed.
//
// Accounts are added with `npm run add-user -- <username> <password>`
// (see scripts/add-user.js), which never prints or stores a plaintext password.

var fs = require("fs");
var path = require("path");

var LOCAL_FILE = path.join(__dirname, "..", "users.local.json");

function loadLocal() {
  try {
    var raw = fs.readFileSync(LOCAL_FILE, "utf8");
    var data = JSON.parse(raw);
    return Array.isArray(data) ? data : [];
  } catch (e) {
    return [];
  }
}

function saveLocal(users) {
  fs.writeFileSync(LOCAL_FILE, JSON.stringify(users, null, 2) + "\n", "utf8");
}

function load() {
  if (process.env.USERS_JSON) {
    try {
      var fromEnv = JSON.parse(process.env.USERS_JSON);
      if (Array.isArray(fromEnv)) return fromEnv;
    } catch (e) {
      console.error("USERS_JSON is set but is not valid JSON. No one can log in until this is fixed.");
    }
    return [];
  }
  return loadLocal();
}

function findByUsername(username) {
  var users = load();
  var u = String(username || "").trim().toLowerCase();
  return users.find(function (x) { return String(x.username).toLowerCase() === u; }) || null;
}

// Local-dev only (no-op conceptually once USERS_JSON is the active source:
// add-user.js prints the updated JSON so it can be pasted into the host's
// USERS_JSON env var instead).
function upsertLocal(username, passwordHash) {
  var users = loadLocal();
  var u = String(username || "").trim();
  var idx = users.findIndex(function (x) { return String(x.username).toLowerCase() === u.toLowerCase(); });
  var rec = { username: u, passwordHash: passwordHash };
  if (idx === -1) users.push(rec); else users[idx] = rec;
  saveLocal(users);
  return users;
}

function list() {
  return load().map(function (x) { return x.username; });
}

module.exports = { load: load, findByUsername: findByUsername, upsertLocal: upsertLocal, list: list };
