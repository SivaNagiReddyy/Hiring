// The only module that knows about the actual Excel file. Everything else
// (routes/index.js) calls the functions below and never touches Graph or
// drive/item IDs directly.
//
// Design:
//  - Two Excel Tables in the workbook: "Candidates" and "Scorecards". Tables
//    (not raw ranges) are required for the row add/update/delete Graph API
//    to work safely.
//  - A row is always found by scanning its "Id" column fresh, not by a
//    cached row index: someone could sort/filter the sheet by hand in
//    Excel Online between saves, and a stale index would silently corrupt
//    the wrong row.
//  - Writes go through a one-at-a-time queue so two near-simultaneous saves
//    from different panelists can't race each other inside a single table.
//  - An in-memory cache backs GET /api/state so most reads are instant; it
//    refreshes on a timer and can be forced.

var graph = require("./graph");
var FACTORS = require("./factors");

var CANDIDATES_TABLE = "Candidates";
var SCORECARDS_TABLE = "Scorecards";

var CAND_HEADERS = [
  "S.No", "Name", "Role", "Designation", "Location", "Day", "Time", "Mode",
  "Panelist", "Phone", "Email", "Resume link", "L1 status", "L2 status",
  "Id", "Updated At"
];
var CARD_HEADERS = ["Candidate", "Role", "Round", "Panelist", "Date", "Outcome"]
  .concat(FACTORS.reduce(function (acc, f) { return acc.concat([f.name, f.name + " - Feedback"]); }, []))
  .concat(["Overall comments / L2 write-up", "Id", "Candidate Id", "Updated At"]);

var driveId = null;
var itemId = null;
var cache = { candidates: {}, cards: {}, lastPull: 0 };
var writeQueue = Promise.resolve();

function enqueue(fn) {
  var result = writeQueue.then(fn, fn);
  // Swallow so one failed write doesn't jam the queue for everyone after it.
  writeQueue = result.catch(function () {});
  return result;
}

function itemBase() {
  if (!driveId || !itemId) throw new Error("sheet.js used before init() resolved the SharePoint file.");
  return "/drives/" + driveId + "/items/" + itemId + "/workbook";
}

async function init() {
  var url = process.env.SHAREPOINT_FILE_URL;
  if (!url) throw new Error("SHAREPOINT_FILE_URL is not set.");
  var shareId = graph.encodeShareUrl(url);
  var item = await graph.graphFetch("/shares/" + shareId + "/driveItem?$select=id,parentReference");
  driveId = item.parentReference.driveId;
  itemId = item.id;
  await ensureTable(CANDIDATES_TABLE, CAND_HEADERS);
  await ensureTable(SCORECARDS_TABLE, CARD_HEADERS);
  await pullAll(true);
}

async function ensureTable(name, headers) {
  var tables = await graph.graphFetch(itemBase() + "/tables");
  var exists = (tables.value || []).some(function (t) { return t.name === name; });
  if (exists) return;

  // Table doesn't exist yet, create it on its own worksheet with just the
  // header row, so a brand-new/blank file works out of the box.
  await graph.graphFetch(itemBase() + "/worksheets/add", { method: "POST", body: { name: name } });
  var lastCol = colLetter(headers.length);
  await graph.graphFetch(itemBase() + "/worksheets('" + name + "')/range(address='A1:" + lastCol + "1')", {
    method: "PATCH",
    body: { values: [headers] }
  });
  var created = await graph.graphFetch(itemBase() + "/worksheets('" + name + "')/tables/add", {
    method: "POST",
    body: { address: name + "!A1:" + lastCol + "1", hasHeaders: true }
  });
  // Graph auto-names new tables (e.g. "Table1"), rename it, addressing it by
  // its stable id (not the name we're in the middle of changing).
  await graph.graphFetch(itemBase() + "/tables('" + created.id + "')/name", {
    method: "PATCH",
    body: { name: name }
  });
}

function colLetter(n) {
  var s = "";
  while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

// ---- Candidates ----

function candidateFromRow(values) {
  var v = values || [];
  return {
    sno: v[0] || "", name: v[1] || "", role: v[2] || "engineer", designation: v[3] || "",
    location: v[4] || "", day: v[5] || "", time: v[6] || "", mode: v[7] || "F2F",
    panelist: v[8] || "", phone: v[9] || "", email: v[10] || "", resume: v[11] || "",
    l1Status: v[12] || "Scheduled", l2Status: v[13] || "Scheduled",
    id: v[14] || "", updatedAt: Number(v[15]) || 0
  };
}
function rowFromCandidate(c) {
  return [
    c.sno || "", c.name || "", c.role || "engineer", c.designation || "", c.location || "",
    c.day || "", c.time || "", c.mode || "F2F", c.panelist || "", c.phone || "", c.email || "",
    c.resume || "", c.l1Status || "Scheduled", c.l2Status || "Scheduled",
    c.id, c.updatedAt || Date.now()
  ];
}

// ---- Scorecards ----

function cardFromRow(values) {
  var v = values || [];
  var i = 6; // first factor column
  var ratings = {}, notes = {};
  FACTORS.forEach(function (f) {
    var rating = v[i], note = v[i + 1];
    if (rating !== "" && rating != null) ratings[f.key] = Number(rating);
    if (note) notes[f.key] = note;
    i += 2;
  });
  return {
    candidate: v[0] || "", role: v[1] || "engineer", round: (v[2] || "l1").toLowerCase(),
    panelist: v[3] || "", date: v[4] || "", overall: v[5] || null,
    ratings: ratings, notes: notes, comments: v[i] || "",
    id: v[i + 1] || "", candidateId: v[i + 2] || "", updatedAt: Number(v[i + 3]) || 0
  };
}
function rowFromCard(c) {
  var row = [c.candidate || "", c.role || "engineer", (c.round || "l1").toUpperCase(), c.panelist || "", c.date || "", c.overall || ""];
  FACTORS.forEach(function (f) {
    var r = c.ratings ? c.ratings[f.key] : null;
    row.push(typeof r === "number" ? r : "");
    row.push((c.notes && c.notes[f.key]) || "");
  });
  row.push(c.comments || "", c.id, c.candidateId || "", c.updatedAt || Date.now());
  return row;
}

// ---- Generic table row helpers ----

async function getRows(tableName) {
  var res = await graph.graphFetch(itemBase() + "/tables('" + tableName + "')/rows?$select=index,values");
  return res.value || [];
}

async function findRowIndex(tableName, idColumnIndex, id) {
  var rows = await getRows(tableName);
  for (var i = 0; i < rows.length; i++) {
    var v = rows[i].values && rows[i].values[0];
    if (v && String(v[idColumnIndex]) === String(id)) return rows[i].index;
  }
  return -1;
}

async function upsertRow(tableName, idColumnIndex, id, values) {
  return enqueue(async function () {
    var index = await findRowIndex(tableName, idColumnIndex, id);
    if (index === -1) {
      await graph.graphFetch(itemBase() + "/tables('" + tableName + "')/rows/add", {
        method: "POST", body: { values: [values] }
      });
    } else {
      await graph.graphFetch(itemBase() + "/tables('" + tableName + "')/rows/itemAt(index=" + index + ")", {
        method: "PATCH", body: { values: [values] }
      });
    }
  });
}

async function deleteRow(tableName, idColumnIndex, id) {
  return enqueue(async function () {
    var index = await findRowIndex(tableName, idColumnIndex, id);
    if (index === -1) return; // already gone
    await graph.graphFetch(itemBase() + "/tables('" + tableName + "')/rows/itemAt(index=" + index + ")/delete", {
      method: "POST"
    });
  });
}

// ---- Public API used by routes/index.js ----

async function pullAll(force) {
  var stale = Date.now() - cache.lastPull > (Number(process.env.SHEET_POLL_INTERVAL_MS) || 30000);
  if (!force && !stale && cache.lastPull) return cache;

  var candRows = await getRows(CANDIDATES_TABLE);
  var cardRows = await getRows(SCORECARDS_TABLE);
  var candidates = {}, cards = {};
  candRows.forEach(function (r) {
    var c = candidateFromRow(r.values && r.values[0]);
    if (c.id) candidates[c.id] = c;
  });
  cardRows.forEach(function (r) {
    var c = cardFromRow(r.values && r.values[0]);
    if (c.id) cards[c.id] = c;
  });
  cache = { candidates: candidates, cards: cards, lastPull: Date.now() };
  return cache;
}

function getCache() { return cache; }

async function upsertCandidate(rec) {
  rec.updatedAt = Date.now();
  await upsertRow(CANDIDATES_TABLE, 14, rec.id, rowFromCandidate(rec));
  cache.candidates[rec.id] = rec;
  return rec;
}
async function deleteCandidate(id) {
  await deleteRow(CANDIDATES_TABLE, 14, id);
  delete cache.candidates[id];
}
async function upsertScorecard(rec) {
  rec.updatedAt = Date.now();
  var i = 6 + FACTORS.length * 2 + 1;
  await upsertRow(SCORECARDS_TABLE, i, rec.id, rowFromCard(rec));
  cache.cards[rec.id] = rec;
  return rec;
}
async function deleteScorecard(id) {
  var i = 6 + FACTORS.length * 2 + 1;
  await deleteRow(SCORECARDS_TABLE, i, id);
  delete cache.cards[id];
}

// Bulk roster replace, mirrors the frontend's clearRoster()+persistCand loop.
async function importRoster(list) {
  return enqueue(async function () {
    var existing = await getRows(CANDIDATES_TABLE);
    for (var i = existing.length - 1; i >= 0; i--) {
      await graph.graphFetch(itemBase() + "/tables('" + CANDIDATES_TABLE + "')/rows/itemAt(index=" + existing[i].index + ")/delete", { method: "POST" });
    }
    cache.candidates = {};
    for (var j = 0; j < list.length; j++) {
      var rec = list[j];
      rec.updatedAt = Date.now();
      await graph.graphFetch(itemBase() + "/tables('" + CANDIDATES_TABLE + "')/rows/add", {
        method: "POST", body: { values: [rowFromCandidate(rec)] }
      });
      cache.candidates[rec.id] = rec;
    }
  });
}

module.exports = {
  init: init,
  pullAll: pullAll,
  getCache: getCache,
  upsertCandidate: upsertCandidate,
  deleteCandidate: deleteCandidate,
  upsertScorecard: upsertScorecard,
  deleteScorecard: deleteScorecard,
  importRoster: importRoster
};
