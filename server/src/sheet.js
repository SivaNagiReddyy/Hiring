// The only module that knows about the actual Excel file. Everything else
// (routes in index.js) calls the functions below and never touches Google
// Drive or the workbook directly.
//
// Design (kept deliberately simple):
//  - One .xlsx file in Google Drive, with two sheets/tabs: "Candidates" and
//    "Scorecards" (same shape the app's own "Export to Excel" already uses).
//  - Every write downloads the whole file, edits it in memory with the same
//    SheetJS library the frontend already uses, and re-uploads the whole
//    file. Simple, and plenty fast at this scale (a hiring roster is at
//    most a few hundred rows), no partial-range/table API to get wrong.
//  - A row is always found by scanning its "Id" column fresh on every
//    write, not a cached position, someone could resort the sheet by
//    hand in Google Sheets between saves.
//  - Writes go through a one-at-a-time queue so two near-simultaneous
//    saves from different panelists can't clobber each other's edit to
//    the file.
//  - An in-memory cache backs GET /api/state so most reads are instant;
//    it refreshes on a timer and can be forced.

var XLSX = require("xlsx");
var drive = require("./drive");
var FACTORS = require("./factors");

var CANDIDATES_SHEET = "Candidates";
var SCORECARDS_SHEET = "Scorecards";

var CAND_HEADERS = [
  "S.No", "Name", "Role", "Designation", "Location", "Day", "Time", "Mode",
  "Panelist", "Phone", "Email", "Resume link", "L1 status", "L2 status",
  "Id", "Updated At"
];
var CARD_HEADERS = ["Candidate", "Role", "Round", "Panelist", "Date", "Outcome"]
  .concat(FACTORS.reduce(function (acc, f) { return acc.concat([f.name, f.name + " - Feedback"]); }, []))
  .concat(["Overall comments / L2 write-up", "Id", "Candidate Id", "Updated At"]);

var CAND_ID_COL = 14;
var CARD_ID_COL = 6 + FACTORS.length * 2 + 1;

var fileId = null;
var cache = { candidates: {}, cards: {}, lastPull: 0 };
var writeQueue = Promise.resolve();

function enqueue(fn) {
  var result = writeQueue.then(fn, fn);
  writeQueue = result.catch(function () {}); // one failed write shouldn't jam the queue
  return result;
}

async function init() {
  var url = process.env.GOOGLE_DRIVE_FILE_URL;
  if (!url) throw new Error("GOOGLE_DRIVE_FILE_URL is not set.");
  fileId = drive.extractFileId(url);
  if (!fileId) throw new Error("Could not find a file id in GOOGLE_DRIVE_FILE_URL, paste the file's normal Google Drive share link.");

  await withWorkbook(function (wb) {
    var changed = false;
    if (ensureSheet(wb, CANDIDATES_SHEET, CAND_HEADERS)) changed = true;
    if (ensureSheet(wb, SCORECARDS_SHEET, CARD_HEADERS)) changed = true;
    return changed;
  });
  await pullAll(true);
}

function ensureSheet(wb, name, headers) {
  if (wb.Sheets[name]) return false;
  var ws = XLSX.utils.aoa_to_sheet([headers]);
  XLSX.utils.book_append_sheet(wb, ws, name);
  return true;
}

// Downloads the file, hands the parsed workbook to mutateFn, and re-uploads
// only if mutateFn returns true. Every call is queued so two writes never
// race each other. Also used read-only at startup (see ensureSheet above)
// simply by returning false.
async function withWorkbook(mutateFn) {
  return enqueue(async function () {
    var buf = await drive.downloadFile(fileId);
    var wb = XLSX.read(buf, { type: "buffer" });
    var changed = mutateFn(wb);
    if (changed) await drive.uploadFile(fileId, XLSX.write(wb, { type: "buffer", bookType: "xlsx" }));
    return wb;
  });
}

function sheetToAoa(wb, name) {
  var ws = wb.Sheets[name];
  if (!ws) return [];
  return XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
}
function aoaToSheet(wb, name, aoa) {
  wb.Sheets[name] = XLSX.utils.aoa_to_sheet(aoa);
}

// ---- Candidates ----

function candidateFromRow(v) {
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

function cardFromRow(v) {
  var i = 6; // first factor column
  var ratings = {}, notes = {};
  FACTORS.forEach(function (f) {
    var rating = v[i], note = v[i + 1];
    if (rating !== "" && rating != null) ratings[f.key] = Number(rating);
    if (note) notes[f.key] = note;
    i += 2;
  });
  return {
    candidate: v[0] || "", role: v[1] || "engineer", round: String(v[2] || "l1").toLowerCase(),
    panelist: v[3] || "", date: v[4] || "", overall: v[5] || null,
    ratings: ratings, notes: notes, comments: v[i] || "",
    id: v[i + 1] || "", candidateId: v[i + 2] || "", updatedAt: Number(v[i + 3]) || 0
  };
}
function rowFromCard(c) {
  var row = [c.candidate || "", c.role || "engineer", String(c.round || "l1").toUpperCase(), c.panelist || "", c.date || "", c.overall || ""];
  FACTORS.forEach(function (f) {
    var r = c.ratings ? c.ratings[f.key] : null;
    row.push(typeof r === "number" ? r : "");
    row.push((c.notes && c.notes[f.key]) || "");
  });
  row.push(c.comments || "", c.id, c.candidateId || "", c.updatedAt || Date.now());
  return row;
}

function isBlankRow(row) {
  return !row || row.every(function (c) { return c === "" || c == null; });
}

// ---- Public API used by routes in index.js ----

async function pullAll(force) {
  var stale = Date.now() - cache.lastPull > (Number(process.env.SHEET_POLL_INTERVAL_MS) || 30000);
  if (!force && !stale && cache.lastPull) return cache;

  var buf = await drive.downloadFile(fileId);
  var wb = XLSX.read(buf, { type: "buffer" });
  var candAoa = sheetToAoa(wb, CANDIDATES_SHEET);
  var cardAoa = sheetToAoa(wb, SCORECARDS_SHEET);

  var candidates = {}, cards = {};
  for (var i = 1; i < candAoa.length; i++) {
    if (isBlankRow(candAoa[i])) continue;
    var c = candidateFromRow(candAoa[i]);
    if (c.id) candidates[c.id] = c;
  }
  for (var j = 1; j < cardAoa.length; j++) {
    if (isBlankRow(cardAoa[j])) continue;
    var sc = cardFromRow(cardAoa[j]);
    if (sc.id) cards[sc.id] = sc;
  }
  cache = { candidates: candidates, cards: cards, lastPull: Date.now() };
  return cache;
}

function getCache() { return cache; }

async function upsertCandidate(rec) {
  rec.updatedAt = Date.now();
  await withWorkbook(function (wb) {
    var aoa = sheetToAoa(wb, CANDIDATES_SHEET);
    if (!aoa.length) aoa = [CAND_HEADERS];
    var rowIdx = -1;
    for (var i = 1; i < aoa.length; i++) { if (String(aoa[i][CAND_ID_COL]) === String(rec.id)) { rowIdx = i; break; } }
    var row = rowFromCandidate(rec);
    if (rowIdx === -1) aoa.push(row); else aoa[rowIdx] = row;
    aoaToSheet(wb, CANDIDATES_SHEET, aoa);
    return true;
  });
  cache.candidates[rec.id] = rec;
  return rec;
}

async function deleteCandidate(id) {
  await withWorkbook(function (wb) {
    var aoa = sheetToAoa(wb, CANDIDATES_SHEET);
    var next = aoa.filter(function (row, i) { return i === 0 || String(row[CAND_ID_COL]) !== String(id); });
    if (next.length === aoa.length) return false;
    aoaToSheet(wb, CANDIDATES_SHEET, next);
    return true;
  });
  delete cache.candidates[id];
}

async function upsertScorecard(rec) {
  rec.updatedAt = Date.now();
  await withWorkbook(function (wb) {
    var aoa = sheetToAoa(wb, SCORECARDS_SHEET);
    if (!aoa.length) aoa = [CARD_HEADERS];
    var rowIdx = -1;
    for (var i = 1; i < aoa.length; i++) { if (String(aoa[i][CARD_ID_COL]) === String(rec.id)) { rowIdx = i; break; } }
    var row = rowFromCard(rec);
    if (rowIdx === -1) aoa.push(row); else aoa[rowIdx] = row;
    aoaToSheet(wb, SCORECARDS_SHEET, aoa);
    return true;
  });
  cache.cards[rec.id] = rec;
  return rec;
}

async function deleteScorecard(id) {
  await withWorkbook(function (wb) {
    var aoa = sheetToAoa(wb, SCORECARDS_SHEET);
    var next = aoa.filter(function (row, i) { return i === 0 || String(row[CARD_ID_COL]) !== String(id); });
    if (next.length === aoa.length) return false;
    aoaToSheet(wb, SCORECARDS_SHEET, next);
    return true;
  });
  delete cache.cards[id];
}

// Bulk roster replace, mirrors the frontend's clearRoster()+persistCand loop.
async function importRoster(list) {
  await withWorkbook(function (wb) {
    var aoa = [CAND_HEADERS];
    list.forEach(function (rec) {
      rec.updatedAt = Date.now();
      aoa.push(rowFromCandidate(rec));
    });
    aoaToSheet(wb, CANDIDATES_SHEET, aoa);
    return true;
  });
  cache.candidates = {};
  list.forEach(function (rec) { cache.candidates[rec.id] = rec; });
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
