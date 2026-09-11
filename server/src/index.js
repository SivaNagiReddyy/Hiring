require("dotenv").config();

var express = require("express");
var cors = require("cors");
var auth = require("./auth");
var sheet = require("./sheet");

var app = express();
app.use(express.json({ limit: "2mb" }));

var corsOrigin = process.env.CORS_ORIGIN || "";
app.use(cors({ origin: corsOrigin ? corsOrigin.split(",").map(function (s) { return s.trim(); }) : false }));

app.get("/health", function (req, res) { res.json({ ok: true }); });

app.post("/auth/login", auth.login);

var api = express.Router();
api.use(auth.requireAuth);

api.use(function (req, res, next) {
  if (!sheetReady) return res.status(503).json({ error: "Not connected to the SharePoint sheet yet. Check the server logs." });
  next();
});

api.get("/state", async function (req, res) {
  try {
    var force = req.query.force === "1";
    var data = await sheet.pullAll(force);
    res.json({ candidates: Object.values(data.candidates), cards: Object.values(data.cards) });
  } catch (e) {
    console.error(e);
    res.status(502).json({ error: "Could not read the sheet: " + e.message });
  }
});

api.put("/candidates/:id", async function (req, res) {
  try {
    var rec = req.body || {};
    rec.id = req.params.id;
    await sheet.upsertCandidate(rec);
    res.json({ ok: true });
  } catch (e) {
    console.error(e);
    res.status(502).json({ error: "Could not save the candidate: " + e.message });
  }
});

api.delete("/candidates/:id", async function (req, res) {
  try {
    await sheet.deleteCandidate(req.params.id);
    res.json({ ok: true });
  } catch (e) {
    console.error(e);
    res.status(502).json({ error: "Could not delete the candidate: " + e.message });
  }
});

api.put("/scorecards/:id", async function (req, res) {
  try {
    var rec = req.body || {};
    rec.id = req.params.id;
    await sheet.upsertScorecard(rec);
    res.json({ ok: true });
  } catch (e) {
    console.error(e);
    res.status(502).json({ error: "Could not save the scorecard: " + e.message });
  }
});

api.delete("/scorecards/:id", async function (req, res) {
  try {
    await sheet.deleteScorecard(req.params.id);
    res.json({ ok: true });
  } catch (e) {
    console.error(e);
    res.status(502).json({ error: "Could not delete the scorecard: " + e.message });
  }
});

api.post("/import", async function (req, res) {
  try {
    var list = Array.isArray(req.body) ? req.body : (req.body && req.body.candidates) || [];
    await sheet.importRoster(list);
    res.json({ ok: true, count: list.length });
  } catch (e) {
    console.error(e);
    res.status(502).json({ error: "Could not replace the roster: " + e.message });
  }
});

app.use("/api", api);

var PORT = process.env.PORT || 8787;
var sheetReady = false;

// The server listens immediately: /health and /auth/login work even if the
// SharePoint connection is down or misconfigured, so login problems and
// sheet-connection problems are never confused with each other, and a
// transient Graph outage degrades to "sheet routes return 503" rather than
// taking the whole service down.
app.listen(PORT, function () { console.log("Listening on :" + PORT); });

function connectSheet() {
  sheet.init().then(function () {
    sheetReady = true;
    console.log("Connected to the SharePoint workbook.");
    // Background refresh so GET /api/state is normally served from a warm cache.
    setInterval(function () {
      sheet.pullAll(true).catch(function (e) { console.error("Background sheet refresh failed:", e.message); });
    }, Number(process.env.SHEET_POLL_INTERVAL_MS) || 30000);
  }).catch(function (e) {
    console.error("Could not connect to the SharePoint workbook (will retry in 15s):", e.message);
    setTimeout(connectSheet, 15000);
  });
}
connectSheet();
