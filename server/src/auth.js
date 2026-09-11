var bcrypt = require("bcryptjs");
var jwt = require("jsonwebtoken");
var users = require("./users");

var JWT_SECRET = process.env.JWT_SECRET;
var JWT_EXPIRES_IN = process.env.JWT_EXPIRES_IN || "12h";

if (!JWT_SECRET) {
  console.error("JWT_SECRET is not set. Refusing to start with an insecure default. Set it in your .env / host env vars.");
  process.exit(1);
}

function login(req, res) {
  var username = req.body && req.body.username;
  var password = req.body && req.body.password;
  if (!username || !password) {
    return res.status(400).json({ error: "Username and password are required." });
  }
  var user = users.findByUsername(username);
  if (!user || !bcrypt.compareSync(password, user.passwordHash)) {
    // Same message either way, don't reveal whether the username exists.
    return res.status(401).json({ error: "Incorrect username or password." });
  }
  var token = jwt.sign({ sub: user.username }, JWT_SECRET, { expiresIn: JWT_EXPIRES_IN });
  res.json({ token: token, username: user.username, expiresIn: JWT_EXPIRES_IN });
}

// Express middleware: requires a valid "Authorization: Bearer <token>" header.
function requireAuth(req, res, next) {
  var header = req.headers.authorization || "";
  var m = /^Bearer\s+(.+)$/.exec(header);
  if (!m) return res.status(401).json({ error: "Missing bearer token." });
  try {
    var payload = jwt.verify(m[1], JWT_SECRET);
    req.user = payload.sub;
    next();
  } catch (e) {
    return res.status(401).json({ error: "Invalid or expired token." });
  }
}

module.exports = { login: login, requireAuth: requireAuth };
