#!/usr/bin/env node
// Add or update a login account.
//   npm run add-user -- siva "a real password"
// (the "--" is required so npm passes the arguments through)
//
// Writes the account (as a bcrypt hash, never the plaintext) to
// server/users.local.json for local testing, and prints the FULL account
// list as JSON so you can paste it into your hosting platform's USERS_JSON
// environment variable for the real deployment (see server/README.md).

var bcrypt = require("bcryptjs");
var users = require("../src/users");

var username = process.argv[2];
var password = process.argv[3];

if (!username || !password) {
  console.error("Usage: npm run add-user -- <username> <password>");
  process.exit(1);
}
if (password.length < 8) {
  console.error("Pick a password with at least 8 characters.");
  process.exit(1);
}

var hash = bcrypt.hashSync(password, 10);
var all = users.upsertLocal(username, hash);

console.log("Saved to server/users.local.json (local dev only).\n");
console.log("For the real deployment, set this as the USERS_JSON environment");
console.log("variable on your hosting platform (paste it as one line, no line breaks):\n");
console.log(JSON.stringify(all));
