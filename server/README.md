# Hiring tracker backend

The only thing in this whole project that is allowed to touch the shared Excel file in Google Drive. The frontend (`../data-hiring-guidelines.html`) never sees the Drive credentials, only a JWT it gets from `/auth/login`.

## 1. One-time Google Cloud setup (no admin needed, anyone can do this)

1. Go to [console.cloud.google.com](https://console.cloud.google.com), create a project (or reuse one), free.
2. **APIs & Services > Library**, search **Google Drive API**, click **Enable**.
3. **APIs & Services > Credentials > Create Credentials > Service account**. Give it any name, e.g. "hiring-tracker-backend". No roles needed, skip that step.
4. Open the new service account, **Keys > Add key > Create new key > JSON**. This downloads a `.json` file, keep it somewhere safe, never commit it.
5. From that file:
   - `client_email` becomes `GOOGLE_CLIENT_EMAIL`.
   - `private_key` becomes `GOOGLE_PRIVATE_KEY` (paste it exactly as it appears in the file, including the `\n` sequences and the `-----BEGIN PRIVATE KEY-----`/`-----END PRIVATE KEY-----` lines).

That's it, no admin approval, no tenant consent screen. This is the whole setup.

## 2. The Excel file

1. Put (or create) the roster/scorecard `.xlsx` file in Google Drive, anywhere you like.
2. **Share** it with the service account's email (the `client_email` from step 1, looks like `hiring-tracker-backend@your-project.iam.gserviceaccount.com`), give it **Editor** access, exactly like sharing with a coworker. The file does not need to be public. Only your service account and whoever you share it with can reach it.
3. Copy its normal Drive share link (Share > Copy link) and paste that as `GOOGLE_DRIVE_FILE_URL`. A plain file id works too.

The backend creates two sheets/tabs in it the first time it connects, if they don't already exist: **Candidates** and **Scorecards**, with headers matching the app's roster/scorecard columns. If you're starting from a blank file this just works; if you already have data from a previous export, leave it and it's picked up as-is (as long as those two sheet names hold the same column layout the app writes).

## 3. Local setup

```
cd server
npm install
cp .env.example .env       # fill in the values from steps 1-2, plus a random JWT_SECRET
npm run add-user -- yourname "a real password"
npm run dev
```

`add-user` writes to `users.local.json` (gitignored, local-only) and also prints the account list as JSON, for local dev that's enough; for a real deployment, see step 5.

Visit `http://localhost:8787/health`, it should return `{"ok":true}` immediately, even before the Drive connection finishes (login and the sheet connection are independent; if the sheet fails to connect it retries every 15s and logs why, without crashing the process).

## 4. Point the frontend at it

In `../data-hiring-guidelines.html`, set:
```js
var API_BASE = "http://localhost:8787"; // or your deployed URL, see below
```
(There's a single clearly-marked constant near the top of the app's `<script>` for this.)

## 5. Deploying (example: Render, free tier)

Any Node host works the same way; Render is used here because it's the simplest to set up.

1. Push this repo to GitHub (already done).
2. On [render.com](https://render.com): **New > Web Service**, connect the `Hiring` repo.
3. **Root Directory**: `server`. **Build Command**: `npm install`. **Start Command**: `npm start`.
4. Under **Environment**, add every variable from `.env.example`, with real values, plus:
   - `USERS_JSON`: paste the JSON array `add-user` printed (add more people by running `add-user` again locally and re-pasting the updated array, there's no admin UI on purpose, this is a small internal tool).
   - `CORS_ORIGIN`: set to your GitHub Pages origin exactly, e.g. `https://sivanagireddyy.github.io` (no trailing slash, no path).
   - `GOOGLE_PRIVATE_KEY`: paste it as one line with literal `\n` in it (most hosts, Render included, handle a multi-line paste fine too, either works).
5. Deploy. Render gives you a URL like `https://hiring-tracker-backend.onrender.com`, put that in the frontend's `API_BASE`.

Free-tier Render services sleep after inactivity and take a few seconds to wake on the next request. The frontend's existing "sync failed, retrying" toast pattern already covers that, no extra work needed, but the first request after a lull will feel slow.

## Known trade-offs, on purpose

- **Whole-file read/modify/write on every save**, not a partial-range API. Simple and reliable at this scale (a hiring roster is at most a few hundred rows); would need a smarter approach if this ever grew to a genuinely large sheet.
- **Last-write-wins per row**, exactly like the app's existing `mergeStores()` logic. If two panelists save the very same row within the same refresh window, the later save wins and the earlier one is silently gone. Scorecard rows are naturally isolated per panelist, so this mostly only matters for the Candidates roster.
- **No admin UI for accounts.** `USERS_JSON` is a manually-maintained env var. Fine for a handful of panelists; revisit if the team using this grows a lot.
- **The file itself isn't a public link.** Only the service account (and whoever you've shared it with directly) can open it, which is more private than the original "anyone with the link" idea, with no extra setup cost.
- **The `xlsx` (SheetJS) package has a known, unpatched high-severity advisory** (prototype pollution / ReDoS when parsing a crafted file). The same library, same risk, is already used client-side in the app today. The realistic exposure here is low (only people who already have edit access to this one internal file could craft a malicious workbook), but it's worth knowing about, `npm audit` will keep flagging it.
