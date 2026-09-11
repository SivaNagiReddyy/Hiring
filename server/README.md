# Hiring tracker backend

The only thing in this whole project that is allowed to talk to Microsoft Graph and the shared SharePoint Excel file. The frontend (`../data-hiring-guidelines.html`) never sees the Graph credentials, only a JWT it gets from `/auth/login`.

## 1. One-time Azure setup (someone with Azure AD admin rights)

This is unavoidable no matter how the file is shared: Microsoft Graph never accepts a fully anonymous call, so the backend needs its own registered identity.

1. Go to **portal.azure.com > Azure Active Directory > App registrations > New registration**.
   - Name: anything, e.g. "Hiring Tracker Backend".
   - Supported account types: "Accounts in this organizational directory only".
   - No redirect URI needed (this app never signs in as a user).
2. On the app's **Overview** page, copy the **Application (client) ID** and **Directory (tenant) ID**. These become `GRAPH_CLIENT_ID` and `GRAPH_TENANT_ID`.
3. Go to **Certificates & secrets > New client secret**. Copy the secret's **Value** immediately (it's hidden after you leave the page). This becomes `GRAPH_CLIENT_SECRET`.
4. Go to **API permissions > Add a permission > Microsoft Graph > Application permissions** and add `Sites.Selected`. Click **Grant admin consent** for it.
5. `Sites.Selected` alone doesn't grant access to any specific site yet, it has to be pointed at just this one SharePoint site. Run this once (the easiest way is the **Graph Explorer** at aka.ms/ge, signed in as an admin):

   ```
   POST https://graph.microsoft.com/v1.0/sites/{site-id}/permissions
   Content-Type: application/json

   {
     "roles": ["write"],
     "grantedToIdentities": [{
       "application": { "id": "{GRAPH_CLIENT_ID}", "displayName": "Hiring Tracker Backend" }
     }]
   }
   ```

   To find `{site-id}`: `GET https://graph.microsoft.com/v1.0/sites/{tenant}.sharepoint.com:/sites/{site-name}`.

This is the least-privilege setup: the app can only read/write this one site, nothing else in the tenant.

## 2. The Excel file

Get its **"Anyone with the link can edit"** sharing URL from SharePoint (Share, then Copy link, set to Anyone with the link, Edit). Paste it as `SHAREPOINT_FILE_URL`.

The backend creates two Excel Tables in it the first time it connects, if they don't already exist: **Candidates** and **Scorecards**, with headers matching the app's existing roster/scorecard columns. If you already have data in the file from a previous export, it's fine to leave it, just make sure whichever sheet holds it doesn't already have tables named exactly `Candidates` or `Scorecards` with a different column layout, or point `SHAREPOINT_FILE_URL` at a fresh file for the first run.

## 3. Local setup

```
cd server
npm install
cp .env.example .env       # fill in the values from steps 1-2, plus a random JWT_SECRET
npm run add-user -- yourname "a real password"
npm run dev
```

`add-user` writes to `users.local.json` (gitignored, local-only) and also prints the account list as JSON. For local dev that's enough; for a real deployment, see step 5.

Visit `http://localhost:8787/health`, it should return `{"ok":true}` immediately, even before the SharePoint connection finishes (login and the sheet connection are independent; if the sheet fails to connect it retries every 15s and logs why, without crashing the process).

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
   - `USERS_JSON`: paste the JSON array `add-user` printed (add more people by running `add-user` again locally and re-pasting the updated array; there's no admin UI on purpose, this is a small internal tool).
   - `CORS_ORIGIN`: set to your GitHub Pages origin exactly, e.g. `https://sivanagireddyy.github.io` (no trailing slash, no path).
5. Deploy. Render gives you a URL like `https://hiring-tracker-backend.onrender.com`, put that in the frontend's `API_BASE`.

Free-tier Render services sleep after inactivity and take a few seconds to wake on the next request. The frontend's existing "sync failed, retrying" toast pattern already covers that, no extra work needed, but the first request after a lull will feel slow.

## Known trade-offs, on purpose

- **No file-level Excel encryption.** The JWT login plus `Sites.Selected` scoping already means only this backend can reach the file at all, which was judged sufficient without also taking on a fragile encrypt/decrypt pipeline (Node's support for password-protected `.xlsx` is thin). Revisit if the sharing link itself is ever suspected to have leaked.
- **Last-write-wins per row**, exactly like the app's existing `mergeStores()` logic. If two panelists save the very same row within the same refresh window, the later save wins and the earlier one is silently gone. Scorecard rows are naturally isolated per panelist, so this mostly only matters for the Candidates roster.
- **No admin UI for accounts.** `USERS_JSON` is a manually-maintained env var. Fine for a handful of panelists; revisit if the team using this grows a lot.
