# GoldSoil Payroll Calculator

A static, browser-based payroll calculator that reads commission data live from
a Google Sheet. Lucia opens the URL, clicks one button, sees commissions per
person. No CSV uploads, no install.

> **Live URL** (after deployment): `https://<your-github-username>.github.io/goldsoil-payroll/`

---

## Table of Contents

- [How Lucia uses it](#how-lucia-uses-it)
- [Architecture](#architecture)
- [One-time deployment](#one-time-deployment)
  - [Part A — Google Apps Script backend](#part-a--google-apps-script-backend)
  - [Part B — GitHub Pages front-end](#part-b--github-pages-front-end)
- [Maintenance](#maintenance)
- [v1 scope and known divergences](#v1-scope-and-known-divergences)
- [Updating the canonical Drive docs](#updating-the-canonical-drive-docs)

---

## How Lucia uses it

On the 1st of each month, after the Salesforce → Google Sheet connectors have
refreshed:

1. Open the calculator URL.
2. Period auto-displays as the previous full month (e.g., on May 3rd it shows `2026-04`).
3. Click **Calculate Commissions**.
4. Page fetches the 4 tabs from the master Google Sheet, runs the rules, shows:
   - Per-person totals (sorted, with grand total)
   - Expandable per-person line-by-line breakdown
   - **Flags** panel — manual review items and ineligibles
5. Click **Copy TSV** → paste into the master rules sheet's `Pay_Periods` tab.
6. Process payroll on the 20th of M+1.

Sheet data is fetched directly into the browser via the Apps Script web app.
No CSVs leave anyone's machine.

---

## Architecture

```
   Salesforce reports
          │
          ▼ (Sheets connector, runs nightly/per-cycle)
   Google Sheet (master) ─── 4 tabs ─────────────────────┐
     • Monthly Comm Outreach / Sales  (rollover)         │
     • Approved Leads                  (last month only) │
     • Sales                           (last month only) │
     • Contracted Deals                (rollover)        │
                                                         ▼
                                            Apps Script web app
                                            (apps-script.gs)
                                                         │
                                                         ▼ (JSON over HTTPS)
                                            Front-end (GitHub Pages)
                                              index.html / style.css / script.js
                                                         │
                                                         ▼
                                                runCommissions(period, data)
                                                         │
                                                         ▼
                                          Per-rule calculators →
                                          Render: summary + detail + flags
```

**External dependencies (loaded from CDN, no install):**
- IBM Plex font family — typography (Google Fonts)

**No build step. Pure HTML/CSS/JS.**

---

## One-time deployment

### Part A — Google Apps Script backend

1. **Get the Sheet ID.** Open the master Google Sheet. The ID is the long string
   in the URL between `/d/` and `/edit`:
   `https://docs.google.com/spreadsheets/d/`**`<SHEET_ID>`**`/edit`

2. **Create a shared secret token.** Any long random string. Generate one with:
   ```bash
   openssl rand -hex 32
   ```
   or just use a strong password generator. Save this string somewhere — you'll
   paste it into two places.

3. **Open the Apps Script editor.** In the Google Sheet, go to
   **Extensions → Apps Script**. This opens a new Apps Script project bound to
   the sheet.

4. **Paste the script.** Replace the empty `Code.gs` with the contents of
   `apps-script.gs` from this repo. Fill in:
   ```javascript
   const SHEET_ID = 'paste-sheet-id-here';
   const TOKEN    = 'paste-shared-secret-here';
   ```

5. **Save and test.** Save the file (Ctrl/Cmd+S). In the editor, select the
   `_test_doGet` function from the function dropdown and click **Run**. First
   run, Google will ask for permissions — grant them (it needs to read your
   sheet). The log should show JSON output for the first few rows.

6. **Deploy as web app.** Click **Deploy → New deployment**. Choose type
   **Web app**. Settings:
   - **Description:** "GoldSoil payroll v1"
   - **Execute as:** Me (your account)
   - **Who has access:** Anyone
   - Click **Deploy**.

7. **Copy the deployment URL.** Looks like
   `https://script.google.com/macros/s/AKfycbz.../exec`. You'll paste this into
   `script.js` next.

> ⚠ **Why "Access: Anyone"?** Cross-origin fetch from GitHub Pages can't follow
> Google login redirects cleanly, so domain-restricted access doesn't work in
> practice. Privacy comes from the URL + token both being secrets. Don't share
> the URL or token in chat, screenshots, or commits.
>
> If you need stronger auth later, the alternative is to host the front-end
> *inside* Apps Script (HTMLService) instead of GitHub Pages — same-origin to
> the sheet, Google login enforced. Out of scope for v1.

### Part B — GitHub Pages front-end

1. **Edit `script.js`** in this folder. Set the two config values at the top:
   ```javascript
   const APPS_SCRIPT_URL = 'paste-deployment-url-here';
   const APPS_SCRIPT_TOKEN = 'paste-same-shared-secret-here';
   ```

2. **Create the GitHub repo:**
   - Sign in to https://github.com
   - Click **New repository**
   - Name: `goldsoil-payroll`
   - Visibility: **Public** (required for free GitHub Pages)
   - Click **Create repository**

3. **Upload the files:**
   - On the new empty repo page, click **uploading an existing file**
   - Drag `index.html`, `style.css`, `script.js`, `README.md` (and optionally
     `apps-script.gs` for reference) in
   - Commit message: "Initial deploy"
   - Click **Commit changes**

4. **Enable GitHub Pages:**
   - Repo → **Settings** → **Pages**
   - Source: branch `main`, folder `/ (root)`
   - Save. The URL appears at the top — typically
     `https://<your-username>.github.io/goldsoil-payroll/`. First deploy takes 1–2 min.

5. **Test:** open the URL, click Calculate. You should see per-person results.

6. **Share the URL with Lucia.**

---

## Maintenance

All rule data is hard-coded in `script.js`. Updates = edit the file, commit,
GitHub Pages redeploys in ~30s.

**When someone is hired or leaves:** Edit the `PEOPLE` object near the top of
`script.js`:
```javascript
const PEOPLE = {
  'New Hire Name':   { hireDate: '2026-05-15', separationDate: null, manager: null },
  'Person Who Left': { hireDate: '2024-11-01', separationDate: '2026-03-31', manager: '...' },
};
```
For Outreach Team Managers (qualifies them for the 20% override), add
`isOutreachManager: true`. For direct reports, set `manager: 'Manager Name'`.

**When monthly financials close:** Add to `FINANCE` for Art and Leslie's calcs:
```javascript
const FINANCE = {
  '2026-04': { revenue: 8500000, noi: 1200000 },
};
```

**When a commission rule changes:** Update the canonical Drive doc *first*,
then mirror in `script.js`. Constants are clearly labeled at the top.

**When the Apps Script secret rotates:** Update both `apps-script.gs` (redeploy
as a new version) and `script.js` (commit). Both must match.

**When the sheet adds a new tab:** Add the tab name to the `TABS` array in
`apps-script.gs`, redeploy, then add fetch/use logic in `script.js`.

---

## v1 scope and known divergences

### In scope
LAM Advance/Retraction · LAM Final · LTC · LIA/LIM Component A · LIA/LIM
Component B · LPA · LLP · LPM Portfolio · Outreach Team Manager override ·
NOI execs (Art, Leslie)

### Out of scope for v1
| Item | Status | Disposition |
|---|---|---|
| **LAA** | Skipped per Anshul | Not staffed currently; reinstate when needed |
| **LTM** | Skipped per Anshul | Not staffed currently |
| **LTC owner attribution** | Simplified | All closed-funded rows → Pracy-Ann (sole active LTC). Add LTC column to the Outreach Sales tab when more LTCs are hired. |
| **LSM** | Deferred | Scope ambiguity with LIM; not implemented |
| **LPA/LLP Tier 1 Subdivide Reward** | Manual | $300/month (or $200 each on tie) to top subdivide identifier. Runs off the auditor's HR report, not Salesforce. Lucia adds manually at payroll time. |
| **Regular Payroll (hourly salaries)** | Planned | Will be added as a second tab on this site when ready |

### Divergences from canonical docs (flagged in output)

**LPA/LLP — Quality of Comps ≥95% gate not enforced.** The
`LAP_LLP Commissions` canonical doc gates Tier 2 payouts on ≥95% QoC. v1 of
this calculator pays LPA/LLP at the documented rates (0.3% / 0.15% × AGP)
**regardless of QoC**, and emits a `REVIEW` flag on every LPA/LLP line saying
"QoC ≥95% gate NOT enforced — verify analyst hit the KPI before payout."
Lucia must check each analyst manually before running payroll.

**LAA canonical doc — AGP, not GP.** Anshul confirmed (May 2026) that all
profit-based calcs use AGP. The LAA canonical doc still says GP. LAA isn't
implemented in v1, but when it's reinstated the doc must be updated to v2.1
("LAA = LAM plan, applied to AGP not GP") before code goes in.

---

## Updating the canonical Drive docs

If you change a rule in code, update the canonical doc in Drive too.
Drive is the source of truth; code mirrors it.

- LAM → `/Commission Calculator/LAM Commissions/`
- LTC → `/Commission Calculator/LTC Commissions/`
- LPA / LLP → `/Commission Calculator/LAP_LLP Commissions/`
- LPM → `/Commission Calculator/LPM Commissions/`
- LIA/LIM (Components A + B) → `/Commission Calculator/LIM-LIA Commissions/`
- DOA PH (Art) → `/Commission Calculator/DOA PH Commissions/`
- DOT (Leslie) → `/Commission Calculator/DOT Commissions/`

---

## License

Internal tool. Not for redistribution.
