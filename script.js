/* ============================================================
   GoldSoil Payroll Calculator — script.js (v1)
   ----------------------------------------------------------------
   Fetches 4 tabs from a Google Sheet (via Apps Script web app),
   runs commission rules, renders per-person payouts.
   See README.md for deployment + maintenance.

   v1 scope (per Anshul, May 2026):
     - LAM Advance/Retraction (Contracted Deals)
     - LAM Final (Monthly Comm Outreach / Sales)
     - LTC (defaults to Pracy-Ann for all closed-funded rows)
     - LPA (0.3% × AGP — QoC gate NOT enforced)
     - LLP (0.15% × AGP — QoC gate NOT enforced)
     - LIA/LIM Component A (Approved Leads)
     - LIA/LIM Component B (Monthly Comm Outreach / Sales)
     - LPM portfolio (Sales tab)
     - Outreach Team Manager override (Component A only)
     - NOI execs (Art Oaing, Leslie Bernolo) — pending revision

   OUT OF SCOPE for v1:
     - LAA, LTM (skipped per Anshul)
     - LSM (deferred — scope ambiguity with LIM)
     - LPA/LLP Tier 1 Subdivide Reward (manual via HR auditor report)
     - QoC ≥95% gate enforcement for LPA/LLP (skipped — diverges from
       LAP_LLP canonical doc; flagged on every LPA/LLP line)

   AGP-everywhere policy: all profit-based calcs use Adj Gross Profit,
   not GP. Code matches LAA canonical doc v2.1 (pending Drive update).
   ============================================================ */

/* ------------------------------------------------------------
   0. CONFIG — fill these in before first deploy
   ------------------------------------------------------------ */
const APPS_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbwrvtRktEP1nylsf3WaRRbpN4NtIIjrlARINsn0DKGFoYa1ipqgJKybjOrnFVcVCLpzLA/exec';
const APPS_SCRIPT_TOKEN = '9EMY8RYSE5WUISFY2ZOBADP9F2FIOCDL';  // must match TOKEN in apps-script.gs

// CC_RECIPIENT_NAME must match exactly how Anshul appears in the
// "Person Name" column of the People tab. The send-emails flow looks
// up his email from that row — no hardcoded address. If his name in
// the sheet isn't literally "Anshul", change this.
const CC_RECIPIENT_NAME = 'Anshul Sharma';

// Tab names — must match the Google Sheet exactly
const TAB_OUTREACH_SALES = 'Monthly Comm Outreach / Sales';
const TAB_APPROVED_LEADS = 'Approved Leads';
const TAB_SALES          = 'Sales';
const TAB_CONTRACTED     = 'Contracted Deals';
const TAB_PEOPLE         = 'People';
const TAB_FINANCE        = 'Finance';
const TAB_LAM_ADVANCE    = 'LAM Advance';        // historical advances for Sept 2025–Mar 2026 signings
const TAB_HISTORY        = 'Commission_History';  // optional — populated by writeCommissionHistory

// Trends panel default window: how many recent months of history to
// include in charts/metrics. 12 = rolling year, the most common ask.
// If you want longer/shorter views later, surface this as a UI toggle
// rather than changing this constant.
const TREND_MONTHS = 12;

// Color palette for the Trends stacked bar chart. Stays close to the
// GoldSoil brand neutrals + accents so the chart looks like it belongs
// to the rest of the dashboard rather than a Chart.js demo. Keys must
// match Department / Role values written into Commission_History
// (see buildHistoryRows). Unknown keys fall back to the (Unassigned)
// gray at render time.
const DEPT_COLORS = {
  'Acquisitions': '#9A7820',
  'Sales':        '#2E5D3D',
  'Transactions': '#4A5468',
  'Operations':   '#C4A551',
  'Marketing':    '#8B2E2E',
  '(Unassigned)': '#8A8E97',
};

const ROLE_COLORS = {
  'LAM':         '#9A7820',
  'LIM / LIA':   '#C4A551',
  'LPA / LLP':   '#4A5468',
  'LTC':         '#2E5D3D',
  'LPM':         '#6B4F8F',
  'Outreach Mgr':'#8B2E2E',
  '':            '#8A8E97',  // role blank on People tab
  '(No Role)':   '#8A8E97',
};

/* ------------------------------------------------------------
   1. PEOPLE — populated at runtime from the People tab in the sheet
   ----------------------------------------------------------------
   v1.1: HR maintains the roster directly in Google Sheets (no code
   commits needed). Sheet columns:
     Person Name | Status | Manager | Is Outreach Manager

   Eligibility model (replaces old hire/separation-date logic):
     - In roster + Status=Active   → paid
     - In roster + Status=Inactive → INELIGIBLE ($0)
     - NOT in roster               → INELIGIBLE ($0) + "add to People
                                      tab if they should be paid" flag

   Status=Inactive and not-in-roster behave the same way at payout
   time, but the flag text differs so HR can see new names that may
   have been overlooked vs. people who were intentionally excluded.
   ------------------------------------------------------------ */
let PEOPLE = {};

const LTC_DEFAULT_PERSON = 'Pracy Ann Pryce';

function parseStatus(v) {
  const s = (v == null ? '' : String(v)).trim().toLowerCase();
  // Default to Active if blank or unrecognized — the row's presence in
  // the roster signals intent to track them; blank Status means HR added
  // them but hasn't filled in everything yet.
  return s === 'inactive' ? 'Inactive' : 'Active';
}

function parseManager(v) {
  if (v == null) return null;
  const s = normalizeName(v);
  if (!s) return null;
  const low = s.toLowerCase();
  if (low === 'n/a' || low === 'na' || low === 'none' || low === '-') return null;
  return s;
}

function parseOutreachMgr(v) {
  if (v === true) return true;
  if (v === false || v == null) return false;
  const s = String(v).trim().toLowerCase();
  return s === 'true' || s === 'yes' || s === 'y' || s === '1';
}

// Normalize a person's name for roster lookup. Trims leading/trailing
// whitespace AND collapses internal multi-spaces to a single space.
// Case is preserved. This catches the most common typos that would
// otherwise make a real roster entry fail to match its SF attribution
// (e.g., "Whitney  Simpson" with a double space vs. "Whitney Simpson").
function normalizeName(s) {
  if (s == null) return '';
  return String(s).replace(/\s+/g, ' ').trim();
}

// Resolve an attribution name (Salesforce role-owner field, POD column, etc.)
// to a roster KEY. Tiers, first hit wins:
//   1. exact normalized match (whitespace differences)
//   2. case-insensitive match
//   3. unique tolerant match — first name equal + last token compatible
//      (equal / initial / prefix), via tolerantNameMatch. Accepted ONLY when
//      EXACTLY ONE roster member matches; 0 or 2+ candidates → no match, so an
//      ambiguous or genuinely-unknown name still flags (REVIEW/INELIGIBLE)
//      instead of being silently mis-paid.
// The tolerant tier handles SF spelling variants we can't change because the
// roster Person Name is tied to Deel — e.g. "Erl Timothy Gutierrez" → roster
// "Erl Gutierrez", "Art Justine Oaing" → roster "Art Oaing".
// (tolerantNameMatch is a hoisted function declaration defined in section 11b.)
function resolveRosterKey(rawName) {
  if (!rawName) return null;
  const normalized = normalizeName(rawName);
  if (!normalized) return null;
  if (PEOPLE[normalized]) return normalized;
  const lower = normalized.toLowerCase();
  for (const key of Object.keys(PEOPLE)) {
    if (key.toLowerCase() === lower) return key;
  }
  let match = null, count = 0;
  for (const key of Object.keys(PEOPLE)) {
    if (tolerantNameMatch(normalized, key)) { match = key; count++; if (count > 1) break; }
  }
  return count === 1 ? match : null;
}

// Look up a person's roster entry (or null if unresolved). Backed by
// resolveRosterKey so SF spelling variants resolve to the right person.
function lookupPerson(rawName) {
  const key = resolveRosterKey(rawName);
  return key ? PEOPLE[key] : null;
}

// Canonical roster Person Name for an attribution string, so spelling variants
// aggregate / pay / display under one identity. Falls back to the normalized raw
// name when unresolved (keeps unknown names visible and flaggable).
function canonicalName(rawName) {
  return resolveRosterKey(rawName) || normalizeName(rawName);
}

function buildPeopleFromSheet(peopleRows) {
  const people = {};
  for (const row of (peopleRows || [])) {
    // Be flexible on the header name — accept Person Name, Name, etc.
    const name = normalizeName(row['Person Name'] || row['Name'] || row['Person'] || '');
    if (!name) continue;
    people[name] = {
      status: parseStatus(row['Status']),
      manager: parseManager(row['Manager']),
      isOutreachManager: parseOutreachMgr(row['Is Outreach Manager'] || row['Outreach Manager']),
      department: normalizeName(row['Department'] || '') || null,
      role: normalizeName(row['Role'] || '') || null,
      email: normalizeName(row['Email'] || '') || null,
      // Payment Mode (e.g. "Deel", "OnPay") gates the CSV download — only
      // rows matching the selected toggle end up in the file. Blank = the
      // person is silently excluded from every download.
      paymentMode: normalizeName(row['Payment Mode'] || row['PaymentMode'] || '') || null,
      // Contract Name = how the person is registered in their payment
      // platform (Deel / OnPay). It's the first column in the download.
      // If blank, the row is dropped from the CSV and the count appears
      // in the post-download status line.
      contractName: normalizeName(row['Contract Name'] || row['ContractName'] || '') || null,
    };
  }
  return people;
}

/* ------------------------------------------------------------
   2. FINANCE — populated at runtime from the Finance tab in the sheet
   ----------------------------------------------------------------
   v1.1: monthly revenue and NOI moved out of script.js into the sheet.
   Sheet columns:
     Period | Revenue | NOI

   Period is a YYYY-MM string. If you accidentally enter '2026-04-01'
   (Sheets auto-formats some inputs as dates), the parser strips it to
   '2026-04' so calcs still work.

   An empty Finance tab is fine — calcNOITiered emits a REVIEW flag
   noting the missing period entry; it doesn't crash.
   ------------------------------------------------------------ */
let FINANCE = {};

// Last fetched Commission_History. Hydrated by fetchSheetData on every
// Calculate, then locally updated by writeCommissionHistoryAsync to
// include the just-computed period. The trends panel renders from
// this — no DOM caching, no fetching between renders.
let lastHistoryData = [];

function parsePeriodKey(v) {
  if (v == null) return null;
  // Apps Script serializes Date cells as 'YYYY-MM-DD'. Either way, slice to YYYY-MM.
  const s = String(v).trim();
  const m = s.match(/^(\d{4}-\d{2})/);
  return m ? m[1] : null;
}

function buildFinanceFromSheet(rows) {
  const out = {};
  for (const row of (rows || [])) {
    const key = parsePeriodKey(row['Period'] || row['Month']);
    if (!key) continue;
    const revenue = parseMoney(row['Revenue']);
    const noi = parseMoney(row['NOI'] || row['Net Operating Income']);
    out[key] = { revenue, noi };
  }
  return out;
}

/* ------------------------------------------------------------
   2b. LAM ADVANCE LOOKUP — populated from the "LAM Advance" tab
   ----------------------------------------------------------------
   Historical advances paid Sept 2025 – Mar 2026 under the old tier
   schedule. Keyed by Contract Name. Used by both calcLAMAdvances
   (retraction path) and calcLAMFinals (advance-deduction path) when
   the contract was signed inside the historical window.

   Sheet columns expected:
     Signed Contract Received Date | Contract Name | Seller Name |
     Transaction Name | LAM Advance Amount

   Only Contract Name and LAM Advance Amount are read here — the
   other columns are kept on the sheet for human reference.

   If a contract from the historical window is NOT on this tab,
   the calculator flags REVIEW rather than guessing — we'd rather
   pause than pay/claw back the wrong amount.
   ------------------------------------------------------------ */
let LAM_ADVANCE_LOOKUP = {};

function buildLAMAdvanceLookup(rows) {
  const out = {};
  for (const row of (rows || [])) {
    const contractName = normalizeName(row['Contract Name'] || row['Contract'] || '');
    if (!contractName) continue;
    const amount = parseMoney(row['LAM Advance Amount'] || row['Advance Amount'] || row['Advance']);
    out[contractName] = amount;
  }
  return out;
}

/* ------------------------------------------------------------
   3. COMMISSION RULE TABLES
   ------------------------------------------------------------ */
const LAM_ADVANCE_TIERS = [
  { min: 0,      max: 19999,    advance: 50  },
  { min: 20000,  max: 49999,    advance: 100 },
  { min: 50000,  max: 99999,    advance: 150 },
  { min: 100000, max: 199999,   advance: 200 },
  { min: 200000, max: Infinity, advance: 250 },
];

// LAM advance program has two distinct rule periods:
//
//   [1] Sept 1, 2025 – Mar 31, 2026 (HISTORICAL):
//       Advances were paid under an earlier, different tier schedule.
//       Actual paid amounts are kept on the "LAM Advance" sheet tab and
//       looked up by Contract Name. This calculator does NOT generate new
//       advance payments for this window (those were paid manually months
//       ago) — but for any contract from this window that cancels OR
//       closes now, we use the LAM Advance tab as the source of truth
//       for the retraction amount and for the advance to net out of the
//       LAM Final payout.
//
//   [2] Apr 1, 2026 onward (CURRENT):
//       Advances use LAM_ADVANCE_TIERS below ($50–$250 by Deal Spread).
//       This is the only window where the calculator generates new
//       advance entries; retractions and finals for these contracts
//       also use the tier table.
//
// Contracts signed BEFORE Sept 1, 2025 received no advance under either
// rule, so they generate no retraction and have no advance to deduct
// from the final.
const LAM_ADVANCE_START_DATE     = '2025-09-01';  // first day any advance was paid (historical window starts)
const LAM_ADVANCE_NEW_RULE_DATE  = '2026-04-01';  // first day current tier schedule applies

// LAM v2: tier rates applied to AGP (not GP)
const LAM_FINAL_TIERS = [
  { min: 0,      max: 19999,    rate: 0.06 },
  { min: 20000,  max: 49999,    rate: 0.07 },
  { min: 50000,  max: 99999,    rate: 0.08 },
  { min: 100000, max: 199999,   rate: 0.09 },
  { min: 200000, max: Infinity, rate: 0.10 },
];

const LTC_FLAT_BONUS = 50;

// LPA/LLP — Tier 2 only (Tier 1 Subdivide Reward handled manually)
// Per LAP_LLP canonical doc, both rules gate on Quality of Comps ≥95%.
// v1 does NOT enforce the gate — every LPA/LLP line flagged for manual
// review so Lucia confirms QoC before payout.
const LPA_RATE = 0.003;   // 0.3% × AGP
const LLP_RATE = 0.0015;  // 0.15% × AGP

const LPM_MATRIX = [
  { domMin: 0,  domMax: 29,       parHigh: 0.0150, parMid: 0.0100, parLow: 0.0075   },
  { domMin: 30, domMax: 45,       parHigh: 0.0125, parMid: 0.0083, parLow: 0.00625  },
  { domMin: 46, domMax: 60,       parHigh: 0.0100, parMid: 0.0067, parLow: 0.0050   },
  { domMin: 61, domMax: Infinity, parHigh: 0.0075, parMid: 0.0050, parLow: 0.00375  },
];

const LEAD_BONUSES = {
  'Phone Fully Q': { split: { lia: 1, lim: 4 }, samePerson: { lia: 0, lim: 5 } },
  'Phone Semi Q':  { split: { lia: 1, lim: 2 }, samePerson: { lia: 0, lim: 3 } },
  'Text Fully Q':  { split: { lia: 1, lim: 0 }, samePerson: { lia: 1, lim: 0 } },
  'LaLa Land':     { split: { lia: 0, lim: 0 }, samePerson: { lia: 0, lim: 0 } },
  'Fantasy Land':  { split: { lia: 0, lim: 0 }, samePerson: { lia: 0, lim: 0 } },
};

const LIM_MANAGER_OVERRIDE_RATE = 0.20;
const LIA_LIM_CLOSING_RATE = 0.003;  // 0.3% × AGP, split per LEAD_BONUSES proportions

const NOI_TIERS = {
  'Art Oaing':      [ { revMax: 9999999, rate: 0.005 }, { revMax: 19999999, rate: 0.006 }, { revMax: Infinity, rate: 0.007 } ],
  'Leslie Bernolo': [ { revMax: 9999999, rate: 0.010 }, { revMax: 19999999, rate: 0.011 }, { revMax: Infinity, rate: 0.012 } ],
};

const LPM_POD_MAPPING = {
  'Gabriel': 'Gabriel Santos',
  'Jose': 'Jose Membreño',
};
function mapLPMPod(pod) {
  const t = (pod || '').trim();
  return LPM_POD_MAPPING[t] || t;
}

/* ------------------------------------------------------------
   4. HELPERS
   ------------------------------------------------------------ */
function parseDate(s) {
  if (!s) return null;
  if (s instanceof Date) return s;
  const trimmed = String(s).trim();
  if (!trimmed) return null;
  // Date-only YYYY-MM-DD: parse as LOCAL time. JavaScript's `new Date("2026-04-01")`
  // interprets a bare date as UTC midnight, which in timezones west of UTC rolls
  // back to the previous day. That silently dropped first-of-month entries like
  // SC-066142 (signed 2026-04-01) out of the April period. Bare date strings
  // should mean "this day in the user's locale", not "midnight UTC".
  if (/^\d{4}-\d{2}-\d{2}$/.test(trimmed)) {
    const [y, m, d] = trimmed.split('-').map(n => parseInt(n, 10));
    return new Date(y, m - 1, d);
  }
  // ISO with time component (e.g. "2026-04-01T05:00:00.000Z") — JS parses correctly.
  if (/^\d{4}-\d{2}-\d{2}T/.test(trimmed)) return new Date(trimmed);
  if (/^\d{1,2}\/\d{1,2}\/\d{4}/.test(trimmed)) {
    const [m, d, y] = trimmed.split('/').map(n => parseInt(n, 10));
    return new Date(y, m - 1, d);
  }
  const dt = new Date(trimmed);
  return isNaN(dt) ? null : dt;
}

function ymKey(date) {
  if (!date) return null;
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, '0');
  return `${y}-${m}`;
}

// Display-friendly period name: '2026-04' → 'April 2026'.
// Internal/storage format stays YYYY-MM for sortability + CSV filenames.
function fmtPeriodName(period) {
  if (!period) return '';
  const [yearStr, monthStr] = String(period).split('-');
  const m = parseInt(monthStr, 10);
  if (isNaN(m) || m < 1 || m > 12) return period;
  const months = ['January', 'February', 'March', 'April', 'May', 'June',
                  'July', 'August', 'September', 'October', 'November', 'December'];
  return `${months[m - 1]} ${yearStr}`;
}

function inPeriod(dateStr, period) {
  const d = parseDate(dateStr);
  return d && ymKey(d) === period;
}

function parseMoney(s) {
  if (s === null || s === undefined || s === '') return 0;
  if (typeof s === 'number') return s;
  return parseFloat(String(s).replace(/[$,\s]/g, '')) || 0;
}

function fmtUSD(amount) {
  return new Intl.NumberFormat('en-US', {
    style: 'currency', currency: 'USD', minimumFractionDigits: 2
  }).format(amount);
}

function fmtPct(rate) {
  return (rate * 100).toFixed(rate < 0.01 ? 3 : 2) + '%';
}

function tierFor(amount, tiers) {
  for (const t of tiers) {
    const min = t.min ?? t.agpMin;
    const max = t.max ?? t.agpMax;
    if (amount >= min && amount <= max) return t;
  }
  return null;
}

// v1.1 eligibility check — reads Status from the People tab.
//   In roster + Active   → paid normally
//   In roster + Inactive → INELIGIBLE ($0, "marked Inactive" flag)
//   NOT in roster        → INELIGIBLE ($0, "add to People tab" flag)
//
// The dealDate parameter is preserved for compatibility (some callers
// pass it) but is no longer used — temporal forfeit logic was replaced
// with status-based logic per Anshul, May 2026. If forfeit-on-separation
// with a specific date is ever needed again, add a Separation Date
// column to the People tab and reintroduce date math here.
function checkSeparationOnly(personName, dealDate) {
  const p = lookupPerson(personName);
  if (!p) {
    return {
      eligible: false,
      reason: `Not in People tab. Add "${personName}" with Status=Active to pay; otherwise leave off (treated as inactive).`,
      flag: 'INELIGIBLE',
    };
  }
  if (p.status === 'Inactive') {
    return {
      eligible: false,
      reason: `Marked Inactive in People tab.`,
      flag: 'INELIGIBLE',
    };
  }
  return { eligible: true, reason: null, flag: 'OK' };
}

// Non-person values that may appear in role-owner fields
const NON_PERSON_VALUES = new Set([
  '', 'na', 'n/a', 'none', 'no assigned closer', 'no assigned agent',
  'no assigned setter', 'unassigned', 'tbd', '?',
]);
function isNonPerson(v) {
  if (!v) return true;
  return NON_PERSON_VALUES.has(String(v).trim().toLowerCase());
}

/* ------------------------------------------------------------
   5. COMMISSION CALCULATORS
   ------------------------------------------------------------ */

// LAM Advance + Retraction — reads Contracted Deals (rollover).
// Program effective for contracts signed on/after LAM_ADVANCE_START_DATE.
// Resolve the advance amount for a single contract, given its signed date,
// Contract Name, and Deal Spread. Returns:
//   { amount, source, error }
// where source ∈ {'historical-tab', 'tier-table', 'pre-program'} and error
// is non-null only when the historical-window contract is missing from the
// LAM Advance tab (caller should flag REVIEW in that case).
//
// This is the single source of truth for "how much was this LAM's advance"
// used by both the retraction path and the LAM Final advance-deduction path.
function resolveLAMAdvanceForContract(signedDateParsed, contractName, spread) {
  const historicalStart = parseDate(LAM_ADVANCE_START_DATE);     // 2025-09-01
  const newRuleStart    = parseDate(LAM_ADVANCE_NEW_RULE_DATE);  // 2026-04-01

  if (!signedDateParsed || signedDateParsed < historicalStart) {
    return { amount: 0, source: 'pre-program', error: null };
  }

  if (signedDateParsed < newRuleStart) {
    // Sept 2025 – Mar 2026 historical window: look up actual paid amount
    // from the LAM Advance tab. Missing = REVIEW, don't guess.
    const key = normalizeName(contractName);
    if (key && Object.prototype.hasOwnProperty.call(LAM_ADVANCE_LOOKUP, key)) {
      return { amount: LAM_ADVANCE_LOOKUP[key], source: 'historical-tab', error: null };
    }
    return {
      amount: 0,
      source: 'historical-tab',
      error: `Contract signed ${ymKey(signedDateParsed)} (historical advance window) but "${contractName}" not found on LAM Advance tab. Add the row with the actual advance paid, then re-run.`,
    };
  }

  // Apr 2026+: current tier schedule.
  const tier = tierFor(spread, LAM_ADVANCE_TIERS);
  if (!tier) {
    return {
      amount: 0,
      source: 'tier-table',
      error: `Spread ${fmtUSD(spread)} outside tier table`,
    };
  }
  return { amount: tier.advance, source: 'tier-table', error: null };
}

function calcLAMAdvances(contractedDeals, period) {
  const entries = [];
  const historicalStart = parseDate(LAM_ADVANCE_START_DATE);
  const newRuleStart    = parseDate(LAM_ADVANCE_NEW_RULE_DATE);
  for (const d of contractedDeals) {
    const contractId = d['Contract Name'] || d.Contract_Name || d.Deal_ID || '?';
    const owner = (d['Closer (Acq)'] || d.Closer_Acq || d.LAM_Owner || d.LAM || '').trim();
    const stage = (d['Contract Stage'] || d.Contract_Stage || '').trim();
    const spread = parseMoney(d['Deal Spread'] || d.Deal_Spread);
    const signedDate = d['Signed Contract Received Date'] || d.Signed_Contract_Received_Date || d.Contract_Signed_Date;
    const cxDate = d['Date Cancelled'] || d.Date_Cancelled || d.Cancellation_Date;

    // Pre-program contracts (signed before Sept 2025) received no advance
    // under any rule, so they generate neither an advance nor a retraction.
    const signedDateParsed = parseDate(signedDate);
    const isCoveredByProgram = signedDateParsed && historicalStart && signedDateParsed >= historicalStart;
    if (!isCoveredByProgram) continue;

    // Whether this contract falls in the historical window (Sept 2025 – Mar 2026)
    // or the new-rule window (Apr 2026+). Used by both paths below.
    const isHistorical = signedDateParsed < newRuleStart;

    // Advance path: signed in target month.
    // The calculator only generates NEW advance entries for the current
    // tier-rule window (Apr 2026+). Historical-window advances were paid
    // manually back in 2025/early 2026 — we don't re-pay them here even
    // if a historical signing happens to fall inside `period`.
    if (inPeriod(signedDate, period) && !isHistorical) {
      if (isNonPerson(owner)) {
        entries.push({
          person: '—', role: 'LAM', period, source: contractId, type: 'LAM Advance',
          amount: 0, calc: '—',
          notes: `Closer (Acq) is non-person ("${owner || 'blank'}"). Fix SF record.`,
          flag: 'REVIEW',
        });
      } else {
        const tier = tierFor(spread, LAM_ADVANCE_TIERS);
        if (!tier) {
          entries.push({
            person: owner, role: 'LAM', period, source: contractId, type: 'LAM Advance',
            amount: 0, calc: `Spread ${fmtUSD(spread)} outside tiers`,
            notes: 'Check Deal Spread', flag: 'REVIEW',
          });
        } else {
          const elig = checkSeparationOnly(owner, signedDate);
          if (!elig.eligible) {
            entries.push({
              person: owner, role: 'LAM', period, source: contractId, type: 'LAM Advance',
              amount: 0, calc: '—', notes: elig.reason, flag: 'INELIGIBLE',
            });
          } else {
            entries.push({
              person: owner, role: 'LAM', period, source: contractId, type: 'LAM Advance',
              amount: tier.advance,
              calc: `Spread ${fmtUSD(spread)} → tier ${fmtUSD(tier.min)}–${tier.max === Infinity ? '∞' : fmtUSD(tier.max)} → ${fmtUSD(tier.advance)}`,
              notes: '', flag: 'OK',
            });
          }
        }
      }
    }

    // Retraction path: cancelled in target month, stage = Cancelled.
    // Uses resolveLAMAdvanceForContract so historical-window cancellations
    // claw back the actual paid amount from the LAM Advance tab, while
    // new-rule cancellations claw back the tier-table amount.
    if (inPeriod(cxDate, period) && stage === 'Cancelled') {
      if (isNonPerson(owner)) {
        entries.push({
          person: '—', role: 'LAM', period, source: contractId, type: 'LAM Retraction',
          amount: 0, calc: '—',
          notes: `Cancelled but Closer (Acq) non-person ("${owner || 'blank'}") — nothing to claw back`,
          flag: 'REVIEW',
        });
        continue;
      }
      const resolved = resolveLAMAdvanceForContract(signedDateParsed, contractId, spread);
      if (resolved.error) {
        entries.push({
          person: owner, role: 'LAM', period, source: contractId, type: 'LAM Retraction',
          amount: 0, calc: '—',
          notes: resolved.error, flag: 'REVIEW',
        });
        continue;
      }
      if (resolved.amount === 0) {
        // Either a $0 entry was explicitly recorded on the LAM Advance tab,
        // or (very rare) a tier returned $0. Either way, nothing to retract.
        entries.push({
          person: owner, role: 'LAM', period, source: contractId, type: 'LAM Retraction',
          amount: 0, calc: '—',
          notes: `No advance to retract (${resolved.source === 'historical-tab' ? 'LAM Advance tab shows $0' : 'tier table returned $0'}).`,
          flag: 'OK',
        });
        continue;
      }
      const p = lookupPerson(owner);
      if (!p || p.status === 'Inactive') {
        const reason = !p
          ? `${owner} not in People tab — treated as inactive, not chasing former employees.`
          : `${owner} marked Inactive in People tab — not chasing former employees.`;
        // Attribute to '—' (not the owner) so the renderer suppresses this line —
        // it's a $0 no-payout skip, and surfacing a former employee as a person row
        // under (Unassigned) is review noise. Matches the non-person branch above.
        // The owner's name is preserved in the note for anyone scanning entries.
        entries.push({
          person: '—', role: 'LAM', period, source: contractId, type: 'LAM Retraction',
          amount: 0, calc: '—', notes: `Retraction skipped — ${reason}`, flag: 'REVIEW',
        });
        continue;
      }
      const reason = (d['Transaction Termination Reason'] || d.Transaction_Termination_Reason || '').trim();
      const calcNote = resolved.source === 'historical-tab'
        ? `Cancelled (signed ${ymKey(signedDateParsed)}, historical window) → retract ${fmtUSD(resolved.amount)} from LAM Advance tab`
        : `Cancelled (spread ${fmtUSD(spread)}) → retract advance ${fmtUSD(resolved.amount)}`;
      entries.push({
        person: owner, role: 'LAM', period, source: contractId, type: 'LAM Retraction',
        amount: -resolved.amount,
        calc: calcNote,
        notes: reason ? `Cancellation reason: ${reason}` : '',
        flag: 'OK',
      });
    }
  }
  return entries;
}

// LAM Final — reads Monthly Comm Outreach / Sales.
//   - Tier rate applied to AGP (max 0 — negative AGP floors to $0)
//   - Advance is deducted using `resolveLAMAdvanceForContract`:
//       * Signed Sept 2025 – Mar 2026 → look up actual amount from
//         the LAM Advance tab (historical schedule, different tiers)
//       * Signed Apr 2026+ → use current LAM_ADVANCE_TIERS
//       * Signed before Sept 2025 → no advance ever paid, deduct $0
//   - Floored at $0 on negative-AGP closes (LAM v2 §5).
//
// Signed date is read directly from the "AB Contract Signed Date" column
// on Outreach Sales (added per Anshul). If the column is blank/missing
// on a row, fall back to looking up the same deal in Contracted Deals.
// If still no match, default to "no advance deducted" — safer to slightly
// overpay than to claw back money the LAM was never actually advanced.
function calcLAMFinals(salesDeals, contractedDeals, period) {
  const entries = [];

  // Fallback index: deal-ID → { signedDate, contractName } from Contracted Deals.
  // The Sales tab uses Deal Settlement as its identifier; we index Contracted Deals
  // by every plausible identifier so we can look up either the signed date or the
  // Contract Name (needed for the LAM Advance tab lookup) when the Sales row
  // doesn't carry them directly.
  const contractMetaByDealId = {};
  for (const c of (contractedDeals || [])) {
    const contractName = (c['Contract Name'] || c.Contract_Name || '').toString().trim();
    const dealSettlement = (c['Deal Settlement'] || c.Deal_Settlement || c.Deal_ID || c['Deal ID'] || '').toString().trim();
    const signed = c['Signed Contract Received Date']
                || c.Signed_Contract_Received_Date
                || c.Contract_Signed_Date;
    const signedParsed = parseDate(signed);
    const meta = { signedDate: signedParsed, contractName: contractName || null };
    // Index under both keys so either lookup hits.
    if (contractName) contractMetaByDealId[contractName] = meta;
    if (dealSettlement && !contractMetaByDealId[dealSettlement]) {
      contractMetaByDealId[dealSettlement] = meta;
    }
  }

  for (const d of salesDeals) {
    const dealId = (d['Deal Settlement'] || d.Deal_Settlement || d.Deal_ID || '?').toString().trim();
    const lam = (d.LAM || d['LAM'] || '').trim();
    const closeDate = d['BC Close Date'] || d.BC_Close_Date || d.Closed_Funded_Date;
    if (!inPeriod(closeDate, period)) continue;

    const agpRaw = parseMoney(d['Adj Gross Profit'] || d.Adj_Gross_Profit || d.Adjusted_Gross_Profit);
    const spread = parseMoney(d['Deal Spread'] || d.Deal_Spread);
    const agp = Math.max(agpRaw, 0);

    if (isNonPerson(lam)) {
      entries.push({
        person: '—', role: 'LAM', period, source: dealId, type: 'LAM Final',
        amount: 0, calc: '—',
        notes: `LAM field is non-person ("${lam || 'blank'}") — fix SF record`,
        flag: 'REVIEW',
      });
      continue;
    }

    const elig = checkSeparationOnly(lam, closeDate);
    if (!elig.eligible) {
      entries.push({
        person: lam, role: 'LAM', period, source: dealId, type: 'LAM Final',
        amount: 0, calc: '—', notes: elig.reason, flag: 'INELIGIBLE',
      });
      continue;
    }

    const tier = tierFor(agp, LAM_FINAL_TIERS);
    if (!tier) {
      entries.push({
        person: lam, role: 'LAM', period, source: dealId, type: 'LAM Final',
        amount: 0, calc: `AGP ${fmtUSD(agp)} outside tiers`,
        notes: 'Check Adj Gross Profit', flag: 'REVIEW',
      });
      continue;
    }
    const gross = tier.rate * agp;

    // Resolve signed date AND contract name (for LAM Advance tab lookup).
    //
    // Signed date: prefer the row's own column; fall back to Contracted
    // Deals index by Deal Settlement.
    //
    // Contract name (LAM Advance lookup key): read directly from the
    // "Seller Transaction: Transaction Name" column, which formats as
    // TR-SC-XXXXXX. Strip the "TR-" prefix to get the SC-XXXXXX key.
    // Per Anshul (May 2026), this column is never blank — if it is, or
    // if the value isn't in TR-SC-XXXXXX form, we flag REVIEW instead
    // of guessing. The old Contract-Name / Contracted-Deals fallback
    // is no longer used here.
    const signedDateRaw = d['AB Contract Signed Date']
                       || d.AB_Contract_Signed_Date
                       || d['Contract Signed Date']
                       || d.Contract_Signed_Date;
    const fallbackMeta = contractMetaByDealId[dealId] || null;
    const signedDate = parseDate(signedDateRaw) || (fallbackMeta && fallbackMeta.signedDate) || null;

    const txnNameRaw = (d['Seller Transaction: Transaction Name']
                     || d.Seller_Transaction_Transaction_Name
                     || '').toString().trim();
    // Strip a leading "TR-" (case-insensitive) to derive the SC-XXXXXX
    // key. If the value doesn't carry the prefix but already looks like
    // SC-XXXXXX, accept it as-is — bare values aren't strictly per spec
    // but they're recoverable; total garbage falls through to REVIEW.
    let contractName = null;
    let txnNameError = null;
    if (!txnNameRaw) {
      txnNameError = `"Seller Transaction: Transaction Name" is blank on this row — required for LAM Advance lookup. Expected format: TR-SC-XXXXXX.`;
    } else {
      const m = txnNameRaw.match(/^\s*(?:TR-)?(SC-\S+)\s*$/i);
      if (m) {
        // Preserve the SC- portion's original casing rather than uppercasing
        // — the LAM Advance tab keys are also user-entered, so a case-sensitive
        // exact match is the safest comparison. (normalizeName inside the
        // resolver handles whitespace.)
        contractName = m[1];
      } else {
        txnNameError = `"Seller Transaction: Transaction Name" = "${txnNameRaw}" doesn't match expected format TR-SC-XXXXXX — cannot derive LAM Advance lookup key.`;
      }
    }

    // If the new column failed (blank or malformed), flag REVIEW and skip
    // the advance math entirely — we won't pay a LAM Final without knowing
    // what advance to deduct.
    if (txnNameError) {
      entries.push({
        person: lam, role: 'LAM', period, source: dealId, type: 'LAM Final',
        amount: 0,
        calc: `${fmtPct(tier.rate)} × ${fmtUSD(agp)} AGP = ${fmtUSD(gross)}, but advance amount could not be resolved.`,
        notes: txnNameError, flag: 'REVIEW',
      });
      continue;
    }

    // Resolve the advance to deduct. Three paths inside the helper:
    //   pre-program → $0; historical-tab → LAM Advance lookup; tier-table → current tiers.
    let origAdvance = 0;
    let resolvedSource = 'pre-program';
    let resolveError = null;
    if (signedDate) {
      const resolved = resolveLAMAdvanceForContract(signedDate, contractName, spread);
      origAdvance = resolved.amount;
      resolvedSource = resolved.source;
      resolveError = resolved.error;
    }

    // If the historical-window lookup missed (contract not on LAM Advance tab),
    // flag REVIEW and don't pay — better than over/underpaying by the wrong amount.
    if (resolveError) {
      entries.push({
        person: lam, role: 'LAM', period, source: dealId, type: 'LAM Final',
        amount: 0,
        calc: `${fmtPct(tier.rate)} × ${fmtUSD(agp)} AGP = ${fmtUSD(gross)}, but advance amount could not be resolved.`,
        notes: resolveError, flag: 'REVIEW',
      });
      continue;
    }

    const net = gross - origAdvance;

    let calcNote;
    if (agpRaw < 0) {
      calcNote = `Negative AGP ${fmtUSD(agpRaw)} floored to $0; advance ${fmtUSD(origAdvance)} kept (deal closed). Net: ${fmtUSD(net)}`;
    } else if (resolvedSource === 'historical-tab') {
      calcNote = `${fmtPct(tier.rate)} × ${fmtUSD(agp)} AGP = ${fmtUSD(gross)}; minus advance ${fmtUSD(origAdvance)} (contract signed ${ymKey(signedDate)}, historical window — from LAM Advance tab) = ${fmtUSD(net)}`;
    } else if (resolvedSource === 'tier-table') {
      calcNote = `${fmtPct(tier.rate)} × ${fmtUSD(agp)} AGP = ${fmtUSD(gross)}; minus advance ${fmtUSD(origAdvance)} (spread ${fmtUSD(spread)}, contract signed ${ymKey(signedDate)}) = ${fmtUSD(net)}`;
    } else {
      const reason = signedDate
        ? `contract signed ${ymKey(signedDate)} — before advance program (${LAM_ADVANCE_START_DATE})`
        : 'no AB Contract Signed Date on row and no match in Contracted Deals — assuming no advance';
      calcNote = `${fmtPct(tier.rate)} × ${fmtUSD(agp)} AGP = ${fmtUSD(gross)} (no advance deducted: ${reason})`;
    }

    const finalAmount = (agpRaw < 0 && net < 0) ? 0 : net;
    const flooredNote = (agpRaw < 0 && net < 0)
      ? ' Net floored to $0 — advance NOT clawed back on negative-AGP close (LAM v2 §5).'
      : '';

    entries.push({
      person: lam, role: 'LAM', period, source: dealId, type: 'LAM Final',
      amount: finalAmount,
      calc: calcNote + flooredNote,
      notes: '',
      flag: 'OK',
    });
  }
  return entries;
}

// LTC — every closed-funded row on Monthly Comm Outreach / Sales = $50 to Pracy-Ann.
// When other LTCs are hired, add an LTC owner column to the report and update this calc
// to read it (with fallback to LTC_DEFAULT_PERSON).
function calcLTC(salesDeals, period) {
  const entries = [];
  let count = 0;
  const dealIds = [];
  for (const d of salesDeals) {
    const closeDate = d['BC Close Date'] || d.BC_Close_Date || d.Closed_Funded_Date;
    if (!inPeriod(closeDate, period)) continue;
    count += 1;
    dealIds.push(d['Deal Settlement'] || d.Deal_Settlement || d.Deal_ID || '?');
  }
  if (count === 0) return entries;

  const elig = checkSeparationOnly(LTC_DEFAULT_PERSON, period + '-15');
  if (!elig.eligible) {
    entries.push({
      person: LTC_DEFAULT_PERSON, role: 'LTC', period,
      source: `${count} closed-funded deal(s)`,
      type: 'LTC Closing Bonus',
      amount: 0, calc: '—', notes: elig.reason, flag: 'INELIGIBLE',
    });
    return entries;
  }

  const amount = count * LTC_FLAT_BONUS;
  entries.push({
    person: LTC_DEFAULT_PERSON, role: 'LTC', period,
    source: `${count} closed-funded deal(s)`,
    type: 'LTC Closing Bonus',
    amount,
    calc: `${count} × ${fmtUSD(LTC_FLAT_BONUS)} flat bonus = ${fmtUSD(amount)}`,
    notes: `All closed-funded deals attributed to ${LTC_DEFAULT_PERSON} (sole active LTC). Deal IDs: ${dealIds.join(', ')}`,
    flag: 'OK',
  });
  return entries;
}

// LPA — 0.3% × AGP per closed-funded deal qualified by the LPA.
// Reads LPA column on Monthly Comm Outreach / Sales.
// QoC ≥95% gate NOT enforced (per Anshul May 2026) — flagged REVIEW on every line
// so Lucia confirms accuracy before payout.
function calcLPA(salesDeals, period) {
  return calcAnalystShare(salesDeals, period, {
    role: 'LPA',
    columnNames: ['LPA', 'LPA_Owner', 'LPA Owner'],
    rate: LPA_RATE,
  });
}

// LLP — 0.15% × AGP per closed-funded deal qualified by the LLP. Same gate caveat.
function calcLLP(salesDeals, period) {
  return calcAnalystShare(salesDeals, period, {
    role: 'LLP',
    columnNames: ['LLP', 'LLP_Owner', 'LLP Owner'],
    rate: LLP_RATE,
  });
}

function calcAnalystShare(salesDeals, period, opts) {
  const entries = [];
  for (const d of salesDeals) {
    const dealId = d['Deal Settlement'] || d.Deal_Settlement || d.Deal_ID || '?';
    const closeDate = d['BC Close Date'] || d.BC_Close_Date || d.Closed_Funded_Date;
    if (!inPeriod(closeDate, period)) continue;

    let owner = null;
    for (const col of opts.columnNames) {
      if (d[col] && String(d[col]).trim()) { owner = String(d[col]).trim(); break; }
    }
    if (!owner || isNonPerson(owner)) continue;  // no analyst attributed → skip silently

    const agpRaw = parseMoney(d['Adj Gross Profit'] || d.Adj_Gross_Profit || d.Adjusted_Gross_Profit);
    const agp = Math.max(agpRaw, 0);
    const amount = opts.rate * agp;

    const elig = checkSeparationOnly(owner, closeDate);
    if (!elig.eligible) {
      entries.push({
        person: owner, role: opts.role, period, source: dealId, type: opts.role,
        amount: 0, calc: '—', notes: elig.reason, flag: 'INELIGIBLE',
      });
      continue;
    }

    entries.push({
      person: owner, role: opts.role, period, source: dealId, type: opts.role,
      amount,
      calc: `${fmtPct(opts.rate)} × AGP ${fmtUSD(agp)} = ${fmtUSD(amount)}${agpRaw < 0 ? ` (negative AGP ${fmtUSD(agpRaw)} floored to $0)` : ''}`,
      notes: 'QoC ≥95% gate NOT enforced — verify analyst hit the KPI before payout (canonical: LAP_LLP doc Tier 2).',
      flag: 'REVIEW',
    });
  }
  return entries;
}

// LIA/LIM Component A — Per-Lead Approval Bonus from Approved Leads.
// v4.1 dedup by Lead ID; v4 §4A.1 blank-Q rule; forfeit-on-separation.
function calcLIMLIA(approvedLeads, period) {
  const entries = [];

  const LEAD_ID_KEYS = ['Lead_ID', 'Lead ID', 'Seller Name: Lead ID (18)', 'Seller_Name_Lead_ID_18'];
  const seenLeadIds = new Set();
  const deduped = [];
  for (const lead of approvedLeads) {
    let leadId = null;
    for (const k of LEAD_ID_KEYS) {
      if (lead[k] && String(lead[k]).trim()) { leadId = String(lead[k]).trim(); break; }
    }
    if (leadId) {
      if (seenLeadIds.has(leadId)) continue;
      seenLeadIds.add(leadId);
    }
    deduped.push(lead);
  }

  for (const lead of deduped) {
    const leadId = lead.Lead_ID || lead['Lead ID'] || lead['Seller Name: Lead ID (18)']
      || lead['Property Contract: Contract Name'] || '?';
    const lia = (lead.LIA_Texter || lead['LIA Texter'] || lead.Sourcer || lead['Sourcer']
      || lead['Seller Name: Sourcer (Text)'] || lead['Seller Name: Sourcer'] || '').trim() || null;
    const lim = (lead.LIM_Phone_Qualifier || lead['LIM Phone Qualifier'] || lead.Qualifier
      || lead['Qualifier'] || lead['Seller Name: Qualifier'] || '').trim() || null;
    const category = (lead.Lead_Category || lead['Lead Category'] || lead.AILeadCategory
      || lead['AILeadCategory'] || '').trim();
    const status = (lead.Approval_Status || lead['Approval Status'] || '').trim().toLowerCase();
    const ad = lead.Approval_Date || lead['Approval Date'];
    if (ad && !inPeriod(ad, period)) continue;
    if (status && status !== 'approved') continue;

    const bonus = LEAD_BONUSES[category];
    if (!bonus) {
      entries.push({
        person: lia || lim || '?', role: 'LIA/LIM', period, source: leadId, type: 'Per-Lead',
        amount: 0, calc: '—',
        notes: `Unknown AILeadCategory: "${category}"`, flag: 'REVIEW',
      });
      continue;
    }

    const isPhoneCat = category === 'Phone Fully Q' || category === 'Phone Semi Q';

    // v4 §4A.1: Phone Fully/Semi Q with blank Qualifier → $1 LIA only
    if (isPhoneCat && !lim && lia) {
      const elig = checkSeparationOnly(lia, ad);
      entries.push({
        person: lia, role: 'LIA', period, source: leadId, type: `Per-Lead (${category} → blank-Q)`,
        amount: elig.eligible ? 1 : 0,
        calc: elig.eligible ? `${category} but Qualifier blank → $1 LIA only (v4 §4A.1)` : '—',
        notes: elig.flag === 'OK' ? '' : elig.reason,
        flag: elig.flag,
      });
      continue;
    }

    const samePerson = lia && lim && lia === lim;
    const scenario = samePerson ? bonus.samePerson : bonus.split;

    if (lia && scenario.lia > 0 && !samePerson) {
      const elig = checkSeparationOnly(lia, ad);
      entries.push({
        person: lia, role: 'LIA', period, source: leadId, type: `Per-Lead (${category})`,
        amount: elig.eligible ? scenario.lia : 0,
        calc: elig.eligible ? `${category} (split): LIA gets ${fmtUSD(scenario.lia)}` : '—',
        notes: elig.flag === 'OK' ? '' : elig.reason, flag: elig.flag,
      });
    }
    if (lim && scenario.lim > 0) {
      const elig = checkSeparationOnly(lim, ad);
      entries.push({
        person: lim, role: 'LIM', period, source: leadId, type: `Per-Lead (${category})`,
        amount: elig.eligible ? scenario.lim : 0,
        calc: elig.eligible
          ? (samePerson
              ? `${category} (same LIM did both): ${fmtUSD(scenario.lim)}`
              : `${category} (split): LIM gets ${fmtUSD(scenario.lim)}`)
          : '—',
        notes: elig.flag === 'OK' ? '' : elig.reason, flag: elig.flag,
      });
    }
    if (samePerson && scenario.lia > 0 && scenario.lim === 0) {
      const elig = checkSeparationOnly(lia, ad);
      entries.push({
        person: lia, role: 'LIA', period, source: leadId, type: `Per-Lead (${category})`,
        amount: elig.eligible ? scenario.lia : 0,
        calc: elig.eligible ? `${category} (same person): ${fmtUSD(scenario.lia)}` : '—',
        notes: elig.flag === 'OK' ? '' : elig.reason, flag: elig.flag,
      });
    }
  }
  return entries;
}

// LIA/LIM Component B — closing commission on Monthly Comm Outreach / Sales.
// Pool = 0.3% × max(AGP, 0). Split LEAD_BONUSES proportions.
// Blank slot → that share forfeited (not redistributed).
function calcLIALIMClosing(salesDeals, period) {
  const entries = [];
  for (const d of salesDeals) {
    const dealId = d.Deal_Settlement || d['Deal Settlement'] || d.Deal_ID || '?';
    const cat = (d.AILeadCategory || d['AILeadCategory'] || '').trim();
    const agp = parseMoney(d.Adj_Gross_Profit || d['Adj Gross Profit']);
    const sourcer = (d['Sourcer (Acq)'] || d.Sourcer_Acq || '').trim() || null;
    const qualifier = (d['Qualifier (Acq)'] || d.Qualifier_Acq || '').trim() || null;
    const closeDate = d.BC_Close_Date || d['BC Close Date'] || d.Closed_Funded_Date;
    if (closeDate && !inPeriod(closeDate, period)) continue;

    const pool = LIA_LIM_CLOSING_RATE * Math.max(agp, 0);

    if (pool === 0) {
      entries.push({
        person: '—', role: 'LIA/LIM Closing', period, source: dealId, type: 'LIA/LIM Closing',
        amount: 0, calc: `AGP ${fmtUSD(agp)} → pool $0`,
        notes: agp < 0 ? 'Negative AGP — floored to $0 (v4 §4B.2)' : 'AGP is zero', flag: 'OK',
      });
      continue;
    }

    if (!sourcer && !qualifier) {
      entries.push({
        person: '—', role: 'LIA/LIM Closing', period, source: dealId, type: 'LIA/LIM Closing',
        amount: 0, calc: `Pool ${fmtUSD(pool)} unallocated`,
        notes: 'Sourcer (Acq) AND Qualifier (Acq) blank — full pool forfeited.',
        flag: 'REVIEW',
      });
      continue;
    }

    const bonus = LEAD_BONUSES[cat];
    if (!bonus) {
      // Legacy fallback (Anshul, May 2026): pre-categorization deals
      // and any unrecognized value default to paying the Sourcer the
      // full pool. The rules currently define only Phone Fully Q,
      // Phone Semi Q, and Text Fully Q (plus LaLa/Fantasy as $0
      // intentional excludes — those go through the LEAD_BONUSES path
      // below, not here). Treating blanks/unknowns as Text Fully Q
      // -equivalent (sourcer gets everything) is the closest match to
      // "who brought in this lead" for legacy uncategorized deals.
      if (!sourcer) {
        // No sourcer to fall back to — nothing payable. Flag so it
        // gets attention if the category gets backfilled later.
        entries.push({
          person: '—', role: 'LIA/LIM Closing', period, source: dealId,
          type: 'LIA/LIM Closing', amount: 0, calc: '—',
          notes: `AILeadCategory "${cat || '(blank)'}" not in defined rules AND no Sourcer (Acq) — legacy fallback cannot apply. Backfill category or sourcer.`,
          flag: 'REVIEW',
        });
        continue;
      }
      const elig = checkSeparationOnly(sourcer, closeDate);
      if (!elig.eligible) {
        entries.push({
          person: sourcer, role: 'LIA', period, source: dealId, type: 'LIA Closing',
          amount: 0, calc: '—', notes: elig.reason, flag: elig.flag,
        });
        continue;
      }
      const catLabel = cat || '(blank)';
      entries.push({
        person: sourcer, role: 'LIA', period, source: dealId, type: 'LIA Closing',
        amount: pool,
        calc: `Legacy fallback: AILeadCategory "${catLabel}" not in {Phone Fully Q, Phone Semi Q, Text Fully Q} → full pool ${fmtUSD(pool)} to Sourcer`,
        notes: `0.3% × AGP ${fmtUSD(agp)}${elig.flag === 'REVIEW' ? '; ' + elig.reason : ''}`,
        flag: elig.flag,
      });
      continue;
    }

    const samePerson = sourcer && qualifier && sourcer === qualifier;
    const proportions = samePerson ? bonus.samePerson : bonus.split;
    const totalUnits = proportions.lia + proportions.lim;
    if (totalUnits === 0) continue;

    const liaShare = (proportions.lia / totalUnits) * pool;
    const limShare = (proportions.lim / totalUnits) * pool;

    if (sourcer && liaShare > 0) {
      const elig = checkSeparationOnly(sourcer, closeDate);
      entries.push({
        person: sourcer, role: 'LIA', period, source: dealId, type: 'LIA Closing',
        amount: elig.eligible ? liaShare : 0,
        calc: elig.eligible
          ? `${cat} ${samePerson ? 'same-person' : 'split'}: LIA ${(proportions.lia / totalUnits * 100).toFixed(2)}% × pool ${fmtUSD(pool)} = ${fmtUSD(liaShare)}`
          : '—',
        notes: `0.3% × AGP ${fmtUSD(agp)}${elig.flag === 'REVIEW' ? '; ' + elig.reason : ''}`,
        flag: elig.flag,
      });
    } else if (!sourcer && liaShare > 0) {
      entries.push({
        person: '—', role: 'LIA Closing', period, source: dealId, type: 'LIA Closing',
        amount: 0, calc: `LIA share ${fmtUSD(liaShare)} forfeited`,
        notes: 'Sourcer (Acq) blank — LIA share not paid (v4 §4B.4)', flag: 'REVIEW',
      });
    }

    if (qualifier && limShare > 0) {
      const elig = checkSeparationOnly(qualifier, closeDate);
      entries.push({
        person: qualifier, role: 'LIM', period, source: dealId, type: 'LIM Closing',
        amount: elig.eligible ? limShare : 0,
        calc: elig.eligible
          ? `${cat} ${samePerson ? 'same-person (full pool)' : 'split'}: LIM ${(proportions.lim / totalUnits * 100).toFixed(2)}% × pool ${fmtUSD(pool)} = ${fmtUSD(limShare)}`
          : '—',
        notes: `0.3% × AGP ${fmtUSD(agp)}${elig.flag === 'REVIEW' ? '; ' + elig.reason : ''}`,
        flag: elig.flag,
      });
    } else if (!qualifier && limShare > 0) {
      entries.push({
        person: '—', role: 'LIM Closing', period, source: dealId, type: 'LIM Closing',
        amount: 0, calc: `LIM share ${fmtUSD(limShare)} forfeited`,
        notes: 'Qualifier (Acq) blank — LIM share not paid (v4 §4B.4)', flag: 'REVIEW',
      });
    }
  }
  return entries;
}

// LPM v2.2 — portfolio aggregation from the Sales tab.
function calcLPM(salesData, period) {
  const entries = [];
  const byOwner = {};

  for (const d of salesData) {
    const listingId = d['Main Listing: Listing #'] || d.Listing_Number || d['Listing #'] || '?';
    const podRaw = (d.POD || d.LPM || '').trim();
    const owner = mapLPMPod(podRaw);
    const seller = d['Seller Name'] || '?';

    if (!owner) {
      entries.push({
        person: '—', role: 'LPM', period, source: listingId, type: 'LPM (excluded)',
        amount: 0, calc: '—',
        notes: `POD column blank — cannot attribute commission. Seller: ${seller}`,
        flag: 'REVIEW',
      });
      continue;
    }

    const listPrice = parseMoney(d['Curr List Price'] || d.Curr_List_Price || d.Current_List_Price);
    const contractPrice = parseMoney(d['BC Under Contract Price'] || d.BC_Under_Contract_Price);
    const agpRaw = parseMoney(d['Adj Gross Profit'] || d.Adj_Gross_Profit || d.Adjusted_Gross_Profit);
    const domStr = d['Main Listing: Days on Market (DOM)'] || d['DOM'] || d.Days_on_Market;
    // DOM of 0 (or blank → 0) is valid: a freshly listed / same-day-under-contract
    // property legitimately has 0 days on market and must NOT be dropped from the
    // portfolio. Only the two prices are required (PAR can't be computed without them).
    let dom = parseInt(String(domStr == null ? '' : domStr).trim(), 10);
    if (isNaN(dom)) dom = 0;

    if (!listPrice || !contractPrice) {
      entries.push({
        person: owner, role: 'LPM', period, source: listingId, type: 'LPM (excluded)',
        amount: 0, calc: '—',
        notes: `Missing Curr List Price or BC Under Contract Price. Seller: ${seller}`,
        flag: 'REVIEW',
      });
      continue;
    }

    if (!byOwner[owner]) byOwner[owner] = { listings: [] };
    byOwner[owner].listings.push({ listingId, dom, listPrice, contractPrice, agpRaw, seller });
  }

  for (const [owner, info] of Object.entries(byOwner)) {
    const elig = checkSeparationOnly(owner, period + '-15');
    if (!elig.eligible) {
      entries.push({
        person: owner, role: 'LPM', period,
        source: `Portfolio (${info.listings.length} listings)`, type: 'LPM Portfolio',
        amount: 0, calc: '—', notes: elig.reason, flag: 'INELIGIBLE',
      });
      continue;
    }

    const n = info.listings.length;
    const sumDOM = info.listings.reduce((s, x) => s + x.dom, 0);
    const sumList = info.listings.reduce((s, x) => s + x.listPrice, 0);
    const sumContract = info.listings.reduce((s, x) => s + x.contractPrice, 0);
    const sumAGP = info.listings.reduce((s, x) => s + x.agpRaw, 0);
    const avgDOM = sumDOM / n;
    const par = sumList > 0 ? sumContract / sumList : 0;

    const matrixRow = LPM_MATRIX.find(r => avgDOM >= r.domMin && avgDOM <= r.domMax);
    if (!matrixRow) {
      entries.push({
        person: owner, role: 'LPM', period,
        source: `Portfolio (${n} listings)`, type: 'LPM Portfolio',
        amount: 0, calc: `Avg DOM ${avgDOM.toFixed(1)} outside matrix`,
        notes: 'Check DOM values', flag: 'REVIEW',
      });
      continue;
    }

    let rate, parTier;
    if (par >= 0.85)      { rate = matrixRow.parHigh; parTier = '≥85%';      }
    else if (par >= 0.75) { rate = matrixRow.parMid;  parTier = '75%–84.9%'; }
    else                  { rate = matrixRow.parLow;  parTier = '<75%';      }

    const grossCommission = rate * sumAGP;
    const amount = Math.max(grossCommission, 0);
    const domTier = matrixRow.domMax === Infinity ? '>60d' : `${matrixRow.domMin}–${matrixRow.domMax}d`;
    const listingDetail = info.listings
      .map(x => `${x.listingId} (DOM ${x.dom}, AGP ${fmtUSD(x.agpRaw)})`)
      .join('; ');
    const floorNote = sumAGP < 0
      ? ` Portfolio Σ(AGP) ${fmtUSD(sumAGP)} negative → commission floored to $0 (LPM v2.2 §5).`
      : '';

    entries.push({
      person: owner, role: 'LPM', period,
      source: `Portfolio (${n} listings)`, type: 'LPM Portfolio',
      amount,
      calc: `Avg DOM ${avgDOM.toFixed(1)} (${domTier}) × Portfolio PAR ${(par * 100).toFixed(2)}% (${parTier}) → ${fmtPct(rate)} × Σ(AGP) ${fmtUSD(sumAGP)} = ${fmtUSD(amount)}.${floorNote}`,
      notes: listingDetail,
      flag: 'OK',
    });
  }
  return entries;
}

// Outreach Team Manager override — 20% × direct reports' Component A only (not Component B).
function calcLIMManagerOverride(allEntries) {
  const managerByName = {};
  for (const [name, p] of Object.entries(PEOPLE)) {
    if (p.isOutreachManager === true) {
      managerByName[name] = { totalReports: 0, reports: [] };
    }
  }
  for (const [reportName, reportInfo] of Object.entries(PEOPLE)) {
    if (!reportInfo.manager || !managerByName[reportInfo.manager]) continue;
    const reportTotal = allEntries
      .filter(e => normalizeName(e.person) === reportName
        && (e.role === 'LIA' || e.role === 'LIM')
        && e.flag === 'OK'
        && !/Closing/i.test(e.type))
      .reduce((sum, e) => sum + e.amount, 0);
    managerByName[reportInfo.manager].totalReports += reportTotal;
    managerByName[reportInfo.manager].reports.push({ name: reportName, total: reportTotal });
  }
  const entries = [];
  for (const [mgrName, info] of Object.entries(managerByName)) {
    if (info.totalReports === 0) {
      entries.push({
        person: mgrName, role: 'Outreach Mgr', period: null, source: 'override',
        type: 'Outreach Team Mgr Override',
        amount: 0, calc: '—',
        notes: 'No direct reports earned Component A this period.',
        flag: 'REVIEW',
      });
      continue;
    }
    const amount = LIM_MANAGER_OVERRIDE_RATE * info.totalReports;
    const detail = info.reports.map(r => `${r.name}: ${fmtUSD(r.total)}`).join('; ');
    entries.push({
      person: mgrName, role: 'Outreach Mgr', period: null, source: 'override',
      type: 'Outreach Team Mgr Override',
      amount,
      calc: `20% × ${fmtUSD(info.totalReports)} (sum of direct reports' Component A) = ${fmtUSD(amount)}`,
      notes: detail, flag: 'OK',
    });
  }
  return entries;
}

// NOI execs — Art Oaing (DOA PH) and Leslie Bernolo (DOT). Pending revision.
function calcNOITiered(period) {
  const entries = [];
  const fin = FINANCE[period];
  if (!fin) {
    entries.push({
      person: '—', role: 'DOA PH / DOT', period, source: 'monthly',
      type: 'NOI Tiered',
      amount: 0, calc: '—',
      notes: `No Finance entry for period ${period}. Add a row to the Finance tab (Period: ${period}, Revenue: …, NOI: …) and re-run.`,
      flag: 'REVIEW',
    });
    return entries;
  }
  for (const [person, tiers] of Object.entries(NOI_TIERS)) {
    const tier = tiers.find(t => fin.revenue <= t.revMax);
    if (!tier) continue;
    const amount = tier.rate * fin.noi;
    entries.push({
      person, role: 'NOI-Tiered Exec', period, source: 'monthly',
      type: 'NOI-Tiered Exec',
      amount,
      calc: `Revenue ${fmtUSD(fin.revenue)} → rate ${fmtPct(tier.rate)} × NOI ${fmtUSD(fin.noi)} = ${fmtUSD(amount)}`,
      notes: 'Pending revision per Anshul', flag: 'REVIEW',
    });
  }
  return entries;
}

/* ------------------------------------------------------------
   6. ORCHESTRATOR
   ------------------------------------------------------------ */
function runCommissions(period, data) {
  const outreachSales = data[TAB_OUTREACH_SALES] || [];
  const approvedLeads = data[TAB_APPROVED_LEADS] || [];
  const sales = data[TAB_SALES] || [];
  const contracted = data[TAB_CONTRACTED] || [];

  const all = [];
  all.push(...calcLAMAdvances(contracted, period));
  all.push(...calcLAMFinals(outreachSales, contracted, period));
  all.push(...calcLTC(outreachSales, period));
  all.push(...calcLPA(outreachSales, period));
  all.push(...calcLLP(outreachSales, period));
  all.push(...calcLIMLIA(approvedLeads, period));
  all.push(...calcLIALIMClosing(outreachSales, period));
  all.push(...calcLPM(sales, period));
  // Canonicalize attributed names to their roster Person Name BEFORE the
  // override and history/render run, so SF spelling variants aggregate and pay
  // under one identity. '—' left untouched; unresolved names keep their spelling.
  for (const e of all) {
    if (e.person && e.person !== '—') e.person = canonicalName(e.person);
  }
  all.push(...calcLIMManagerOverride(all));
  all.push(...calcNOITiered(period));
  return all;
}

/* ------------------------------------------------------------
   7. FETCH
   ------------------------------------------------------------ */
async function fetchSheetData() {
  if (APPS_SCRIPT_URL.startsWith('PASTE_')) {
    throw new Error('APPS_SCRIPT_URL not configured. Edit script.js to set your deployment URL.');
  }
  const url = `${APPS_SCRIPT_URL}?token=${encodeURIComponent(APPS_SCRIPT_TOKEN)}`;
  const res = await fetch(url, { method: 'GET' });
  if (!res.ok) throw new Error(`Apps Script returned HTTP ${res.status}`);
  const body = await res.json();
  if (body.error) throw new Error(`Apps Script: ${body.error}`);
  // Per-tab errors
  for (const tab of [TAB_OUTREACH_SALES, TAB_APPROVED_LEADS, TAB_SALES, TAB_CONTRACTED, TAB_PEOPLE, TAB_FINANCE, TAB_LAM_ADVANCE]) {
    const v = body[tab];
    if (v && v.error) throw new Error(`Tab "${tab}": ${v.error}`);
    if (!v) throw new Error(`Tab "${tab}" missing from response`);
  }
  // Hydrate PEOPLE roster from the People tab BEFORE any calc runs.
  PEOPLE = buildPeopleFromSheet(body[TAB_PEOPLE]);
  if (Object.keys(PEOPLE).length === 0) {
    throw new Error(
      `People tab is empty. Add at least one row (Person Name + Status=Active) ` +
      `or the calculator will mark every commission line INELIGIBLE.`
    );
  }
  // Hydrate FINANCE from the Finance tab. Empty is OK — calcNOITiered flags the gap.
  FINANCE = buildFinanceFromSheet(body[TAB_FINANCE]);
  // Hydrate LAM_ADVANCE_LOOKUP from the LAM Advance tab. Empty is OK if no
  // historical-window (Sept 2025 – Mar 2026) contracts are ever processed —
  // but if one shows up in retraction/finals with no lookup match, those
  // lines will flag REVIEW.
  LAM_ADVANCE_LOOKUP = buildLAMAdvanceLookup(body[TAB_LAM_ADVANCE]);
  // Hydrate lastHistoryData from the optional Commission_History tab.
  // Absence is normal before the very first writeCommissionHistory
  // call ever runs — first Calculate after Step 3 deploys will create
  // and populate it. After that, this hydration keeps the trends panel
  // current with what's already on disk.
  const historyRaw = body[TAB_HISTORY];
  if (historyRaw && !historyRaw.error) {
    lastHistoryData = Array.isArray(historyRaw) ? historyRaw : [];
  } else {
    lastHistoryData = [];
    if (historyRaw && historyRaw.error) {
      console.warn(`Commission_History tab present but errored: ${historyRaw.error}`);
    }
  }
  return body;
}

/* ------------------------------------------------------------
   8. RENDERING
   ------------------------------------------------------------ */

// Fixed display order for departments. Anything not in this list
// (custom departments HR may add later, plus the "(Unassigned)"
// catch-all bucket) sorts after these, alphabetically, with
// "(Unassigned)" last so it reads as an exception.
const DEPT_ORDER = ['Acquisitions', 'Transactions', 'Marketing', 'Sales', 'Operations'];

// Produce a short, human-readable explanation of why someone earned $0
// this period. Used for zero-commission rows on the dashboard. Two
// specific roles get specific copy (LSM and Marketing) because we know
// those plans aren't producing any payout right now regardless of
// activity; everyone else gets a generic "no activity" line that
// names the role for context.
function noCommissionReason(person, period) {
  const p = lookupPerson(person);
  const role = (p && p.role) ? p.role : null;
  const periodName = fmtPeriodName(period);
  if (role) {
    const roleUpper = role.toUpperCase();
    if (roleUpper === 'LSM') {
      return `No commission this period — LSM commission plan is currently deferred.`;
    }
    if (roleUpper.includes('MARKETING')) {
      return `No commission this period — ${role} commission plan is pending definition.`;
    }
    return `No commissionable activity for ${periodName}. Role: ${role}.`;
  }
  return `No commissionable activity for ${periodName}.`;
}

function renderResults(entries, period) {
  const summaryEl = document.getElementById('summary');

  // -- 1. Aggregate by person from entries. Skip:
  //   - INELIGIBLE entries (not surfaced anywhere now that the Flags
  //     panel has been removed)
  //   - Unattributed entries (person === '—'); SF data hygiene issues,
  //     not real payouts.
  const byPerson = {};
  for (const e of entries) {
    if (e.flag === 'INELIGIBLE') continue;
    if (!e.person || e.person === '—') continue;
    if (!byPerson[e.person]) byPerson[e.person] = { total: 0, count: 0, hasReview: false, byType: {} };
    byPerson[e.person].total += e.amount;
    byPerson[e.person].count += 1;
    if (e.flag === 'REVIEW') byPerson[e.person].hasReview = true;
    const t = e.type || '(unknown)';
    byPerson[e.person].byType[t] = (byPerson[e.person].byType[t] || 0) + e.amount;
  }

  // -- 2. Add zero-commission rows for Active roster members who didn't
  // appear in any entry. We match case-insensitively against the roster
  // so a typo in SF (e.g. "anna rivers" vs "Anna Rivers") doesn't cause
  // someone to be double-listed.
  const withEntries = new Set();
  for (const personName of Object.keys(byPerson)) {
    const norm = normalizeName(personName);
    if (PEOPLE[norm]) {
      withEntries.add(norm);
      continue;
    }
    const lower = norm.toLowerCase();
    for (const key of Object.keys(PEOPLE)) {
      if (key.toLowerCase() === lower) {
        withEntries.add(key);
        break;
      }
    }
  }
  for (const [name, p] of Object.entries(PEOPLE)) {
    if (p.status !== 'Active') continue;
    if (withEntries.has(name)) continue;
    byPerson[name] = { total: 0, count: 0, hasReview: false, byType: {}, isZero: true };
  }

  // -- 3. Group by department.
  const byDepartment = {};
  for (const [person, info] of Object.entries(byPerson)) {
    const p = lookupPerson(person);
    const dept = (p && p.department) ? p.department : '(Unassigned)';
    if (!byDepartment[dept]) byDepartment[dept] = [];
    byDepartment[dept].push([person, info]);
  }
  // Sort departments using the fixed DEPT_ORDER; unknowns alphabetically;
  // (Unassigned) last. Comparison is case-insensitive so HR can enter
  // either "Operations" or "operations" without affecting order.
  const orderLower = DEPT_ORDER.map(d => d.toLowerCase());
  const deptNames = Object.keys(byDepartment).sort((a, b) => {
    const ai = orderLower.indexOf(a.toLowerCase());
    const bi = orderLower.indexOf(b.toLowerCase());
    if (ai !== -1 && bi !== -1) return ai - bi;
    if (ai !== -1) return -1;
    if (bi !== -1) return 1;
    if (a === '(Unassigned)') return 1;
    if (b === '(Unassigned)') return -1;
    return a.localeCompare(b);
  });
  for (const dept of deptNames) {
    byDepartment[dept].sort((a, b) => a[0].localeCompare(b[0]));
  }

  const grandTotal = Object.values(byPerson).reduce((s, v) => s + v.total, 0);

  // -- 4. Helper: render a single person row. Two shapes:
  //   - Zero-commission rows: a non-expandable div with the reason
  //   - Regular rows: an expandable <details> with the full drill-down
  const renderOne = (person, v) => {
    if (v.isZero) {
      const reason = noCommissionReason(person, period);
      return `<div class="person-row-zero">`
        + `<span class="name">${escapeHTML(person)}</span>`
        + `<span class="breakdown-col">${escapeHTML(reason)}</span>`
        + `<span class="num">${fmtUSD(0)}</span>`
        + `</div>`;
    }
    const personEntries = entries.filter(e => e.person === person && e.flag !== 'INELIGIBLE');
    if (personEntries.length === 0) return '';

    // Per-type breakdown for the summary row. Sort by absolute amount
    // descending so the biggest line items come first.
    const breakdownParts = Object.entries(v.byType)
      .sort((a, b) => Math.abs(b[1]) - Math.abs(a[1]))
      .map(([type, amount]) => `${type} ${fmtUSD(amount)}`);
    const breakdownStr = breakdownParts.join(', ');

    let h = `<details class="person-row${v.hasReview ? ' has-review' : ''}">`;
    h += `<summary>`;
    h += `<span class="name">${escapeHTML(person)}</span>`;
    h += `<span class="breakdown-col">${escapeHTML(breakdownStr)}</span>`;
    h += `<span class="num">${fmtUSD(v.total)}</span>`;
    h += `</summary>`;

    // Drill-down: per-deal rows for Contracted Deals, Sales, and Outreach
    // Sales. For Approved Leads (Per-Lead bonuses) aggregate by type — a
    // LIM with 50+ approved leads listed individually is noise.
    const perLeadByType = {};
    const individualEntries = [];
    for (const e of personEntries) {
      if (/^Per-Lead/.test(e.type)) {
        if (!perLeadByType[e.type]) perLeadByType[e.type] = [];
        perLeadByType[e.type].push(e);
      } else {
        individualEntries.push(e);
      }
    }
    individualEntries.sort((a, b) =>
      a.type.localeCompare(b.type) || String(a.source || '').localeCompare(String(b.source || ''))
    );
    const perLeadTypes = Object.keys(perLeadByType).sort();

    h += '<div class="person-row-body">';
    h += '<table class="detail-table"><thead><tr>';
    h += '<th>Type</th><th>Source</th><th class="num">Amount</th><th>Calculation</th>';
    h += '</tr></thead><tbody>';

    for (const e of individualEntries) {
      const flagClass = e.flag === 'REVIEW' ? 'flag-review' : 'flag-ok';
      h += `<tr class="${flagClass}">`;
      h += `<td>${escapeHTML(e.type)}</td>`;
      h += `<td class="src-cell">${escapeHTML(e.source)}</td>`;
      h += `<td class="num">${fmtUSD(e.amount)}</td>`;
      h += `<td class="summary-cell">${escapeHTML(e.calc)}</td>`;
      h += `</tr>`;
    }

    for (const type of perLeadTypes) {
      const ents = perLeadByType[type];
      const total = ents.reduce((s, e) => s + e.amount, 0);
      const count = ents.length;
      const hasReview = ents.some(e => e.flag === 'REVIEW');
      const flagClass = hasReview ? 'flag-review' : 'flag-ok';
      const amounts = [...new Set(ents.map(e => e.amount))];
      let calc;
      if (amounts.length === 1) {
        calc = `${count} approved lead${count === 1 ? '' : 's'} × ${fmtUSD(amounts[0])} = ${fmtUSD(total)}`;
      } else {
        calc = `${count} approved leads, varying amounts = ${fmtUSD(total)}`;
      }
      h += `<tr class="${flagClass}">`;
      h += `<td>${escapeHTML(type)}</td>`;
      h += `<td class="src-cell">${count} lead${count === 1 ? '' : 's'}</td>`;
      h += `<td class="num">${fmtUSD(total)}</td>`;
      h += `<td class="summary-cell">${escapeHTML(calc)}</td>`;
      h += `</tr>`;
    }

    h += '</tbody></table>';

    h += '</div></details>';
    return h;
  };

  // -- 5. Render header + top grand-total banner.
  let html = `<div class="period-banner">Period: <strong>${escapeHTML(fmtPeriodName(period))}</strong> &nbsp;·&nbsp; Grand Total: <strong>${fmtUSD(grandTotal)}</strong></div>`;
  html += '<div class="people-list">';
  html += '<div class="people-header"><span>Person</span><span>Breakdown</span><span class="num">Total</span></div>';

  // -- 6. Render each department.
  for (const dept of deptNames) {
    const peopleInDept = byDepartment[dept];
    const deptTotal = peopleInDept.reduce((s, [, v]) => s + v.total, 0);
    const headCount = peopleInDept.length;

    html += `<div class="dept-header">`;
    html += `<span class="dept-name">${escapeHTML(dept)}</span>`;
    html += `<span class="dept-count">${headCount} ${headCount === 1 ? 'person' : 'people'}</span>`;
    html += `</div>`;

    // Sub-group people inside this department by their Role. Applied
    // to every department, not just Operations, so the visual hierarchy
    // (Department → Role → Person) is consistent across the dashboard.
    // People with no Role on the People tab bucket under "(No Role)",
    // which sorts last.
    const byRole = {};
    for (const [person, info] of peopleInDept) {
      const p = lookupPerson(person);
      const role = (p && p.role) ? p.role : '(No Role)';
      if (!byRole[role]) byRole[role] = [];
      byRole[role].push([person, info]);
    }
    const roleNames = Object.keys(byRole).sort((a, b) => {
      if (a === '(No Role)') return 1;
      if (b === '(No Role)') return -1;
      return a.localeCompare(b);
    });
    for (const role of roleNames) {
      byRole[role].sort((a, b) => a[0].localeCompare(b[0]));
    }
    for (const role of roleNames) {
      const peopleInRole = byRole[role];
      const roleCount = peopleInRole.length;
      html += `<div class="role-header">`;
      html += `<span class="role-name">${escapeHTML(role)}</span>`;
      html += `<span class="role-count">${roleCount} ${roleCount === 1 ? 'person' : 'people'}</span>`;
      html += `</div>`;
      for (const [person, v] of peopleInRole) {
        html += renderOne(person, v);
      }
    }

    // Department subtotal row.
    html += `<div class="dept-subtotal">`;
    html += `<span class="dept-subtotal-label">Subtotal — ${escapeHTML(dept)}</span>`;
    html += `<span></span>`;
    html += `<span class="num">${fmtUSD(deptTotal)}</span>`;
    html += `</div>`;
  }

  // -- 7. Bottom grand-total row.
  html += `<div class="people-total"><span><strong>GRAND TOTAL</strong></span><span></span><span class="num"><strong>${fmtUSD(grandTotal)}</strong></span></div>`;
  html += '</div>';  // close .people-list

  summaryEl.innerHTML = html;
  document.getElementById('results').classList.remove('hidden');
}

function escapeHTML(s) {
  if (s === null || s === undefined) return '';
  return String(s).replace(/[&<>"']/g, c => ({
    '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'
  }[c]));
}

/* ------------------------------------------------------------
   9. EXPORT — Commission CSV download for payment platforms (Deel / OnPay)
   ----------------------------------------------------------------
   Outputs a CSV file (NOT clipboard TSV) containing Contract Name +
   Amount for every person whose Payment Mode in the People tab matches
   the active filter toggle. The CSV is what Lucia uploads to the
   selected payment platform; column order and naming are minimal on
   purpose so the file is acceptable to both Deel and OnPay regardless
   of which exact template they prompt for.

   Filtering rules:
     - Payment Mode must equal the selected filter exactly (case-insensitive)
     - $0 amounts excluded (no point uploading them)
     - Blank Contract Name excluded (row would have no identifier)
     - INELIGIBLE entries excluded (matches dashboard behavior)
     - Unattributed entries (person='—') excluded
     - People not in the roster excluded

   Every excluded person is counted into a `skipped` breakdown so the
   status line below the download button can surface what was dropped
   and why. The CSV download is silent if zero rows match — the status
   line is the only signal Lucia gets, so it has to be informative.

   Payroll download (separate from this Commission download) is intentionally
   not yet implemented — see the disabled #download-payroll-btn in index.html.
   When that's defined, add a parallel function here that uses the same
   filter logic but a different column set / aggregation.
   ------------------------------------------------------------ */

// Valid Payment Mode values, in display order. These drive the toggle
// buttons in index.html and the filter dropdown. Adding a new value here
// (e.g. "WiseBusiness") AND adding the matching button to index.html is
// all that's needed to support a new payment platform.
const PAYMENT_MODES = ['Deel', 'OnPay'];

// Tracks the currently selected Payment Mode filter. Updated by the
// toggle button click handler in DOMContentLoaded; read by the download
// functions to decide which roster rows to include.
let activePaymentMode = PAYMENT_MODES[0];  // default: 'Deel'

// Escape a CSV cell. Wraps in double quotes only when needed (comma,
// quote, or newline in the value). Doubles internal quotes per RFC 4180.
function csvEscape(value) {
  const s = value == null ? '' : String(value);
  if (/[",\r\n]/.test(s)) return '"' + s.replace(/"/g, '""') + '"';
  return s;
}

// Trigger a file download in the browser. Uses a Blob + temporary <a>
// element with the download attribute — works in every modern browser
// and doesn't require any external library. The setTimeout-revoke is
// belt-and-suspenders for older Chromium versions that drop the click
// before they've finished writing the file to disk.
function downloadFile(content, filename, mimeType) {
  const blob = new Blob([content], { type: mimeType });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  a.style.display = 'none';
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 100);
}

function downloadCommissionCSV(entries, paymentMode) {
  // Aggregate entries into per-person totals — same logic the dashboard
  // uses, so what's in the file matches what Lucia just audited.
  const byPerson = {};
  for (const e of entries) {
    if (e.flag === 'INELIGIBLE') continue;
    if (!e.person || e.person === '—') continue;
    byPerson[e.person] = (byPerson[e.person] || 0) + e.amount;
  }

  // Per-person filter pass. Each skip-reason is tracked separately so
  // the status line can give Lucia an actionable summary ("3 missing
  // Contract Name" tells her to fix the People tab; "5 not Deel
  // Payment Mode" is expected and just informational).
  const rows = [];
  const skipped = { wrongMode: 0, noMode: 0, noContractName: 0, notInRoster: 0, zeroAmount: 0 };
  const wrongModeNames = [];
  const noModeNames = [];
  const noContractNameNames = [];
  const notInRosterNames = [];

  for (const [person, total] of Object.entries(byPerson)) {
    const p = lookupPerson(person);
    if (!p) {
      skipped.notInRoster += 1;
      if (notInRosterNames.length < 5) notInRosterNames.push(person);
      continue;
    }
    if (!p.paymentMode) {
      skipped.noMode += 1;
      if (noModeNames.length < 5) noModeNames.push(person);
      continue;
    }
    if (String(p.paymentMode).toLowerCase() !== String(paymentMode).toLowerCase()) {
      skipped.wrongMode += 1;
      if (wrongModeNames.length < 5) wrongModeNames.push(person);
      continue;
    }
    if (total === 0) { skipped.zeroAmount += 1; continue; }
    if (!p.contractName) {
      skipped.noContractName += 1;
      if (noContractNameNames.length < 5) noContractNameNames.push(person);
      continue;
    }
    rows.push([p.contractName, total.toFixed(2)]);
  }

  // Sort by Contract Name so the file is stable across runs — easier
  // for Lucia to spot-check against last month's file.
  rows.sort((a, b) => a[0].localeCompare(b[0]));

  // Build the CSV body. Headers are fixed — change here AND check that
  // the upload template at Deel/OnPay accepts them.
  const headers = ['Contract Name', 'Amount'];
  const lines = [headers, ...rows].map(r => r.map(csvEscape).join(','));
  // CRLF line endings are friendlier for Windows + Excel users opening
  // the file directly, and harmless on macOS/Linux. RFC 4180 calls for
  // CRLF anyway.
  const csv = lines.join('\r\n');

  // Filename: include platform + period so files don't collide in the
  // Downloads folder across months. Lowercase platform for consistency.
  const periodStr = lastPeriod || ymKey(new Date());
  const filename = `goldsoil_commission_${paymentMode.toLowerCase()}_${periodStr}.csv`;
  downloadFile(csv, filename, 'text/csv;charset=utf-8;');

  return { rows: rows.length, skipped, wrongModeNames, noModeNames, noContractNameNames, notInRosterNames };
}

// Update the status line below the download buttons with the outcome
// of the most recent download. Called by the click handler after
// downloadCommissionCSV returns.
function showDownloadStatus(result, paymentMode) {
  const el = document.getElementById('csv-download-status');
  if (!el) return;

  const parts = [];
  if (result.skipped.noContractName) {
    parts.push(`${result.skipped.noContractName} missing Contract Name `
      + `(${result.noContractNameNames.slice(0, 3).join(', ')}`
      + (result.skipped.noContractName > 3 ? `, +${result.skipped.noContractName - 3} more` : '')
      + ')');
  }
  if (result.skipped.noMode) {
    parts.push(`${result.skipped.noMode} with no Payment Mode set`);
  }
  if (result.skipped.notInRoster) {
    parts.push(`${result.skipped.notInRoster} not in People tab`);
  }
  // wrongMode and zeroAmount are expected — surface them only when no
  // rows were downloaded, to explain why the file might look empty.
  if (result.rows === 0) {
    if (result.skipped.wrongMode) {
      parts.push(`${result.skipped.wrongMode} are on a different Payment Mode (not ${paymentMode})`);
    }
    if (result.skipped.zeroAmount) {
      parts.push(`${result.skipped.zeroAmount} earned $0 this period`);
    }
  }

  let msg;
  if (result.rows === 0) {
    msg = `No rows matched filter "${paymentMode}". ` +
      (parts.length ? `Skipped: ${parts.join('; ')}.` : 'No commission earners this period.');
    el.className = 'csv-status error';
  } else if (parts.length) {
    msg = `Downloaded ${result.rows} row${result.rows === 1 ? '' : 's'} for ${paymentMode}. `
      + `Skipped: ${parts.join('; ')}.`;
    el.className = 'csv-status warn';
  } else {
    msg = `Downloaded ${result.rows} row${result.rows === 1 ? '' : 's'} for ${paymentMode}.`;
    el.className = 'csv-status success';
  }
  el.textContent = msg;
}

/* ------------------------------------------------------------
   9b. EMAIL — build payload, preview, send
   ----------------------------------------------------------------
   Flow:
     1. buildEmailPayload(entries, period) — aggregates per person,
        attaches email from PEOPLE tab, skips $0, returns:
          { recipients: [...], missingEmails: [...], ccEmail }
     2. If missingEmails.length > 0 → show blocker in modal.
        Otherwise → show confirmation modal.
     3. On confirm → POST to Apps Script doPost.
     4. Render results (sent / skipped / failed).
   ------------------------------------------------------------ */

function buildEmailPayload(entries, period) {
  const byPerson = {};
  for (const e of entries) {
    if (e.flag === 'INELIGIBLE') continue;
    if (!e.person || e.person === '—') continue;
    if (!byPerson[e.person]) byPerson[e.person] = { total: 0, byType: {}, entries: [] };
    byPerson[e.person].total += e.amount;
    const t = e.type || '(unknown)';
    byPerson[e.person].byType[t] = (byPerson[e.person].byType[t] || 0) + e.amount;
    byPerson[e.person].entries.push(e);
  }

  const recipients = [];
  const missingEmails = [];
  for (const [name, info] of Object.entries(byPerson)) {
    if (info.total === 0) continue;   // skip $0 — they see it on the dashboard
    const p = lookupPerson(name);
    const email = p && p.email ? p.email : null;
    if (!email) {
      missingEmails.push({ name, total: info.total });
      continue;
    }
    // Per-type rollup, sorted by |amount| desc — drives the type-header
    // order in the email layout.
    const breakdown = Object.entries(info.byType)
      .sort((a, b) => Math.abs(b[1]) - Math.abs(a[1]))
      .map(([type, amount]) => ({ type, amount }));
    // Per-line detail (mirrors the dashboard drill-down).
    const lines = buildEmailLines(info.entries);
    recipients.push({ name, email, total: info.total, breakdown, lines });
  }
  recipients.sort((a, b) => a.name.localeCompare(b.name));

  const ccPerson = lookupPerson(CC_RECIPIENT_NAME);
  const ccEmail = ccPerson && ccPerson.email ? ccPerson.email : null;

  return { recipients, missingEmails, ccEmail };
}

// Convert raw entries into the line objects the email template renders.
// Per-Lead types get aggregated (a LIM with 50 approved leads collapses
// to one line) — same behavior as the dashboard drill-down. Everything
// else is one line per entry.
function buildEmailLines(personEntries) {
  const perLeadByType = {};
  const individual = [];
  for (const e of personEntries) {
    if (/^Per-Lead/.test(e.type)) {
      if (!perLeadByType[e.type]) perLeadByType[e.type] = [];
      perLeadByType[e.type].push(e);
    } else {
      individual.push({
        type: e.type,
        source: e.source,
        amount: e.amount,
        calc: e.calc || '',
      });
    }
  }
  const aggregated = [];
  for (const [type, ents] of Object.entries(perLeadByType)) {
    const total = ents.reduce((s, x) => s + x.amount, 0);
    const count = ents.length;
    const amounts = [...new Set(ents.map(e => e.amount))];
    const calc = amounts.length === 1
      ? `${count} approved lead${count === 1 ? '' : 's'} × ${fmtUSD(amounts[0])} = ${fmtUSD(total)}`
      : `${count} approved leads, varying amounts = ${fmtUSD(total)}`;
    aggregated.push({
      type,
      source: `${count} lead${count === 1 ? '' : 's'}`,
      amount: total,
      calc,
    });
  }
  const all = [...individual, ...aggregated];
  all.sort((a, b) =>
    String(a.type).localeCompare(String(b.type)) ||
    String(a.source || '').localeCompare(String(b.source || ''))
  );
  return all;
}

function openSendEmailsModal() {
  const period = lastPeriod;
  const payload = buildEmailPayload(lastEntries, period);
  const modal = document.getElementById('email-modal');
  const body = document.getElementById('email-modal-body');
  const confirmBtn = document.getElementById('email-confirm-btn');
  const statusEl = document.getElementById('email-modal-status');

  statusEl.textContent = '';
  statusEl.className = 'hint';
  let html = '';

  if (payload.missingEmails.length > 0) {
    html += '<div class="modal-blocker">';
    html += '<h3>Cannot send — missing email addresses</h3>';
    html += '<p>The following ' + payload.missingEmails.length + ' people earn commission this period but have no Email in the People tab:</p>';
    html += '<ul>';
    for (const m of payload.missingEmails) {
      html += '<li><strong>' + escapeHTML(m.name) + '</strong> — ' + fmtUSD(m.total) + '</li>';
    }
    html += '</ul>';
    html += '<p>Add their email addresses to the People tab and re-run Calculate.</p>';
    html += '</div>';
    confirmBtn.disabled = true;
  } else if (!payload.ccEmail) {
    html += '<div class="modal-blocker">';
    html += '<h3>Cannot send — CC recipient not found</h3>';
    html += '<p>Could not find an email for <code>' + escapeHTML(CC_RECIPIENT_NAME) + '</code> in the People tab. ';
    html += 'Either add him with an Email, or update <code>CC_RECIPIENT_NAME</code> in script.js to match how he\'s listed.</p>';
    html += '</div>';
    confirmBtn.disabled = true;
  } else if (payload.recipients.length === 0) {
    html += '<div class="modal-blocker">';
    html += '<h3>Nothing to send</h3>';
    html += '<p>No one earned commission this period.</p>';
    html += '</div>';
    confirmBtn.disabled = true;
  } else {
    confirmBtn.disabled = false;
    html += '<p class="modal-lede">About to send <strong>' + payload.recipients.length + '</strong> email' + (payload.recipients.length === 1 ? '' : 's') + ' for <strong>' + escapeHTML(fmtPeriodName(period)) + '</strong>.<br>';
    html += 'CC: <code>' + escapeHTML(payload.ccEmail) + '</code></p>';

    html += '<table class="modal-recipient-table"><thead><tr>';
    html += '<th>Name</th><th>Email</th><th class="num">Total</th>';
    html += '</tr></thead><tbody>';
    for (const r of payload.recipients) {
      html += '<tr>';
      html += '<td>' + escapeHTML(r.name) + '</td>';
      html += '<td><code>' + escapeHTML(r.email) + '</code></td>';
      html += '<td class="num">' + fmtUSD(r.total) + '</td>';
      html += '</tr>';
    }
    html += '</tbody></table>';

    const sample = payload.recipients[0];
    html += '<details class="email-preview"><summary>Preview email (' + escapeHTML(sample.name) + ')</summary>';
    html += '<div class="email-preview-body">';
    html += '<div class="email-preview-meta">';
    html += '<div><strong>To:</strong> ' + escapeHTML(sample.email) + '</div>';
    html += '<div><strong>CC:</strong> ' + escapeHTML(payload.ccEmail) + '</div>';
    html += '<div><strong>Subject:</strong> Your GoldSoil Commission — ' + escapeHTML(fmtPeriodName(period)) + '</div>';
    html += '</div>';
    html += renderEmailPreviewHTML(sample, fmtPeriodName(period));
    html += '</div></details>';
  }

  body.innerHTML = html;
  modal.classList.remove('hidden');

  modal.dataset.payload = JSON.stringify({
    period: period,
    periodName: fmtPeriodName(period),
    ccEmail: payload.ccEmail,
    recipients: payload.recipients,
  });
}

function closeSendEmailsModal() {
  document.getElementById('email-modal').classList.add('hidden');
}

// Client-side preview that mirrors Apps Script's renderEmailHTML.
// If you change the template in Apps Script, eyeball this one too.
function renderEmailPreviewHTML(r, periodLabel) {
  const firstName = escapeHTML(String(r.name).split(' ')[0]);
  const groups = groupLinesByTypeForEmail(r.lines || [], r.breakdown || []);
  let rows = '';
  for (const g of groups) {
    // Type header
    rows += '<tr>'
      + '<td colspan="2" style="font-weight:600;padding:10px 0 5px 0;border-bottom:1px solid #D7CFBF;font-size:14px;">' + escapeHTML(g.type) + '</td>'
      + '<td style="text-align:right;font-family:monospace;font-weight:600;padding:10px 0 5px 0;border-bottom:1px solid #D7CFBF;font-size:14px;">' + fmtUSD(g.subtotal) + '</td>'
      + '</tr>';
    // Per-line rows under this type
    for (const ln of g.lines) {
      rows += '<tr>'
        + '<td style="padding:6px 8px 2px 16px;font-family:monospace;font-size:13px;color:#4A5468;width:32%;vertical-align:top;">' + escapeHTML(ln.source || '—') + '</td>'
        + '<td style="padding:6px 4px 2px 0;font-size:12px;color:#8A8E97;font-style:italic;vertical-align:top;">' + escapeHTML(ln.calc || '') + '</td>'
        + '<td style="padding:6px 0 2px 0;text-align:right;font-family:monospace;font-size:13px;width:18%;vertical-align:top;">' + fmtUSD(ln.amount) + '</td>'
        + '</tr>';
    }
  }
  return '<div style="font-family:sans-serif;color:#0F1B2D;line-height:1.55;font-size:14px;padding:1rem;background:white;border-radius:2px;">'
    + '<p>Hi ' + firstName + ',</p>'
    + '<p>Hope you\'re doing well.</p>'
    + '<p>Great news — you\'ve earned <strong style="color:#2E5D3D;">' + fmtUSD(r.total) + '</strong> in commissions for <strong>' + escapeHTML(periodLabel) + '</strong>. Here\'s your breakdown:</p>'
    + '<table style="border-collapse:collapse;width:100%;margin-top:8px;">' + rows
    + '<tr>'
    + '<td colspan="2" style="font-weight:600;font-size:15px;padding:14px 0 0 0;border-top:2px solid #D7CFBF;">Total</td>'
    + '<td style="text-align:right;font-family:monospace;font-weight:600;font-size:15px;padding:14px 0 0 0;border-top:2px solid #D7CFBF;">' + fmtUSD(r.total) + '</td>'
    + '</tr>'
    + '</table>'
    + '<p style="margin-top:24px;">Payment will be processed on the 20th of this month per the standard schedule.</p>'
    + '<p>Thank you for your ongoing commitment, consistency, and great work — your contribution is a key part of our success, and we\'re grateful to have you on the team.</p>'
    + '<p>If you have any questions about the calculation, please reply to this email — Anshul is CC\'d.</p>'
    + '<p style="color:#8A8E97;font-size:13px;margin-top:32px;">— GoldSoil HR</p>'
    + '</div>';
}

// Group lines by type, ordered to match the per-type breakdown. Same
// logic in Apps Script's groupLinesByType — keep them in sync.
function groupLinesByTypeForEmail(allLines, breakdown) {
  const linesByType = {};
  for (const ln of allLines) {
    if (!linesByType[ln.type]) linesByType[ln.type] = [];
    linesByType[ln.type].push(ln);
  }
  const groups = [];
  for (const b of breakdown) {
    if (linesByType[b.type]) {
      groups.push({ type: b.type, subtotal: b.amount, lines: linesByType[b.type] });
    }
  }
  // Defensive: any line whose type isn't in breakdown still shows up.
  const seen = new Set(breakdown.map(b => b.type));
  for (const [type, lines] of Object.entries(linesByType)) {
    if (seen.has(type)) continue;
    const subtotal = lines.reduce((s, x) => s + x.amount, 0);
    groups.push({ type, subtotal, lines });
  }
  return groups;
}

async function confirmAndSendEmails() {
  const modal = document.getElementById('email-modal');
  const confirmBtn = document.getElementById('email-confirm-btn');
  const statusEl = document.getElementById('email-modal-status');
  const payload = JSON.parse(modal.dataset.payload);

  confirmBtn.disabled = true;
  statusEl.className = 'hint';
  statusEl.textContent = 'Sending ' + payload.recipients.length + ' email(s)…';

  try {
    const res = await fetch(APPS_SCRIPT_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },  // text/plain avoids CORS preflight
      body: JSON.stringify({
        token: APPS_SCRIPT_TOKEN,
        action: 'sendCommissionEmails',
        ...payload,
      }),
    });
    const result = await res.json();

    if (result.error) {
      statusEl.className = 'hint error';
      statusEl.textContent = 'Error: ' + result.error;
      confirmBtn.disabled = false;
      return;
    }

    const sent = (result.sent || []).length;
    const skipped = (result.skipped || []).length;
    const failed = (result.failed || []).length;

    let msg = '✓ Sent ' + sent + ' email' + (sent === 1 ? '' : 's') + '.';
    if (skipped > 0) msg += ' Skipped ' + skipped + ' (already sent for this period).';
    if (failed > 0) {
      msg += ' Failed: ' + failed + '.';
      statusEl.className = 'hint error';
    } else {
      statusEl.className = 'hint success';
    }
    statusEl.textContent = msg;

    if (failed > 0) {
      const detail = result.failed.map(f => f.name + ' (' + f.email + '): ' + f.error).join('\n');
      console.error('Email failures:\n' + detail);
    }
  } catch (err) {
    statusEl.className = 'hint error';
    statusEl.textContent = 'Network error: ' + err.message;
    confirmBtn.disabled = false;
  }
}

/* ------------------------------------------------------------
   9c. COMMISSION HISTORY — auto-written after every Calculate
   ----------------------------------------------------------------
   After a successful Calculate, the client aggregates lastEntries
   into (Period, Person, Type) rows and POSTs them to the Apps Script
   writeCommissionHistory endpoint. The server writes them to the
   Commission_History tab with replace-on-write semantics, so re-runs
   for the same period are idempotent.

   Failure model: BEST-EFFORT. A history write failure does NOT fail
   the calc — the dashboard has already rendered by the time this
   fires, and the user has what they came for. Failures log to console
   and produce a soft warning in the status line; the next successful
   Calculate will rewrite the period cleanly.
   ------------------------------------------------------------ */

// Aggregate lastEntries into the Commission_History row shape. Groups
// by (person, type) — same dimensions Lucia sees in the dashboard's
// per-type breakdown column — so a person earning across multiple
// types produces one row per type, not one row per source deal.
//
// Skips:
//   - INELIGIBLE entries (already $0, not real payouts)
//   - Unattributed entries (person === '—'; data-quality issues, not
//     people earning money)
//   - Zero-amount aggregates (after grouping, if a type nets to $0
//     across all sources for a person, we don't write it — Commission_
//     History should reflect actual paid commissions, not the dashboard
//     drill-down)
//
// Attaches Department + Role from the PEOPLE roster so the trends
// panel can group by either axis without doing roster joins client-side
// on every render. If a person isn't in the roster (which would mean
// they had INELIGIBLE entries that were already filtered out above,
// but defensively), they get "(Unassigned)" / "" — matches the
// dashboard's department-grouping default.
function buildHistoryRows(entries, period) {
  const byKey = {};
  for (const e of entries) {
    if (e.flag === 'INELIGIBLE') continue;
    if (!e.person || e.person === '—') continue;
    const key = e.person + '|||' + e.type;
    if (!byKey[key]) byKey[key] = { person: e.person, type: e.type, amount: 0 };
    byKey[key].amount += e.amount;
  }
  const rows = [];
  for (const k of Object.keys(byKey)) {
    const { person, type, amount } = byKey[k];
    if (amount === 0) continue;
    const p = lookupPerson(person);
    rows.push({
      period,
      person,
      department: (p && p.department) || '(Unassigned)',
      role: (p && p.role) || '',
      type,
      // Round to cents to avoid float drift like 99.99999998 ending up
      // in the sheet. The server stores raw values; rounding here keeps
      // the on-disk numbers clean.
      amount: Math.round(amount * 100) / 100,
    });
  }
  // Stable order — Person ASC, then Type ASC — so re-runs of the same
  // period produce identically-ordered rows. Cosmetic; the calculator
  // doesn't depend on order, but it makes Commission_History readable
  // when someone opens it in Sheets.
  rows.sort((a, b) =>
    a.person.localeCompare(b.person) || a.type.localeCompare(b.type)
  );
  return rows;
}

// POST aggregated history rows to the Apps Script. Returns a result
// object — never throws — so the caller can fold the outcome into the
// status line without try/catch. Updates lastHistoryData in-memory on
// success so the trends panel reflects today's calc immediately
// without needing a round-trip back to doGet.
async function writeCommissionHistoryAsync(period, entries) {
  const rows = buildHistoryRows(entries, period);
  try {
    const res = await fetch(APPS_SCRIPT_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },  // text/plain avoids CORS preflight
      body: JSON.stringify({
        token: APPS_SCRIPT_TOKEN,
        action: 'writeCommissionHistory',
        period,
        rows,
      }),
    });
    if (!res.ok) {
      return { ok: false, error: `HTTP ${res.status}`, written: 0, replaced: 0 };
    }
    const result = await res.json();
    if (result.error) {
      return { ok: false, error: result.error, written: 0, replaced: 0 };
    }

    // Update in-memory history with the just-written period. Drop
    // existing rows for this period, then append the new ones — mirrors
    // the server-side replace-on-write semantics. Shape matches what
    // doGet would return so any downstream code reading lastHistoryData
    // doesn't have to special-case "rows written this session".
    lastHistoryData = lastHistoryData.filter(r => String(r.Period) !== period);
    const ts = new Date().toISOString();
    for (const row of rows) {
      lastHistoryData.push({
        Period: row.period,
        Person: row.person,
        Department: row.department,
        Role: row.role,
        Type: row.type,
        Amount: row.amount,
        'Calc Timestamp': ts,
      });
    }

    return {
      ok: true,
      error: null,
      written: result.written || 0,
      replaced: result.replaced || 0,
    };
  } catch (err) {
    return { ok: false, error: err.message, written: 0, replaced: 0 };
  }
}

/* ------------------------------------------------------------
   9d. TRENDS PANEL — chart, metric cards, top earners
   ----------------------------------------------------------------
   Reads from `lastHistoryData` (hydrated by fetchSheetData and
   appended-to by writeCommissionHistoryAsync). All rendering is
   pure-function — no fetching, no recomputation of commissions —
   so tab switches and mode toggles are free.

   The current chart instance is held in `trendChartInstance` so
   we can destroy() it before a re-render. Chart.js mutates DOM and
   leaks listeners if you stack instances on the same canvas.
   ------------------------------------------------------------ */

let trendChartInstance = null;
let currentTrendMode = 'department';  // 'department' | 'role' — set by toggle

// Top-level entry point. Decides between empty state and full render,
// then dispatches to the three subroutines. Safe to call repeatedly —
// e.g. after every Calculate, on tab switch, on mode toggle.
function renderTrends() {
  const emptyEl = document.getElementById('trends-empty');
  const bodyEl = document.getElementById('trends-body');
  if (!emptyEl || !bodyEl) return;  // tab pane not yet in DOM (shouldn't happen post-DOMContentLoaded)

  // Filter to numeric-amount rows defensively — the Amount column can
  // come back as a string if the sheet was hand-edited.
  const rows = (lastHistoryData || []).filter(r => r && r.Period && r.Person);

  if (rows.length === 0) {
    emptyEl.hidden = false;
    bodyEl.hidden = true;
    if (trendChartInstance) { trendChartInstance.destroy(); trendChartInstance = null; }
    return;
  }

  emptyEl.hidden = true;
  bodyEl.hidden = false;

  renderTrendsMetrics(rows);
  renderTrendsChart(rows, currentTrendMode);
  renderTrendsTopEarners(rows);
}

// Slice to most recent TREND_MONTHS distinct periods. Returns the rows
// belonging to those periods plus the sorted list of period keys.
// History rows for periods outside the window are dropped from every
// chart/table/metric — keeps the trends panel scoped to "recent" even
// after years of accumulation.
function trimToTrendWindow(rows) {
  const periodSet = new Set();
  for (const r of rows) periodSet.add(String(r.Period));
  const allPeriods = [...periodSet].sort();
  const windowPeriods = allPeriods.slice(-TREND_MONTHS);
  const windowSet = new Set(windowPeriods);
  const windowRows = rows.filter(r => windowSet.has(String(r.Period)));
  return { rows: windowRows, periods: windowPeriods };
}

// Sum Amount by Period — used for grand-total-per-month metric and
// for the per-month totals shown on chart tooltips.
function totalsByPeriod(rows) {
  const out = {};
  for (const r of rows) {
    const p = String(r.Period);
    out[p] = (out[p] || 0) + (Number(r.Amount) || 0);
  }
  return out;
}

// Sum Amount by Period × key (key = 'Department' or 'Role'). Returns:
//   { periods: ['2026-04', ...], series: { 'Acquisitions': [3000, ...], 'Sales': [200, ...] } }
// Every series array is aligned to `periods` (zeros where the key
// didn't earn in that month) so Chart.js datasets line up cleanly.
function aggregateByPeriodAndKey(rows, periods, keyName) {
  const byPeriod = {};
  for (const r of rows) {
    const period = String(r.Period);
    let key = r[keyName];
    if (key == null || String(key).trim() === '') {
      key = keyName === 'Role' ? '(No Role)' : '(Unassigned)';
    }
    if (!byPeriod[period]) byPeriod[period] = {};
    byPeriod[period][key] = (byPeriod[period][key] || 0) + (Number(r.Amount) || 0);
  }
  const keySet = new Set();
  for (const p of periods) {
    if (!byPeriod[p]) continue;
    for (const k of Object.keys(byPeriod[p])) keySet.add(k);
  }
  // Order keys by total descending so the biggest stack is at the bottom
  // (Chart.js stacks in dataset order, bottom-up). Most-impactful at the
  // bottom = visually anchored, easier to compare across months.
  const keys = [...keySet].sort((a, b) => {
    const ta = periods.reduce((s, p) => s + ((byPeriod[p] && byPeriod[p][a]) || 0), 0);
    const tb = periods.reduce((s, p) => s + ((byPeriod[p] && byPeriod[p][b]) || 0), 0);
    return tb - ta;
  });
  const series = {};
  for (const k of keys) {
    series[k] = periods.map(p => (byPeriod[p] && byPeriod[p][k]) || 0);
  }
  return { periods, series };
}

/* ---------- Metric cards ---------- */

function renderTrendsMetrics(rows) {
  const { rows: windowRows, periods } = trimToTrendWindow(rows);
  const totals = totalsByPeriod(windowRows);
  const latestPeriod = periods[periods.length - 1];
  const latestTotal = totals[latestPeriod] || 0;

  // Rolling avg over the window. Stays muted (italic, smaller) if we
  // only have one month of data — "average" of one is just the same
  // number, which is misleading.
  const avgWindow = periods.reduce((s, p) => s + (totals[p] || 0), 0) / Math.max(periods.length, 1);
  const showAvg = periods.length >= 2;

  // Active earners in the LATEST period only — what the next payroll
  // run will actually pay out. Not the cumulative-window count.
  const latestEarners = new Set();
  for (const r of windowRows) {
    if (String(r.Period) !== latestPeriod) continue;
    if ((Number(r.Amount) || 0) === 0) continue;
    latestEarners.add(r.Person);
  }

  // Top 3 concentration for the latest period — "what % of this month's
  // payout went to the top 3 earners". High concentration = single point
  // of risk; useful management signal.
  const personTotalsLatest = {};
  for (const r of windowRows) {
    if (String(r.Period) !== latestPeriod) continue;
    const p = r.Person;
    personTotalsLatest[p] = (personTotalsLatest[p] || 0) + (Number(r.Amount) || 0);
  }
  const sortedLatest = Object.values(personTotalsLatest).sort((a, b) => b - a);
  const top3Sum = sortedLatest.slice(0, 3).reduce((s, v) => s + v, 0);
  const top3Share = latestTotal > 0 ? (top3Sum / latestTotal * 100) : 0;

  const html = `
    <div class="metric-card">
      <div class="metric-card-label">Latest Month</div>
      <div class="metric-card-value">${fmtUSD(latestTotal)}</div>
      <div class="metric-card-sub">${escapeHTML(fmtPeriodName(latestPeriod))}</div>
    </div>
    <div class="metric-card">
      <div class="metric-card-label">${periods.length}-Month Avg</div>
      <div class="metric-card-value ${showAvg ? '' : 'muted'}">${showAvg ? fmtUSD(avgWindow) : '—'}</div>
      <div class="metric-card-sub">${showAvg ? 'across rolling window' : 'need ≥2 months'}</div>
    </div>
    <div class="metric-card">
      <div class="metric-card-label">Top 3 Share</div>
      <div class="metric-card-value">${top3Share.toFixed(1)}%</div>
      <div class="metric-card-sub">of latest month's payout</div>
    </div>
    <div class="metric-card">
      <div class="metric-card-label">Active Earners</div>
      <div class="metric-card-value">${latestEarners.size}</div>
      <div class="metric-card-sub">in ${escapeHTML(fmtPeriodName(latestPeriod))}</div>
    </div>
  `;
  document.getElementById('trends-metrics').innerHTML = html;
}

/* ---------- Stacked bar chart ---------- */

function renderTrendsChart(rows, mode) {
  const { rows: windowRows, periods } = trimToTrendWindow(rows);
  const keyName = mode === 'role' ? 'Role' : 'Department';
  const { series } = aggregateByPeriodAndKey(windowRows, periods, keyName);
  const palette = mode === 'role' ? ROLE_COLORS : DEPT_COLORS;

  const datasets = Object.entries(series).map(([key, values]) => ({
    label: key,
    data: values,
    backgroundColor: palette[key] || '#8A8E97',
    borderWidth: 0,
  }));

  const labels = periods.map(p => fmtPeriodName(p));

  const ctx = document.getElementById('trend-chart-canvas');
  if (!ctx) return;

  if (trendChartInstance) trendChartInstance.destroy();
  trendChartInstance = new Chart(ctx, {
    type: 'bar',
    data: { labels, datasets },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      // Slightly tighter bars when there's only 1-2 periods so the
      // chart doesn't look like a giant single column.
      barPercentage: periods.length <= 2 ? 0.45 : 0.85,
      scales: {
        x: {
          stacked: true,
          grid: { display: false },
          ticks: { color: '#4A5468', font: { family: 'Inter, sans-serif', size: 12 } },
        },
        y: {
          stacked: true,
          beginAtZero: true,
          grid: { color: '#E8E2D4' },
          ticks: {
            color: '#4A5468',
            font: { family: 'IBM Plex Mono, monospace', size: 11 },
            callback: (v) => fmtUSD(v),
          },
        },
      },
      plugins: {
        legend: {
          position: 'bottom',
          labels: {
            color: '#0F1B2D',
            font: { family: 'Inter, sans-serif', size: 12 },
            boxWidth: 14,
            boxHeight: 14,
            padding: 12,
          },
        },
        tooltip: {
          backgroundColor: '#0F1B2D',
          titleFont: { family: 'Inter, sans-serif', size: 13 },
          bodyFont: { family: 'IBM Plex Mono, monospace', size: 12 },
          callbacks: {
            label: (ctx) => `${ctx.dataset.label}: ${fmtUSD(ctx.parsed.y)}`,
            // Total line at the bottom of each tooltip — useful when
            // there are many stacked segments and the eye can't add them up.
            footer: (items) => {
              const total = items.reduce((s, x) => s + x.parsed.y, 0);
              return `Total: ${fmtUSD(total)}`;
            },
          },
        },
      },
    },
  });
}

/* ---------- Top earners table ---------- */

function renderTrendsTopEarners(rows) {
  const { rows: windowRows, periods } = trimToTrendWindow(rows);

  // Aggregate cumulative totals across the visible window.
  const personData = {};
  for (const r of windowRows) {
    const p = r.Person;
    if (!personData[p]) {
      personData[p] = { total: 0, department: r.Department || '(Unassigned)' };
    }
    personData[p].total += Number(r.Amount) || 0;
    // Most-recent Department wins (in case someone changed department
    // mid-window — typical for new hires getting reassigned).
    personData[p].department = r.Department || personData[p].department;
  }
  const sorted = Object.entries(personData)
    .filter(([, v]) => v.total !== 0)
    .sort((a, b) => b[1].total - a[1].total);
  const grandTotal = sorted.reduce((s, [, v]) => s + v.total, 0);
  const top10 = sorted.slice(0, 10);

  if (top10.length === 0) {
    document.getElementById('trends-top-earners').innerHTML =
      '<p class="hint">No earners in the visible window.</p>';
    return;
  }

  const windowLabel = periods.length === 1
    ? fmtPeriodName(periods[0])
    : `${fmtPeriodName(periods[0])} – ${fmtPeriodName(periods[periods.length - 1])}`;

  let html = `<table class="top-earners-table">`;
  html += `<thead><tr>`;
  html += `<th class="rank">#</th>`;
  html += `<th>Person</th>`;
  html += `<th>Department</th>`;
  html += `<th class="num">${escapeHTML(windowLabel)} Total</th>`;
  html += `<th class="num">Share</th>`;
  html += `</tr></thead><tbody>`;
  top10.forEach(([person, info], i) => {
    const share = grandTotal > 0 ? (info.total / grandTotal * 100).toFixed(1) : '0.0';
    html += `<tr>`;
    html += `<td class="rank">${i + 1}</td>`;
    html += `<td class="name">${escapeHTML(person)}</td>`;
    html += `<td><span class="dept-tag">${escapeHTML(info.department)}</span></td>`;
    html += `<td class="num">${fmtUSD(info.total)}</td>`;
    html += `<td class="num">${share}%</td>`;
    html += `</tr>`;
  });
  html += `</tbody></table>`;

  document.getElementById('trends-top-earners').innerHTML = html;
}

/* ------------------------------------------------------------
   10. UI WIRING
   ------------------------------------------------------------ */
let lastEntries = [];
let lastPeriod = null;

function previousMonthKey() {
  const now = new Date();
  const lastMonth = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  return ymKey(lastMonth);
}

/* ============================================================
   11. REGULAR PAYROLL — Time Doctor hours review
   ------------------------------------------------------------
   Self-contained, separate from the commission calc. Lucia uploads the
   Time Doctor "Hours Tracked" CSV; we parse the per-day decimal-hour
   columns and flag two things she manually checks today:
     (a) any WEEKDAY over the daily limit (default 8h), and
     (b) any WEEKEND work, since staff are scheduled Mon–Fri only.
   Output is a review list (flags, not deductions) plus a CSV export.
   No Google Sheet, no network — everything runs in the browser.
   ============================================================ */

// Module state — last parsed/analyzed payroll run, kept so the CSV
// download button can re-emit exactly what's on screen.
let lastPayrollAnalysis = null;
let lastPayrollMeta = null;

// Minimal RFC-4180 CSV parser. Time Doctor headers contain quoted commas
// (e.g. "Fri, May 1 (Decimal)"), so a naive split(',') corrupts the
// columns — we have to honor quotes. Returns an array of row arrays.
function parsePayrollCSV(text) {
  const rows = [];
  let row = [], field = '', inQuotes = false;
  // Normalize newlines so \r\n and \r both behave like \n.
  const s = text.replace(/\r\n?/g, '\n');
  for (let i = 0; i < s.length; i++) {
    const c = s[i];
    if (inQuotes) {
      if (c === '"') {
        if (s[i + 1] === '"') { field += '"'; i++; }  // escaped quote
        else inQuotes = false;
      } else field += c;
    } else if (c === '"') {
      inQuotes = true;
    } else if (c === ',') {
      row.push(field); field = '';
    } else if (c === '\n') {
      row.push(field); rows.push(row); row = []; field = '';
    } else field += c;
  }
  // Flush trailing field/row (file may not end in a newline).
  if (field.length > 0 || row.length > 0) { row.push(field); rows.push(row); }
  // Drop fully-empty trailing rows.
  return rows.filter(r => r.some(cell => cell.trim() !== ''));
}

// From the header row, locate the per-day DECIMAL columns. Time Doctor
// emits each day twice: "<Day>, Mon N (Hours & minutes)" and
// "<Day>, Mon N (Decimal)". We only want the Decimal ones, and we must
// NOT catch the period "Total (Decimal)" column. The day-of-week prefix
// before the first comma tells us weekday vs weekend.
const WEEKEND_DAYS = new Set(['Sat', 'Sun']);
function findPayrollDayColumns(header) {
  const cols = [];
  for (let i = 0; i < header.length; i++) {
    const h = (header[i] || '').trim();
    const m = h.match(/^(Mon|Tue|Wed|Thu|Fri|Sat|Sun),\s*(.+?)\s*\(Decimal\)$/);
    if (m) {
      cols.push({
        idx: i,
        dayName: m[1],
        label: (m[2] || '').trim(),          // e.g. "May 1"
        fullLabel: `${m[1]}, ${(m[2] || '').trim()}`,
        isWeekend: WEEKEND_DAYS.has(m[1]),
      });
    }
  }
  return cols;
}

// Locate a named column case-insensitively; returns -1 if absent.
function findHeaderIdx(header, name) {
  const target = name.toLowerCase();
  return header.findIndex(h => (h || '').trim().toLowerCase() === target);
}

// Parse a decimal-hours cell to a number, tolerating blanks / stray text.
function parsePayrollHours(v) {
  if (v == null) return 0;
  const n = parseFloat(String(v).trim());
  return Number.isFinite(n) ? n : 0;
}

// Core analysis. opts = { dailyLimit, graceMin, flagWeekend }.
// Returns { dayColumns, flagged: [...], clean: [...], dateRange }.
// Each flagged person: { name, email, totalHrs, overDays:[{label,dayName,hrs,excessMin}],
//                        weekendDays:[{label,dayName,hrs}], totalOverHrs, worstDayHrs, weekendHrs }.
function analyzePayroll(rows, opts) {
  const dailyLimit = opts.dailyLimit;
  const graceHrs = (opts.graceMin || 0) / 60;
  const threshold = dailyLimit + graceHrs;   // a weekday over THIS is flagged
  const flagWeekend = opts.flagWeekend;

  const header = rows[0];
  const dayCols = findPayrollDayColumns(header);
  if (dayCols.length === 0) {
    throw new Error('No daily "(Decimal)" hour columns found — is this a Time Doctor "Hours Tracked" export?');
  }
  const nameIdx = Math.max(0, findHeaderIdx(header, 'Name'));
  const emailIdx = findHeaderIdx(header, 'Email');
  const totalDecIdx = findHeaderIdx(header, 'Total (Decimal)');

  const flagged = [], clean = [];

  for (let r = 1; r < rows.length; r++) {
    const row = rows[r];
    const name = (row[nameIdx] || '').trim();
    if (!name) continue;
    const email = emailIdx >= 0 ? (row[emailIdx] || '').trim() : '';
    const totalHrs = totalDecIdx >= 0 ? parsePayrollHours(row[totalDecIdx]) : 0;

    const overDays = [], weekendDays = [];
    let totalOverHrs = 0, worstDayHrs = 0, weekendHrs = 0;

    for (const col of dayCols) {
      const hrs = parsePayrollHours(row[col.idx]);
      if (col.isWeekend) {
        if (flagWeekend && hrs > 0) {
          weekendDays.push({ label: col.fullLabel, dayName: col.dayName, hrs });
          weekendHrs += hrs;
        }
      } else if (hrs > threshold) {
        const excessHrs = hrs - dailyLimit;       // report excess vs the LIMIT, not the grace-padded threshold
        overDays.push({ label: col.fullLabel, dayName: col.dayName, hrs, excessMin: excessHrs * 60 });
        totalOverHrs += excessHrs;
        if (hrs > worstDayHrs) worstDayHrs = hrs;
      }
    }

    const rec = { name, email, totalHrs, overDays, weekendDays, totalOverHrs, worstDayHrs, weekendHrs };
    if (overDays.length > 0 || weekendDays.length > 0) flagged.push(rec);
    else clean.push(rec);
  }

  // Sort flagged by severity: total weekday excess + weekend hours, desc.
  flagged.sort((a, b) => (b.totalOverHrs + b.weekendHrs) - (a.totalOverHrs + a.weekendHrs));
  clean.sort((a, b) => b.totalHrs - a.totalHrs);

  const dateRange = dayCols.length
    ? `${dayCols[0].fullLabel} – ${dayCols[dayCols.length - 1].fullLabel}`
    : '';

  return { dayColumns: dayCols, flagged, clean, dateRange };
}

// Format minutes-over compactly, e.g. "+1h 32m" or "+18m".
function fmtOverMin(min) {
  const r = Math.round(min);
  if (r >= 60) {
    const h = Math.floor(r / 60), m = r % 60;
    return m ? `+${h}h ${m}m` : `+${h}h`;
  }
  return `+${r}m`;
}

function fmtHrs(h) { return `${h.toFixed(2)}h`; }

function renderPayrollResults(analysis, meta) {
  const wrap = document.getElementById('payroll-summary');
  const resultsSection = document.getElementById('payroll-results');
  const { flagged, clean } = analysis;
  const totalPeople = flagged.length + clean.length;

  let html = '';

  // Banner: what was analyzed + the rule that was applied.
  const graceNote = meta.graceMin > 0 ? ` · ${meta.graceMin}m grace` : '';
  html += `<div class="period-banner">`
        + `<span>Report: <strong>${escapeHTML(analysis.dateRange || meta.fileName)}</strong></span>`
        + `<span>Limit: <strong>${meta.dailyLimit}h/day${graceNote}</strong> · ${flagged.length} of ${totalPeople} flagged</span>`
        + `</div>`;

  if (flagged.length === 0) {
    html += `<div class="payroll-clean-banner">✓ No one exceeded ${meta.dailyLimit}h on a weekday${meta.flagWeekend ? ' and no weekend work was logged' : ''} in this period.</div>`;
    wrap.innerHTML = html;
    resultsSection.classList.remove('hidden');
    return;
  }

  html += `<table class="detail-table payroll-table">`
        + `<thead><tr>`
        + `<th>Name</th>`
        + `<th class="num">Total</th>`
        + `<th class="num">Days&nbsp;&gt;&nbsp;limit</th>`
        + `<th class="num">Worst day</th>`
        + `<th class="num">Total over</th>`
        + `<th class="num">Weekend</th>`
        + `</tr></thead><tbody>`;

  for (const p of flagged) {
    const worst = p.worstDayHrs > 0 ? fmtHrs(p.worstDayHrs) : '—';
    const over = p.totalOverHrs > 0 ? `+${p.totalOverHrs.toFixed(2)}h` : '—';
    const wknd = p.weekendHrs > 0 ? fmtHrs(p.weekendHrs) : '—';

    // Build the per-day detail rows for the expandable section.
    let detail = '';
    if (p.overDays.length) {
      detail += `<div class="payroll-detail-group"><div class="payroll-detail-head">Weekday over ${meta.dailyLimit}h</div>`;
      for (const d of p.overDays) {
        detail += `<div class="payroll-detail-row"><span>${escapeHTML(d.label)}</span>`
                + `<span class="mono">${fmtHrs(d.hrs)} <em>${fmtOverMin(d.excessMin)}</em></span></div>`;
      }
      detail += `</div>`;
    }
    if (p.weekendDays.length) {
      detail += `<div class="payroll-detail-group"><div class="payroll-detail-head">Weekend work</div>`;
      for (const d of p.weekendDays) {
        detail += `<div class="payroll-detail-row"><span>${escapeHTML(d.label)}</span>`
                + `<span class="mono">${fmtHrs(d.hrs)}</span></div>`;
      }
      detail += `</div>`;
    }

    const wkndBadge = p.weekendDays.length ? ` <span class="payroll-badge">weekend</span>` : '';

    html += `<tr class="flag-review payroll-row">`
          + `<td><details class="payroll-person"><summary><strong>${escapeHTML(p.name)}</strong>${wkndBadge}`
          + `<span class="payroll-email">${escapeHTML(p.email)}</span></summary>`
          + `<div class="payroll-detail">${detail}</div></details></td>`
          + `<td class="num">${fmtHrs(p.totalHrs)}</td>`
          + `<td class="num">${p.overDays.length || '—'}</td>`
          + `<td class="num">${worst}</td>`
          + `<td class="num">${over}</td>`
          + `<td class="num">${wknd}</td>`
          + `</tr>`;
  }
  html += `</tbody></table>`;

  // Clean list — collapsed, for completeness so Lucia sees everyone was reviewed.
  if (clean.length) {
    html += `<details class="payroll-clean"><summary>${clean.length} within limits — no flags</summary>`
          + `<div class="payroll-clean-list">`
          + clean.map(p => `<span>${escapeHTML(p.name)} <em>${fmtHrs(p.totalHrs)}</em></span>`).join('')
          + `</div></details>`;
  }

  wrap.innerHTML = html;
  resultsSection.classList.remove('hidden');
}

// CSV export of the flag list — one row per flagged day, so Lucia can
// drop it into her review notes or attach it to an approval request.
function downloadPayrollReport(analysis, meta) {
  const lines = [['Name', 'Email', 'Date', 'Day', 'Hours', 'Flag', 'Over (min)'].join(',')];
  for (const p of analysis.flagged) {
    for (const d of p.overDays) {
      lines.push([
        csvEscape(p.name), csvEscape(p.email), csvEscape(d.label), csvEscape(d.dayName),
        d.hrs.toFixed(3), `Weekday over ${meta.dailyLimit}h`, Math.round(d.excessMin),
      ].join(','));
    }
    for (const d of p.weekendDays) {
      lines.push([
        csvEscape(p.name), csvEscape(p.email), csvEscape(d.label), csvEscape(d.dayName),
        d.hrs.toFixed(3), 'Weekend work', '',
      ].join(','));
    }
  }
  const csv = lines.join('\r\n');
  const stamp = (analysis.dateRange || meta.fileName || 'report').replace(/[^A-Za-z0-9]+/g, '_');
  downloadFile(csv, `payroll_hours_flags_${stamp}.csv`, 'text/csv;charset=utf-8;');
}

/* ------------------------------------------------------------
   11b. CUSTOM EXPORT (long format) — full three-control review
   ------------------------------------------------------------
   The Time Doctor Custom Export with "Daily breakdown per user" gives one
   row per person per day, with Date, Time tracked, Paid/Unpaid Break time,
   and Start/End times. That lets us run all of Lucia's controls in one pass:
     • Hours  — weekday Time tracked over the daily limit (break INCLUDED,
                per company rule: the 8h limit counts the paid break).
     • Breaks — daily paid+unpaid break over the policy limit (30 min).
     • Window — start before / end after the operational window (8–5 Central).
     • Weekend — any Sat/Sun work.
   ------------------------------------------------------------ */

// Detect which Time Doctor export was uploaded. Custom Export (long) has a
// per-row "Date" column AND a "Time tracked" column; the Hours Tracked
// (wide) export instead has per-day "(Decimal)" columns and no "Date".
function detectPayrollFormat(rows) {
  const header = (rows[0] || []).map(h => (h || '').trim().toLowerCase());
  if (header.includes('date') && header.includes('time tracked')) return 'custom';
  return 'wide';
}

const MONTH_NAMES = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
const DOW_NAMES = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];

// Parse a Time Doctor "M/D/YY" date → { dayName, isWeekend, label, sort }.
function parsePayrollDate(s) {
  const m = String(s || '').trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
  if (!m) return null;
  let [_, mo, da, yr] = m;
  mo = parseInt(mo, 10); da = parseInt(da, 10); yr = parseInt(yr, 10);
  if (yr < 100) yr += 2000;
  const d = new Date(yr, mo - 1, da);
  const dow = d.getDay();
  return {
    dayName: DOW_NAMES[dow],
    isWeekend: dow === 0 || dow === 6,
    label: `${MONTH_NAMES[mo - 1]} ${da}`,
    sort: yr * 10000 + mo * 100 + da,
  };
}

// "HH:MM" → minutes since midnight; blank/invalid → null.
function parseClock(s) {
  const m = String(s || '').trim().match(/^(\d{1,2}):(\d{2})$/);
  if (!m) return null;
  return parseInt(m[1], 10) * 60 + parseInt(m[2], 10);
}

// "07:34" passthrough for display (kept as-is from the export).
function clockStr(s) { return String(s || '').trim(); }

/* ------------------------------------------------------------
   Additional Hours Request (approvals) — suppression source
   ------------------------------------------------------------
   A tab on the Commissions Report sheet, synced from Salesforce, logs
   requests for extra/weekend work. When a request is Approved (status
   column), the matching flag is suppressed instead of shown. We match on
   EMAIL first (reliable across systems) and fall back to normalized name,
   since Time Doctor uses informal display names (e.g. "Art O") while the
   sheet likely uses full names.

   Column detection is intentionally flexible — we don't hard-code the
   sheet's exact headers (which still need confirming). We look for a
   status column, a date column (plus optional end-date for ranges), an
   email column, a name column, and an optional type/category column.
   ------------------------------------------------------------ */
const TAB_ADDITIONAL_HOURS = 'Additional Hours Request';

// Find the first header index whose lowercased text contains any token.
function findColByTokens(headerLower, tokens) {
  for (let i = 0; i < headerLower.length; i++) {
    if (tokens.some(t => headerLower[i].includes(t))) return i;
  }
  return -1;
}

// Parse a date value → calendar sort key (yyyymmdd int) or null.
// Reads the calendar date from the STRING and never applies a timezone shift:
// an ISO datetime like "2026-05-02T00:00:00.000Z" (how Apps Script serializes a
// real date cell) must stay May 2, not roll back to May 1 in a western locale.
// Accepts M/D/YY[YYYY] (with or without trailing time), YYYY-MM-DD[Thh:mm…],
// and a Date object (harmless if one is ever passed through).
function parseApprovalDateKey(s) {
  if (s instanceof Date && !isNaN(s)) return s.getFullYear()*10000+(s.getMonth()+1)*100+s.getDate();
  const str = String(s || '').trim();
  if (!str) return null;
  let m = str.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);        // ISO date or datetime → take date part literally
  if (m) return (+m[1])*10000+(+m[2])*100+(+m[3]);
  m = str.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})/);        // M/D/YYYY (optionally followed by a time)
  if (m) { let mo=+m[1], da=+m[2], yr=+m[3]; if (yr<100) yr+=2000; return yr*10000+mo*100+da; }
  const d = new Date(str);
  if (!isNaN(d)) return d.getFullYear()*10000+(d.getMonth()+1)*100+d.getDate();
  return null;
}

// Classify a request STATUS into approved / pending / rejected.
//   approved → excess hours are included in billable
//   pending  → a request exists but isn't approved yet → excess held as pending
//   rejected → treated like no request → excess removed
// A blank status on an existing request row is treated as pending (a row was
// filed but no decision recorded). Adjust the keyword sets here if the
// Additional Hours Request tab uses different status labels.
function classifyApprovalStatus(s) {
  const t = String(s || '').trim().toLowerCase();
  if (!t) return 'pending';
  if (t.includes('approve')) return 'approved';
  if (/(reject|den|declin|cancel|void|withdraw)/.test(t)) return 'rejected';
  return 'pending';
}

// Parse a requested-hours cell into a number of HOURS, or null if blank/
// unparseable (null = "no stated limit" → approved excess isn't capped).
// Accepts "2", "2h", "2 hrs", "1.5", "1:30" (→1.5h), "90m"/"90 min" (→1.5h).
function parseRequestedHours(v) {
  if (v == null) return null;
  if (typeof v === 'number') return isFinite(v) && v > 0 ? v : null;
  let s = String(v).trim().toLowerCase();
  if (!s) return null;
  let m = s.match(/^(\d{1,2}):(\d{2})$/);                 // H:MM
  if (m) return parseInt(m[1],10) + parseInt(m[2],10)/60;
  m = s.match(/(\d+(?:\.\d+)?)\s*(m|min|mins|minutes)\b/); // minutes
  if (m) return parseFloat(m[1]) / 60;
  m = s.match(/(\d+(?:\.\d+)?)/);                          // bare number / "2h" / "2 hrs" → hours
  if (m) { const n = parseFloat(m[1]); return isFinite(n) && n > 0 ? n : null; }
  return null;
}

// Classify a request "type" string into a flag family it can clear.
//   'weekend' → clears weekend flags;  'hours' → clears over-hours flags;
//   '' (generic/unknown) → clears both.  Window flags are never auto-cleared.
function classifyApprovalType(s) {
  const t = String(s || '').toLowerCase();
  if (!t) return '';
  if (t.includes('weekend')) return 'weekend';
  if (/(hour|overtime|\bot\b|additional|extra)/.test(t)) return 'hours';
  return '';
}

// Tokenize a name → lowercased word tokens, punctuation stripped.
function nameTokens(s) {
  return String(s || '').toLowerCase()
    .replace(/[^a-z\s'-]/g, ' ').replace(/['-]/g, '')
    .split(/\s+/).filter(Boolean);
}

// Tolerant name match between a Time Doctor display name (often first-name or
// first + last-initial, e.g. "Jose", "Art O", "Leslie B") and the request
// tab's full legal name ("Jose Membreno", "Art Oaing"). Deliberately
// conservative: first name must match, and for multi-token names the last
// token must be compatible (equal, an initial, or a prefix). This errs toward
// NOT matching when unsure — a missed clear just leaves an item flagged for
// manual review (safe), whereas a wrong clear would hide unapproved work.
function tolerantNameMatch(tdName, tabName) {
  const A = nameTokens(tdName), B = nameTokens(tabName);
  if (!A.length || !B.length) return false;
  if (A[0] !== B[0]) return false;                 // first name must match
  if (A.length === 1 || B.length === 1) return true;
  const al = A[A.length - 1], bl = B[B.length - 1];
  if (al === bl) return true;
  if (al.length === 1 && bl[0] === al) return true;   // "B" vs "Bernolo"
  if (bl.length === 1 && al[0] === bl) return true;
  if (al.startsWith(bl) || bl.startsWith(al)) return true;  // "O" vs "Oaing"
  return false;
}

// Build an approval index from a header array + data row arrays.
// Returns { byEmail, byName, list, detected } or null if no usable columns.
//   byEmail/byName: Map(key -> Map(dateKey -> Set(family)))
//   list: approved entries for transparency display
function buildApprovalIndex(header, dataRows) {
  const hl = header.map(h => String(h || '').trim().toLowerCase());
  const iStatus = findColByTokens(hl, ['status', 'approval']);
  const iEmail  = findColByTokens(hl, ['email']);
  // Name: prefer an explicit "owner/employee/person" column over a generic
  // "...name" (several columns here share an "HR Requests Hub:" prefix).
  let iName = findColByTokens(hl, ['owner name', 'employee name', 'person name', 'requested by', 'staff name', 'owner', 'employee']);
  if (iName < 0) iName = findColByTokens(hl, ['name', 'person', 'staff']);
  // Type: 'type'/'category' only — NOT 'request', which would falsely match
  // the shared "HR Requests Hub:" prefix on other columns.
  const iType   = findColByTokens(hl, ['type', 'category', 'reason']);
  const iDate   = findColByTokens(hl, ['date', 'day']);
  // Requested-hours column — the # of additional hours the person asked for.
  // Approved excess is capped at this. Tries specific labels first, then a
  // bare "hours" as a fallback. Guarded below so it can't collide.
  let iReq = findColByTokens(hl, ['hours requested', 'requested hours', 'hours approved',
                                   'approved hours', 'additional hours', '# of hours',
                                   'number of hours', 'no. of hours', 'extra hours']);
  if (iReq < 0) iReq = findColByTokens(hl, ['hours', 'duration', 'qty', 'quantity']);
  // an end-date column for ranges (must differ from the start-date col)
  let iEnd = -1;
  for (let i = 0; i < hl.length; i++) {
    if (i !== iDate && /(end|to)\b/.test(hl[i]) && hl[i].includes('date')) { iEnd = i; break; }
  }
  // Defensive: type must not collapse onto another detected column.
  let iTypeSafe = iType;
  if (iTypeSafe === iName || iTypeSafe === iStatus || iTypeSafe === iDate || iTypeSafe === iEmail) iTypeSafe = -1;
  const iTypeFinal = iTypeSafe;
  // Defensive: requested-hours must not collapse onto another detected column.
  let iReqFinal = iReq;
  if (iReqFinal === iName || iReqFinal === iStatus || iReqFinal === iDate ||
      iReqFinal === iEmail || iReqFinal === iEnd || iReqFinal === iTypeFinal) iReqFinal = -1;
  if (iStatus < 0 || iDate < 0 || (iEmail < 0 && iName < 0)) {
    return { byEmail: new Map(), byName: new Map(), entries: [], list: [],
             detected: { ok: false, iStatus, iEmail, iName, iType: iTypeFinal, iDate, iEnd, iReq: iReqFinal } };
  }

  const byEmail = new Map(), byName = new Map(), list = [], entries = [], pendingEntries = [];
  const add = (map, key, dateKey, family) => {
    if (!key) return;
    if (!map.has(key)) map.set(key, new Map());
    const m = map.get(key);
    if (!m.has(dateKey)) m.set(dateKey, new Set());
    m.get(dateKey).add(family);
  };

  for (const row of dataRows) {
    const decision = classifyApprovalStatus(row[iStatus]);
    if (decision === 'rejected') continue;             // rejected = no effect
    const startKey = parseApprovalDateKey(row[iDate]);
    if (startKey == null) continue;
    const endKey = iEnd >= 0 ? parseApprovalDateKey(row[iEnd]) : null;
    const family = classifyApprovalType(iTypeFinal >= 0 ? row[iTypeFinal] : '');
    const email = iEmail >= 0 ? String(row[iEmail] || '').trim().toLowerCase() : '';
    const rawName = iName >= 0 ? normalizeName(row[iName] || '') : '';
    const name = rawName.toLowerCase();

    // Expand a date range (inclusive, capped) into individual day keys.
    const keys = [startKey];
    if (endKey != null && endKey > startKey) {
      let y = Math.floor(startKey/10000), mo = Math.floor((startKey%10000)/100), da = startKey%100;
      const cur = new Date(y, mo-1, da); let guard = 0;
      while (guard++ < 120) {
        cur.setDate(cur.getDate()+1);
        const k = cur.getFullYear()*10000+(cur.getMonth()+1)*100+cur.getDate();
        if (k > endKey) break;
        keys.push(k);
      }
    }

    const entry = { email, rawName, dateKeys: new Set(keys), family,
                    reqHrs: iReqFinal >= 0 ? parseRequestedHours(row[iReqFinal]) : null };
    if (decision === 'approved') {
      for (const k of keys) {
        if (email) add(byEmail, email, k, family);
        if (name)  add(byName, name, k, family);
      }
      entries.push(entry);
      list.push({ email, name: rawName, type: iTypeFinal >= 0 ? String(row[iTypeFinal]||'').trim() : '',
                  dateKey: startKey, endKey, reqHrs: entry.reqHrs });
    } else {
      // pending — does NOT suppress weekend flags or clear over-hours, but
      // lets the billable calc hold the excess as "pending approval" rather
      // than removing it.
      pendingEntries.push(entry);
    }
  }
  return { byEmail, byName, entries, pendingEntries, list,
           detected: { ok: true, iStatus, iEmail, iName, iType: iTypeFinal, iDate, iEnd, iReq: iReqFinal,
                       headers: { status: header[iStatus], date: header[iDate],
                                  email: iEmail>=0?header[iEmail]:null, name: iName>=0?header[iName]:null,
                                  type: iTypeFinal>=0?header[iTypeFinal]:null, end: iEnd>=0?header[iEnd]:null,
                                  req: iReqFinal>=0?header[iReqFinal]:null } } };
}

// Adapter: Apps Script returns each tab as an array of objects keyed by header.
function buildApprovalIndexFromObjects(objRows) {
  if (!Array.isArray(objRows) || objRows.length === 0)
    return buildApprovalIndex([], []);
  const header = Object.keys(objRows[0]);
  const dataRows = objRows.map(o => header.map(h => o[h]));
  return buildApprovalIndex(header, dataRows);
}
// Adapter: an uploaded approvals CSV parsed into row arrays.
function buildApprovalIndexFromCSV(text) {
  const rows = parsePayrollCSV(text);
  if (rows.length < 2) return buildApprovalIndex([], []);
  return buildApprovalIndex(rows[0], rows.slice(1));
}

// Is a person+date approved for a given flag family? Scans approved entries,
// matching on email when available, else a tolerant name match. family is
// 'weekend' or 'hours'; a generic ('') approval clears either.
function approvalClears(approvals, email, personName, dateKey, family) {
  if (!approvals || !approvals.entries) return false;
  const em = (email || '').toLowerCase();
  for (const e of approvals.entries) {
    if (!e.dateKeys.has(dateKey)) continue;
    if (!(e.family === '' || e.family === family)) continue;
    const emailMatch = em && e.email && em === e.email;
    const nameMatch = !emailMatch && e.rawName && tolerantNameMatch(personName, e.rawName);
    if (emailMatch || nameMatch) return true;
  }
  return false;
}

// Three-state lookup for the daily billable cap: is there an APPROVED hours
// request, a PENDING (submitted-not-approved) one, or NONE for this person+date?
// Only 'hours' / generic-family requests count (weekend approvals don't lift
// the daily cap). Approved beats pending. For approved matches, also returns
// the total requested hours (reqHrs) so the caller can cap the included excess
// at what was actually requested; reqHrs is null when no/blank hours figure
// was on the request (→ treated as "no stated limit", include full excess).
function hoursRequestStatus(approvals, email, personName, dateKey) {
  if (!approvals) return { status: 'none', reqHrs: null };
  const em = (email || '').toLowerCase();
  const matches = (e) => {
    if (!e.dateKeys.has(dateKey)) return false;
    if (!(e.family === '' || e.family === 'hours')) return false;
    const emailMatch = em && e.email && em === e.email;
    if (emailMatch) return true;
    return e.rawName && tolerantNameMatch(personName, e.rawName);
  };
  const appr = (approvals.entries || []).filter(matches);
  if (appr.length) {
    // Sum requested hours across matching approved requests; if ANY lacks a
    // stated figure, treat the whole day as unlimited (null) — we won't
    // enforce a cap we can't read.
    let total = 0, anyNull = false;
    for (const e of appr) { if (e.reqHrs == null) anyNull = true; else total += e.reqHrs; }
    return { status: 'approved', reqHrs: anyNull ? null : total };
  }
  if ((approvals.pendingEntries || []).some(matches)) return { status: 'pending', reqHrs: null };
  return { status: 'none', reqHrs: null };
}

// Display lookup: total additional hours requested for a person on a date,
// across matching hours-family requests (approved OR pending). Returns a
// number of hours, or null when no request matches (→ shown as "—"). A
// matching request with a blank hours figure contributes 0.
function requestedHoursForDay(approvals, email, personName, dateKey) {
  if (!approvals) return null;
  const em = (email || '').toLowerCase();
  const matches = (e) => {
    if (!e.dateKeys.has(dateKey)) return false;
    if (!(e.family === '' || e.family === 'hours')) return false;
    const emailMatch = em && e.email && em === e.email;
    if (emailMatch) return true;
    return e.rawName && tolerantNameMatch(personName, e.rawName);
  };
  let total = 0, found = false;
  for (const e of (approvals.entries || []))        if (matches(e)) { found = true; if (e.reqHrs != null) total += e.reqHrs; }
  for (const e of (approvals.pendingEntries || [])) if (matches(e)) { found = true; if (e.reqHrs != null) total += e.reqHrs; }
  return found ? total : null;
}

/* ------------------------------------------------------------
   Billable-hours model (Anshul, Jun 2026)
   ----------------------------------------------------------------
   "Time tracked" INCLUDES the paid break, so actual work = tracked − paid break.

   Rule 1 — Paid break & excess (relaxed Anshul Jun 2026):
     The 30-min paid break is always earned, up to the 30-min allowance —
     there is NO full-day work requirement. Any paid break logged BEYOND
     30 min is "excess" and is removed from billable immediately (e.g. 8h40m
     tracked with a 40m break → 10m excess removed → 8h30m billable).
     "Break removed" = paid break that didn't count toward billable, which
     under this relaxed rule means only the excess over 30 min.

   Rule 2 — Daily cap + additional-hours gate (weekdays):
     Billable is capped at 8h30m/day (8h work + 30m paid break). Excess over
     the cap is gated on an Additional Hours Request:
       approved  → excess included (billable = full earned, pending = 0)
       submitted → excess held     (billable = 8.5h, pending = excess)
       none/rej. → excess removed   (billable = 8.5h, pending = 0)

   Weekends — same pipeline, different values:
     Weekend work must be REQUESTED. With an approved weekend request the day
     is billable up to the weekend cap (a full weekend day = an additional
     8h); without one, the day is flagged and contributes 0 billable. Weekend
     constants below are independent so they can diverge from weekdays.

   Total tracked always reflects raw Time tracked regardless.
   ------------------------------------------------------------ */
const WORK_FULL_DAY_HRS   = 8.0;   // per-day work cap (no longer gates the paid break)
const PAID_BREAK_MAX_HRS  = 0.5;   // 30-min paid break, capped
const DAILY_BILLABLE_CAP  = WORK_FULL_DAY_HRS + PAID_BREAK_MAX_HRS;  // 8.5h
const HRS_EPS = 1e-6;              // float tolerance so 8.5 isn't "over 8.5"

// Weekend model — full weekend day = an additional 8h, and weekend work is
// only billable with an approved weekend request. Defaults: no paid break on
// weekends (set WEEKEND_PAID_BREAK_MAX_HRS = 0.5 to mirror weekdays), cap 8h.
const WEEKEND_WORK_FULL_DAY_HRS  = 8.0;
const WEEKEND_PAID_BREAK_MAX_HRS = 0.0;
const WEEKEND_BILLABLE_CAP       = 8.0;

const WEEKDAY_CFG = { full: WORK_FULL_DAY_HRS,         brkMax: PAID_BREAK_MAX_HRS,         cap: DAILY_BILLABLE_CAP };
const WEEKEND_CFG = { full: WEEKEND_WORK_FULL_DAY_HRS, brkMax: WEEKEND_PAID_BREAK_MAX_HRS, cap: WEEKEND_BILLABLE_CAP };

// Compute the pre-cap billable for one day from raw tracked + paid-break hours.
//   tt        = Time tracked (INCLUDES paid break)
//   paidBrk   = paid break time, in hours
//   cfg       = { full, brkMax, cap }  (WEEKDAY_CFG or WEEKEND_CFG)
// Returns { work, earnedBreak, billableUncapped, breakStripped, excessBreak, breakRemoved }.
//   breakStripped = paid break lost because work < full day (not earned)
//   excessBreak   = paid break beyond the brkMax allowance (over the limit)
//   breakRemoved  = total paid break NOT billable = paidBrk − earnedBreak
function computeBillableDay(tt, paidBrk, cfg) {
  const work = Math.max(0, tt - paidBrk);
  // Relaxed rule (Anshul Jun 2026): the paid break is always earned up to the
  // allowance, regardless of whether work reached a full day. The full-day gate
  // was removed; cfg.brkMax = 0 (e.g. weekends) still yields no paid break.
  const earnedBreak = Math.min(paidBrk, cfg.brkMax);
  return { work, earnedBreak, billableUncapped: work + earnedBreak,
           breakStripped: 0, excessBreak: Math.max(0, paidBrk - cfg.brkMax),
           breakRemoved: paidBrk - earnedBreak };
}

function analyzePayrollCustom(rows, opts, approvals) {
  const header = rows[0];
  const col = {};
  header.forEach((h, i) => { col[(h || '').trim().toLowerCase()] = i; });
  const need = ['name', 'date', 'time tracked'];
  for (const n of need) if (!(n in col)) throw new Error(`Custom Export is missing the "${n}" column.`);

  const get = (row, name) => { const i = col[name]; return i == null ? '' : row[i]; };
  const num = (row, name) => parsePayrollHours(get(row, name));

  const dailyLimit = opts.dailyLimit;
  const hourThresh = dailyLimit + (opts.graceMin || 0) / 60;
  const breakLimitMin = opts.breakLimitMin;
  const breakThreshMin = breakLimitMin + (opts.breakGraceMin || 0);
  const winStart = opts.winStartMin, winEnd = opts.winEndMin, winGrace = opts.winGraceMin || 0;
  const flagWeekend = opts.flagWeekend;

  // Group rows by person.
  const byPerson = new Map();
  let minSort = Infinity, maxSort = -Infinity, minLabel = '', maxLabel = '';
  for (let r = 1; r < rows.length; r++) {
    const row = rows[r];
    const name = (get(row, 'name') || '').trim();
    if (!name) continue;
    const dt = parsePayrollDate(get(row, 'date'));
    if (!dt) continue;
    if (dt.sort < minSort) { minSort = dt.sort; minLabel = dt.label; }
    if (dt.sort > maxSort) { maxSort = dt.sort; maxLabel = dt.label; }
    if (!byPerson.has(name)) {
      byPerson.set(name, { name, email: (get(row, 'email') || '').trim(),
                           group: (get(row, 'user groups') || '').trim(), days: [] });
    }
    byPerson.get(name).days.push({
      dt,
      tt: num(row, 'time tracked'),
      paidBrkHrs: num(row, 'paid break time'),
      brkMin: (num(row, 'paid break time') + num(row, 'unpaid break time')) * 60,
      start: get(row, 'start time'),
      end: get(row, 'end time'),
    });
  }

  const flagged = [], clean = [];
  const suppressed = [];          // approved items we cleared, for transparency
  for (const p of byPerson.values()) {
    const weekendDays = [], breakDays = [], windowDays = [];
    const pendingDays = [];       // submitted-not-approved over-cap → excess held as pending
    const droppedDays = [];       // over cap, no/rejected request → excess auto-removed
    const breakAdjDays = [];      // paid break stripped because work < full day (informational)
    const excessBreakDays = [];   // paid break over the 30m allowance, removed (informational)
    const approvedOverDays = [];  // over cap but approved → included (informational)
    const overReqDays = [];       // approved but logged MORE over-cap than requested → surplus removed
    const weekendBillDays = [];   // approved weekend days that earned billable hours
    const dayRows = [];           // one row per worked day → columnar detail table
    let weekendHrs = 0;           // weekend hours flagged (unrequested) — not billable
    let totalTracked = 0, totalBillable = 0, totalPending = 0, totalBreakRemoved = 0, totalPaidBreak = 0, totalRequested = 0;
    let worstBillableHrs = 0;
    const nameLower = (p.name || '').toLowerCase();

    for (const d of p.days) {
      if (d.tt <= 0) continue;       // not a worked day
      totalTracked += d.tt;
      totalPaidBreak += d.paidBrkHrs;

      // ===================== WEEKEND =====================
      if (d.dt.isWeekend) {
        if (!flagWeekend) continue;  // weekend checks disabled
        const approved = approvalClears(approvals, p.email, nameLower, d.dt.sort, 'weekend');
        if (!approved) {
          // No approved weekend request → flagged, contributes 0 billable.
          weekendDays.push({ label: d.dt.label, dayName: d.dt.dayName, hrs: d.tt });
          weekendHrs += d.tt;
          const wreq = requestedHoursForDay(approvals, p.email, nameLower, d.dt.sort);
          if (wreq != null) totalRequested += wreq;
          dayRows.push({ label: d.dt.label, dayName: d.dt.dayName, weekend: true,
                         tracked: d.tt, paidBrkMin: d.paidBrkHrs*60, work: Math.max(0, d.tt - d.paidBrkHrs),
                         billable: 0, pending: 0, removedMin: 0, reqHrs: wreq, flags: ['W'],
                         notes: ['weekend — not requested'], note_cls: 'warn' });
          continue;
        }
        // Approved weekend work → billable under the weekend model.
        const wb = computeBillableDay(d.tt, d.paidBrkHrs, WEEKEND_CFG);
        let wbill = wb.billableUncapped;
        let wremoved = 0;
        const wover = wbill - WEEKEND_CFG.cap;
        if (wover > HRS_EPS) { wremoved = wover; wbill = WEEKEND_CFG.cap; }  // over weekend cap → removed
        totalBillable += wbill;
        totalBreakRemoved += wb.breakRemoved;
        if (wbill > worstBillableHrs) worstBillableHrs = wbill;
        weekendBillDays.push({ label: d.dt.label, dayName: d.dt.dayName, tt: d.tt,
                               billable: wbill, breakRemoved: wb.breakRemoved, capRemoved: wremoved });
        suppressed.push({ name: p.name, email: p.email, label: d.dt.label, dayName: d.dt.dayName,
                          kind: 'Weekend work (approved)', detail: `billable ${wbill.toFixed(2)}h` });
        const wn = ['weekend — approved'];
        if (wover > HRS_EPS) wn.push(`over ${WEEKEND_BILLABLE_CAP}h cap`);
        const wreq2 = requestedHoursForDay(approvals, p.email, nameLower, d.dt.sort);
        if (wreq2 != null) totalRequested += wreq2;
        dayRows.push({ label: d.dt.label, dayName: d.dt.dayName, weekend: true,
                       tracked: d.tt, paidBrkMin: d.paidBrkHrs*60, work: wb.work,
                       billable: wbill, pending: 0, removedMin: wb.breakRemoved*60, reqHrs: wreq2, flags: ['Wk'],
                       notes: wn, note_cls: 'ok' });
        continue;
      }

      // ===================== WEEKDAY =====================
      // ---- Rule 1: paid-break eligibility + excess → pre-cap billable ----
      const b = computeBillableDay(d.tt, d.paidBrkHrs, WEEKDAY_CFG);
      let billable = b.billableUncapped;
      let pending = 0;
      const notes = [];
      const dflags = [];
      let noteCls = '';
      totalBreakRemoved += b.breakRemoved;
      if (b.breakStripped > HRS_EPS) {
        breakAdjDays.push({ label: d.dt.label, dayName: d.dt.dayName,
                            tt: d.tt, work: b.work, stripped: b.breakStripped, billable });
        notes.push(`break not earned (work < ${WORK_FULL_DAY_HRS}h)`);
        noteCls = noteCls || 'warn';
      } else if (b.excessBreak > HRS_EPS) {
        excessBreakDays.push({ label: d.dt.label, dayName: d.dt.dayName,
                               tt: d.tt, paidBrk: d.paidBrkHrs, excess: b.excessBreak, billable });
        notes.push(`excess break −${Math.round(b.excessBreak*60)}m`);
        noteCls = noteCls || 'warn';
      }

      // ---- Rule 2: daily cap + additional-hours gate --------------------
      const over = billable - DAILY_BILLABLE_CAP;
      if (over > HRS_EPS) {
        const req = hoursRequestStatus(approvals, p.email, nameLower, d.dt.sort);
        const status = req.status;
        if (status === 'approved') {
          // Include the excess, but only up to the # of hours requested.
          // reqHrs == null → no stated limit on the request → include all.
          const allowed = (req.reqHrs == null) ? over : Math.max(0, req.reqHrs);
          const included = Math.min(over, allowed);
          const beyond = over - included;            // logged beyond what was requested
          billable = DAILY_BILLABLE_CAP + included;
          approvedOverDays.push({ label: d.dt.label, dayName: d.dt.dayName, tt: d.tt,
                                  billable, included, reqHrs: req.reqHrs, beyond });
          if (beyond > HRS_EPS) {
            // Logged more over-cap hours than were requested → trim the surplus.
            overReqDays.push({ label: d.dt.label, dayName: d.dt.dayName, tt: d.tt,
                               billable, reqHrs: req.reqHrs, logged: over, beyond });
            notes.push(`over ${DAILY_BILLABLE_CAP}h — approved ${req.reqHrs}h, ${beyond.toFixed(2)}h beyond request removed`);
            dflags.push('R');
            noteCls = 'flag';
          } else {
            suppressed.push({ name: p.name, email: p.email, label: d.dt.label, dayName: d.dt.dayName,
                              kind: `Over ${DAILY_BILLABLE_CAP}h (approved)`, detail: `billable ${billable.toFixed(2)}h` });
            notes.push(req.reqHrs == null
              ? `over ${DAILY_BILLABLE_CAP}h — approved`
              : `over ${DAILY_BILLABLE_CAP}h — approved ${req.reqHrs}h`);
            noteCls = 'ok';
          }
        } else if (status === 'pending') {
          pending = over;
          billable = DAILY_BILLABLE_CAP;
          pendingDays.push({ label: d.dt.label, dayName: d.dt.dayName, tt: d.tt,
                             billable, pending, earned: b.billableUncapped });
          notes.push(`over ${DAILY_BILLABLE_CAP}h — pending`);
          dflags.push('P');
          noteCls = 'pending';
        } else {
          billable = DAILY_BILLABLE_CAP;   // excess auto-removed
          droppedDays.push({ label: d.dt.label, dayName: d.dt.dayName, tt: d.tt,
                             billable, removed: over, earned: b.billableUncapped });
          notes.push(`over ${DAILY_BILLABLE_CAP}h — removed`);
          dflags.push('H');
          noteCls = 'flag';
        }
      }

      totalBillable += billable;
      totalPending  += pending;
      if (billable > worstBillableHrs) worstBillableHrs = billable;

      // ---- Independent controls -----------------------------------------
      if (d.brkMin > breakThreshMin) {
        breakDays.push({ label: d.dt.label, dayName: d.dt.dayName, brkMin: d.brkMin, overMin: d.brkMin - breakLimitMin });
        notes.push(`break ${Math.round(d.brkMin)}m`);
        dflags.push('B');
        noteCls = noteCls || 'flag';
      }
      const sMin = parseClock(d.start), eMin = parseClock(d.end);
      const early = sMin != null && sMin < winStart - winGrace;
      const late = eMin != null && eMin > winEnd + winGrace;
      if (early || late) {
        windowDays.push({ label: d.dt.label, dayName: d.dt.dayName, start: clockStr(d.start), end: clockStr(d.end), early, late });
        notes.push(`outside window (${clockStr(d.start)}–${clockStr(d.end)})`);
        dflags.push('O');
        noteCls = noteCls || 'flag';
      }

      const reqHrsDay = requestedHoursForDay(approvals, p.email, nameLower, d.dt.sort);
      if (reqHrsDay != null) totalRequested += reqHrsDay;
      dayRows.push({ label: d.dt.label, dayName: d.dt.dayName, weekend: false,
                     tracked: d.tt, paidBrkMin: d.paidBrkHrs*60, work: b.work,
                     billable, pending, removedMin: b.breakRemoved*60, reqHrs: reqHrsDay, flags: dflags,
                     notes, note_cls: noteCls });
    }

    const rec = { name: p.name, email: p.email, group: p.group,
                  totalTracked, totalBillable, totalPending, totalBreakRemoved, totalPaidBreak, totalRequested, worstBillableHrs,
                  weekendDays, breakDays, windowDays, dayRows,
                  pendingDays, droppedDays, breakAdjDays, excessBreakDays,
                  approvedOverDays, overReqDays, weekendBillDays, weekendHrs,
                  // flagCount = items that need Lucia's attention. breakAdj,
                  // excessBreak, approvedOver, weekendBill are informational.
                  flagCount: droppedDays.length + pendingDays.length + overReqDays.length
                           + breakDays.length + windowDays.length + weekendDays.length };
    if (rec.flagCount > 0) flagged.push(rec); else clean.push(rec);
  }

  flagged.sort((a, b) =>
    (b.totalPending - a.totalPending) ||
    ((b.droppedDays.length + b.weekendHrs) - (a.droppedDays.length + a.weekendHrs)) ||
    (b.flagCount - a.flagCount));
  clean.sort((a, b) => b.totalBillable - a.totalBillable);
  suppressed.sort((a, b) => a.name.localeCompare(b.name));

  const dateRange = minLabel && maxLabel ? `${minLabel} – ${maxLabel}` : '';
  return { format: 'custom', flagged, clean, suppressed, dateRange };
}

// Render the billable-hours + controls custom-export results.
function renderPayrollCustom(analysis, meta) {
  const wrap = document.getElementById('payroll-summary');
  const resultsSection = document.getElementById('payroll-results');
  const { flagged, clean } = analysis;
  const total = flagged.length + clean.length;

  const win = `${meta.winStartStr}–${meta.winEndStr}`;
  const cap = DAILY_BILLABLE_CAP;
  let html = `<div class="period-banner">`
           + `<span>Report: <strong>${escapeHTML(analysis.dateRange || meta.fileName)}</strong></span>`
           + `<span>${cap}h/day billable cap · paid break ${Math.round(PAID_BREAK_MAX_HRS*60)}m · window ${win} CT · ${flagged.length} of ${total} flagged</span>`
           + `</div>`;

  // Legend.
  html += `<div class="payroll-flag-legend">`
        + `<span class="pf pf-hours">H · over ${cap}h, no request (excess removed)</span>`
        + `<span class="pf pf-overreq">R · approved, but logged beyond requested (surplus removed)</span>`
        + `<span class="pf pf-pending">P · over ${cap}h, pending approval</span>`
        + `<span class="pf pf-break">B · break &gt; ${meta.breakLimitMin}m</span>`
        + `<span class="pf pf-window">O · outside ${win}</span>`
        + `<span class="pf pf-weekend">W · weekend</span>`
        + `</div>`;

  // Approvals status + the items they cleared (transparency: cleared ≠ hidden).
  if (meta.approvalsNote) {
    const cls = meta.approvalsOk ? 'payroll-appr-note ok' : 'payroll-appr-note warn';
    html += `<div class="${cls}">${escapeHTML(meta.approvalsNote)}</div>`;
  }
  const sup = analysis.suppressed || [];
  if (sup.length) {
    html += `<details class="payroll-suppressed"><summary>✓ ${sup.length} item${sup.length===1?'':'s'} cleared by approved requests</summary>`
          + `<div class="payroll-detail">`;
    for (const s of sup) {
      html += `<div class="payroll-detail-row"><span>${escapeHTML(s.name)} · ${escapeHTML(s.dayName)} ${escapeHTML(s.label)} <em class="muted">${escapeHTML(s.kind)}</em></span><span class="mono">${escapeHTML(s.detail)}</span></div>`;
    }
    html += `</div></details>`;
  }

  // ONE table. Each person is a clickable header row showing their totals;
  // their day rows + a subtotal row live in the SAME table (same columns,
  // aligned), hidden until the person row is clicked.
  const fmtMin = (m) => { const r = Math.round(m); return r ? `${r}m` : '—'; };
  const dayFlagChip = (code) => {
    const map = { H: 'pf-hours', R: 'pf-overreq', P: 'pf-pending', B: 'pf-break', O: 'pf-window', W: 'pf-weekend', Wk: 'pf-wknd-ok' };
    return `<span class="pf ${map[code] || ''}">${code}</span>`;
  };

  const all = flagged.concat(clean);
  const fmtReq = (v) => (v != null && v > HRS_EPS) ? fmtHrs(v) : '—';
  let html2 = `<table class="detail-table payroll-table payroll-billable-table"><thead><tr>`
            + `<th>Name / Day</th><th>Flags</th>`
            + `<th class="num">Tracked</th><th class="num">Break</th>`
            + `<th class="num">Excess Brk</th><th class="num">Requested</th><th class="num">Pending</th>`
            + `<th class="num">Billable</th>`
            + `</tr></thead><tbody>`;

  all.forEach((p, idx) => {
    const chips = [];
    if (p.droppedDays.length)  chips.push(`<span class="pf pf-hours">H${p.droppedDays.length}</span>`);
    if (p.overReqDays.length)  chips.push(`<span class="pf pf-overreq">R${p.overReqDays.length}</span>`);
    if (p.pendingDays.length)  chips.push(`<span class="pf pf-pending">P${p.pendingDays.length}</span>`);
    if (p.breakDays.length)    chips.push(`<span class="pf pf-break">B${p.breakDays.length}</span>`);
    if (p.windowDays.length)   chips.push(`<span class="pf pf-window">O${p.windowDays.length}</span>`);
    if (p.weekendDays.length)  chips.push(`<span class="pf pf-weekend">W${p.weekendDays.length}</span>`);
    const groupTag = p.group ? `<span class="payroll-group-tag">${escapeHTML(p.group)}</span>` : '';
    const pendingCell = p.totalPending > HRS_EPS
      ? `<span class="pf pf-pending">${fmtHrs(p.totalPending)}</span>` : '—';

    // Person header row.
    html2 += `<tr class="payroll-person-row${p.flagCount ? ' flag-review' : ''}" data-pid="${idx}" tabindex="0" role="button" aria-expanded="false">`
           + `<td><span class="payroll-toggle">▸</span> <strong>${escapeHTML(p.name)}</strong>${groupTag}`
           + `<span class="payroll-email">${escapeHTML(p.email)}</span></td>`
           + `<td>${chips.join(' ') || '<span class="muted">—</span>'}</td>`
           + `<td class="num">${fmtHrs(p.totalTracked)}</td>`
           + `<td class="num">${fmtMin((p.totalPaidBreak || 0) * 60)}</td>`
           + `<td class="num">${fmtMin((p.totalBreakRemoved || 0) * 60)}</td>`
           + `<td class="num">${fmtReq(p.totalRequested)}</td>`
           + `<td class="num">${pendingCell}</td>`
           + `<td class="num"><strong>${fmtHrs(p.totalBillable)}</strong></td>`
           + `</tr>`;

    // Day rows (hidden until expanded).
    for (const d of (p.dayRows || [])) {
      const dchips = (d.flags && d.flags.length) ? d.flags.map(dayFlagChip).join(' ') : '<span class="muted">—</span>';
      html2 += `<tr class="payroll-child pid-${idx}${d.weekend ? ' payroll-day-weekend' : ''}" hidden>`
             + `<td class="payroll-day-cell">${escapeHTML(d.dayName)} ${escapeHTML(d.label)}</td>`
             + `<td>${dchips}</td>`
             + `<td class="num">${fmtHrs(d.tracked)}</td>`
             + `<td class="num">${fmtMin(d.paidBrkMin)}</td>`
             + `<td class="num">${fmtMin(d.removedMin)}</td>`
             + `<td class="num">${fmtReq(d.reqHrs)}</td>`
             + `<td class="num">${d.pending > HRS_EPS ? fmtHrs(d.pending) : '—'}</td>`
             + `<td class="num"><strong>${fmtHrs(d.billable)}</strong></td>`
             + `</tr>`;
    }
    // Subtotal row (hidden until expanded) — echoes the person totals.
    html2 += `<tr class="payroll-child payroll-subtotal pid-${idx}" hidden>`
           + `<td><strong>Subtotal — ${escapeHTML(p.name)}</strong></td><td></td>`
           + `<td class="num">${fmtHrs(p.totalTracked)}</td>`
           + `<td class="num">${fmtMin((p.totalPaidBreak || 0) * 60)}</td>`
           + `<td class="num">${fmtMin((p.totalBreakRemoved || 0) * 60)}</td>`
           + `<td class="num">${fmtReq(p.totalRequested)}</td>`
           + `<td class="num">${p.totalPending > HRS_EPS ? fmtHrs(p.totalPending) : '—'}</td>`
           + `<td class="num"><strong>${fmtHrs(p.totalBillable)}</strong></td>`
           + `</tr>`;
  });

  // Grand total row, inside the same table.
  const sum = (k) => all.reduce((a, p) => a + (p[k] || 0), 0);
  html2 += `<tr class="payroll-grand-row">`
         + `<td><strong>GRAND TOTAL</strong> <span class="muted">— ${all.length} ${all.length === 1 ? 'person' : 'people'}</span></td><td></td>`
         + `<td class="num">${fmtHrs(sum('totalTracked'))}</td>`
         + `<td class="num">${fmtMin(sum('totalPaidBreak') * 60)}</td>`
         + `<td class="num">${fmtMin(sum('totalBreakRemoved') * 60)}</td>`
         + `<td class="num">${fmtReq(sum('totalRequested'))}</td>`
         + `<td class="num">${sum('totalPending') > HRS_EPS ? fmtHrs(sum('totalPending')) : '—'}</td>`
         + `<td class="num"><strong>${fmtHrs(sum('totalBillable'))}</strong></td>`
         + `</tr>`;
  html2 += `</tbody></table>`;
  html += html2;

  wrap.innerHTML = html;
  resultsSection.classList.remove('hidden');

  // Wire up expand/collapse: clicking a person row toggles its child rows.
  wrap.querySelectorAll('.payroll-person-row').forEach((row) => {
    const toggle = () => {
      const pid = row.getAttribute('data-pid');
      const open = row.getAttribute('aria-expanded') === 'true';
      row.setAttribute('aria-expanded', open ? 'false' : 'true');
      const tog = row.querySelector('.payroll-toggle');
      if (tog) tog.textContent = open ? '▸' : '▾';
      wrap.querySelectorAll('.pid-' + pid).forEach((r) => { r.hidden = open; });
    };
    row.addEventListener('click', toggle);
    row.addEventListener('keydown', (e) => {
      if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); toggle(); }
    });
  });
}

function downloadPayrollCustomReport(analysis, meta) {
  const win = `${meta.winStartStr}-${meta.winEndStr}`;
  const cap = DAILY_BILLABLE_CAP;
  const all = analysis.flagged.concat(analysis.clean || []);

  // Sheet 1 content: per-person summary (tracked / billable / pending / break removed).
  const lines = [['Name', 'Email', 'User group', 'Total Tracked (h)', 'Total Billable (h)', 'Excess Break Removed (min)', 'Additional Hrs Requested (h)', 'Pending Approval (h)', 'Flags'].join(',')];
  for (const p of all) {
    const flags = [
      p.droppedDays.length ? `H${p.droppedDays.length}` : '',
      p.overReqDays.length ? `R${p.overReqDays.length}` : '',
      p.pendingDays.length ? `P${p.pendingDays.length}` : '',
      p.breakDays.length ? `B${p.breakDays.length}` : '',
      p.windowDays.length ? `O${p.windowDays.length}` : '',
      p.weekendDays.length ? `W${p.weekendDays.length}` : '',
    ].filter(Boolean).join(' ');
    lines.push([
      csvEscape(p.name), csvEscape(p.email), csvEscape(p.group),
      p.totalTracked.toFixed(2), p.totalBillable.toFixed(2),
      Math.round(p.totalBreakRemoved * 60), (p.totalRequested || 0).toFixed(2),
      p.totalPending.toFixed(2),
      csvEscape(flags),
    ].join(','));
  }

  // Sheet 2 content: per-day exception detail, appended after a blank line.
  lines.push('');
  lines.push(['Name', 'Email', 'Date', 'Day', 'Exception', 'Detail'].join(','));
  const push = (p, label, dayName, flag, detail) =>
    lines.push([csvEscape(p.name), csvEscape(p.email), csvEscape(label), csvEscape(dayName), csvEscape(flag), csvEscape(detail)].join(','));
  for (const p of analysis.flagged.concat(analysis.clean || [])) {
    for (const d of p.droppedDays) push(p, d.label, d.dayName, `Over ${cap}h (no request, removed)`, `tracked ${d.tt.toFixed(2)}h, earned ${d.earned.toFixed(2)}h, billable ${d.billable.toFixed(2)}h, removed ${d.removed.toFixed(2)}h`);
    for (const d of p.pendingDays) push(p, d.label, d.dayName, `Over ${cap}h (pending approval)`, `tracked ${d.tt.toFixed(2)}h, billable ${d.billable.toFixed(2)}h, pending ${d.pending.toFixed(2)}h`);
    for (const d of p.approvedOverDays) push(p, d.label, d.dayName, `Over ${cap}h (approved${d.reqHrs!=null?` ${d.reqHrs}h`:''})`, `tracked ${d.tt.toFixed(2)}h, billable ${d.billable.toFixed(2)}h${d.beyond>1e-6?`, ${d.beyond.toFixed(2)}h beyond request removed`:''}`);
    for (const d of p.overReqDays) push(p, d.label, d.dayName, `Over ${cap}h (approved ${d.reqHrs}h, surplus removed)`, `tracked ${d.tt.toFixed(2)}h, logged over-cap ${d.logged.toFixed(2)}h, requested ${d.reqHrs}h, removed ${d.beyond.toFixed(2)}h, billable ${d.billable.toFixed(2)}h`);
    for (const d of p.excessBreakDays) push(p, d.label, d.dayName, `Excess paid break removed (over ${Math.round(PAID_BREAK_MAX_HRS*60)}m)`, `tracked ${d.tt.toFixed(2)}h, break ${Math.round(d.paidBrk*60)}m, removed ${Math.round(d.excess*60)}m, billable ${d.billable.toFixed(2)}h`);
    for (const d of p.breakAdjDays) push(p, d.label, d.dayName, `Paid break not earned (work < ${WORK_FULL_DAY_HRS}h)`, `tracked ${d.tt.toFixed(2)}h, work ${d.work.toFixed(2)}h, billable ${d.billable.toFixed(2)}h`);
    for (const d of p.weekendBillDays) push(p, d.label, d.dayName, `Weekend work (approved, billable)`, `tracked ${d.tt.toFixed(2)}h, billable ${d.billable.toFixed(2)}h${d.breakRemoved>1e-6?`, break removed ${Math.round(d.breakRemoved*60)}m`:''}${d.capRemoved>1e-6?`, over cap ${d.capRemoved.toFixed(2)}h`:''}`);
    for (const d of p.breakDays) push(p, d.label, d.dayName, `Break over ${meta.breakLimitMin}m`, `${Math.round(d.brkMin)}m (+${Math.round(d.overMin)}m)`);
    for (const d of p.windowDays) push(p, d.label, d.dayName, `Outside ${win} CT`, `${d.start}-${d.end}${d.early ? ' early' : ''}${d.late ? ' late' : ''}`);
    for (const d of p.weekendDays) push(p, d.label, d.dayName, 'Weekend work (unrequested, not billable)', `${d.hrs.toFixed(2)}h`);
  }
  const csv = lines.join('\r\n');
  const stamp = (analysis.dateRange || meta.fileName || 'report').replace(/[^A-Za-z0-9]+/g, '_');
  downloadFile(csv, `payroll_billable_${stamp}.csv`, 'text/csv;charset=utf-8;');
}

// Fetch the Additional Hours Request tab from the same Apps Script the
// commission side uses. Uses the &only= fast path so the server returns just
// this one tab instead of serializing the whole workbook (which was timing
// out). Tolerant: never throws — returns {ok, rows|reason} so a missing tab
// or offline sheet just means "nothing suppressed".
async function fetchApprovedHours() {
  if (!APPS_SCRIPT_URL || APPS_SCRIPT_URL.startsWith('PASTE_'))
    return { ok: false, reason: 'sheet not configured' };
  const TIMEOUT_MS = 20000;
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), TIMEOUT_MS);
  try {
    const url = `${APPS_SCRIPT_URL}?token=${encodeURIComponent(APPS_SCRIPT_TOKEN)}`
              + `&only=${encodeURIComponent(TAB_ADDITIONAL_HOURS)}`;
    const res = await fetch(url, { method: 'GET', signal: controller.signal });
    if (!res.ok) return { ok: false, reason: `HTTP ${res.status}` };
    const body = await res.json();
    if (body.error) return { ok: false, reason: body.error };
    const tab = body[TAB_ADDITIONAL_HOURS];
    if (!tab) return { ok: false, reason: `tab "${TAB_ADDITIONAL_HOURS}" not in sheet response` };
    if (tab.error) return { ok: false, reason: tab.error };
    return { ok: true, rows: tab };
  } catch (e) {
    return { ok: false, reason: e.name === 'AbortError' ? `timed out after ${TIMEOUT_MS/1000}s` : e.message };
  } finally {
    clearTimeout(timer);
  }
}

// Format-aware dispatchers used by the UI handler.
function runPayrollAnalysis(rows, meta, approvals) {
  const fmt = detectPayrollFormat(rows);
  if (fmt === 'custom') return analyzePayrollCustom(rows, meta, approvals);
  return analyzePayroll(rows, meta);   // legacy wide Hours Tracked (no approvals)
}
function renderPayroll(analysis, meta) {
  if (analysis.format === 'custom') return renderPayrollCustom(analysis, meta);
  return renderPayrollResults(analysis, meta);
}
function downloadPayroll(analysis, meta) {
  if (analysis.format === 'custom') return downloadPayrollCustomReport(analysis, meta);
  return downloadPayrollReport(analysis, meta);
}

document.addEventListener('DOMContentLoaded', () => {
  const periodDisplay = document.getElementById('period-display');
  const calcBtn = document.getElementById('calc-btn');
  const statusLine = document.getElementById('status-line');

  const period = previousMonthKey();
  lastPeriod = period;
  periodDisplay.textContent = fmtPeriodName(period);

  calcBtn.addEventListener('click', async () => {
    calcBtn.classList.add('loading');
    calcBtn.disabled = true;
    statusLine.className = '';
    statusLine.textContent = 'Fetching sheet data…';
    try {
      const data = await fetchSheetData();
      statusLine.textContent = 'Running calculations…';
      lastEntries = runCommissions(period, data);
      renderResults(lastEntries, period);

      // Calc is done — render is on screen. From here, the history
      // write is best-effort: a slow/failed write should not block
      // the user or roll back what they just saw. We do await it
      // (so the final status line is accurate), but on failure we
      // log the error and surface a soft warning, not a hard error.
      statusLine.className = 'success';
      statusLine.textContent = `Done. ${lastEntries.length} commission lines computed. Saving history…`;
      const hist = await writeCommissionHistoryAsync(period, lastEntries);
      if (hist.ok) {
        const replacedNote = hist.replaced > 0 ? ` (replaced ${hist.replaced} prior row${hist.replaced === 1 ? '' : 's'})` : '';
        statusLine.textContent = `Done. ${lastEntries.length} commission lines computed. History saved: ${hist.written} row${hist.written === 1 ? '' : 's'}${replacedNote}.`;
      } else {
        // Don't promote to error class — the calc itself succeeded.
        // Keep the success styling so it doesn't look like the dashboard
        // is broken; just note the soft failure for transparency.
        console.warn('Commission_History write failed:', hist.error);
        statusLine.textContent = `Done. ${lastEntries.length} commission lines computed. (History save failed — see console. Next Calculate will retry.)`;
      }

      // Refresh the trends panel even though it may be hidden behind
      // the Summary tab — cheap, and means switching tabs is instant.
      renderTrends();
    } catch (err) {
      console.error(err);
      statusLine.className = 'error';
      statusLine.textContent = `Error: ${err.message}`;
    } finally {
      calcBtn.classList.remove('loading');
      calcBtn.disabled = false;
    }
  });

  // Payment Mode toggle + Download Commission CSV. The toggle controls
  // which Payment Mode value filters the download; clicking a toggle
  // button doesn't trigger a download — Lucia clicks one of the
  // download buttons separately. The Payroll button is rendered but
  // disabled — see #download-payroll-btn in index.html — pending
  // definition of what a Payroll CSV should contain.
  const paymentToggleBtns = document.querySelectorAll('.payment-toggle-btn');
  paymentToggleBtns.forEach(btn => {
    btn.addEventListener('click', () => {
      paymentToggleBtns.forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      activePaymentMode = btn.dataset.mode;
      // Clear any leftover status from a previous download — the filter
      // change might have invalidated those skip counts.
      const el = document.getElementById('csv-download-status');
      if (el) { el.textContent = ''; el.className = 'csv-status'; }
    });
  });

  const downloadBtn = document.getElementById('download-csv-btn');
  if (downloadBtn) {
    downloadBtn.addEventListener('click', () => {
      const result = downloadCommissionCSV(lastEntries, activePaymentMode);
      showDownloadStatus(result, activePaymentMode);
    });
  }
  // #download-payroll-btn is intentionally not wired — the disabled
  // attribute in the HTML prevents clicks, and there's nothing to do
  // when Payroll is defined we'll add a downloadPayrollCSV function
  // following the same pattern as downloadCommissionCSV and wire it here.

  // Email-sending wiring. Guarded — the HTML elements come in a later
  // step, so if they're not in the DOM yet this is a no-op and the rest
  // of the dashboard continues to work normally.
  const sendBtn = document.getElementById('send-emails-btn');
  if (sendBtn) {
    sendBtn.addEventListener('click', openSendEmailsModal);
    document.getElementById('email-cancel-btn').addEventListener('click', closeSendEmailsModal);
    document.getElementById('email-confirm-btn').addEventListener('click', confirmAndSendEmails);
    document.getElementById('email-modal').addEventListener('click', (e) => {
      if (e.target.id === 'email-modal') closeSendEmailsModal();  // click backdrop to close
    });
  }

  // Tab navigation — Summary / Trends. Pure show/hide; data lives in
  // memory and is rendered eagerly on every Calculate (renderResults +
  // renderTrends both fire), so switching tabs is instant. When user
  // first lands on Trends without ever having Calculated, the empty
  // state copy in #trends-empty explains what to do.
  const tabBtns = document.querySelectorAll('.tab-btn');
  tabBtns.forEach(btn => {
    btn.addEventListener('click', () => {
      const target = btn.dataset.tab;
      tabBtns.forEach(b => {
        const isActive = b.dataset.tab === target;
        b.classList.toggle('active', isActive);
        b.setAttribute('aria-selected', isActive ? 'true' : 'false');
      });
      document.querySelectorAll('.tab-pane').forEach(pane => {
        pane.hidden = (pane.id !== target + '-pane');
      });
      // Re-render trends on activation. Two reasons:
      //   1. If the trends pane was hidden when its Chart.js instance
      //      was created, the canvas had 0 width and Chart.js will have
      //      rendered an invisible chart. Re-render = correct sizing.
      //   2. Cheap insurance against any state drift between renders.
      if (target === 'trends') renderTrends();
    });
  });

  // Chart breakdown toggle — Department vs Role. Persists in
  // currentTrendMode (module scope) so re-renders from elsewhere
  // (Calculate, tab switch) keep whichever mode the user picked.
  const trendModeBtns = document.querySelectorAll('.trend-mode-btn');
  trendModeBtns.forEach(btn => {
    btn.addEventListener('click', () => {
      trendModeBtns.forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      currentTrendMode = btn.dataset.mode;
      // Only the chart depends on mode — metrics and top earners are
      // mode-independent. Just re-render the chart to avoid flicker.
      if (lastHistoryData && lastHistoryData.length > 0) {
        renderTrendsChart(lastHistoryData.filter(r => r && r.Period && r.Person), currentTrendMode);
      }
    });
  });

  /* ---- MODE SWITCH: Commissions ⇄ Regular Payroll ----
     Pure show/hide of #commission-view and #payroll-view. Each view owns
     its own inputs/results; switching never re-runs either calc. */
  const modeBtns = document.querySelectorAll('.mode-btn');
  const views = { commission: document.getElementById('commission-view'),
                  payroll:    document.getElementById('payroll-view') };
  modeBtns.forEach(btn => {
    btn.addEventListener('click', () => {
      const target = btn.dataset.mode;
      modeBtns.forEach(b => {
        const on = b.dataset.mode === target;
        b.classList.toggle('active', on);
        b.setAttribute('aria-selected', on ? 'true' : 'false');
      });
      Object.entries(views).forEach(([mode, el]) => { if (el) el.hidden = (mode !== target); });
    });
  });

  /* ---- REGULAR PAYROLL wiring ---- */
  const payrollFile = document.getElementById('payroll-file');
  const payrollFileName = document.getElementById('payroll-file-name');
  const payrollAnalyzeBtn = document.getElementById('payroll-analyze-btn');
  const payrollStatus = document.getElementById('payroll-status-line');
  const payrollPeriodDisplay = document.getElementById('payroll-period-display');
  const payrollDownloadBtn = document.getElementById('payroll-download-btn');
  let payrollFileText = null;

  if (payrollFile) {
    payrollFile.addEventListener('change', () => {
      const f = payrollFile.files && payrollFile.files[0];
      if (!f) return;
      payrollFileName.textContent = f.name;
      payrollPeriodDisplay.textContent = f.name.replace(/\.csv$/i, '');
      payrollStatus.className = 'hint center';
      payrollStatus.textContent = 'Reading file…';
      const reader = new FileReader();
      reader.onload = () => {
        payrollFileText = reader.result;
        payrollAnalyzeBtn.disabled = false;
        payrollStatus.textContent = 'File ready. Click Analyze Hours.';
      };
      reader.onerror = () => {
        payrollStatus.className = 'hint center error';
        payrollStatus.textContent = 'Could not read that file. Try re-exporting the CSV.';
      };
      reader.readAsText(f);
    });
  }

  // Optional approvals upload (overrides the synced sheet when present).
  const approvalsFile = document.getElementById('payroll-approvals-file');
  const approvalsFileName = document.getElementById('payroll-approvals-name');
  let approvalsFileText = null;
  if (approvalsFile) {
    approvalsFile.addEventListener('change', () => {
      const f = approvalsFile.files && approvalsFile.files[0];
      if (!f) { approvalsFileText = null; if (approvalsFileName) approvalsFileName.textContent = 'Approved hours CSV (optional)'; return; }
      if (approvalsFileName) approvalsFileName.textContent = f.name;
      const reader = new FileReader();
      reader.onload = () => { approvalsFileText = reader.result; };
      reader.readAsText(f);
    });
  }

  if (payrollAnalyzeBtn) {
    payrollAnalyzeBtn.addEventListener('click', async () => {
      if (!payrollFileText) return;
      payrollAnalyzeBtn.disabled = true;
      payrollAnalyzeBtn.classList.add('loading');
      payrollStatus.className = 'hint center';
      payrollStatus.textContent = 'Analyzing…';
      try {
        const val = (id, def) => { const el = document.getElementById(id); const v = parseFloat(el && el.value); return Number.isFinite(v) ? v : def; };
        const winStartStr = (document.getElementById('payroll-window-start') || {}).value || '08:00';
        const winEndStr = (document.getElementById('payroll-window-end') || {}).value || '17:00';
        const meta = {
          dailyLimit: val('payroll-limit', 8),
          graceMin: val('payroll-grace', 0),
          breakLimitMin: val('payroll-break-limit', 30),
          breakGraceMin: val('payroll-break-grace', 2),
          winStartStr, winEndStr,
          winStartMin: parseClock(winStartStr) ?? 480,
          winEndMin: parseClock(winEndStr) ?? 1020,
          winGraceMin: val('payroll-window-grace', 5),
          flagWeekend: (document.getElementById('payroll-flag-weekend') || {}).checked ?? true,
          fileName: (payrollFile.files[0] && payrollFile.files[0].name) || 'report',
        };
        const rows = parsePayrollCSV(payrollFileText);
        if (rows.length < 2) throw new Error('File has no data rows.');
        const isCustom = detectPayrollFormat(rows) === 'custom';

        // Helper to render an analysis + update the status line.
        const show = (analysis) => {
          lastPayrollAnalysis = analysis;
          lastPayrollMeta = meta;
          renderPayroll(analysis, meta);
          const n = analysis.flagged.length;
          const supN = (analysis.suppressed || []).length;
          const fmtNote = analysis.format === 'custom' ? '' : ' (hours-only — upload a Custom Export for break & window checks)';
          const supNote = supN ? ` ${supN} cleared by approvals.` : '';
          payrollStatus.className = 'hint center success';
          payrollStatus.textContent = (n === 0
            ? `Done. No flags.${fmtNote}`
            : `Done. ${n} ${n === 1 ? 'person' : 'people'} flagged for review.${fmtNote}`) + supNote;
        };

        // 1) Uploaded approvals (instant) take precedence.
        let approvals = null;
        if (isCustom && approvalsFileText) {
          approvals = buildApprovalIndexFromCSV(approvalsFileText);
          meta.approvalsOk = approvals.detected.ok;
          meta.approvalsNote = approvals.detected.ok
            ? `Approvals: ${approvals.list.length} approved request${approvals.list.length===1?'':'s'} loaded from uploaded file.`
            : `Approvals file uploaded but columns weren't recognized (need a status, date, and name/email column) — nothing suppressed.`;
        } else if (isCustom) {
          // We'll fetch the synced tab in the background (see step 3).
          meta.approvalsNote = `Checking "${TAB_ADDITIONAL_HOURS}" for approvals…`;
          meta.approvalsOk = true;
        }

        // 2) Render flags IMMEDIATELY — the network never blocks this.
        show(runPayrollAnalysis(rows, meta, approvals));

        // 3) Background: pull synced approvals, then re-render with suppression.
        //    Not awaited, so the button frees up and results are already visible.
        if (isCustom && !approvalsFileText) {
          fetchApprovedHours().then((r) => {
            if (r.ok) {
              const appr = buildApprovalIndexFromObjects(r.rows);
              meta.approvalsOk = appr.detected.ok;
              meta.approvalsNote = appr.detected.ok
                ? `Approvals: ${appr.list.length} approved request${appr.list.length===1?'':'s'} synced from "${TAB_ADDITIONAL_HOURS}".`
                : `"${TAB_ADDITIONAL_HOURS}" loaded but columns weren't recognized — nothing suppressed.`;
              show(appr.detected.ok ? analyzePayrollCustom(rows, meta, appr) : lastPayrollAnalysis);
            } else {
              meta.approvalsOk = false;
              meta.approvalsNote = `Approvals not loaded (${r.reason}) — nothing suppressed. Upload an approvals CSV to apply them.`;
              renderPayroll(lastPayrollAnalysis, meta);   // refresh just the note
            }
          }).catch((e) => {
            meta.approvalsOk = false;
            meta.approvalsNote = `Approvals not loaded (${e.message}) — nothing suppressed.`;
            renderPayroll(lastPayrollAnalysis, meta);
          });
        }
      } catch (err) {
        console.error(err);
        payrollStatus.className = 'hint center error';
        payrollStatus.textContent = `Error: ${err.message}`;
      } finally {
        payrollAnalyzeBtn.disabled = false;
        payrollAnalyzeBtn.classList.remove('loading');
      }
    });
  }

  if (payrollDownloadBtn) {
    payrollDownloadBtn.addEventListener('click', () => {
      if (lastPayrollAnalysis && lastPayrollMeta) downloadPayroll(lastPayrollAnalysis, lastPayrollMeta);
    });
  }
});
