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

// Tab names — must match the Google Sheet exactly
const TAB_OUTREACH_SALES = 'Monthly Comm Outreach / Sales';
const TAB_APPROVED_LEADS = 'Approved Leads';
const TAB_SALES          = 'Sales';
const TAB_CONTRACTED     = 'Contracted Deals';
const TAB_PEOPLE         = 'People';
const TAB_FINANCE        = 'Finance';

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

// Look up a person in the PEOPLE roster. Tries:
//   1. exact normalized match (handles whitespace differences)
//   2. case-insensitive match (handles casing differences)
// Returns the PEOPLE entry or null if no match.
function lookupPerson(rawName) {
  if (!rawName) return null;
  const normalized = normalizeName(rawName);
  if (!normalized) return null;
  if (PEOPLE[normalized]) return PEOPLE[normalized];
  const lower = normalized.toLowerCase();
  for (const key of Object.keys(PEOPLE)) {
    if (key.toLowerCase() === lower) return PEOPLE[key];
  }
  return null;
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
   3. COMMISSION RULE TABLES
   ------------------------------------------------------------ */
const LAM_ADVANCE_TIERS = [
  { min: 0,      max: 19999,    advance: 50  },
  { min: 20000,  max: 49999,    advance: 100 },
  { min: 50000,  max: 99999,    advance: 150 },
  { min: 100000, max: 199999,   advance: 200 },
  { min: 200000, max: Infinity, advance: 250 },
];

// LAM advance program is effective only for contracts signed ON OR AFTER
// this date. Contracts signed before this generate no advance and (if
// later cancelled) no retraction either — there was nothing to pay back.
// Per Anshul: the actual program start was September 2025.
const LAM_ADVANCE_START_DATE = '2025-09-01';

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
  if (/^\d{4}-\d{2}-\d{2}/.test(trimmed)) return new Date(trimmed);
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
function calcLAMAdvances(contractedDeals, period) {
  const entries = [];
  const cutoff = parseDate(LAM_ADVANCE_START_DATE);
  for (const d of contractedDeals) {
    const contractId = d['Contract Name'] || d.Contract_Name || d.Deal_ID || '?';
    const owner = (d['Closer (Acq)'] || d.Closer_Acq || d.LAM_Owner || d.LAM || '').trim();
    const stage = (d['Contract Stage'] || d.Contract_Stage || '').trim();
    const spread = parseMoney(d['Deal Spread'] || d.Deal_Spread);
    const signedDate = d['Signed Contract Received Date'] || d.Signed_Contract_Received_Date || d.Contract_Signed_Date;
    const cxDate = d['Date Cancelled'] || d.Date_Cancelled || d.Cancellation_Date;

    // Program eligibility: contracts signed BEFORE the program start date
    // generate neither an advance nor a retraction.
    const signedDateParsed = parseDate(signedDate);
    const isCoveredByProgram = signedDateParsed && cutoff && signedDateParsed >= cutoff;
    if (!isCoveredByProgram) continue;

    // Advance path: signed in target month
    if (inPeriod(signedDate, period)) {
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

    // Retraction path: cancelled in target month, stage = Cancelled
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
      const tier = tierFor(spread, LAM_ADVANCE_TIERS);
      if (!tier) {
        entries.push({
          person: owner, role: 'LAM', period, source: contractId, type: 'LAM Retraction',
          amount: 0, calc: `Spread ${fmtUSD(spread)} outside tiers`,
          notes: 'Check Deal Spread', flag: 'REVIEW',
        });
        continue;
      }
      const p = lookupPerson(owner);
      if (!p || p.status === 'Inactive') {
        const reason = !p
          ? `${owner} not in People tab — treated as inactive, not chasing former employees.`
          : `${owner} marked Inactive in People tab — not chasing former employees.`;
        entries.push({
          person: owner, role: 'LAM', period, source: contractId, type: 'LAM Retraction',
          amount: 0, calc: '—', notes: `Retraction skipped — ${reason}`, flag: 'REVIEW',
        });
        continue;
      }
      const reason = (d['Transaction Termination Reason'] || d.Transaction_Termination_Reason || '').trim();
      entries.push({
        person: owner, role: 'LAM', period, source: contractId, type: 'LAM Retraction',
        amount: -tier.advance,
        calc: `Cancelled (spread ${fmtUSD(spread)}) → retract advance ${fmtUSD(tier.advance)}`,
        notes: reason ? `Cancellation reason: ${reason}` : '',
        flag: 'OK',
      });
    }
  }
  return entries;
}

// LAM Final — reads Monthly Comm Outreach / Sales.
//   - Tier rate applied to AGP (max 0 — negative AGP floors to $0)
//   - Advance is deducted ONLY when the contract was signed on/after
//     LAM_ADVANCE_START_DATE. For pre-program contracts no advance was
//     ever paid, so there's nothing to net out.
//   - Floored at $0 on negative-AGP closes (LAM v2 §5).
//
// Signed date is read directly from the "AB Contract Signed Date" column
// on Outreach Sales (added per Anshul). If the column is blank/missing
// on a row, fall back to looking up the same deal in Contracted Deals.
// If still no match, default to "no advance deducted" — safer to slightly
// overpay than to claw back money the LAM was never actually advanced.
function calcLAMFinals(salesDeals, contractedDeals, period) {
  const entries = [];
  const cutoff = parseDate(LAM_ADVANCE_START_DATE);

  // Fallback index: deal-ID → signed-date from Contracted Deals
  const signedByDealId = {};
  for (const c of (contractedDeals || [])) {
    const id = (c['Contract Name'] || c.Contract_Name
              || c['Deal Settlement'] || c.Deal_Settlement
              || c.Deal_ID || c['Deal ID'] || '').toString().trim();
    if (!id) continue;
    const signed = c['Signed Contract Received Date']
                || c.Signed_Contract_Received_Date
                || c.Contract_Signed_Date;
    const signedParsed = parseDate(signed);
    if (signedParsed) signedByDealId[id] = signedParsed;
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

    // Pick up signed date — preferred source first, then fallback.
    const signedDateRaw = d['AB Contract Signed Date']
                       || d.AB_Contract_Signed_Date
                       || d['Contract Signed Date']
                       || d.Contract_Signed_Date;
    const signedDate = parseDate(signedDateRaw) || signedByDealId[dealId] || null;
    const advanceWasPaid = signedDate && cutoff && signedDate >= cutoff;

    let origAdvance = 0;
    if (advanceWasPaid) {
      const advTier = tierFor(spread, LAM_ADVANCE_TIERS);
      origAdvance = advTier ? advTier.advance : 0;
    }
    const net = gross - origAdvance;

    let calcNote;
    if (agpRaw < 0) {
      calcNote = `Negative AGP ${fmtUSD(agpRaw)} floored to $0; advance ${fmtUSD(origAdvance)} kept (deal closed). Net: ${fmtUSD(net)}`;
    } else if (advanceWasPaid) {
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
      || lead['Seller Name: Sourcer'] || '').trim() || null;
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
      entries.push({
        person: sourcer || qualifier || '?', role: 'LIA/LIM Closing', period, source: dealId,
        type: 'LIA/LIM Closing', amount: 0, calc: '—',
        notes: `Unknown AILeadCategory: "${cat}"`, flag: 'REVIEW',
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
    const dom = parseInt(String(domStr).trim(), 10);

    if (!listPrice || !contractPrice || isNaN(dom)) {
      entries.push({
        person: owner, role: 'LPM', period, source: listingId, type: 'LPM (excluded)',
        amount: 0, calc: '—',
        notes: `Missing one of: Curr List Price, BC Under Contract Price, DOM. Seller: ${seller}`,
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
  for (const tab of [TAB_OUTREACH_SALES, TAB_APPROVED_LEADS, TAB_SALES, TAB_CONTRACTED, TAB_PEOPLE, TAB_FINANCE]) {
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
  return body;
}

/* ------------------------------------------------------------
   8. RENDERING
   ------------------------------------------------------------ */
function renderResults(entries, period) {
  const summaryEl = document.getElementById('summary');
  const detailEl  = document.getElementById('detail');
  const flagsEl   = document.getElementById('flags');

  // Aggregate by person. Skip:
  //   - INELIGIBLE entries (handled by the Flags panel)
  //   - Unattributed entries (person === '—'); these are SF data hygiene
  //     issues, not real payouts, and shouldn't clutter the table.
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
  // Alphabetical sort so payroll review reads top-down by name, not by
  // amount. Locale-aware so accented characters sort sensibly.
  const sortedPeople = Object.entries(byPerson).sort((a, b) => a[0].localeCompare(b[0]));
  const grandTotal = sortedPeople.reduce((s, [, v]) => s + v.total, 0);

  let html = `<div class="period-banner">Period: <strong>${escapeHTML(fmtPeriodName(period))}</strong> &nbsp;·&nbsp; Grand Total: <strong>${fmtUSD(grandTotal)}</strong></div>`;

  // Summary table — Person · Review · Breakdown · Total.
  // Breakdown shows per-type subtotals (e.g., "LAM Advance $50, LAM Final $565.80")
  // so payroll review can see at a glance what makes up someone's total
  // without expanding the drill-down.
  html += '<div class="people-list">';
  html += '<div class="people-header">';
  html += '<span>Person</span>';
  html += '<span class="flag-col" title="Manual review needed">⚑</span>';
  html += '<span>Breakdown</span>';
  html += '<span class="num">Total</span>';
  html += '</div>';

  for (const [person, v] of sortedPeople) {
    const personEntries = entries.filter(e => e.person === person && e.flag !== 'INELIGIBLE');
    if (personEntries.length === 0) continue;
    // Build the breakdown string from per-type subtotals. Sort by absolute
    // amount descending so the biggest line items come first.
    const breakdownParts = Object.entries(v.byType)
      .sort((a, b) => Math.abs(b[1]) - Math.abs(a[1]))
      .map(([type, amount]) => `${type} ${fmtUSD(amount)}`);
    const breakdownStr = breakdownParts.join(', ');
    html += `<details class="person-row${v.hasReview ? ' has-review' : ''}">`;
    html += `<summary>`;
    html += `<span class="name">${escapeHTML(person)}</span>`;
    html += `<span class="flag-col">${v.hasReview ? '🚩' : ''}</span>`;
    html += `<span class="breakdown-col">${escapeHTML(breakdownStr)}</span>`;
    html += `<span class="num">${fmtUSD(v.total)}</span>`;
    html += `</summary>`;

    // Drill-down: aggregate by type, not one row per line. For Juan-style
    // cases (38 separate $1 leads), this collapses to "38 × $1 = $38" in
    // a single row.
    html += '<div class="person-row-body">';
    html += '<table class="detail-table"><thead><tr>';
    html += '<th>Type</th><th class="num">Lines</th><th class="num">Total</th><th>Summary</th>';
    html += '</tr></thead><tbody>';
    for (const g of groupByType(personEntries)) {
      const flagClass = g.hasReview ? 'flag-review' : 'flag-ok';
      html += `<tr class="${flagClass}">`;
      html += `<td>${escapeHTML(g.type)}</td>`;
      html += `<td class="num">${g.count}</td>`;
      html += `<td class="num">${fmtUSD(g.total)}</td>`;
      html += `<td class="summary-cell">${escapeHTML(describeGroup(g))}</td>`;
      html += `</tr>`;
    }
    html += '</tbody></table>';

    // Per-person manual review list — appears under the calc table so
    // each person's flags are right next to their payout.
    // Grouped by (type, note) so duplicate notes (e.g., LPA's QoC
    // verification reminder repeated across N lines) collapse into one line.
    const reviewItems = personEntries.filter(e => e.flag === 'REVIEW');
    if (reviewItems.length > 0) {
      const reviewGroups = {};
      for (const e of reviewItems) {
        const message = e.notes || e.calc || '';
        const key = `${e.type}||${message}`;
        if (!reviewGroups[key]) {
          reviewGroups[key] = { type: e.type, message, count: 0, sources: [] };
        }
        reviewGroups[key].count += 1;
        if (reviewGroups[key].sources.length < 5) reviewGroups[key].sources.push(e.source);
      }
      const groups = Object.values(reviewGroups);
      const totalReview = reviewItems.length;
      html += `<div class="person-review-list">`;
      html += `<div class="person-review-h">🚩 Manual Review Needed (${totalReview})</div>`;
      html += `<ul>`;
      for (const g of groups) {
        let sources;
        if (g.count === 1) {
          sources = g.sources[0] ? ` <span class="src">(${escapeHTML(g.sources[0])})</span>` : '';
        } else {
          const shown = g.sources.join(', ');
          const more = g.count > g.sources.length ? `, +${g.count - g.sources.length} more` : '';
          sources = ` <span class="src">(${g.count} lines: ${escapeHTML(shown + more)})</span>`;
        }
        html += `<li><strong>${escapeHTML(g.type)}</strong>${sources}: ${escapeHTML(g.message)}</li>`;
      }
      html += `</ul></div>`;
    }

    html += '</div></details>';
  }

  html += `<div class="people-total"><span><strong>TOTAL</strong></span><span></span><span></span><span class="num"><strong>${fmtUSD(grandTotal)}</strong></span></div>`;
  html += '</div>';  // close .people-list

  summaryEl.innerHTML = html;
  if (detailEl) detailEl.innerHTML = '';  // legacy container, now empty

  // Bottom flags panel — INELIGIBLE only. REVIEW items now live inline
  // under each person's drill-down (right next to their payout).
  const ineligibleEntries = entries.filter(e => e.flag === 'INELIGIBLE');
  let flagsHTML = '';
  if (ineligibleEntries.length > 0) {
    flagsHTML += `<h3 class="flag-h">— Ineligible (${ineligibleEntries.length})</h3><ul>`;
    for (const e of ineligibleEntries) {
      flagsHTML += `<li><strong>${escapeHTML(e.person)} / ${escapeHTML(e.type)}</strong> — ${escapeHTML(e.source)}: ${escapeHTML(e.notes)}</li>`;
    }
    flagsHTML += '</ul>';
  }
  if (!flagsHTML) flagsHTML = '<p class="empty">No ineligibles. Manual-review items (if any) are shown under each person above.</p>';
  flagsEl.innerHTML = flagsHTML;

  document.getElementById('results').classList.remove('hidden');
}

// Group a person's entries by type. For drill-down summary.
// Tracks whether any entry in the group is REVIEW (for the flag indicator).
function groupByType(entries) {
  const groups = {};
  for (const e of entries) {
    const key = e.type || '(unknown)';
    if (!groups[key]) {
      groups[key] = {
        type: key,
        count: 0,
        total: 0,
        hasReview: false,
        amounts: new Set(),
        firstCalc: e.calc,
        firstNotes: e.notes,
      };
    }
    groups[key].count += 1;
    groups[key].total += e.amount;
    if (e.flag === 'REVIEW') groups[key].hasReview = true;
    groups[key].amounts.add(Math.round(e.amount * 100));  // dedupe to cents
  }
  return Object.values(groups).sort((a, b) => Math.abs(b.total) - Math.abs(a.total));
}

// Produce a concise human-readable summary for a group of entries.
//   - Single entry: use its full calc string
//   - Multiple entries, all same amount: "N × $X = $total"
//   - Multiple entries, varying amounts: brief "N lines, varying" note
// REVIEW-flagged groups get an inline note pulled from the first entry's
// notes/calc (e.g., the QoC gate caveat for LPA/LLP).
function describeGroup(g) {
  if (g.count === 1) return g.firstCalc || '—';
  if (g.amounts.size === 1) {
    const each = g.total / g.count;
    let s = `${g.count} × ${fmtUSD(each)} = ${fmtUSD(g.total)}`;
    if (g.hasReview && g.firstNotes) s += ` — ${g.firstNotes}`;
    return s;
  }
  let s = `${g.count} lines, varying amounts`;
  if (g.hasReview && g.firstNotes) s += ` — ${g.firstNotes}`;
  return s;
}

function escapeHTML(s) {
  if (s === null || s === undefined) return '';
  return String(s).replace(/[&<>"']/g, c => ({
    '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'
  }[c]));
}

/* ------------------------------------------------------------
   9. EXPORTS
   ------------------------------------------------------------ */
function entriesToTable(entries) {
  const headers = ['Person', 'Role', 'Period', 'Source', 'Type', 'Amount', 'Calculation', 'Notes', 'Flag'];
  const rows = entries.map(e => [
    e.person, e.role, e.period || '', e.source, e.type,
    e.amount.toFixed(2), e.calc, e.notes, e.flag,
  ]);
  return [headers, ...rows];
}

function downloadCSV(entries, period) {
  const table = entriesToTable(entries);
  const csv = table.map(row => row.map(cell => {
    const s = String(cell ?? '');
    return /[",\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
  }).join(',')).join('\n');
  const blob = new Blob([csv], { type: 'text/csv;charset=utf-8' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `commissions_${period}.csv`;
  a.click();
  URL.revokeObjectURL(url);
}

function copyTSV(entries) {
  const table = entriesToTable(entries);
  const tsv = table.map(row => row.join('\t')).join('\n');
  navigator.clipboard.writeText(tsv).then(() => {
    const btn = document.getElementById('copy-tsv-btn');
    const orig = btn.textContent;
    btn.textContent = '✓ Copied!';
    setTimeout(() => { btn.textContent = orig; }, 2000);
  });
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
      statusLine.className = 'success';
      statusLine.textContent = `Done. ${lastEntries.length} commission lines computed.`;
    } catch (err) {
      console.error(err);
      statusLine.className = 'error';
      statusLine.textContent = `Error: ${err.message}`;
    } finally {
      calcBtn.classList.remove('loading');
      calcBtn.disabled = false;
    }
  });

  document.getElementById('download-csv-btn').addEventListener('click', () =>
    downloadCSV(lastEntries, lastPeriod));
  document.getElementById('copy-tsv-btn').addEventListener('click', () =>
    copyTSV(lastEntries));
});
