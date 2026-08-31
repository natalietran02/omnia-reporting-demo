/**
 * Omnia Reporting — Apps Script Web App
 *
 * Deploy: Extensions → Apps Script → Deploy → New deployment → Web app
 *   Execute as: Me (the deploying account)
 *   Who has access: Anyone
 *
 * Access control is NOT the Google deployment setting above (the allowed
 * domains are Microsoft/Entra tenants, so nobody has a Google account on
 * their work email) — it's the app-layer email OTP login in the "OTP /
 * session auth" section below, gated by ALLOWED_DOMAINS.
 *
 * API key: Store in Project Settings → Script Properties → ANTHROPIC_API_KEY
 * GitHub token (for the code-review "create fix PR" flow): Script Properties → GITHUB_TOKEN
 * Never hardcode either key in this file or in Index.html.
 *
 * ── Data model ──────────────────────────────────────────────────────
 * This file is bound to the "Omnia Reporting Sources" spreadsheet. The 8
 * TABS.* entries below are Connected Sheets extract tabs (Power BI /
 * HubSpot / roster exports) — update the right-hand side to match the
 * EXACT tab names at the bottom of your spreadsheet if they ever get
 * renamed. memos/memoVersions/editors/pages/config are plain Sheets tabs
 * this app reads and writes directly (board memo + role management).
 */

// ---------- Config ----------
const TABS = {
  // Connected-sheet data extracts.
  keyMetrics:     'data_key_metrics_monthly',   // CENTRALISED SOURCE OF TRUTH for monthly headline KPIs — period/month_date/entity/division/billable_revenue/utilisation_pct/margin/active_clients/billable_fte/cummulative_billable_revenue. Prefer this for any headline/summary query; use unifiedFact only for project/client/BU-level drill-down.
  unifiedFact:    'data_powerbi_unified_fact',  // month_date/entity/business_unit/project_*/headcount/total_revenue/total_cost/gross_margin/margin_pct/avg_daily_rate/monthly_revenue_target
  hubspotDeals:   'data_hubspot_deals',         // deal pipeline (HubSpot) — Deal_Stage/Amount/Close_Date/Omnia_Company/Confidence
  refProjects:    'data_ref_projects',          // Project_Code/Project_Name/Start_Date/End_Date/Client/Industry/Omnia_Company
  employeeRoster: 'data_employee_roster',       // userId/role_group/business_unit/employment_type/FTE/Omnia_Company
  revenueTarget:  'data_revenue_target',        // month_date/entity/total_revenue/revenue_target
  utilisation:    'data_utilisation',           // userId/month/fiscal_year/available_days/available_days_full/invoiced_days/leave_days/holiday_days/is_billable_staff/employment_basis/is_contractor/fte/business_unit/role_group/location/role_code/daily_rate/billed_amount/utilisation_pct/utilisation_pct_full/Omnia_Company (covers Solace Advisory/Kestrel Group/Northlight AI for all periods, Fernway Data only from FERNWAY_STANDARD_SOURCE_CUTOFF_PERIOD onward)
  certifications: 'data_certifications',        // CorpEmail/TrainingType/ActivationStatus/ExpirationDate/TrainingCompletionDate
  targetRates:    'data_target_daily_rate',     // FY/Grade/Utilisation_Target/Daily_Rate_Target
  pl:             'data_pl',                    // company-level P&L budget/actual/forecast, monthly (all OpCos including Fernway Data)
  // elysiumPlAnnual:        'data_elysium_pl_annual',         // Elysium's own P&L source — annual, one row per FY
  elysiumPlMonthly:       'data_elysium_pl_monthly',        // Elysium's own P&L source — monthly, Actual/Forecast(=Budget) rows
  // elysiumPlByLineMonthly: 'data_elysium_pl_byline_monthly', // Elysium's own P&L source — by revenue stream, monthly
  // elysiumPlByLineAnnual:  'data_elysium_pl_byline_annual',  // Elysium's own P&L source — by revenue stream, annual (FY24..FYnn)

  // App-owned tabs (board memo + role management).
  memos:        'memos',
  memoVersions: 'memo_versions',
  editors:      'editors',
  pages:        'pages',
  config:       'config',
  activityLog:  'activity_log',
  accessUsers:  'access_users',
  feedback:     'feedback',
};

const CLAUDE_MODEL = 'claude-sonnet-4-6';

const MONTH_ABBR_TO_NUM = { jan:1, feb:2, mar:3, apr:4, may:5, jun:6, jul:7, aug:8, sep:9, oct:10, nov:11, dec:12 };

// From this period onwards, Fernway Data's P&L/utilisation/FTE/day-rate
// come from the same standard sources as every other OpCo (data_pl,
// data_utilisation, data_powerbi_unified_fact, data_revenue_target,
// data_key_metrics_monthly) instead of its own dedicated tables. Periods
// before this still route to Elysium's own dedicated source
// (data_elysium_pl_monthly). Mirrors Index.html's
// FERNWAY_STANDARD_SOURCE_CUTOFF (new Date(2026, 4, 1) = May 2026) — keep
// both in sync if this ever changes.
// Employee master data (roster, certifications) is NOT affected by this
// cutoff — Elysium has always shared data_employee_roster/data_certifications
// with every other OpCo, for all periods.
const FERNWAY_STANDARD_SOURCE_CUTOFF_PERIOD = '2026-05';

// Parses a Date object, a 'YYYY-MM-DD'-ish string, or a 'Mon-YYYY' text label
// (e.g. data_certifications.Month = "Feb-2026") into a 'YYYY-MM' period label.
// Returns '' if it can't be parsed. Sheet cells come back as Date objects
// from getValues(); readTab() below normalises those to local 'YYYY-MM-DD'
// strings before this ever sees them, but the Date branch stays as a
// defensive fallback for any caller passing a raw Date directly. The day-of-
// month is irrelevant here — data_hubspot_deals.Month uses end-of-month
// dates (e.g. 6/30/2024) while every other table uses start-of-month
// (7/1/2023) — both normalise to the same 'YYYY-MM' period either way.
function periodOf(v) {
  if (!v) return '';
  if (v instanceof Date) return v.getFullYear() + '-' + String(v.getMonth() + 1).padStart(2, '0');
  const s = String(v);
  const iso = s.match(/^(\d{4})-(\d{2})/);
  if (iso) return iso[1] + '-' + iso[2];
  const abbr = s.match(/^([A-Za-z]{3})-(\d{4})$/);
  if (abbr) {
    const num = MONTH_ABBR_TO_NUM[abbr[1].toLowerCase()];
    return num ? abbr[2] + '-' + String(num).padStart(2, '0') : '';
  }
  // Fallback for a numeric 'M/D/YYYY' cell that landed as plain TEXT instead
  // of a real Date type (readTab() only auto-converts actual Date objects —
  // a mixed-type column, e.g. one manually-edited data_hubspot_deals row
  // among mostly-real-Date ones, silently fell through to '' here and was
  // invisible to every period-scoped tool call regardless of which month was
  // queried). M/D order matches this table's established convention (see
  // the Month-column comment above, e.g. '6/30/2024').
  const mdy = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (mdy) return mdy[3] + '-' + mdy[1].padStart(2, '0');

  return '';
}

// Normalises a percentage value that might be stored as a 0-1 fraction
// (e.g. 0.42) or already as a 0-100 number (e.g. 42) — sheet exports for
// *_pct columns are inconsistent about this. Anything <= 1.5 is assumed to
// be a fraction (no real margin/utilisation % is between 0 and 1.5%).
function asPercent(v) {
  const n = Number(v) || 0;
  return Math.abs(n) <= 1.5 ? n * 100 : n;
}

// Converts a Date cell to a 'YYYY-MM-DD' string using its LOCAL calendar
// components — NOT Date.prototype.toISOString(), which converts to UTC and
// can shift the date backward by a day depending on the spreadsheet's
// timezone (e.g. a 7/1/2023 local-midnight cell in an Australia/Sydney
// spreadsheet becomes "2023-06-30T14:00:00.000Z" under toISOString()).
// Every *_date/month column across the connected sheets (month_date,
// data_hubspot_deals.Month, data_utilisation.month, ...) is a real Date
// cell at month-start or month-end, so this keeps periodOf()'s 'YYYY-MM'
// extraction correct regardless of which day-of-month the sheet uses.
function dateToLocalISO(d) {
  return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
}

// ---------- Entry points ----------
// Google Sign-In can't gate this page for us — the allowed domains are
// Microsoft/Entra tenants, not Google Workspace ones, so nobody has a Google
// account on their work email. Access is deployed as ANYONE (anonymous, no
// Google login), and auth instead happens at the application layer via an
// email OTP (see "OTP / session auth" below) — every RPC call below carries
// a session token instead of relying on Session.getActiveUser().
function doGet(e) {
  if (!e || !e.parameter || !e.parameter.action) {
    return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Omnia Reporting')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  return handle(e, 'GET');
}

function doPost(e) {
  return handle(e, 'POST');
}

// Called from the client via google.script.run — bypasses CORS entirely.
// request_otp/verify_otp are the login flow itself, so they run before any
// token exists; every other action requires params.token to resolve a user.
function apiAction(action, params) {
  params = params || {};
  const lower = (action || '').toLowerCase();
  if (lower === 'request_otp') return requestOtp(params.email);
  if (lower === 'verify_otp')  return verifyOtpLogged(params.email, params.code);
  const user   = userFromToken(params.token);
  const routes = routeHandlers(params, user);
  if (!routes[lower]) throw new Error(`Unknown action: ${action}`);
  try {
    const result = JSON.parse(JSON.stringify(routes[lower]()));
    logActivity(user, lower, 'ok', params);
    return result;
  } catch (err) {
    logActivity(user, lower, 'error', params);
    throw err;
  }
}

function handle(e, method) {
  let action = '', p = {}, user = '';
  try {
    action = (e.parameter.action || '').toLowerCase();
    let body = {};
    if (method === 'POST' && e.postData) {
      try {
        body = JSON.parse(e.postData.contents || '{}');
      } catch (parseErr) {
        return json({ error: 'Request body is not valid JSON: ' + String(parseErr && parseErr.message || parseErr) });
      }
    }
    p = Object.assign({}, e.parameter, body);
    if (action === 'request_otp') return json(requestOtp(p.email));
    if (action === 'verify_otp')  return json(verifyOtpLogged(p.email, p.code));
    user = userFromToken(p.token);
    const routes = routeHandlers(p, user);
    if (!routes[action]) return json({ error: `Unknown action: ${action}` });
    const result = routes[action]();
    logActivity(user, action, 'ok', p);
    return json(result);
  } catch (err) {
    logActivity(user, action, 'error', p);
    return json({ error: String(err && err.message || err) });
  }
}

// Thin direct RPC (no apiAction round trip needed) — lets the client fetch
// role/session info with a single google.script.run call, the same way it
// already calls getAllReportingData().
function getWhoAmI(token) {
  const email = userFromToken(token);
  return whoami(email);
}

// ---------- Shared route handlers ----------
// Single source of truth for all API routes reachable via apiAction/doGet/doPost.
// p = unified params (query string + POST body merged); user = caller email.
function routeHandlers(p, user) {
  return {
    whoami: () => whoami(user),
    pages:  () => readTab(TABS.pages, { status: 'active' }),
    config: () => configMap(),

    // Board memo
    memo:         () => readOne(TABS.memos, { page_slug: p.page, period: p.period }),
    memos:        () => readTab(TABS.memos, p.period ? { period: p.period } : {}),
    save_memo:    () => saveMemo(user, p),
    publish_memo: () => publishMemo(user, p),

    // Editor / role management (admin only) — who can edit/publish the board memo.
    list_editors:  () => listEditors(user),
    save_editor:   () => saveEditor(user, p),
    remove_editor: () => removeEditor(user, p),

    // App access management (access-admin only) — who is allowed into the
    // app at all, separate from memo-editor permissions above. See
    // "App access (share-like allowlist)" section.
    list_access_users: () => listAccessUsers(user),
    share_access:       () => shareAccess(user, p),
    revoke_access:      () => revokeAccess(user, p),

    // Fired by the client whenever a user clicks a data-download button —
    // no actual work here, this route exists purely so the standard
    // apiAction()/handle() logActivity() call below records who downloaded
    // what (p.filename), for the admin activity log.
    log_download: () => ({ ok: true }),

    // Free-text feedback, submitted from the feedback button on every page.
    submit_feedback: () => submitFeedback(user, p),

    // Activity log (admin only) — most recent first, capped so a large sheet
    // doesn't get pulled whole into a single response.
    activity_log: () => {
      requireRole(user, 'admin');
      const limit = Math.min(Number(p.limit) || 200, 1000);
      return readTab(TABS.activityLog).slice(-limit).reverse();
    },

    // Health / monitoring (admin only — exposes AI spend, data internals,
    // and lets a caller trigger real emails/AI calls, so this isn't safe
    // to leave open to every logged-in user).
    healthcheck: () => { requireRole(user, 'admin'); return runHealthCheck(); },
    monitor:     () => { requireRole(user, 'admin'); return { health: runHealthCheck(), checks: runDataChecks() }; },
    ai_commentary: () => { requireRole(user, 'admin'); return getClaudeCommentary(6); },
    send_test_email: () => {
      requireRole(user, 'admin');
      var email = String(p.email || '').trim();
      if (!email || !email.includes('@')) throw new Error('Valid email address required');
      return sendMonitorEmail(email);
    },

    // Runs the real health check then injects a synthetic issue so you can
    // verify the alert email format without breaking the sheet.
    test_data_alert: () => {
      requireRole(user, 'admin');
      var email = String(p.email || '').trim();
      if (!email || !email.includes('@')) throw new Error('Valid email address required');
      var health     = runHealthCheck();
      var dataChecks = runDataChecks();
      var SCENARIOS = {
        missing_data: 'TEST — revenue data missing for current reporting period (no rows in last 30 days)',
        stale_data:   'TEST — data_revenue_target last updated 47 days ago (threshold: 35 days)',
        bad_gp:       'TEST — gross margin is 142% for reporting period — likely a sign error in data_powerbi_unified_fact',
        missing_tab:  'TEST — data_hubspot_deals tab not found in spreadsheet',
      };
      var scenario = String(p.scenario || 'missing_data');
      dataChecks.issues.push(SCENARIOS[scenario] || SCENARIOS.missing_data);
      dataChecks.warnings.push('TEST MODE — triggered manually via Health Check page to verify alert delivery');
      return sendMonitorEmail(email, health, dataChecks);
    },

    // Emails the findings from the in-browser code review to the given address.
    email_code_review: () => {
      requireRole(user, 'admin');
      var toEmail  = String(p.email || user).trim();
      var findings = typeof p.findings === 'string' ? JSON.parse(p.findings) : p.findings;
      if (!Array.isArray(findings)) throw new Error('findings must be an array');
      return emailCodeReviewFindings(toEmail, findings, {
        reviewedAt:    p.reviewed_at,
        filesReviewed: Array.isArray(p.files_reviewed) ? p.files_reviewed : ['Code.js'],
      });
    },

    code_review: () => {
      requireRole(user, 'admin');
      return runCodeReview(p.include_html === 'true' || p.include_html === true);
    },
    targeted_code_review: () => {
      requireRole(user, 'admin');
      return runTargetedCodeReview(p.description, p.include_html === 'true' || p.include_html === true);
    },
    code_fix: () => {
      requireRole(user, 'admin');
      if (!p.finding) throw new Error('finding object required');
      var finding = typeof p.finding === 'string' ? JSON.parse(p.finding) : p.finding;
      return generateCodeFix(finding);
    },
    create_fix_pr: () => {
      requireRole(user, 'admin');
      if (!p.finding) throw new Error('finding object required');
      if (!p.patch)   throw new Error('patch object required');
      var finding = typeof p.finding === 'string' ? JSON.parse(p.finding) : p.finding;
      var patch   = typeof p.patch   === 'string' ? JSON.parse(p.patch)   : p.patch;
      return createFixPR(finding, patch);
    },
    list_fix_prs: () => {
      requireRole(user, 'admin');
      return listFixPRs();
    },
    merge_fix_pr: () => {
      requireRole(user, 'admin');
      if (!p.pr_number) throw new Error('pr_number required');
      return mergeFixPR(Number(p.pr_number), p.commit_title);
    },
    close_fix_pr: () => {
      requireRole(user, 'admin');
      if (!p.pr_number) throw new Error('pr_number required');
      return closeFixPR(Number(p.pr_number));
    },
  };
}

// ---------- Auth ----------
const ALLOWED_DOMAINS = [
  'edgered.com.au',
  'elysiumdigital.com.au',
  'boundinteractive.com',
  'theonset.com.au',
  'theoc.ai',
];

function assertAllowedDomain(email) {
  const domain = (email || '').split('@')[1] || '';
  if (!ALLOWED_DOMAINS.includes(domain.toLowerCase())) {
    throw new Error(`Access denied for ${email || '(no email)'}`);
  }
}

// ---------- App access (share-like allowlist) ----------
// Second, independent layer on top of ALLOWED_DOMAINS: being on an allowed
// domain is necessary but no longer sufficient — a domain match just means
// "could plausibly be an Omnia Collective person", the access_users tab is
// the actual allowlist of who has been granted the app, similar to how
// sharing a Google Sheet works. Seed with seedInitialAccessUsers() (run once
// from the Apps Script editor); after that, access-admins add/remove people
// via share_access/revoke_access (UI: "Manage access").
// role: 'admin' can share/revoke access for others; 'user' just gets in.
const ACCESS_USERS_HEADERS = ['email', 'role', 'status', 'added_by', 'added_at', 'updated_at', 'updated_by'];

function accessUsersSheet() {
  return ensureSheet(TABS.accessUsers, ACCESS_USERS_HEADERS);
}

function accessRowFor(email) {
  accessUsersSheet(); // ensure the tab exists before readTab() below
  const rows = readTab(TABS.accessUsers, { status: 'active' });
  return rows.find(r => (r.email || '').toLowerCase() === (email || '').toLowerCase()) || null;
}

function assertHasAccess(email) {
  if (!accessRowFor(email)) {
    throw new Error(`${email || '(no email)'} has not been given access to Omnia Reporting yet. Ask an admin to share access with you.`);
  }
}

function requireAccessAdmin(email) {
  const row = accessRowFor(email);
  if (!row || row.role !== 'admin') throw new Error('Forbidden: access-admin required');
}

function listAccessUsers(user) {
  requireAccessAdmin(user);
  accessUsersSheet();
  return readTab(TABS.accessUsers);
}

// Grants (or re-activates) access for `body.email`. Only existing
// access-admins can call this — mirrors "share" on a Google Sheet. New
// people default to role 'user' (can use the app, can't share it onward);
// pass role: 'admin' to also let them share/revoke access for others.
function shareAccess(user, body) {
  requireAccessAdmin(user);
  const email = String(body.email || '').trim().toLowerCase();
  if (!email) throw new Error('email required');
  assertAllowedDomain(email); // still restricted to Omnia Collective domains
  const role = body.role === 'admin' ? 'admin' : 'user';
  return upsertRow(TABS.accessUsers, { email }, {
    email, role, status: 'active', added_by: user,
  }, user).row;
}

function revokeAccess(user, body) {
  requireAccessAdmin(user);
  const email = String(body.email || '').trim().toLowerCase();
  if (!email) throw new Error('email required');
  if (email === String(user || '').trim().toLowerCase()) throw new Error("You can't revoke your own access.");
  const result = upsertRow(TABS.accessUsers, { email }, { status: 'inactive' }, user).row;
  revokeSessionsForEmail(email);
  return result;
}

// One-time setup: in the Apps Script editor, select seedInitialAccessUsers
// in the function dropdown and click Run. Safe to re-run — upsertRow()
// updates existing rows instead of duplicating them.
function seedInitialAccessUsers() {
  accessUsersSheet(); // create the access_users tab (with headers) if this is the first run
  const ADMINS = [
    'natasha@edgered.com.au',
    'nathan@edgered.com.au',
    'wil@edgered.com.au',
    'vanessa@edgered.com.au',
    'veronica.williams@theoc.ai',
  ];
  const USERS = [
    'nargess@edgered.com.au',
    'atharva@edgered.com.au',
  ];
  ADMINS.forEach(email => shareAccessSeed(email, 'admin'));
  USERS.forEach(email => shareAccessSeed(email, 'user'));
  Logger.log('seedInitialAccessUsers: granted access to ' + (ADMINS.length + USERS.length) + ' user(s).');
}

// Seeding bypasses the requireAccessAdmin() gate in shareAccess() since no
// one has access yet on first run.
function shareAccessSeed(email, role) {
  email = String(email).trim().toLowerCase();
  assertAllowedDomain(email);
  return upsertRow(TABS.accessUsers, { email }, {
    email, role, status: 'active', added_by: 'seed',
  }, 'seed').row;
}

// ---------- OTP / session auth ----------
// Replaces Google Sign-In (the allowed domains are Microsoft/Entra tenants,
// so nobody has a Google account on their work email). A user proves control
// of an allowed-domain inbox with a one-time code, and gets back a session
// token the client stores itself (localStorage) and sends back on every
// subsequent call.
const OTP_TTL_SECONDS = 600;                       // 10 min to enter the code
const SESSION_TTL_MS  = 3600 * 1000;                // 1 hour idle timeout — stored in PropertiesService, not CacheService, since Cache tops out at 6h

function requestOtp(email) {
  email = String(email || '').trim().toLowerCase();
  if (!email || !email.includes('@')) throw new Error('Valid email address required');
  assertAllowedDomain(email);
  assertHasAccess(email);
  const code = String(Math.floor(100000 + Math.random() * 900000));
  // PropertiesService, not CacheService — Cache is best-effort and Google's
  // docs allow it to evict entries before their TTL under load, which showed
  // up as "Invalid or expired code" well within the 10-minute window.
  PropertiesService.getScriptProperties().setProperty(
    'otp_' + email,
    JSON.stringify({ code: code, exp: Date.now() + OTP_TTL_SECONDS * 1000 })
  );
  MailApp.sendEmail({
    to:       email,
    subject:  'Your Omnia Reporting login code',
    htmlBody: '<p>Your login code is:</p>' +
      '<p style="font-size:28px;font-weight:700;letter-spacing:6px">' + code + '</p>' +
      '<p style="color:#666">This code expires in 10 minutes. If you didn\'t request it, ignore this email.</p>',
  });
  return { sent: true };
}

function verifyOtp(email, code) {
  email = String(email || '').trim().toLowerCase();
  code  = String(code || '').trim();
  assertAllowedDomain(email);
  assertHasAccess(email);
  const props = PropertiesService.getScriptProperties();
  const key   = 'otp_' + email;
  const raw   = props.getProperty(key);
  const stored = raw ? JSON.parse(raw) : null;
  if (!stored || Date.now() > stored.exp || stored.code !== code) {
    throw new Error('Invalid or expired code');
  }
  props.deleteProperty(key);
  const token = Utilities.getUuid();
  PropertiesService.getScriptProperties().setProperty(
    'session_' + token,
    JSON.stringify({ email: email, exp: Date.now() + SESSION_TTL_MS })
  );
  return { token: token, email: email };
}

// Wraps verifyOtp() with an activity-log entry — login attempts happen
// before any session token exists, so they're outside the handle()/
// apiAction() logActivity() calls that cover every other route.
function verifyOtpLogged(email, code) {
  email = String(email || '').trim().toLowerCase();
  try {
    const result = verifyOtp(email, code);
    logActivity(email, 'login', 'ok');
    return result;
  } catch (err) {
    logActivity(email, 'login', 'error');
    throw err;
  }
}

// Resolves a session token to the email it was issued for, sliding the
// expiry forward so an active user is never logged out mid-session — only
// after 15 days of no requests at all. Stored in PropertiesService (not
// CacheService) since Cache tops out at a 6h TTL; cleanupExpiredSessions()
// below prunes stale entries so this store doesn't grow toward the 500KB
// Script Properties cap.
function userFromToken(token) {
  token = String(token || '').trim();
  if (!token) throw new Error('Not logged in');
  const props = PropertiesService.getScriptProperties();
  const key   = 'session_' + token;
  const raw   = props.getProperty(key);
  if (!raw) throw new Error('Session expired — please log in again');
  const session = JSON.parse(raw);
  if (Date.now() > session.exp) {
    props.deleteProperty(key);
    throw new Error('Session expired — please log in again');
  }
  // Re-checked on every request (not just at login) so revoking access — or
  // rolling this allowlist out against sessions that predate it — takes
  // effect immediately instead of waiting up to SESSION_TTL_MS for expiry.
  assertHasAccess(session.email);
  props.setProperty(key, JSON.stringify({ email: session.email, exp: Date.now() + SESSION_TTL_MS }));
  return session.email;
}

// Deletes every session token issued to `email` — call this when an editor
// is deactivated so their access doesn't linger for up to 15 days after
// removal. Scans all Script Properties, so cost grows with total session
// count; fine at the volumes this app expects (see cleanupExpiredSessions).
function revokeSessionsForEmail(email) {
  email = String(email || '').trim().toLowerCase();
  const props = PropertiesService.getScriptProperties();
  const all = props.getProperties();
  Object.keys(all).forEach(function (key) {
    if (key.indexOf('session_') !== 0) return;
    try {
      if (JSON.parse(all[key]).email === email) props.deleteProperty(key);
    } catch (e) { /* malformed entry — leave for cleanupExpiredSessions to reap */ }
  });
}

// Prunes expired session_ and otp_ entries from Script Properties. Deletes
// keys one at a time (not a bulk setProperties overwrite) so a login
// happening mid-run can't be clobbered by a stale snapshot. Run daily via
// installSessionCleanupTrigger(). otp_ entries also self-invalidate on read
// (verifyOtp checks stored.exp) — this just reclaims the Script Properties
// space from ones nobody ever came back to verify.
function cleanupExpiredSessions() {
  const props = PropertiesService.getScriptProperties();
  const all = props.getProperties();
  const now = Date.now();
  let removed = 0;
  Object.keys(all).forEach(function (key) {
    if (key.indexOf('session_') !== 0 && key.indexOf('otp_') !== 0) return;
    try {
      const entry = JSON.parse(all[key]);
      if (now > entry.exp) { props.deleteProperty(key); removed++; }
    } catch (e) {
      props.deleteProperty(key); // malformed entry, safe to drop
      removed++;
    }
  });
  Logger.log('cleanupExpiredSessions: removed ' + removed + ' expired session/otp entr(y/ies).');
  return { removed: removed };
}

// One-time setup: in the Apps Script editor, select installSessionCleanupTrigger
// in the function dropdown and click Run. Safe to re-run — clears any existing
// trigger for cleanupExpiredSessions first, so re-running never creates duplicates.
function installSessionCleanupTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'cleanupExpiredSessions') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('cleanupExpiredSessions')
    .timeBased()
    .everyDays(1)
    .atHour(3)
    .inTimezone('Australia/Sydney')
    .create();
  Logger.log('Daily trigger installed — cleanupExpiredSessions will run around 3am Australia/Sydney every day.');
}

function whoami(email) {
  const editors = readTab(TABS.editors, { status: 'active' });
  const me = editors.find(r => (r.email || '').toLowerCase() === (email || '').toLowerCase());
  const accessRow = accessRowFor(email);
  return {
    email,
    role:  me ? me.role : 'viewer',
    pages: me ? (me.pages || '').split(',').map(s => s.trim()).filter(Boolean) : [],
    isAccessAdmin: !!(accessRow && accessRow.role === 'admin'), // can share/revoke app access
  };
}

function requireRole(user, minRole, pageSlug) {
  const me = whoami(user);
  const order = { viewer: 0, editor: 1, admin: 2 };
  if ((order[me.role] || 0) < order[minRole]) throw new Error(`Forbidden: ${minRole} required`);
  if (pageSlug && me.role !== 'admin' && !(me.pages.includes('*') || me.pages.includes(pageSlug))) {
    throw new Error(`Forbidden: no access to page ${pageSlug}`);
  }
  return me;
}

// ---------- Sheet helpers ----------
function sheet(name) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(name);
  if (!sh) throw new Error(`Missing tab: ${name}`);
  return sh;
}

// Like sheet(), but creates the tab with the given header row on first use
// instead of throwing — for app-owned tabs (e.g. activity_log) that don't
// need to be pre-provisioned by hand in the spreadsheet.
function ensureSheet(name, headers) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(headers);
  }
  return sh;
}

// ---------- Activity log ----------
// One row per action, appended from the single handle()/apiAction() dispatch
// points below so every route is covered without touching routeHandlers().
// `detail` is a short param summary, not the full request — some routes
// (code_review, create_fix_pr, ...) carry large JSON payloads that don't
// belong in a spreadsheet cell, and params.token must never be logged.
const ACTIVITY_LOG_HEADERS = ['timestamp', 'email', 'action', 'status', 'detail'];

function summarizeParams(p) {
  if (!p) return '';
  const safe = Object.assign({}, p);
  delete safe.token;
  delete safe.findings;   // large JSON blobs — see comment above
  delete safe.finding;
  delete safe.patch;
  const s = JSON.stringify(safe);
  return s.length > 300 ? s.slice(0, 300) + '…' : s;
}

function logActivity(email, action, status, params) {
  try {
    const sh = ensureSheet(TABS.activityLog, ACTIVITY_LOG_HEADERS);
    sh.appendRow([new Date(), email || '', action || '', status || '', summarizeParams(params)]);
  } catch (e) {
    // Logging must never break the actual request.
    Logger.log('logActivity failed: ' + (e && e.message || e));
  }
}

// ---------- Feedback ----------
// Free-text feedback from any logged-in user, submitted via the feedback
// button on every page. Kept as its own tab (not folded into activity_log)
// since it's addressed to a person, not a machine-readable audit trail.
const FEEDBACK_HEADERS = ['timestamp', 'email', 'page', 'message'];

function submitFeedback(user, p) {
  const message = String(p.message || '').trim();
  if (!message) throw new Error('Feedback message is required.');
  const page = String(p.page || '').trim();
  const sh = ensureSheet(TABS.feedback, FEEDBACK_HEADERS);
  sh.appendRow([new Date(), user || '', page, message]);
  return { ok: true };
}

// Reads a tab into an array of row objects keyed by the header row. Handles
// ragged sheets (trailing blank rows/columns) and normalises Date cells to
// ISO strings so results serialize cleanly to the client either via
// JSON.parse(JSON.stringify()) (apiAction) or google.script.run's own
// marshalling (getAllReportingData). `where` (optional) filters rows by
// exact string match on one or more columns.
function readTab(name, where) {
  const sh = sheet(name);
  let values;
  try {
    values = sh.getDataRange().getValues();
  } catch (e) {
    const lastRow = Math.max(sh.getMaxRows(), 1);
    const lastCol = Math.max(sh.getMaxColumns(), 1);
    values = sh.getRange(1, 1, lastRow, lastCol).getValues();
    while (values.length > 1 && values[values.length - 1].every(v => v === '' || v == null)) {
      values.pop();
    }
    const lastHeaderCol = values[0].reduce((acc, v, i) => (v !== '' && v != null ? i + 1 : acc), 0);
    if (lastHeaderCol > 0 && lastHeaderCol < values[0].length) {
      values = values.map(r => r.slice(0, lastHeaderCol));
    }
  }
  if (values.length < 2) return [];
  const headers = values[0].map(h => String(h).trim());
  const rows = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (row.every(cell => cell === '' || cell === null)) continue; // skip fully blank rows
    const obj = {};
    headers.forEach((h, idx) => {
      let v = row[idx];
      if (v instanceof Date) v = dateToLocalISO(v);
      obj[h] = v;
    });
    rows.push(obj);
  }
  if (!where) return rows;
  return rows.filter(r => Object.entries(where).every(([k, v]) => {
    if (v == null) return true;
    return String(r[k]) === String(v);
  }));
}

function readOne(name, where) {
  return readTab(name, where)[0] || null;
}

// Header lookup is case/whitespace-insensitive because the Config tab's header
// row has drifted before (e.g. 'Config Key' vs 'config key') and silently broke
// current_period + all check_* overrides without any error.
function configMap() {
  const rows = readTab(TABS.config);
  const keyOf = (row, wanted) => {
    const match = Object.keys(row).find(h => String(h).trim().toLowerCase() === wanted);
    return match ? row[match] : undefined;
  };
  return Object.fromEntries(
    rows
      .map(r => [keyOf(r, 'config key'), keyOf(r, 'default')])
      .filter(([k]) => k !== undefined && k !== '')
  );
}

// The single source of truth for "the latest period with real actuals" —
// config-driven via current_period (set this once the month closes), falling
// back to this calendar month if unset. Every read-only summary (AI
// commentary) and sanity check (runDataChecks) must filter periods through
// this before aggregating/displaying, otherwise future budget/forecast rows —
// which carry a target revenue but no real cost/margin yet — get treated as
// actuals and produce nonsense numbers (e.g. GP% pinned at a placeholder
// value) or nonsense "anomalies".
function currentReportingPeriod() {
  var cfg = {};
  try { cfg = configMap(); } catch (e) {}
  if (cfg.current_period) return String(cfg.current_period);
  var d = new Date();
  return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0');
}

function findRowIndex(name, where) {
  const sh = sheet(name);
  const values = sh.getDataRange().getValues();
  const headers = values[0];
  for (let i = 1; i < values.length; i++) {
    const row = Object.fromEntries(headers.map((h, j) => [h, values[i][j]]));
    if (Object.entries(where).every(([k, v]) => String(row[k]) === String(v))) {
      return { rowIndex: i + 1, headers, row };
    }
  }
  return { rowIndex: -1, headers, row: null };
}

function upsertRow(name, where, patch, user) {
  const sh = sheet(name);
  const { rowIndex, headers, row } = findRowIndex(name, where);
  const now = new Date();
  const next = Object.assign({}, row || {}, where, patch, { updated_at: now, updated_by: user });
  const values = headers.map(h => next[h] != null ? next[h] : '');
  if (rowIndex === -1) {
    sh.appendRow(values);
    return { inserted: true, row: next, prev: null };
  } else {
    sh.getRange(rowIndex, 1, 1, headers.length).setValues([values]);
    return { inserted: false, row: next, prev: row };
  }
}

// ---------- Raw reporting data (Connected Sheets extracts) ----------
// Public functions called from Index.html via google.script.run. Each
// returns the raw row objects for that tab. Aggregation/shaping for
// specific chart sections happens client-side in Index.html so the backend
// stays a thin, reusable data layer.
function getKeyMetrics()     { return readTab(TABS.keyMetrics); }
function getUnifiedFact()    { return readTab(TABS.unifiedFact); }
function getHubspotDeals()   { return readTab(TABS.hubspotDeals); }
function getRefProjects()    { return readTab(TABS.refProjects); }
function getEmployeeRoster() { return readTab(TABS.employeeRoster); }
function getRevenueTarget()  { return readTab(TABS.revenueTarget); }
function getUtilisation()    { return readTab(TABS.utilisation); }
function getCertifications() { return readTab(TABS.certifications); }
function getTargetRates()    { return readTab(TABS.targetRates); }
function getPl()                     { return readTab(TABS.pl); }
// function getElysiumPlAnnual()        { return readTab(TABS.elysiumPlAnnual); }
function getElysiumPlMonthly()       { return readTab(TABS.elysiumPlMonthly); }
// function getElysiumPlByLineMonthly() { return readTab(TABS.elysiumPlByLineMonthly); }
// function getElysiumPlByLineAnnual()  { return readTab(TABS.elysiumPlByLineAnnual); }

// Convenience: fetch everything in one call, so the frontend can do a
// single google.script.run round trip on page load instead of nine.
function getAllReportingData(token) {
  userFromToken(token);
  return {
    keyMetrics:     getKeyMetrics(),
    unifiedFact:    getUnifiedFact(),
    hubspotDeals:   getHubspotDeals(),
    refProjects:    getRefProjects(),
    employeeRoster: getEmployeeRoster(),
    revenueTarget:  getRevenueTarget(),
    utilisation:    getUtilisation(),
    certifications: getCertifications(),
    targetRates:    getTargetRates(),
    pl:                     getPl(),
    // elysiumPlAnnual:        getElysiumPlAnnual(),
    elysiumPlMonthly:       getElysiumPlMonthly(),
    // elysiumPlByLineMonthly: getElysiumPlByLineMonthly(),
    // elysiumPlByLineAnnual:  getElysiumPlByLineAnnual(),
  };
}

// ---------- Memos ----------
function saveMemo(user, body) {
  if (!body.page_slug || !body.period) throw new Error('page_slug and period required');
  requireRole(user, 'editor', body.page_slug);
  const memo_id = `${body.page_slug}_${body.period}`;
  const patch = pick(body, [
    'headline', 'standfirst',
    'bullet_1', 'bullet_2', 'bullet_3', 'bullet_4',
    'callout_label', 'callout_text',
    'status', 'source', 'model', 'prompt_version',
  ]);
  if (!patch.status) patch.status = 'draft';
  if (!patch.source) patch.source = 'human';
  const result = upsertRow(TABS.memos, { memo_id }, Object.assign({ memo_id, page_slug: body.page_slug, period: body.period }, patch), user);
  logVersions(memo_id, result.prev, result.row, user);
  return result.row;
}

function publishMemo(user, body) {
  if (!body.memo_id) throw new Error('memo_id required');
  const existing = readOne(TABS.memos, { memo_id: body.memo_id });
  if (!existing) throw new Error('memo not found');
  requireRole(user, 'editor', existing.page_slug);
  const result = upsertRow(TABS.memos, { memo_id: body.memo_id }, {
    status:       'published',
    published_at: new Date(),
    published_by: user,
  }, user);
  logVersions(body.memo_id, existing, result.row, user);
  return result.row;
}

function logVersions(memo_id, prev, next, user) {
  const sh = sheet(TABS.memoVersions);
  if (sh.getLastColumn() === 0) {
    console.warn('logVersions: memo_versions sheet has no columns, skipping version logging');
    return;
  }
  if (sh.getLastRow() < 1) {
    console.warn('logVersions: memo_versions sheet has no header row, skipping version logging');
    return;
  }
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const firstDataRow = sh.getLastRow() + 1;
  const now = new Date();
  const tracked = ['headline', 'standfirst', 'bullet_1', 'bullet_2', 'bullet_3', 'bullet_4', 'callout_label', 'callout_text', 'status'];
  const rows = [];
  tracked.forEach(field => {
    const oldV = prev ? (prev[field] || '') : '';
    const newV = next[field] || '';
    if (String(oldV) !== String(newV)) {
      const rec = {
        version_id: Utilities.getUuid(),
        memo_id, changed_at: now, changed_by: user,
        field, old_value: oldV, new_value: newV,
      };
      rows.push(headers.map(h => rec[h] != null ? rec[h] : ''));
    }
  });
  if (rows.length) sh.getRange(firstDataRow, 1, rows.length, headers.length).setValues(rows);
}

// ---------- Editor / role management ----------
function listEditors(user) {
  requireRole(user, 'admin');
  return readTab(TABS.editors);
}

function saveEditor(user, body) {
  requireRole(user, 'admin');
  const email = String(body.email || '').trim().toLowerCase();
  if (!email) throw new Error('email required');
  const patch = {
    role:   body.role  || 'viewer',
    pages:  body.pages || '*',
    status: 'active',
  };
  return upsertRow(TABS.editors, { email }, Object.assign({ email }, patch), user).row;
}

function removeEditor(user, body) {
  requireRole(user, 'admin');
  const email = String(body.email || '').trim().toLowerCase();
  if (!email) throw new Error('email required');
  const result = upsertRow(TABS.editors, { email }, { status: 'inactive' }, user).row;
  revokeSessionsForEmail(email); // kill any live 15-day sessions instead of letting access linger
  return result;
}

// ---------- AI cost budget guard ----------
// Protects the Anthropic API key from runaway spend. Rather than capping the
// NUMBER of requests (two requests can differ 10x in cost depending on
// conversation length/context), this tracks actual USD spend — computed from
// the `usage` block every Anthropic response already returns — and blocks
// the NEXT call once a configured daily/monthly threshold is reached. A
// request already in flight always completes; only spend already recorded
// before a call started is checked, so this is a budget ceiling, not a
// hard per-request price cap (which isn't knowable before the call returns).
// Thresholds live in Script Properties (same place as ANTHROPIC_API_KEY):
//   AI_DAILY_BUDGET_USD, AI_MONTHLY_BUDGET_USD — platform-wide, shared by
//     every user; protects the API key itself. Unset/non-numeric = no limit.
//   AI_PER_USER_DAILY_BUDGET_USD — optional additional per-user daily cap,
//     checked ONLY when a caller identifies the user (currently askChat,
//     since that's the only user-facing/high-volume AI entry point — the
//     admin-only Health-page tools stay platform-only). Stops one heavy
//     user from burning the whole platform budget without needing a
//     separate per-user quota system.
// Pricing is USD per 1M tokens (https://www.anthropic.com/pricing) — add an
// entry here if CLAUDE_MODEL ever changes to a different model family.
const AI_MODEL_PRICING_PER_MTOK = {
  'claude-sonnet-4-6': { input: 3, output: 15, cacheWrite: 3.75, cacheRead: 0.30 },
};

function aiCallCostUsd_(usage, model) {
  const pricing = AI_MODEL_PRICING_PER_MTOK[model];
  if (!pricing || !usage) return 0;
  const perMtok = n => (Number(n) || 0) / 1e6;
  return perMtok(usage.input_tokens) * pricing.input
    + perMtok(usage.output_tokens) * pricing.output
    + perMtok(usage.cache_creation_input_tokens) * pricing.cacheWrite
    + perMtok(usage.cache_read_input_tokens) * pricing.cacheRead;
}

function aiSpendDayKey_()   { return 'ai_spend_day_' + Utilities.formatDate(new Date(), 'Etc/UTC', 'yyyy-MM-dd'); }
function aiSpendMonthKey_() { return 'ai_spend_month_' + Utilities.formatDate(new Date(), 'Etc/UTC', 'yyyy-MM'); }
// Per-user key is namespaced by email + day — sanitized since Script
// Properties keys are plain strings and an email contains '@'/'.' safely,
// but this keeps the format predictable if that ever changes.
function aiSpendUserDayKey_(email) {
  return 'ai_spend_user_' + String(email || '').trim().toLowerCase() + '_' + Utilities.formatDate(new Date(), 'Etc/UTC', 'yyyy-MM-dd');
}

// Throws if today's or this month's accumulated spend has already reached
// its configured budget — platform-wide, plus a per-user daily check when
// `email` is provided and AI_PER_USER_DAILY_BUDGET_USD is set. Read is
// lock-free — a stale read just risks one or two extra calls slipping
// through right at the boundary, which is fine for a soft cost guard and
// not worth serializing every AI call to avoid.
function assertAiBudgetOk_(email) {
  const props = PropertiesService.getScriptProperties();
  const dailyLimit = Number(props.getProperty('AI_DAILY_BUDGET_USD'));
  if (dailyLimit > 0) {
    const spent = Number(props.getProperty(aiSpendDayKey_())) || 0;
    if (spent >= dailyLimit) {
      throw new Error('Daily AI budget reached ($' + spent.toFixed(2) + ' of $' + dailyLimit.toFixed(2) + ' used). Try again tomorrow, or ask an admin to raise AI_DAILY_BUDGET_USD in Script Properties.');
    }
  }
  const monthlyLimit = Number(props.getProperty('AI_MONTHLY_BUDGET_USD'));
  if (monthlyLimit > 0) {
    const spent = Number(props.getProperty(aiSpendMonthKey_())) || 0;
    if (spent >= monthlyLimit) {
      throw new Error('Monthly AI budget reached ($' + spent.toFixed(2) + ' of $' + monthlyLimit.toFixed(2) + ' used). Ask an admin to raise AI_MONTHLY_BUDGET_USD in Script Properties.');
    }
  }
  if (email) {
    const perUserDailyLimit = Number(props.getProperty('AI_PER_USER_DAILY_BUDGET_USD'));
    if (perUserDailyLimit > 0) {
      const spent = Number(props.getProperty(aiSpendUserDayKey_(email))) || 0;
      if (spent >= perUserDailyLimit) {
        throw new Error('You’ve reached today’s AI usage limit ($' + spent.toFixed(2) + ' of $' + perUserDailyLimit.toFixed(2) + '). Try again tomorrow.');
      }
    }
  }
}

// Adds costUsd to today's and this month's platform-wide running totals,
// plus today's per-user total when `email` is provided. Uses LockService so
// concurrent requests never lose an increment to a read-modify-write race —
// the whole point of this guard is an accurate running total, so losing
// increments here would silently under-count spend. Best-effort: if the
// lock is busy, this call's cost is dropped rather than blocking/failing
// the user's request over accounting.
function recordAiSpend_(costUsd, email) {
  if (!(costUsd > 0)) return;
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return;
  try {
    const props = PropertiesService.getScriptProperties();
    const dayKey = aiSpendDayKey_();
    const monthKey = aiSpendMonthKey_();
    props.setProperty(dayKey, String((Number(props.getProperty(dayKey)) || 0) + costUsd));
    props.setProperty(monthKey, String((Number(props.getProperty(monthKey)) || 0) + costUsd));
    if (email) {
      const userDayKey = aiSpendUserDayKey_(email);
      props.setProperty(userDayKey, String((Number(props.getProperty(userDayKey)) || 0) + costUsd));
    }
  } finally {
    lock.releaseLock();
  }
}

// Snapshot for the Health page — current spend/limits, or null for a limit
// that's unset (no cap configured).
function aiBudgetStatus_() {
  const props = PropertiesService.getScriptProperties();
  const dailyLimit = Number(props.getProperty('AI_DAILY_BUDGET_USD')) || null;
  const monthlyLimit = Number(props.getProperty('AI_MONTHLY_BUDGET_USD')) || null;
  return {
    dailySpentUsd:   Number(props.getProperty(aiSpendDayKey_())) || 0,
    dailyLimitUsd:   dailyLimit,
    monthlySpentUsd: Number(props.getProperty(aiSpendMonthKey_())) || 0,
    monthlyLimitUsd: monthlyLimit,
  };
}

// ---------- AI Chat / Ask-about-data ----------
// Low-level Anthropic Messages API call, shared by askAboutData and
// getClaudeCommentary. Throws on any transport/API error — callers that
// need a non-throwing contract wrap this themselves. `email` is optional —
// pass it when the caller knows which user triggered the call, so the
// per-user daily budget (if configured) is enforced; admin-only tools that
// don't pass it still get full platform-wide protection.
function anthropicCall(payload, email) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!apiKey) throw new Error('No Anthropic API key configured. Add ANTHROPIC_API_KEY under Project Settings > Script Properties.');

  assertAiBudgetOk_(email);

  const response = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method:      'post',
    contentType: 'application/json',
    headers: {
      'x-api-key':         apiKey,
      'anthropic-version': '2023-06-01',
    },
    payload:            JSON.stringify(payload),
    muteHttpExceptions: true,
  });

  const statusCode = response.getResponseCode();
  const data = JSON.parse(response.getContentText());
  if (statusCode !== 200) {
    const msg = (data && data.error && data.error.message) || `Anthropic API request failed (HTTP ${statusCode})`;
    throw new Error(msg);
  }
  recordAiSpend_(aiCallCostUsd_(data.usage, payload.model), email);
  return data;
}

// Runs the tool-call round trip: calls Claude, and if it asks to use a tool,
// executes it locally (runChatTool) and feeds the result back, up to
// maxRounds times. Returns the final text answer, or throws. `email` is
// optional — forwarded to anthropicCall on every round so a per-user daily
// budget (if configured) applies to the whole tool-call loop, not just its
// first round.
function runAssistantLoop(system, messages, tools, usageOut, email) {
  messages = messages.slice();
  const maxRounds = 5; // hard stop against a runaway tool-call loop

  for (let round = 0; round < maxRounds; round++) {
    const data = anthropicCall({
      model:      CLAUDE_MODEL,
      max_tokens: 1000,
      system:     system,
      messages:   messages,
      tools:      tools || undefined,
    }, email);
    // Optional: caller passes an array to collect per-round token usage
    // (incl. cache_read/cache_creation) for cost debugging, without
    // changing this function's plain-string return contract.
    if (usageOut && data.usage) usageOut.push(data.usage);

    if (data.stop_reason !== 'tool_use') {
      const textBlock = (data.content || []).find(b => b.type === 'text');
      return textBlock ? textBlock.text : '(No answer text returned.)';
    }

    messages.push({ role: 'assistant', content: data.content });
    const toolResults = (data.content || [])
      .filter(b => b.type === 'tool_use')
      .map(b => ({
        type:        'tool_result',
        tool_use_id: b.id,
        content:     runChatTool(b.name, b.input),
      }));
    messages.push({ role: 'user', content: toolResults });
  }

  throw new Error('Assistant made too many tool calls without producing an answer.');
}

const ASK_SYSTEM_PROMPT =
  'You are answering questions about a board reporting dashboard for a professional services company. ' +
  'The panel-level data below is a summary — if it is not enough to answer precisely (e.g. it needs an exact ' +
  'row, a different period, or a full breakdown), call one of the provided tools to fetch exact data instead ' +
  'of guessing. For a headline/summary monthly number ONLY (a single company- or division-level revenue, margin, ' +
  'utilisation, active-clients, or billable-FTE figure, with no further breakdown requested), call ' +
  'get_key_metrics_data FIRST and prefer its numbers over the in-context summary or any other tool — ' +
  'data_key_metrics_monthly is the centralised, agreed source of truth for those metrics. Do NOT call ' +
  'get_key_metrics_data for a question that asks for a metric broken down "by" a dimension (client industry, ' +
  'project stream, business unit, project, client, role, location, etc.) — that table carries no such breakdown. ' +
  'Go straight to the matching drill-down tool (e.g. get_unified_fact_data) instead, and use its `group_by` ' +
  'parameter when one is available, rather than wasting a round on get_key_metrics_data first. ' +
  'For a question about a specific P&L expense line (e.g. wages, travel, rent, recruitment, opex), EBITDA/EBIT/net ' +
  'profit detail, or budget-vs-actual-vs-forecast comparison, use get_pl_data (all OpCos) or, for Fernway Data ' +
  'specifically before ' + FERNWAY_STANDARD_SOURCE_CUTOFF_PERIOD + ', get_elysium_pl_monthly_data — this carries the full P&L detail that ' +
  // 'specifically, get_elysium_pl_monthly_data/get_elysium_pl_annual_data (or the _byline_ variants when the question ' +
  // 'is about a specific revenue stream or needs figures below EBITDA) — these carry the full P&L detail that ' +
  'get_key_metrics_data and get_unified_fact_data do not. ' +
  'ELYSIUM DIGITAL DATA ROUTING: employee master data (headcount, roster, certifications) has always used the same ' +
  'shared sources as every other OpCo (get_employee_roster_data, get_certifications_data) — no special handling ' +
  'needed there, for any period. Everything else about Fernway Data (P&L, revenue, margin, utilisation %, FTE, ' +
  'average day rate) is different: for periods before ' + FERNWAY_STANDARD_SOURCE_CUTOFF_PERIOD + ' (i.e. up to and including April 2026), Elysium\'s numbers ' +
  'live ONLY in its own dedicated table via get_elysium_pl_monthly_data (get_pl_data/get_key_metrics_data/' +
  'get_utilisation_data/get_revenue_target_data will have no Elysium rows for those months). From ' +
  FERNWAY_STANDARD_SOURCE_CUTOFF_PERIOD + ' onward (i.e. May 2026 onward), Elysium switches to the exact same shared sources as every ' +
  'other OpCo (get_pl_data, get_key_metrics_data, get_utilisation_data, get_revenue_target_data) and no longer ' +
  'appears in get_elysium_pl_monthly_data. A question spanning both sides of that cutoff (e.g. a trailing-12-month ' +
  'range, or a fiscal year straddling it) needs a call to BOTH sources and the results combined. ' +
  'CURRENCY: every dollar figure from every tool (revenue, cost, margin, daily_rate, billed_amount, deal amount, ' +
  'etc.) is in AUD, for every entity/OpCo including Kestrel Group and Fernway Data — there is no other currency in this ' +
  'data. Always format money as AUD with the "$" symbol (e.g. $1,234) — never GBP/£, USD/US$, or any other symbol, ' +
  'regardless of what an entity\'s name might suggest. ' +
  'ARITHMETIC: never do a calculation with more than one operation or more than two numbers in your head — call ' +
  'the calculate tool instead and quote its result verbatim. This applies to sums across several P&L lines, ' +
  'multi-month/multi-row averages, growth rates, margin/utilisation percentages derived from two other figures, ' +
  'and anything combining rows fetched from more than one tool call (e.g. across the Elysium cutoff). Mental ' +
  'multi-step arithmetic is unreliable and is the leading cause of wrong numbers in answers — do not skip the tool ' +
  'call to save a round trip. ' +
  'Be concise: 2-4 sentences, plain language, no markdown headers. ' +
  'If neither the data below nor the tools can answer the question, say so plainly.';

// Resolves "quarter Q of fiscal year fyInput" to the 3 calendar months it
// covers (['YYYY-MM', ...]), using the config tab's fy_start_period (e.g.
// '2025-07' => FY26 runs Jul 2025-Jun 2026, named by the year it ENDS in)
// as the single source of truth. This used to be done by asking the model
// to do the FY-label-to-calendar-month arithmetic itself from a prose
// description — that was unreliable (it would silently use a different/
// inconsistent quarter-to-month mapping per question) — so it now happens
// once, deterministically, in code. fyInput accepts 'FY26', 'FY2026', '26'
// or '2026'. Returns null if fy_start_period is missing/malformed or
// quarter isn't 1-4.
function fiscalQuarterPeriods_(fyInput, quarter) {
  const cfg = configMap();
  const m = String(cfg.fy_start_period || '').match(/^(\d{4})-(\d{2})$/);
  if (!m) return null;
  const startMonth = Number(m[2]); // 1-12
  const configEndYear = Number(m[1]) + (startMonth === 1 ? 0 : 1);

  const fyStr = String(fyInput || '').trim().toUpperCase().replace(/^FY/, '');
  if (!/^\d+$/.test(fyStr)) return null;
  const targetEndYear = fyStr.length <= 2
    ? (configEndYear - (configEndYear % 100)) + Number(fyStr)
    : Number(fyStr);
  const targetStartYear = targetEndYear - (startMonth === 1 ? 0 : 1);

  const q = Number(quarter);
  if (!(q >= 1 && q <= 4)) return null;

  const periods = [];
  for (let k = 0; k < 3; k++) {
    const i = (q - 1) * 3 + k;
    const monthIdx = (startMonth - 1 + i) % 12; // 0-based
    const yr = targetStartYear + Math.floor((startMonth - 1 + i) / 12);
    periods.push(yr + '-' + String(monthIdx + 1).padStart(2, '0'));
  }
  return periods;
}

// Human-readable guidance appended to the chat system prompt so the model
// knows FY quarters exist as a tool parameter (fiscal_year + quarter) and
// never has to compute month ranges itself.
function fiscalYearGuidance_() {
  const cfg = configMap();
  const m = String(cfg.fy_start_period || '').match(/^(\d{4})-(\d{2})$/);
  if (!m) return '';
  return 'FISCAL YEAR CONVENTION: Solace Advisory\'s fiscal year is named by the calendar year it ENDS in (config fy_start_period=' + m[0] + '). ' +
    'Do NOT compute which calendar months a quarter covers yourself. Instead, for ANY question naming a quarter ' +
    '("Q2", "Quarter 2", "second quarter") together with a fiscal year ("FY24", "FY2024") or a bare year ("2026") — ' +
    'treat a bare year the same as its FY label — pass that year as `fiscal_year` and the quarter number (1-4) as ' +
    '`quarter` on get_key_metrics_data / get_unified_fact_data / get_revenue_target_data / get_utilisation_data / ' +
    'get_hubspot_deals_data / get_certifications_data / get_employee_roster_data / get_ref_projects_data. ' +
    'The tool resolves the exact 3 months and returns all of them in one call — do not also pass `period`, and do ' +
    'not call the tool once per month.';
}

// Global "Ask Omnia AI" chat widget — called from Index.html's popup via
// google.script.run.askChat(messages, dataContext). `messages` is the
// frontend's running chatHistory (already capped there at the last 20
// entries) so the assistant sees the whole conversation, not just the
// latest question; MAX_HISTORY re-applies that cap server-side too, since
// the client is not a trust boundary. Tool-calling (CHAT_TOOLS/
// runChatTool/runAssistantLoop) is unchanged from the old per-panel
// askAboutData — the model can still pull exact/complete data beyond the
// viewer's current entity/period framing passed in dataContext.
function askChat(messages, dataContext, token) {
  const email = userFromToken(token);
  if (!Array.isArray(messages) || !messages.length) throw new Error('No messages provided.');

  const MAX_HISTORY = 20;
  const history = messages.slice(-MAX_HISTORY).map(m => ({
    role:    m.role === 'assistant' ? 'assistant' : 'user',
    content: String(m.content || ''),
  }));

  let staticSystem = ASK_SYSTEM_PROMPT;
  const fyGuidance = fiscalYearGuidance_();
  if (fyGuidance) staticSystem += '\n\n' + fyGuidance;
  // Split into a static block (identical on every call -> cacheable) and a
  // dynamic block (the viewer's current entity/period, different per call) so
  // Anthropic only re-bills full price for the small DATA block instead of
  // the whole system prompt on every round/question.
  const system = [
    { type: 'text', text: staticSystem, cache_control: { type: 'ephemeral' } },
    { type: 'text', text: 'DATA:\n' + String(dataContext || '') },
  ];

  // Collect per-round token usage so the caller can log/inspect cost —
  // surfaced to the browser console via sendChat() in Index.html.
  const usagePerRound = [];
  const answer = runAssistantLoop(system, history, CHAT_TOOLS, usagePerRound, email);
  const usage = usagePerRound.reduce((acc, u) => ({
    input_tokens:                acc.input_tokens + (u.input_tokens || 0),
    output_tokens:                acc.output_tokens + (u.output_tokens || 0),
    cache_creation_input_tokens: acc.cache_creation_input_tokens + (u.cache_creation_input_tokens || 0),
    cache_read_input_tokens:     acc.cache_read_input_tokens + (u.cache_read_input_tokens || 0),
  }), { input_tokens: 0, output_tokens: 0, cache_creation_input_tokens: 0, cache_read_input_tokens: 0 });

  return { answer, usage, rounds: usagePerRound.length };
}

// Tool definitions exposed to Claude via the Anthropic Messages API "tools"
// param, scoped to the 9 Connected Sheets extract tabs.
const CHAT_TOOLS = [
  {
    name: 'get_key_metrics_data',
    description: 'Returns the curated, agreed headline monthly KPIs (billable_revenue, utilisation_pct, margin, active_clients, billable_fte, cummulative_billable_revenue) from data_key_metrics_monthly, filtered by entity/division and a `period` (YYYY-MM) — REQUIRED, this table spans years of history. For a fiscal-quarter question specifically, pass `fiscal_year`+`quarter` INSTEAD of `period` to get all 3 months in one call. This is the CENTRALISED SOURCE OF TRUTH for monthly headline/summary numbers — call this FIRST for any such question and prefer its numbers over any other tool or the in-context summary. Only use get_unified_fact_data instead when the question needs project/client/BU-level drill-down this table does not carry. Covers Fernway Data only from ' + FERNWAY_STANDARD_SOURCE_CUTOFF_PERIOD + ' onward — for Elysium headline numbers before that, use get_elysium_pl_monthly_data instead.',
    input_schema: {
      type: 'object',
      properties: {
        entity:      { type: 'string', description: 'Company/entity name, e.g. Solace Advisory' },
        division:    { type: 'string', description: 'Division, if filtering below entity level' },
        period:      { type: 'string', description: 'Month in YYYY-MM format, e.g. 2026-06. Required for month-level questions — omit only if using fiscal_year+quarter for a quarter question instead.' },
        fiscal_year: { type: 'string', description: 'Fiscal-quarter questions only — use instead of `period`. Fiscal year label, e.g. FY26, FY2026, or 2026.' },
        quarter:     { type: 'integer', description: 'Fiscal-quarter questions only, paired with fiscal_year. Quarter number 1-4.' },
      },
      required: [],
    },
  },
  {
    name: 'get_unified_fact_data',
    description: 'Returns rows from data_powerbi_unified_fact (revenue, cost, gross margin, headcount, days by project/client/BU/month), filtered by entity, business_unit, and a `period` (YYYY-MM) — REQUIRED, this table spans years of history. For a fiscal-quarter question specifically, pass `fiscal_year`+`quarter` INSTEAD of `period` to get all 3 months in one call. Use only for project/client/BU-level drill-down — for headline/summary numbers, call get_key_metrics_data instead. For a question asking for a metric "by X" or "by X and Y" (e.g. margin by client industry, revenue by project stream) — pass `group_by` to get pre-aggregated summary rows (one per group, with row_count, summed revenue/cost/margin, and a correctly weighted margin_pct) INSTEAD of every raw row — this is far cheaper and the model should not sum raw rows itself.',
    input_schema: {
      type: 'object',
      properties: {
        entity:        { type: 'string', description: 'Company/entity name, e.g. Solace Advisory' },
        business_unit: { type: 'string', description: 'Business unit, if filtering below entity level' },
        period:        { type: 'string', description: 'Month in YYYY-MM format, e.g. 2026-06. Required for month-level questions — omit only if using fiscal_year+quarter for a quarter question instead.' },
        fiscal_year:   { type: 'string', description: 'Fiscal-quarter questions only — use instead of `period`. Fiscal year label, e.g. FY26, FY2026, or 2026.' },
        quarter:       { type: 'integer', description: 'Fiscal-quarter questions only, paired with fiscal_year. Quarter number 1-4.' },
        group_by:      { type: 'string', description: 'Optional. Comma-separated field(s) to aggregate by instead of returning raw rows, e.g. "client_industry,project_stream". Valid fields: period, fiscal_year, year_month, month_name, calendar_year, calendar_quarter, entity, revenue_status, business_unit, role_group, location, project_code, project_name, client_name, project_type, project_stream, client_industry. Returns one summary row per group instead of every raw row — use this for any "by X" / "by X and Y" question.' },
      },
      required: [],
    },
  },
  {
    name: 'get_hubspot_deals_data',
    description: 'Returns deal rows from data_hubspot_deals (filtered on Close_Date) — including deal probability, lost-reason, stalled flag, and weighted open pipeline — optionally further filtered by owning company (entity) and/or deal stage, and a `period` (YYYY-MM) — REQUIRED, this table spans years of history. For a fiscal-quarter question specifically, pass `fiscal_year`+`quarter` INSTEAD of `period` to get all 3 months in one call. For a question asking for deal totals/counts "by X" (e.g. deal amount by deal_stage, by deal_owner) — pass `group_by` for pre-aggregated summary rows (row_count, total amount, avg_amount, total weighted pipeline, avg probability) INSTEAD of every raw deal row.',
    input_schema: {
      type: 'object',
      properties: {
        entity:      { type: 'string', description: 'Owning company (Omnia_Company), e.g. Solace Advisory' },
        stage:       { type: 'string', description: 'Deal_Stage to filter to' },
        period:      { type: 'string', description: 'Month in YYYY-MM format, e.g. 2026-06. Filters on Close_Date. Required — omit only if using fiscal_year+quarter for a quarter question instead.' },
        fiscal_year: { type: 'string', description: 'Fiscal-quarter questions only — use instead of `period`. Fiscal year label, e.g. FY26, FY2026, or 2026.' },
        quarter:     { type: 'integer', description: 'Fiscal-quarter questions only, paired with fiscal_year. Quarter number 1-4.' },
        group_by:    { type: 'string', description: 'Optional. Comma-separated field(s) to aggregate by instead of returning raw rows, e.g. "deal_stage" or "deal_owner,entity". Valid fields: entity, company, brand, deal_owner, deal_team, deal_stage, confidence, source_system, deal_source, is_stalled, month.' },
      },
      required: [],
    },
  },
  {
    name: 'get_revenue_target_data',
    description: 'Returns monthly revenue-vs-target rows from data_revenue_target, filtered by entity and a `period` (YYYY-MM) — REQUIRED, this table spans years of history. For a fiscal-quarter question specifically, pass `fiscal_year`+`quarter` INSTEAD of `period` to get all 3 months in one call. Covers Fernway Data only from ' + FERNWAY_STANDARD_SOURCE_CUTOFF_PERIOD + ' onward — for Elysium revenue/target before that, use get_elysium_pl_monthly_data instead.',
    input_schema: {
      type: 'object',
      properties: {
        entity:      { type: 'string', description: 'Company/entity name, e.g. Solace Advisory' },
        period:      { type: 'string', description: 'Month in YYYY-MM format, e.g. 2026-06. Required for month-level questions — omit only if using fiscal_year+quarter for a quarter question instead.' },
        fiscal_year: { type: 'string', description: 'Fiscal-quarter questions only — use instead of `period`. Fiscal year label, e.g. FY26, FY2026, or 2026.' },
        quarter:     { type: 'integer', description: 'Fiscal-quarter questions only, paired with fiscal_year. Quarter number 1-4.' },
      },
      required: [],
    },
  },
  {
    name: 'get_utilisation_data',
    description: 'Returns per-employee monthly utilisation rows from data_utilisation (available/invoiced/leave/holiday days, fte, daily_rate, billed_amount, utilisation_pct, is_contractor, employment_basis), filtered by business unit and a `period` (YYYY-MM) — REQUIRED, this table spans years of history. For a fiscal-quarter question specifically, pass `fiscal_year`+`quarter` INSTEAD of `period` to get all 3 months in one call. For a question asking for utilisation "by X" (e.g. utilisation by business unit, by role_code, by is_contractor for permanent-vs-contractor) — pass `group_by` for pre-aggregated summary rows (row_count, summed available/invoiced days, and a correctly weighted utilisation_pct) INSTEAD of every raw per-employee row. Covers Solace Advisory/Kestrel Group/Northlight AI for all periods, and Fernway Data only from ' + FERNWAY_STANDARD_SOURCE_CUTOFF_PERIOD + ' onward — for Elysium utilisation/FTE/day-rate before that, use get_elysium_pl_monthly_data instead (its billable_fte/average_day_rate/billable_employee_utilisation_pct columns).',
    input_schema: {
      type: 'object',
      properties: {
        entity:      { type: 'string', description: 'Business unit to filter to' },
        period:      { type: 'string', description: 'Month in YYYY-MM format, e.g. 2026-06. Required for month-level questions — omit only if using fiscal_year+quarter for a quarter question instead.' },
        fiscal_year: { type: 'string', description: 'Fiscal-quarter questions only — use instead of `period`. Fiscal year label, e.g. FY26, FY2026, or 2026.' },
        quarter:     { type: 'integer', description: 'Fiscal-quarter questions only, paired with fiscal_year. Quarter number 1-4.' },
        group_by:    { type: 'string', description: 'Optional. Comma-separated field(s) to aggregate by instead of returning raw rows, e.g. "business_unit" or "role_code,employment_basis". Valid fields: period, fiscal_year, business_unit, role_group, role_code, location, employment_basis, is_contractor.' },
      },
      required: [],
    },
  },
  {
    name: 'get_employee_roster_data',
    description: 'Returns the employee/contractor roster from data_employee_roster AS IT STOOD as of a `period` (YYYY-MM) — REQUIRED, using the SCD effective_from column to resolve one row per employee. For a fiscal-quarter question specifically, pass `fiscal_year`+`quarter` INSTEAD of `period` to resolve as of the end of that quarter. Optionally filtered by company (Omnia_Company) — covers every OpCo including Fernway Data, for ALL periods (no cutoff/date split here, unlike Elysium\'s P&L/utilisation data). For a question asking for headcount/FTE/payroll "by X" (e.g. headcount by role_group and business_unit) — pass `group_by` for pre-aggregated summary rows (row_count as headcount, summed FTE/base_salary, avg tenure/hours/cost_per_hour) INSTEAD of every raw employee row.',
    input_schema: {
      type: 'object',
      properties: {
        entity:      { type: 'string', description: 'Owning company (Omnia_Company), e.g. Solace Advisory' },
        period:      { type: 'string', description: 'Month in YYYY-MM format, e.g. 2026-06. Resolves the roster as of this month. Required — omit only if using fiscal_year+quarter for a quarter question instead.' },
        fiscal_year: { type: 'string', description: 'Fiscal-quarter questions only — use instead of `period`. Fiscal year label, e.g. FY26, FY2026, or 2026.' },
        quarter:     { type: 'integer', description: 'Fiscal-quarter questions only, paired with fiscal_year. Quarter number 1-4.' },
        group_by:    { type: 'string', description: 'Optional. Comma-separated field(s) to aggregate by instead of returning raw rows, e.g. "role_group,business_unit". Valid fields: entity, role_title, role_code, role_group, employment_type, business_unit, location, salary_band_year. row_count in the result is the headcount for that group.' },
      },
      required: [],
    },
  },
  {
    name: 'get_certifications_data',
    description: 'Returns Microsoft partner training/certification completion rows from data_certifications (filtered on Month), for a `period` (YYYY-MM) — REQUIRED, this table spans years of history. Tracked at the Omnia Collective level (all OpCos including Fernway Data together, for ALL periods — no cutoff/date split here). For a fiscal-quarter question specifically, pass `fiscal_year`+`quarter` INSTEAD of `period` to get all 3 months in one call.',
    input_schema: {
      type: 'object',
      properties: {
        period:      { type: 'string', description: 'Month in YYYY-MM format, e.g. 2026-06. Filters on Month. Required — omit only if using fiscal_year+quarter for a quarter question instead.' },
        fiscal_year: { type: 'string', description: 'Fiscal-quarter questions only — use instead of `period`. Fiscal year label, e.g. FY26, FY2026, or 2026.' },
        quarter:     { type: 'integer', description: 'Fiscal-quarter questions only, paired with fiscal_year. Quarter number 1-4.' },
      },
      required: [],
    },
  },
  {
    name: 'get_target_rates_data',
    description: 'Returns FY+Grade target utilisation/daily-rate rows from data_target_daily_rate, optionally filtered by fiscal year and/or grade.',
    input_schema: {
      type: 'object',
      properties: {
        fiscal_year: { type: 'string', description: 'Fiscal year label, e.g. FY26' },
        grade:       { type: 'string', description: 'Role grade' },
      },
      required: [],
    },
  },
  {
    name: 'get_ref_projects_data',
    description: 'Returns project reference rows (client, industry, start/end dates) from data_ref_projects that were active during a `period` (YYYY-MM) — REQUIRED, this table spans years of history — optionally further filtered by project code and/or client name (substring match). For a fiscal-quarter question specifically, pass `fiscal_year`+`quarter` INSTEAD of `period` to match projects active during any month of that quarter.',
    input_schema: {
      type: 'object',
      properties: {
        project_code: { type: 'string', description: 'Exact Project_Code' },
        client:       { type: 'string', description: 'Client name substring' },
        period:       { type: 'string', description: 'Month in YYYY-MM format, e.g. 2026-06. Matches projects active during this month (Start_Date-End_Date). Required — omit only if using fiscal_year+quarter for a quarter question instead.' },
        fiscal_year:  { type: 'string', description: 'Fiscal-quarter questions only — use instead of `period`. Fiscal year label, e.g. FY26, FY2026, or 2026.' },
        quarter:      { type: 'integer', description: 'Fiscal-quarter questions only, paired with fiscal_year. Quarter number 1-4.' },
      },
      required: [],
    },
  },
  {
    name: 'get_pl_data',
    description: 'Returns company-level P&L budget/actual/forecast rows from data_pl (all OpCos, including Fernway Data from ' + FERNWAY_STANDARD_SOURCE_CUTOFF_PERIOD + ' onward — for Elysium periods before that, call get_elysium_pl_monthly_data instead), filtered by company (entity) and a `period` (YYYY-MM) — REQUIRED, this table spans years of history. For a fiscal-quarter question specifically, pass `fiscal_year`+`quarter` INSTEAD of `period` to get all 3 months in one call. For a question asking for a P&L line "by X" (e.g. operating expenses by company) — pass `group_by` for pre-aggregated summary rows (summed cost/revenue lines, averaged margin/rate/utilisation lines) INSTEAD of every raw row.',
    input_schema: {
      type: 'object',
      properties: {
        entity:      { type: 'string', description: 'Company/entity name (Omnia_Company), e.g. Solace Advisory' },
        period:      { type: 'string', description: 'Month in YYYY-MM format, e.g. 2026-06. Required for month-level questions — omit only if using fiscal_year+quarter for a quarter question instead.' },
        fiscal_year: { type: 'string', description: 'Fiscal-quarter questions only — use instead of `period`. Fiscal year label, e.g. FY26, FY2026, or 2026.' },
        quarter:     { type: 'integer', description: 'Fiscal-quarter questions only, paired with fiscal_year. Quarter number 1-4.' },
        group_by:    { type: 'string', description: 'Optional. Comma-separated field(s) to aggregate by instead of returning raw rows, e.g. "entity". Valid fields: period, fiscal_year, entity.' },
      },
      required: [],
    },
  },
  {
    name: 'get_elysium_pl_monthly_data',
    description: 'Returns Fernway Data\'s own monthly P&L rows from data_elysium_pl_monthly (revenue, direct costs, gross profit, opex lines, normalised EBITDA, FTE, day rate, utilisation) for periods BEFORE ' + FERNWAY_STANDARD_SOURCE_CUTOFF_PERIOD + ' only — from ' + FERNWAY_STANDARD_SOURCE_CUTOFF_PERIOD + ' onward Elysium\'s P&L/FTE/day-rate/utilisation come from the same standard sources as every other OpCo instead (get_pl_data, get_key_metrics_data, get_utilisation_data), so call those for that range. Filtered by `type` (Actual or Forecast/Budget) and a `period` (YYYY-MM) — REQUIRED, this table spans years of history. For a fiscal-quarter question specifically, pass `fiscal_year`+`quarter` INSTEAD of `period` to get all 3 months in one call.',
    input_schema: {
      type: 'object',
      properties: {
        entity:      { type: 'string', description: 'Company/entity name (Omnia_Company) — normally Fernway Data' },
        type:        { type: 'string', description: 'Actual or Forecast (Forecast = Budget)' },
        period:      { type: 'string', description: 'Month in YYYY-MM format, e.g. 2026-06. Required for month-level questions — omit only if using fiscal_year+quarter for a quarter question instead.' },
        fiscal_year: { type: 'string', description: 'Fiscal-quarter questions only — use instead of `period`. Fiscal year label, e.g. FY26, FY2026, or 2026.' },
        quarter:     { type: 'integer', description: 'Fiscal-quarter questions only, paired with fiscal_year. Quarter number 1-4.' },
      },
      required: [],
    },
  },
  // {
  //   name: 'get_elysium_pl_annual_data',
  //   description: 'Returns Elysium\'s own annual P&L rows from data_elysium_pl_annual (one row per fiscal year — revenue, direct costs, gross profit, opex lines, EBITDA, YoY growth, average FTE/day-rate/utilisation), optionally filtered by fiscal year.',
  //   input_schema: {
  //     type: 'object',
  //     properties: {
  //       entity:      { type: 'string', description: 'Company/entity name (Omnia_Company) — normally Fernway Data' },
  //       fiscal_year: { type: 'string', description: 'Fiscal year label, e.g. FY26, FY2026, or 2026' },
  //     },
  //     required: [],
  //   },
  // },
  // {
  //   name: 'get_elysium_pl_byline_monthly_data',
  //   description: 'Returns Elysium\'s own monthly P&L rows from data_elysium_pl_byline_monthly, broken down BY REVENUE STREAM (professional services / contract placements / permanent recruitment / other) down through gross profit, opex, EBITDA, EBIT, and net profit after tax. Filtered by `type` (Actual or Forecast/Budget) and a `period` (YYYY-MM) — REQUIRED, this table spans years of history. For a fiscal-quarter question specifically, pass `fiscal_year`+`quarter` INSTEAD of `period` to get all 3 months in one call. Use this instead of get_elysium_pl_monthly_data when the question is about a specific revenue stream or needs figures below EBITDA (EBIT, net profit).',
  //   input_schema: {
  //     type: 'object',
  //     properties: {
  //       entity:      { type: 'string', description: 'Company/entity name (Omnia_Company) — normally Fernway Data' },
  //       type:        { type: 'string', description: 'Actual or Forecast (Forecast = Budget)' },
  //       period:      { type: 'string', description: 'Month in YYYY-MM format, e.g. 2026-06. Required for month-level questions — omit only if using fiscal_year+quarter for a quarter question instead.' },
  //       fiscal_year: { type: 'string', description: 'Fiscal-quarter questions only — use instead of `period`. Fiscal year label, e.g. FY26, FY2026, or 2026.' },
  //       quarter:     { type: 'integer', description: 'Fiscal-quarter questions only, paired with fiscal_year. Quarter number 1-4.' },
  //     },
  //     required: [],
  //   },
  // },
  // {
  //   name: 'get_elysium_pl_byline_annual_data',
  //   description: 'Returns Elysium\'s own annual P&L rows from data_elysium_pl_byline_annual, broken down BY REVENUE STREAM (professional services / contract placements / permanent recruitment / other) down through gross profit, opex, EBITDA, EBIT, and net profit after tax — one row per fiscal year + type. Optionally filtered by fiscal year and/or `type` (Actual or Forecast/Budget). Use this instead of get_elysium_pl_annual_data when the question is about a specific revenue stream or needs figures below EBITDA (EBIT, net profit).',
  //   input_schema: {
  //     type: 'object',
  //     properties: {
  //       entity:      { type: 'string', description: 'Company/entity name (Omnia_Company) — normally Fernway Data' },
  //       type:        { type: 'string', description: 'Actual or Forecast (Forecast = Budget)' },
  //       fiscal_year: { type: 'string', description: 'Fiscal year label, e.g. FY26, FY2026, or 2026' },
  //     },
  //     required: [],
  //   },
  //   // Marks the cache breakpoint for the whole (static, identical-every-call)
  //   // tools array, so Anthropic reuses the cached prefix instead of re-billing
  //   // full price for these 14 tool definitions on every question/round.
  //   cache_control: { type: 'ephemeral' },
  // },
    {
    name: 'calculate',
    description: 'Evaluates an exact arithmetic expression (+ - * / and parentheses only, e.g. "(1250000 - 980000) / 980000 * 100") server-side and returns the precise numeric result. REQUIRED for any calculation that chains more than one operation or more than ~2 numbers pulled from tool results — e.g. summing several P&L expense lines, a multi-month average, a growth-rate or margin-percentage derived from two other numbers, or combining rows across the Elysium cutoff. Never compute such a result mentally in the response text — call this tool instead and quote its result, since multi-step mental arithmetic is the single biggest source of wrong numbers in answers. Fine to skip only for trivial single-operation arithmetic (e.g. just reading one number, or one plain subtraction/percentage already returned by a tool).',
    input_schema: {
      type: 'object',
      properties: {
        expression: { type: 'string', description: 'A numeric expression using only digits, + - * / ( ) and decimal points — no variables or function calls. Substitute the actual numbers from prior tool results in before calling.' },
      },
      required: ['expression'],
    },
  },
];

function runChatTool(name, input) {
  Logger.log('CHAT_TOOL: %s %s', name, JSON.stringify(input));
  try {
    input = input || {};
    const entity = input.entity ? String(input.entity) : null;
    const period = input.period ? String(input.period) : null;
    // fiscal_year + quarter resolve to the 3 months that quarter covers, so
    // a single tool call can answer a quarter question instead of the model
    // computing (and getting wrong) the month range itself.
    const quarterPeriods = (input.fiscal_year && input.quarter != null)
      ? fiscalQuarterPeriods_(input.fiscal_year, input.quarter)
      : null;
    const byEntityPeriod = (rows, entityField, periodField) => rows.filter(r => {
      if (entity && String(r[entityField] || '') !== entity) return false;
      const rowPeriod = periodOf(r[periodField]);
      if (period && rowPeriod !== period) return false;
      if (quarterPeriods && !quarterPeriods.includes(rowPeriod)) return false;
      return true;
    });

    // 8 of the 9 tables are required to be scoped to a period — either a
    // monthly time series (key metrics, unified fact, revenue target,
    // utilisation), a table with its own month/date column to filter on
    // (deals via Close_Date, certifications via Month), a point-in-time
    // roster resolved via its SCD effective_from column, or a date-range
    // table (ref-projects via Start_Date/End_Date). A call with no
    // period/quarter was observed dumping an ENTIRE table into one
    // tool_result (800k+ input tokens on a single question), so require a
    // period instead of silently returning everything.
    const periodMissing = !period && !quarterPeriods;
    const periodRequiredMsg = 'This table spans years of monthly history — specify `period` (YYYY-MM) or `fiscal_year`+`quarter` to scope the query.';

    // Only data_target_daily_rate has no month/date column at all (it's keyed
    // by FY+Grade, not by month) — it stays period-optional and instead uses
    // this row-count backstop.
    const MAX_TOOL_ROWS = 300;
    const bounded = (rows, noneMessage) => {
      if (!rows.length) return { message: noneMessage };
      if (rows.length <= MAX_TOOL_ROWS) return rows;
      return rows.slice(0, MAX_TOOL_ROWS).concat([{
        note: `Only the first ${MAX_TOOL_ROWS} of ${rows.length} matching rows are returned — narrow the request (e.g. entity/stage/client) to see the rest.`,
      }]);
    };

    // Shared server-side aggregation for "X by Y" questions — collapses raw
    // rows into one summary row per group (sums + simple averages + any
    // caller-supplied derived ratios, e.g. weighted margin %) instead of
    // making the model sum thousands of raw rows itself. Returns
    // { error } on an invalid group_by field, or { aggregated } otherwise.
    const aggregateByGroup = (rows, groupByInput, groupableFields, sumFields, avgFields, deriveFn) => {
      const groupFields = String(groupByInput).split(',').map(s => s.trim()).filter(Boolean);
      const invalid = groupFields.filter(f => !groupableFields.includes(f));
      if (invalid.length) return { error: `Invalid group_by field(s): ${invalid.join(', ')}. Valid fields: ${groupableFields.join(', ')}` };
      const groups = {};
      rows.forEach(r => {
        const key = groupFields.map(f => r[f]).join('||');
        if (!groups[key]) {
          groups[key] = { row_count: 0 };
          groupFields.forEach(f => { groups[key][f] = r[f]; });
          sumFields.forEach(f => { groups[key][f] = 0; });
          avgFields.forEach(f => { groups[key]['_' + f + '_sum'] = 0; });
        }
        const g = groups[key];
        g.row_count++;
        sumFields.forEach(f => { g[f] += r[f] || 0; });
        avgFields.forEach(f => { g['_' + f + '_sum'] += r[f] || 0; });
      });
      const aggregated = Object.values(groups).map(g => {
        const out = {};
        groupFields.forEach(f => { out[f] = g[f]; });
        out.row_count = g.row_count;
        sumFields.forEach(f => { out[f] = g[f]; });
        avgFields.forEach(f => { out[f] = g.row_count ? g['_' + f + '_sum'] / g.row_count : null; });
        if (deriveFn) Object.assign(out, deriveFn(g));
        return out;
      });
      return { aggregated };
    };

    if (name === 'get_key_metrics_data') {
      if (periodMissing) return JSON.stringify({ error: periodRequiredMsg });
      let rows = byEntityPeriod(readTab(TABS.keyMetrics), 'entity', 'month_date');
      if (input.division) rows = rows.filter(r => String(r.division || '') === String(input.division));
      const projected = rows.map(r => ({
        period:                       periodOf(r.month_date),
        entity:                       String(r.entity || ''),
        division:                     String(r.division || ''),
        billable_revenue:             Number(r.billable_revenue) || null,
        utilisation_pct:              Number(r.utilisation_pct) || null,
        margin:                       Number(r.margin) || null,
        active_clients:               Number(r.active_clients) || null,
        billable_fte:                 Number(r.billable_fte) || null,
        cummulative_billable_revenue: Number(r.cummulative_billable_revenue) || null,
      }));
      return JSON.stringify(projected.length ? projected : { message: 'No key-metrics rows found for the given filters.' });
    }
    if (name === 'get_unified_fact_data') {
      if (periodMissing) return JSON.stringify({ error: periodRequiredMsg });
      let rows = byEntityPeriod(readTab(TABS.unifiedFact), 'entity', 'month_date');
      if (input.business_unit) rows = rows.filter(r => String(r.business_unit || '') === String(input.business_unit));
      const projected = rows.map(r => ({
        period:                periodOf(r.month_date),
        fiscal_year:           String(r.fiscal_year || ''),
        year_month:            String(r.year_month || ''),
        month_name:            String(r.month_name || ''),
        calendar_year:         String(r.calendar_year || ''),
        calendar_quarter:      String(r.calendar_quarter || ''),
        entity:                String(r.entity || ''),
        revenue_status:        String(r.revenue_status || ''),
        business_unit:         String(r.business_unit || ''),
        role_group:            String(r.role_group || ''),
        location:              String(r.location || ''),
        project_code:          String(r.project_code || ''),
        project_name:          String(r.project_name || ''),
        client_name:           String(r.client_name || ''),
        project_type:          String(r.project_type || ''),
        project_stream:        String(r.project_stream || ''),
        client_industry:       String(r.client_industry || ''),
        headcount:             Number(r.headcount) || null,
        total_days:            Number(r.total_days) || null,
        total_revenue:         Number(r.total_revenue) || null,
        total_cost:            Number(r.total_cost) || null,
        gross_margin:          Number(r.gross_margin) || null,
        margin_pct:            Number(r.margin_pct) || null,
        avg_daily_rate:        Number(r.avg_daily_rate) || null,
        avg_daily_cost:        Number(r.avg_daily_cost) || null,
        monthly_revenue_target: Number(r.monthly_revenue_target) || null,
      }));
      // Optional server-side aggregation — for "X by Y" questions (e.g. margin
      // by client industry and project stream), returning a handful of grouped
      // summary rows instead of every raw project-month row cuts the token
      // cost of the tool_result by 90%+ instead of making the model sum
      // thousands of rows itself. Only kicks in when the model explicitly asks
      // for it — omitting `group_by` returns the same raw rows as before.
      if (input.group_by) {
        const result = aggregateByGroup(
          projected,
          input.group_by,
          ['period', 'fiscal_year', 'year_month', 'month_name', 'calendar_year', 'calendar_quarter', 'entity', 'revenue_status', 'business_unit', 'role_group', 'location', 'project_code', 'project_name', 'client_name', 'project_type', 'project_stream', 'client_industry'],
          ['total_revenue', 'total_cost', 'gross_margin', 'total_days', 'headcount', 'monthly_revenue_target'],
          ['avg_daily_rate', 'avg_daily_cost'],
          // Weighted margin % (sum of margin / sum of revenue) is the correct
          // aggregate — a plain average of per-row margin_pct would overweight
          // small-revenue rows.
          g => ({ margin_pct: g.total_revenue ? (g.gross_margin / g.total_revenue) * 100 : null })
        );
        if (result.error) return JSON.stringify({ error: result.error });
        return JSON.stringify(result.aggregated.length ? result.aggregated : { message: 'No unified-fact rows found for the given filters.' });
      }
      return JSON.stringify(projected.length ? projected : { message: 'No unified-fact rows found for the given filters.' });
    }
    if (name === 'get_hubspot_deals_data') {
      if (periodMissing) return JSON.stringify({ error: periodRequiredMsg + ' (filters on Close_Date.)' });
      let rows = readTab(TABS.hubspotDeals);
      if (entity) rows = rows.filter(r => String(r.Omnia_Company || '') === entity);
      if (input.stage) rows = rows.filter(r => String(r.Deal_Stage || '').toLowerCase() === String(input.stage).toLowerCase());
      if (period) rows = rows.filter(r => periodOf(r.Close_Date) === period);
      if (quarterPeriods) rows = rows.filter(r => quarterPeriods.includes(periodOf(r.Close_Date)));
      const projected = rows.map(r => ({
        record_id:                String(r.Record_ID || ''),
        entity:                   String(r.Omnia_Company || ''),
        company:                  String(r.Company || ''),
        brand:                    String(r.Brand || ''),
        deal_name:                String(r.Deal_Name || ''),
        deal_owner:               String(r.Deal_Owner || ''),
        deal_team:                String(r.Deal_Team || ''),
        deal_stage:               String(r.Deal_Stage || ''),
        amount:                   Number(r.Amount) || null,
        confidence:               String(r.Confidence || ''),
        close_date:               String(r.Close_Date || ''),
        open_date:                String(r.Open_Date || ''),
        last_modified_date:       String(r.Last_Modified_Date || ''),
        last_updated:             String(r.Last_Updated || ''),
        source_system:            String(r.Source_System || ''),
        contributing_opco_1:      String(r.Contributing_OpCo_1 || ''),
        contributing_opco_2:      String(r.Contributing_OpCo_2 || ''),
        contributing_opco_3:      String(r.Contributing_OpCo_3 || ''),
        contributing_opco_4:      String(r.Contributing_OpCo_4 || ''),
        contributing_opco_5:      String(r.Contributing_OpCo_5 || ''),
        deal_source:              String(r.Deal_Source || ''),
        alliance_partners:        String(r.Alliance_Partners || ''),
        logged_in_partner_portal: String(r.Logged_In_Partner_Portal || ''),
        deal_probability:         Number(r.Deal_Probability) || null,
        deal_lost_reason:         String(r.Deal_Lost_Reason || ''),
        why_lost:                 String(r.Why_Lost || ''),
        is_stalled:               String(r.Is_Stalled || ''),
        stage_entered_date:       String(r.Stage_Entered_Date || ''),
        weighted_open_pipeline:   Number(r.Weighted_Open_Pipeline) || null,
        month:                    String(r.Month || ''),
      }));
      if (input.group_by) {
        const result = aggregateByGroup(
          projected,
          input.group_by,
          ['entity', 'company', 'brand', 'deal_owner', 'deal_team', 'deal_stage', 'confidence', 'source_system', 'deal_source', 'is_stalled', 'month'],
          ['amount', 'weighted_open_pipeline'],
          ['deal_probability'],
          g => ({ avg_amount: g.row_count ? g.amount / g.row_count : null })
        );
        if (result.error) return JSON.stringify({ error: result.error });
        return JSON.stringify(result.aggregated.length ? result.aggregated : { message: 'No deals found for the given filters.' });
      }
      return JSON.stringify(projected.length ? projected : { message: 'No deals found for the given filters.' });
    }
    if (name === 'get_revenue_target_data') {
      if (periodMissing) return JSON.stringify({ error: periodRequiredMsg });
      const rows = byEntityPeriod(readTab(TABS.revenueTarget), 'entity', 'month_date');
      const projected = rows.map(r => ({
        period:         periodOf(r.month_date),
        entity:         String(r.entity || ''),
        total_revenue:  Number(r.total_revenue) || null,
        revenue_target: Number(r.revenue_target) || null,
      }));
      return JSON.stringify(projected.length ? projected : { message: 'No revenue-target rows found for the given filters.' });
    }
    if (name === 'get_utilisation_data') {
      if (periodMissing) return JSON.stringify({ error: periodRequiredMsg });
      let rows = readTab(TABS.utilisation);
      if (quarterPeriods) rows = rows.filter(r => quarterPeriods.includes(periodOf(r.month)));
      else if (period) rows = rows.filter(r => periodOf(r.month) === period);
      if (entity) rows = rows.filter(r => String(r.Omnia_Company || '') === entity);
      const projected = rows.map(r => ({
        period:               periodOf(r.month),
        fiscal_year:          String(r.fiscal_year || ''),
        user_id:              String(r.userId || ''),
        entity:               String(r.Omnia_Company || ''),
        business_unit:        String(r.business_unit || ''),
        role_group:           String(r.role_group || ''),
        role_code:            String(r.role_code || ''),
        location:             String(r.location || ''),
        employment_basis:     String(r.employment_basis || ''),
        is_contractor:        String(r.is_contractor || ''),
        is_billable_staff:    String(r.is_billable_staff || ''),
        available_days:       Number(r.available_days) || null,
        available_days_full:  Number(r.available_days_full) || null,
        invoiced_days:        Number(r.invoiced_days) || null,
        leave_days:           Number(r.leave_days) || null,
        holiday_days:         Number(r.holiday_days) || null,
        daily_rate:           Number(r.daily_rate) || null,
        billed_amount:        Number(r.billed_amount) || null,
        fte:                  Number(r.fte) || null,
        utilisation_pct:      Number(r.utilisation_pct) || null,
        utilisation_pct_full: Number(r.utilisation_pct_full) || null,
      }));
      if (input.group_by) {
        const result = aggregateByGroup(
          projected,
          input.group_by,
          ['period', 'fiscal_year', 'business_unit', 'role_group', 'role_code', 'location', 'employment_basis', 'is_contractor'],
          ['available_days', 'available_days_full', 'invoiced_days', 'leave_days', 'holiday_days', 'billed_amount', 'fte'],
          ['daily_rate'],
          // Weighted utilisation % (sum of invoiced / sum of available days) is
          // the correct aggregate — a plain average of per-row utilisation_pct
          // would overweight employees with fewer available days.
          g => ({ utilisation_pct: g.available_days ? (g.invoiced_days / g.available_days) * 100 : null })
        );
        if (result.error) return JSON.stringify({ error: result.error });
        return JSON.stringify(result.aggregated.length ? result.aggregated : { message: 'No utilisation rows found for the given filters.' });
      }
      return JSON.stringify(projected.length ? projected : { message: 'No utilisation rows found for the given filters.' });
    }
    if (name === 'get_employee_roster_data') {
      if (periodMissing) return JSON.stringify({ error: periodRequiredMsg + ' (roster is point-in-time — resolves to the roster as it stood as of this period, using effective_from.)' });
      // Roster rows are SCD (a new row per userId each time role/pay changes),
      // with no effective_to — so "as of period" means: for each userId, the
      // row with the latest effective_from that is still <= the target period.
      const asOf = period || quarterPeriods[quarterPeriods.length - 1];
      const latestByUser = {};
      readTab(TABS.employeeRoster).forEach(r => {
        const ef = periodOf(r.effective_from);
        if (!ef || ef > asOf) return;
        const uid = String(r.userId || '');
        if (!latestByUser[uid] || ef > periodOf(latestByUser[uid].effective_from)) latestByUser[uid] = r;
      });
      let rows = Object.values(latestByUser);
      if (entity) rows = rows.filter(r => String(r.Omnia_Company || '') === entity);
      const projected = rows.map(r => ({
        user_id:                String(r.userId || ''),
        entity:                 String(r.Omnia_Company || ''),
        role_title:             String(r.role_title || ''),
        role_code:              String(r.role_code || ''),
        role_group:             String(r.role_group || ''),
        employment_type:        String(r.employment_type || ''),
        business_unit:          String(r.business_unit || ''),
        location:               String(r.location || ''),
        employment_start_date:  String(r.employment_start_date || ''),
        effective_from:         String(r.effective_from || ''),
        tenure:                 Number(r.tenure) || null,
        tenure_long:            String(r.tenure_long || ''),
        fte:                    Number(r.FTE) || null,
        actual_weekly_hours:    Number(r.actual_weekly_hours) || null,
        salary_band_year:       String(r.salary_band_year || ''),
        base_salary:            Number(r.base_salary) || null,
        cost_per_hour:          Number(r.cost_per_hour) || null,
      }));
      if (input.group_by) {
        const result = aggregateByGroup(
          projected,
          input.group_by,
          ['entity', 'role_title', 'role_code', 'role_group', 'employment_type', 'business_unit', 'location', 'salary_band_year'],
          ['fte', 'base_salary'],
          ['tenure', 'actual_weekly_hours', 'cost_per_hour']
        );
        if (result.error) return JSON.stringify({ error: result.error });
        return JSON.stringify(result.aggregated.length ? result.aggregated : { message: 'No roster rows found for the given filters.' });
      }
      return JSON.stringify(projected.length ? projected : { message: 'No roster rows found for the given filters.' });
    }
    if (name === 'get_certifications_data') {
      if (periodMissing) return JSON.stringify({ error: periodRequiredMsg + ' (filters on Month.)' });
      let rows = readTab(TABS.certifications);
      if (period) rows = rows.filter(r => periodOf(r.Month) === period);
      if (quarterPeriods) rows = rows.filter(r => quarterPeriods.includes(periodOf(r.Month)));
      const projected = rows.map(r => ({
        aad_user_id:            String(r.AADUserId || ''),
        individual_first_name:  String(r.IndividualFirstName || ''),
        individual_last_name:   String(r.IndividualLastName || ''),
        email:                  String(r.Email || ''),
        corp_email:             String(r.CorpEmail || ''),
        partner_name:           String(r.PartnerName || ''),
        partner_country_location: String(r.PartnerCountryLocation || ''),
        partner_city_location:  String(r.PartnerCityLocation || ''),
        mpn_id:                 String(r.MPNId || ''),
        pga_mpn_id:             String(r.PGAMpnId || ''),
        training_activity_id:   String(r.TrainingActivityId || ''),
        training_type:          String(r.TrainingType || ''),
        training_title:         String(r.TrainingTitle || ''),
        activation_status:      String(r.ActivationStatus || ''),
        completion_date:        String(r.TrainingCompletionDate || ''),
        expiration_date:        String(r.ExpirationDate || ''),
        month:                  String(r.Month || ''),
      }));
      return JSON.stringify(projected.length ? projected : { message: 'No certification rows found.' });
    }
    if (name === 'get_target_rates_data') {
      let rows = readTab(TABS.targetRates);
      if (input.fiscal_year) rows = rows.filter(r => String(r.FY || '') === String(input.fiscal_year));
      if (input.grade)       rows = rows.filter(r => String(r.Grade || '') === String(input.grade));
      const projected = rows.map(r => ({
        fiscal_year:       String(r.FY || ''),
        grade:             String(r.Grade || ''),
        utilisation_target: Number(r.Utilisation_Target) || null,
        daily_rate_target:  Number(r.Daily_Rate_Target) || null,
      }));
      return JSON.stringify(bounded(projected, 'No target-rate rows found for the given filters.'));
    }
    if (name === 'get_ref_projects_data') {
      if (periodMissing) return JSON.stringify({ error: periodRequiredMsg + ' (filters to projects active during this period, using Start_Date/End_Date.)' });
      // A project is "active during" a period if the period falls within
      // [Start_Date, End_Date] — an empty End_Date means still ongoing.
      const activeDuring = (r, p) => {
        const start = periodOf(r.Start_Date);
        const end   = periodOf(r.End_Date);
        if (start && p < start) return false;
        if (end && p > end) return false;
        return true;
      };
      let rows = readTab(TABS.refProjects).filter(r =>
        period ? activeDuring(r, period) : quarterPeriods.some(p => activeDuring(r, p))
      );
      if (input.project_code) rows = rows.filter(r => String(r.Project_Code || '') === String(input.project_code));
      if (input.client)       rows = rows.filter(r => String(r.Client || '').toLowerCase().includes(String(input.client).toLowerCase()));
      const projected = rows.map(r => ({
        project_code: String(r.Project_Code || ''),
        project_name: String(r.Project_Name || ''),
        client:       String(r.Client || ''),
        industry:     String(r.Industry || ''),
        start_date:   String(r.Start_Date || ''),
        end_date:     String(r.End_Date || ''),
        entity:       String(r.Omnia_Company || ''),
      }));
      return JSON.stringify(projected.length ? projected : { message: 'No project rows found for the given filters.' });
    }
    if (name === 'get_pl_data') {
      if (periodMissing) return JSON.stringify({ error: periodRequiredMsg });
      // Fernway Data rows before the standard-source cutoff are dropped —
      // its canonical numbers for those months live in data_elysium_pl_monthly
      // instead (call get_elysium_pl_monthly_data for that range).
      const rows = byEntityPeriod(readTab(TABS.pl), 'Omnia_Company', 'month')
        .filter(r => String(r.Omnia_Company || '') !== 'Fernway Data' || periodOf(r.month) >= FERNWAY_STANDARD_SOURCE_CUTOFF_PERIOD);
      const projected = rows.map(r => ({
        period:                              periodOf(r.month),
        fiscal_year:                         String(r.fiscal_year || ''),
        entity:                              String(r.Omnia_Company || ''),
        direct_costs_actual:                 Number(r.direct_costs_actual) || null,
        direct_costs_forecast:               Number(r.direct_costs_forecast) || null,
        wages_salaries_indirect_actual:      Number(r.wages_salaries_indirect_actual) || null,
        wages_salaries_indirect_forecast:    Number(r.wages_salaries_indirect_forecast) || null,
        employee_benefits_actual:            Number(r.employee_benefits_actual) || null,
        employee_benefits_forecast:          Number(r.employee_benefits_forecast) || null,
        sales_marketing_actual:              Number(r.sales_marketing_actual) || null,
        sales_marketing_forecast:            Number(r.sales_marketing_forecast) || null,
        travel_expenses_actual:              Number(r.travel_expenses_actual) || null,
        travel_expenses_forecast:            Number(r.travel_expenses_forecast) || null,
        administrative_expenses_actual:      Number(r.administrative_expenses_actual) || null,
        administrative_expenses_forecast:    Number(r.administrative_expenses_forecast) || null,
        rent_expense_actual:                 Number(r.rent_expense_actual) || null,
        rent_expense_forecast:               Number(r.rent_expense_forecast) || null,
        recruitment_expense_actual:          Number(r.recruitment_expense_actual) || null,
        recruitment_expense_forecast:        Number(r.recruitment_expense_forecast) || null,
        other_operating_expenses_actual:     Number(r.other_operating_expenses_actual) || null,
        other_operating_expenses_forecast:   Number(r.other_operating_expenses_forecast) || null,
        operating_expenses_actual:           Number(r.operating_expenses_actual) || null,
        operating_expenses_forecast:         Number(r.operating_expenses_forecast) || null,
        revenue_actual:                      Number(r.revenue_actual) || null,
        revenue_budget:                      Number(r.revenue_budget) || null,
        gp_budget:                           Number(r.gp_budget) || null,
        ebitda_actual:                       Number(r.ebitda_actual) || null,
        ebitda_budget:                       Number(r.ebitda_budget) || null,
        utilisation_budget:                  Number(r.utilisation_budget) || null,
        headcount_budget:                    Number(r.headcount_budget) || null,
        daily_rate_budget:                   Number(r.daily_rate_budget) || null,
        gross_margin_actual:                 Number(r.gross_margin_actual) || null,
        gross_margin_forecast:               Number(r.gross_margin_forecast) || null,
        operating_margin_actual:             Number(r.operating_margin_actual) || null,
        operating_margin_forecast:           Number(r.operating_margin_forecast) || null,
      }));
      if (input.group_by) {
        const result = aggregateByGroup(
          projected,
          input.group_by,
          ['period', 'fiscal_year', 'entity'],
          ['direct_costs_actual', 'direct_costs_forecast', 'wages_salaries_indirect_actual', 'wages_salaries_indirect_forecast',
           'employee_benefits_actual', 'employee_benefits_forecast', 'sales_marketing_actual', 'sales_marketing_forecast',
           'travel_expenses_actual', 'travel_expenses_forecast', 'administrative_expenses_actual', 'administrative_expenses_forecast',
           'rent_expense_actual', 'rent_expense_forecast', 'recruitment_expense_actual', 'recruitment_expense_forecast',
           'other_operating_expenses_actual', 'other_operating_expenses_forecast', 'operating_expenses_actual', 'operating_expenses_forecast',
           'revenue_actual', 'revenue_budget', 'gp_budget', 'ebitda_actual', 'ebitda_budget', 'headcount_budget'],
          ['utilisation_budget', 'daily_rate_budget', 'gross_margin_actual', 'gross_margin_forecast', 'operating_margin_actual', 'operating_margin_forecast']
        );
        if (result.error) return JSON.stringify({ error: result.error });
        return JSON.stringify(result.aggregated.length ? result.aggregated : { message: 'No P&L rows found for the given filters. If this was for Fernway Data before ' + FERNWAY_STANDARD_SOURCE_CUTOFF_PERIOD + ', call get_elysium_pl_monthly_data instead.' });
      }
      return JSON.stringify(projected.length ? projected : { message: 'No P&L rows found for the given filters. If this was for Fernway Data before ' + FERNWAY_STANDARD_SOURCE_CUTOFF_PERIOD + ', call get_elysium_pl_monthly_data instead.' });
    }
    if (name === 'get_elysium_pl_monthly_data') {
      if (periodMissing) return JSON.stringify({ error: periodRequiredMsg });
      // Only periods before the standard-source cutoff — from that period
      // onward Elysium's numbers come from data_pl instead (get_pl_data).
      let rows = byEntityPeriod(readTab(TABS.elysiumPlMonthly), 'Omnia_Company', 'month_date')
        .filter(r => periodOf(r.month_date) < FERNWAY_STANDARD_SOURCE_CUTOFF_PERIOD);
      if (input.type) rows = rows.filter(r => String(r.type || '').toLowerCase() === String(input.type).toLowerCase());
      const projected = rows.map(r => ({
        period:                                periodOf(r.month_date),
        fiscal_year:                           String(r.fiscal_year || ''),
        type:                                  String(r.type || ''),
        entity:                                String(r.Omnia_Company || ''),
        revenue:                               Number(r.revenue) || null,
        direct_costs:                          Number(r.direct_costs) || null,
        gross_profit:                          Number(r.gross_profit) || null,
        wages_salaries_indirect:               Number(r.wages_salaries_indirect) || null,
        employee_benefits:                     Number(r.employee_benefits) || null,
        sales_marketing:                       Number(r.sales_marketing) || null,
        travel_expenses:                       Number(r.travel_expenses) || null,
        administrative_expenses:               Number(r.administrative_expenses) || null,
        rent_expense:                          Number(r.rent_expense) || null,
        recruitment_expense:                   Number(r.recruitment_expense) || null,
        other_operating_expenses:              Number(r.other_operating_expenses) || null,
        operating_expenses:                    Number(r.operating_expenses) || null,
        normalised_ebitda:                     Number(r.normalised_ebitda) || null,
        gross_profit_margin_pct:               Number(r.gross_profit_margin_pct) || null,
        normalised_ebitda_margin_pct:          Number(r.normalised_ebitda_margin_pct) || null,
        billable_fte:                          Number(r.billable_fte) || null,
        total_fte:                             Number(r.total_fte) || null,
        average_day_rate:                      Number(r.average_day_rate) || null,
        billable_employee_utilisation_pct:     Number(r.billable_employee_utilisation_pct) || null,
        nwd:                                   Number(r.nwd) || null,
      }));
      return JSON.stringify(projected.length ? projected : { message: 'No Elysium monthly P&L rows found for the given filters. If this was for a period from ' + FERNWAY_STANDARD_SOURCE_CUTOFF_PERIOD + ' onward, call get_pl_data instead.' });
    }
    // if (name === 'get_elysium_pl_annual_data') {
    //   let rows = readTab(TABS.elysiumPlAnnual);
    //   if (entity) rows = rows.filter(r => String(r.Omnia_Company || '') === entity);
    //   if (input.fiscal_year) rows = rows.filter(r => String(r.fiscal_year || '') === String(input.fiscal_year));
    //   const projected = rows.map(r => ({
    //     fiscal_year:               String(r.fiscal_year || ''),
    //     entity:                    String(r.Omnia_Company || ''),
    //     revenue:                   Number(r.revenue) || null,
    //     direct_costs:              Number(r.direct_costs) || null,
    //     gross_profit:              Number(r.gross_profit) || null,
    //     wages_salaries_indirect:   Number(r.wages_salaries_indirect) || null,
    //     employee_benefits:         Number(r.employee_benefits) || null,
    //     sales_marketing:           Number(r.sales_marketing) || null,
    //     travel_expenses:           Number(r.travel_expenses) || null,
    //     administrative_expenses:   Number(r.administrative_expenses) || null,
    //     rent_expense:              Number(r.rent_expense) || null,
    //     recruitment_expense:       Number(r.recruitment_expense) || null,
    //     other_operating_expenses:  Number(r.other_operating_expenses) || null,
    //     operating_expenses:        Number(r.operating_expenses) || null,
    //     ebitda:                    Number(r.ebitda) || null,
    //     revenue_growth_yoy_pct:    Number(r.revenue_growth_yoy_pct) || null,
    //     gross_profit_margin_pct:   Number(r.gross_profit_margin_pct) || null,
    //     ebitda_margin_pct:         Number(r.ebitda_margin_pct) || null,
    //     average_billable_fte:     Number(r.average_billable_fte) || null,
    //     average_total_fte:        Number(r.average_total_fte) || null,
    //     average_day_rate:         Number(r.average_day_rate) || null,
    //     blended_utilisation_pct:  Number(r.blended_utilisation_pct) || null,
    //     nwd:                      Number(r.nwd) || null,
    //     indirect_pct:             Number(r.indirect_pct) || null,
    //   }));
    //   return JSON.stringify(bounded(projected, 'No Elysium annual P&L rows found for the given filters.'));
    // }
    // if (name === 'get_elysium_pl_byline_monthly_data') {
    //   if (periodMissing) return JSON.stringify({ error: periodRequiredMsg });
    //   let rows = byEntityPeriod(readTab(TABS.elysiumPlByLineMonthly), 'Omnia_Company', 'month_date');
    //   if (input.type) rows = rows.filter(r => String(r.type || '').toLowerCase() === String(input.type).toLowerCase());
    //   const projected = rows.map(r => ({
    //     period:                                periodOf(r.month_date),
    //     fiscal_year:                           String(r.fiscal_year || ''),
    //     type:                                  String(r.type || ''),
    //     entity:                                String(r.Omnia_Company || ''),
    //     professional_services_revenue:         Number(r.professional_services_revenue) || null,
    //     contract_placements_revenue:           Number(r.contract_placements_revenue) || null,
    //     permanent_recruitment_revenue:         Number(r.permanent_recruitment_revenue) || null,
    //     other_revenue:                         Number(r.other_revenue) || null,
    //     total_revenue:                         Number(r.total_revenue) || null,
    //     professional_services_direct_costs:    Number(r.professional_services_direct_costs) || null,
    //     contract_placements_direct_costs:      Number(r.contract_placements_direct_costs) || null,
    //     permanent_recruitment_direct_costs:    Number(r.permanent_recruitment_direct_costs) || null,
    //     other_direct_costs:                    Number(r.other_direct_costs) || null,
    //     direct_costs:                          Number(r.direct_costs) || null,
    //     gross_profit:                          Number(r.gross_profit) || null,
    //     gross_profit_margin_pct:               Number(r.gross_profit_margin_pct) || null,
    //     wages_salaries_indirect:               Number(r.wages_salaries_indirect) || null,
    //     employee_benefits:                     Number(r.employee_benefits) || null,
    //     sales_marketing:                       Number(r.sales_marketing) || null,
    //     travel_expenses:                       Number(r.travel_expenses) || null,
    //     administrative_expenses:               Number(r.administrative_expenses) || null,
    //     rent_expense:                          Number(r.rent_expense) || null,
    //     recruitment_expense:                   Number(r.recruitment_expense) || null,
    //     other_operating_expenses:              Number(r.other_operating_expenses) || null,
    //     operating_expenses:                    Number(r.operating_expenses) || null,
    //     reported_ebitda:                       Number(r.reported_ebitda) || null,
    //     ebitda_margin_pct:                     Number(r.ebitda_margin_pct) || null,
    //     depreciation:                          Number(r.depreciation) || null,
    //     amortisation:                          Number(r.amortisation) || null,
    //     reported_ebit:                         Number(r.reported_ebit) || null,
    //     interest_income:                       Number(r.interest_income) || null,
    //     interest_expense:                      Number(r.interest_expense) || null,
    //     other_expenses:                        Number(r.other_expenses) || null,
    //     other_income:                          Number(r.other_income) || null,
    //     net_profit_before_tax:                 Number(r.net_profit_before_tax) || null,
    //     income_tax:                            Number(r.income_tax) || null,
    //     net_profit_after_tax:                  Number(r.net_profit_after_tax) || null,
    //   }));
    //   return JSON.stringify(projected.length ? projected : { message: 'No Elysium by-line monthly P&L rows found for the given filters.' });
    // }
    // if (name === 'get_elysium_pl_byline_annual_data') {
    //   let rows = readTab(TABS.elysiumPlByLineAnnual);
    //   if (entity) rows = rows.filter(r => String(r.Omnia_Company || '') === entity);
    //   if (input.fiscal_year) rows = rows.filter(r => String(r.fiscal_year || '') === String(input.fiscal_year));
    //   if (input.type)        rows = rows.filter(r => String(r.type || '').toLowerCase() === String(input.type).toLowerCase());
    //   const projected = rows.map(r => ({
    //     fiscal_year:                           String(r.fiscal_year || ''),
    //     type:                                  String(r.type || ''),
    //     entity:                                String(r.Omnia_Company || ''),
    //     professional_services_revenue:         Number(r.professional_services_revenue) || null,
    //     contract_placements_revenue:           Number(r.contract_placements_revenue) || null,
    //     permanent_recruitment_revenue:         Number(r.permanent_recruitment_revenue) || null,
    //     other_revenue:                         Number(r.other_revenue) || null,
    //     total_revenue:                         Number(r.total_revenue) || null,
    //     professional_services_direct_costs:    Number(r.professional_services_direct_costs) || null,
    //     contract_placements_direct_costs:      Number(r.contract_placements_direct_costs) || null,
    //     permanent_recruitment_direct_costs:    Number(r.permanent_recruitment_direct_costs) || null,
    //     other_direct_costs:                    Number(r.other_direct_costs) || null,
    //     direct_costs:                          Number(r.direct_costs) || null,
    //     gross_profit:                          Number(r.gross_profit) || null,
    //     gross_profit_margin_pct:               Number(r.gross_profit_margin_pct) || null,
    //     wages_salaries_indirect:               Number(r.wages_salaries_indirect) || null,
    //     employee_benefits:                     Number(r.employee_benefits) || null,
    //     sales_marketing:                       Number(r.sales_marketing) || null,
    //     travel_expenses:                       Number(r.travel_expenses) || null,
    //     administrative_expenses:               Number(r.administrative_expenses) || null,
    //     rent_expense:                          Number(r.rent_expense) || null,
    //     recruitment_expense:                   Number(r.recruitment_expense) || null,
    //     other_operating_expenses:              Number(r.other_operating_expenses) || null,
    //     operating_expenses:                    Number(r.operating_expenses) || null,
    //     reported_ebitda:                       Number(r.reported_ebitda) || null,
    //     ebitda_margin_pct:                     Number(r.ebitda_margin_pct) || null,
    //     depreciation:                          Number(r.depreciation) || null,
    //     amortisation:                          Number(r.amortisation) || null,
    //     reported_ebit:                         Number(r.reported_ebit) || null,
    //     interest_income:                       Number(r.interest_income) || null,
    //     interest_expense:                      Number(r.interest_expense) || null,
    //     other_expenses:                        Number(r.other_expenses) || null,
    //     other_income:                          Number(r.other_income) || null,
    //     net_profit_before_tax:                 Number(r.net_profit_before_tax) || null,
    //     income_tax:                            Number(r.income_tax) || null,
    //     net_profit_after_tax:                  Number(r.net_profit_after_tax) || null,
    //   }));
    //   return JSON.stringify(bounded(projected, 'No Elysium by-line annual P&L rows found for the given filters.'));
    // }
      if (name === 'calculate') {
      const expr = String(input.expression || '').trim();
      if (!expr) return JSON.stringify({ error: 'expression required' });
      // Allowlist digits/operators/parens/decimal/whitespace only — Function()
      // below only ever sees a string matching this, so it can't execute
      // arbitrary code even though the model's input drives the expression.
      if (!/^[\d+\-*/().\s]+$/.test(expr)) {
        return JSON.stringify({ error: 'Invalid expression — only digits, + - * / ( ) and decimal points are allowed, no variables or function calls.' });
      }
      try {
        const result = Function('"use strict"; return (' + expr + ');')();
        if (typeof result !== 'number' || !isFinite(result)) {
          return JSON.stringify({ error: 'Expression did not evaluate to a finite number.' });
        }
        return JSON.stringify({ expression: expr, result: result });
      } catch (e) {
        return JSON.stringify({ error: 'Could not evaluate expression: ' + String(e && e.message || e) });
      }
    }
    return JSON.stringify({ error: `Unknown tool: ${name}` });
  } catch (err) {
    return JSON.stringify({ error: String(err && err.message || err) });
  }
}

// ---------- Monitoring ----------

// Returns a JSON snapshot of app health — used by /healthcheck route and runDailyMonitor().
function runHealthCheck() {
  const tabResults = {};
  Object.entries(TABS).forEach(function(entry) {
    var key = entry[0], tabName = entry[1];
    try {
      var rows = readTab(tabName);
      tabResults[key] = { ok: true, rows: rows.length, tab: tabName };
    } catch(e) {
      tabResults[key] = { ok: false, rows: 0, tab: tabName, error: e.message };
    }
  });

  var latestPeriod = null;
  var daysSinceUpdate = null;
  try {
    var target = readTab(TABS.revenueTarget);
    var periods = target.map(function(r) { return periodOf(r.month_date); }).filter(Boolean).sort();
    if (periods.length) {
      latestPeriod = periods[periods.length - 1];
      var parts = latestPeriod.split('-').map(Number);
      // Anchor to first day of the FOLLOWING month so the 40-day clock starts when the period closes.
      var latestDate = new Date(parts[0], parts[1], 1);
      daysSinceUpdate = Math.floor((new Date() - latestDate) / 86400000);
    }
  } catch(e) {}

  var hasApiKey = !!PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  var allTabsOk = Object.values(tabResults).every(function(r) { return r.ok; });
  var dataFresh = daysSinceUpdate !== null && daysSinceUpdate < 40;

  var aiBudget = null;
  try { aiBudget = aiBudgetStatus_(); } catch (e) {}

  return {
    status:          (allTabsOk && dataFresh) ? 'ok' : 'error',
    timestamp:       new Date().toISOString(),
    latestPeriod:    latestPeriod,
    daysSinceUpdate: daysSinceUpdate,
    hasApiKey:       hasApiKey,
    aiBudget:        aiBudget,
    tabs:            tabResults,
  };
}

// Rules-based data sanity checks over the new data model. Returns
// { issues: [], warnings: [] }. First pass — field names/thresholds should
// be verified against the real sheets and adjusted (config-driven via the
// config tab: check_gp_min, check_gp_max, check_util_max, check_mom_drop_pct).
function runDataChecks() {
  var issues = [];
  var warnings = [];

  var cfgMap = {};
  try { cfgMap = configMap(); } catch(e) {}
  var cfgNum = function(key, def) {
    var v = cfgMap[key];
    return (v !== undefined && v !== '' && v !== null) ? Number(v) : def;
  };
  var GP_MIN   = cfgNum('check_gp_min', 0);
  var GP_MAX   = cfgNum('check_gp_max', 80);
  var UTIL_MAX = cfgNum('check_util_max', 110);
  var MOM_DROP = cfgNum('check_mom_drop_pct', 40) / 100;

  // Scope every check below to actuals: FY25 onward, up to the current
  // reporting period (config-driven via current_period, falling back to this
  // month). Without this, budget/forecast rows for future periods — which
  // carry a target revenue but no real cost/margin yet — get flagged as
  // GP%/utilisation anomalies or bogus MoM revenue drops.
  var FY_START  = '2024-07';
  var FY_END    = currentReportingPeriod();
  var inFyScope = function(p) { return p >= FY_START && p <= FY_END; };

  var ENTITIES = ['Solace Advisory', 'Kestrel Group', 'Northlight AI', 'Fernway Data'];

  // Revenue sanity + month-over-month drop, per OpCo, from data_revenue_target.
  // Fernway Data has no rows here before FERNWAY_STANDARD_SOURCE_CUTOFF_PERIOD
  // (see TABS.revenueTarget usage elsewhere) — it just has fewer periods to check
  // until then, not an error.
  try {
    var target = readTab(TABS.revenueTarget);
    ENTITIES.forEach(function(entity) {
      var byPeriod = {};
      target.filter(function(r) { return String(r.entity || '') === entity; }).forEach(function(r) {
        var p = periodOf(r.month_date);
        if (!p) return;
        byPeriod[p] = (byPeriod[p] || 0) + (Number(r.total_revenue) || 0);
      });
      if (Object.keys(byPeriod).length === 0) {
        if (entity !== 'Fernway Data') issues.push('data_revenue_target has no ' + entity + ' rows');
        return;
      }
      var periods = Object.keys(byPeriod).sort().filter(inFyScope);
      for (var i = 1; i < periods.length; i++) {
        var prev = byPeriod[periods[i-1]], curr = byPeriod[periods[i]];
        if (prev > 0 && curr > 0) {
          var change = (curr - prev) / prev;
          if (change < -MOM_DROP) {
            warnings.push(entity + ': revenue dropped ' + Math.abs(change * 100).toFixed(0) + '% from ' + periods[i-1] + ' to ' + periods[i]);
          }
        }
      }
    });
  } catch (e) {
    issues.push('Failed to check data_revenue_target: ' + e.message);
  }

  // GP% bounds, per OpCo, computed as (revenue_actual + direct_costs_actual) /
  // revenue_actual from data_pl — direct_costs_actual is stored as a NEGATIVE
  // number in this sheet, so it's added, not subtracted. Company-level P&L
  // actuals, one row per entity+month (summed defensively in case of multiple
  // rows). This is the agreed GP formula; data_key_metrics_monthly's own
  // `margin` field is NOT used here since it can carry stale/placeholder
  // values for periods without real cost data yet. Rows with a blank/null
  // direct_costs_actual (cost not entered yet for that period) are skipped
  // entirely rather than treated as a $0 cost — folding them in as 0 would
  // silently inflate GP% to ~100%. Fernway Data periods before
  // FERNWAY_STANDARD_SOURCE_CUTOFF_PERIOD have no rows in data_pl — those
  // come from its dedicated data_elysium_pl_monthly instead, using its own
  // gross_profit column directly rather than re-deriving it (sidesteps
  // needing to know that table's own revenue/cost sign convention) — same
  // blank/null skip applies there.
  var isBlank = function(v) { return v === '' || v === null || v === undefined; };
  try {
    var pl = readTab(TABS.pl);
    var elysiumPlGp = readTab(TABS.elysiumPlMonthly)
      .filter(function(r) { return String(r.type || '').toLowerCase() === 'actual'; });

    ENTITIES.forEach(function(entity) {
      var gpByPeriod = {};
      pl.filter(function(r) { return String(r.Omnia_Company || '') === entity; }).forEach(function(r) {
        var p = periodOf(r.month);
        if (!p || isBlank(r.direct_costs_actual)) return;
        gpByPeriod[p] = gpByPeriod[p] || { rev: 0, gp: 0 };
        var rev = Number(r.revenue_actual) || 0;
        gpByPeriod[p].rev += rev;
        gpByPeriod[p].gp  += rev + (Number(r.direct_costs_actual) || 0);
      });

      if (entity === 'Fernway Data') {
        elysiumPlGp.forEach(function(r) {
          var p = periodOf(r.month_date);
          if (!p || p >= FERNWAY_STANDARD_SOURCE_CUTOFF_PERIOD || isBlank(r.direct_costs)) return; // data_pl owns this period instead
          gpByPeriod[p] = { rev: Number(r.revenue) || 0, gp: Number(r.gross_profit) || 0 };
        });
      }

      if (Object.keys(gpByPeriod).length === 0) {
        warnings.push('No P&L data found for ' + entity + ' — GP% check skipped');
        return;
      }

      Object.keys(gpByPeriod).filter(inFyScope).forEach(function(p) {
        var b = gpByPeriod[p];
        if (b.rev <= 0) return;
        var gpPct = (b.gp / b.rev) * 100;
        if (gpPct < GP_MIN || gpPct > GP_MAX) {
          issues.push(entity + ' ' + p + ': GP% is ' + gpPct.toFixed(1) + '% (expected ' + GP_MIN + '–' + GP_MAX + '%)');
        }
      });
    });
  } catch (e) {
    warnings.push('Could not check GP% from data_pl: ' + e.message);
  }

  // Utilisation bounds, per OpCo, from data_key_metrics_monthly (the
  // centralised source of truth for this headline number). Revenue-weighted
  // across rows (the table is per entity+division, so a period can have
  // several rows). asPercent() normalises utilisation_pct whether the sheet
  // stores it as a 0-1 fraction or a 0-100 number. Fernway Data periods
  // before FERNWAY_STANDARD_SOURCE_CUTOFF_PERIOD have no rows here — those
  // come from its dedicated data_elysium_pl_monthly instead (type='Actual'
  // rows only).
  try {
    var km = readTab(TABS.keyMetrics);
    var elysiumPlUtil = readTab(TABS.elysiumPlMonthly)
      .filter(function(r) { return String(r.type || '').toLowerCase() === 'actual'; });

    ENTITIES.forEach(function(entity) {
      var kmByPeriod = {};
      km.filter(function(r) { return String(r.entity || '') === entity; }).forEach(function(r) {
        var p = periodOf(r.month_date || r.period);
        if (!p) return;
        kmByPeriod[p] = kmByPeriod[p] || { rev: 0, utilRevSum: 0 };
        var rev = Number(r.billable_revenue) || 0;
        kmByPeriod[p].rev        += rev;
        kmByPeriod[p].utilRevSum += asPercent(r.utilisation_pct) * rev;
      });

      if (entity === 'Fernway Data') {
        elysiumPlUtil.forEach(function(r) {
          var p = periodOf(r.month_date);
          if (!p || p >= FERNWAY_STANDARD_SOURCE_CUTOFF_PERIOD) return; // standard source owns this period instead
          var rev = Number(r.revenue) || 0;
          kmByPeriod[p] = {
            rev:        rev,
            utilRevSum: asPercent(r.billable_employee_utilisation_pct) * rev,
          };
        });
      }

      if (Object.keys(kmByPeriod).length === 0) {
        warnings.push('No key-metrics data found for ' + entity + ' — utilisation check skipped');
        return;
      }

      Object.keys(kmByPeriod).filter(inFyScope).forEach(function(p) {
        var b = kmByPeriod[p];
        if (b.rev <= 0) return;
        var utilPct = b.utilRevSum / b.rev;
        if (utilPct < 0 || utilPct > UTIL_MAX) {
          issues.push(entity + ' ' + p + ': utilisation is ' + utilPct.toFixed(1) + '% (expected 0–' + UTIL_MAX + '%)');
        }
      });
    });
  } catch (e) {
    warnings.push('Could not check utilisation from data_key_metrics_monthly: ' + e.message);
  }

  // Secondary, exact cross-check of utilisation from data_utilisation
  // (invoiced_days/available_days) — precise per-employee source, independent
  // of the key-metrics rollup above. Covers Solace Advisory/Kestrel Group/Northlight AI for all
  // periods, Fernway Data only from FERNWAY_STANDARD_SOURCE_CUTOFF_PERIOD
  // onward (see TABS.utilisation comment) — its earlier months are already
  // covered by the elysiumPl-derived check above.
  try {
    var util = readTab(TABS.utilisation);
    ENTITIES.forEach(function(entity) {
      var byPeriodUtil = {};
      util.filter(function(r) { return String(r.Omnia_Company || '') === entity; }).forEach(function(r) {
        var p = periodOf(r.month);
        if (!p) return;
        byPeriodUtil[p] = byPeriodUtil[p] || { billed: 0, max: 0 };
        byPeriodUtil[p].billed += Number(r.invoiced_days)   || 0;
        byPeriodUtil[p].max    += Number(r.available_days)  || 0;
      });
      Object.keys(byPeriodUtil).filter(inFyScope).forEach(function(p) {
        var b = byPeriodUtil[p];
        if (b.max <= 0) return;
        var utilPct = (b.billed / b.max) * 100;
        if (utilPct < 0 || utilPct > UTIL_MAX) {
          issues.push(entity + ' ' + p + ': utilisation (from data_utilisation) is ' + utilPct.toFixed(1) + '% (expected 0–' + UTIL_MAX + '%)');
        }
      });
    });
  } catch (e) {
    warnings.push('Could not check utilisation from data_utilisation: ' + e.message);
  }

  return { issues: issues, warnings: warnings };
}

// Aggregates the last N months of headline metrics for every Omnia OpCo
// (Solace Advisory, Kestrel Group, Northlight AI, Fernway Data) and asks Claude for anomaly
// commentary in one call. Returns plain text (throws on failure). Pulls
// billable_revenue/utilisation_pct/billable_fte from data_key_metrics_monthly
// (the centralised source of truth for these headline numbers), GP% from
// data_pl as (revenue_actual + direct_costs_actual) / revenue_actual —
// direct_costs_actual is stored as a NEGATIVE number, so it's added, not
// subtracted; same formula and source as runDataChecks, so the two stay
// consistent — plus average daily rate from data_powerbi_unified_fact (the only source that
// carries ADR; weighted as sum(total_revenue)/sum(total_days) per period
// since that table is project/BU-level, not a single row per period).
// Fernway Data periods before FERNWAY_STANDARD_SOURCE_CUTOFF_PERIOD have no
// rows in any of those (see ELYSIUM DIGITAL DATA ROUTING above) — those
// months are pulled straight from its dedicated data_elysium_pl_monthly instead.
function getClaudeCommentary(months) {
  months = months || 6;
  var ENTITIES = ['Solace Advisory', 'Kestrel Group', 'Northlight AI', 'Fernway Data'];
  var CURRENT_PERIOD = currentReportingPeriod();
  var isBlank = function(v) { return v === '' || v === null || v === undefined; };

  var km = readTab(TABS.keyMetrics);
  var uf = readTab(TABS.unifiedFact);
  var pl = readTab(TABS.pl);
  var elysiumPl = readTab(TABS.elysiumPlMonthly)
    .filter(function(r) { return String(r.type || '').toLowerCase() === 'actual'; });

  var sections = ENTITIES.map(function(entity) {
    var byPeriod = {};
    km.filter(function(r) { return String(r.entity || '') === entity; }).forEach(function(r) {
      var p = periodOf(r.month_date || r.period);
      if (!p) return;
      byPeriod[p] = byPeriod[p] || { revenue: 0, utilRevSum: 0, fte: 0 };
      var rev = Number(r.billable_revenue) || 0;
      byPeriod[p].revenue    += rev;
      byPeriod[p].utilRevSum += asPercent(r.utilisation_pct) * rev;
      byPeriod[p].fte         = Math.max(byPeriod[p].fte, Number(r.billable_fte) || 0);
    });

    var adrByPeriod = {};
    uf.filter(function(r) { return String(r.entity || '') === entity; }).forEach(function(r) {
      var p = periodOf(r.month_date);
      if (!p) return;
      adrByPeriod[p] = adrByPeriod[p] || { revenue: 0, days: 0 };
      adrByPeriod[p].revenue += Number(r.total_revenue) || 0;
      adrByPeriod[p].days    += Number(r.total_days)    || 0;
    });

    var gpByPeriod = {};
    pl.filter(function(r) { return String(r.Omnia_Company || '') === entity; }).forEach(function(r) {
      var p = periodOf(r.month);
      if (!p || isBlank(r.direct_costs_actual)) return; // cost not entered yet — don't treat as $0
      gpByPeriod[p] = gpByPeriod[p] || { rev: 0, gp: 0 };
      var rev = Number(r.revenue_actual) || 0;
      gpByPeriod[p].rev += rev;
      gpByPeriod[p].gp  += rev + (Number(r.direct_costs_actual) || 0); // direct_costs_actual is negative
    });

    if (entity === 'Fernway Data') {
      elysiumPl.forEach(function(r) {
        var p = periodOf(r.month_date);
        if (!p || p >= FERNWAY_STANDARD_SOURCE_CUTOFF_PERIOD || isBlank(r.direct_costs)) return; // standard sources own this period instead
        var rev = Number(r.revenue) || 0;
        byPeriod[p] = {
          revenue:    rev,
          utilRevSum: asPercent(r.billable_employee_utilisation_pct) * rev,
          fte:        Number(r.billable_fte) || 0,
        };
        adrByPeriod[p] = { revenue: Number(r.average_day_rate) || 0, days: 1 };
        gpByPeriod[p]  = { rev: rev, gp: Number(r.gross_profit) || 0 };
      });
    }

    var periods = Object.keys(byPeriod)
      .filter(function(p) { return p <= CURRENT_PERIOD; }) // drop future budget/forecast rows
      .sort()
      .slice(-months);
    var lines = periods.map(function(p) {
      var b   = byPeriod[p];
      var adr = adrByPeriod[p];
      var gp  = gpByPeriod[p];
      var gpPct   = gp && gp.rev > 0 ? ((gp.gp / gp.rev) * 100).toFixed(1) : 'N/A';
      var utilPct = b.revenue > 0 ? (b.utilRevSum / b.revenue).toFixed(1) : 'N/A';
      var adrVal  = adr && adr.days > 0 ? Math.round(adr.revenue / adr.days) : 'N/A';
      return p + ': Revenue $' + Math.round(b.revenue / 1000) + 'k, GP ' + gpPct +
        '%, HC ' + (b.fte || 'N/A') + ', Utilisation ' + utilPct +
        '%, ADR $' + adrVal;
    });
    return entity + ':\n' + (lines.join('\n') || '(no data)');
  });

  var dataStr = sections.join('\n\n');

  return runAssistantLoop(
    'You are a data quality analyst for Omnia, a group of Australian professional-services OpCos ' +
    '(Solace Advisory, Kestrel Group, Northlight AI, Fernway Data). Review these ' + months + ' months of business ' +
    'metrics for each OpCo and flag anything anomalous or potentially incorrect, either within an OpCo ' +
    'or across OpCos. Be concise and specific, and group findings by OpCo. If everything looks normal, say so.',
    [{ role: 'user', content: dataStr || 'No data available.' }],
    null
  );
}

// Builds and sends the monitor email to the given address.
// Called by runDailyMonitor (scheduled) and send_test_email (manual UI trigger).
// Accepts optional pre-computed health/dataChecks to avoid redundant sheet reads.
function sendMonitorEmail(toEmail, health, dataChecks) {
  health     = health     || runHealthCheck();
  dataChecks = dataChecks || runDataChecks();
  var overallOk  = health.status === 'ok' && dataChecks.issues.length === 0;

  var claudeCommentary = '(not available)';
  try {
    claudeCommentary = getClaudeCommentary(6);
  } catch(e) {
    claudeCommentary = 'Claude call failed: ' + e.message;
  }

  var ts = new Date().toLocaleString('en-AU', { timeZone: 'Australia/Sydney' });
  var statusColor = overallOk ? '#2e7d32' : '#c62828';
  var statusLabel = overallOk ? '✅ All Clear' : '❌ Issues Detected';

  var html = '<div style="font-family:sans-serif;max-width:680px">';
  html += '<h2 style="margin-bottom:4px">Omnia Reporting — Daily Health Check</h2>';
  html += '<p style="margin-top:0;color:#555">' + ts + '</p>';
  html += '<p><b>Status:</b> <span style="color:' + statusColor + '">' + statusLabel + '</span></p>';
  html += '<hr>';

  html += '<h3>Sheet Tabs</h3>';
  html += '<table border="1" cellpadding="5" cellspacing="0" style="border-collapse:collapse;width:100%">';
  html += '<tr style="background:#f5f5f5"><th>Key</th><th>Tab</th><th>Status</th><th>Rows</th></tr>';
  Object.entries(health.tabs).forEach(function(entry) {
    var key = entry[0], info = entry[1];
    var c = info.ok ? '#2e7d32' : '#c62828';
    var s = info.ok ? '✅ OK' : '❌ ' + (info.error || 'missing');
    html += '<tr><td>' + key + '</td><td>' + info.tab + '</td>' +
      '<td style="color:' + c + '">' + s + '</td><td>' + info.rows + '</td></tr>';
  });
  html += '</table>';
  html += '<p><b>Latest data period:</b> ' + (health.latestPeriod || 'unknown') +
    ' (' + (health.daysSinceUpdate !== null ? health.daysSinceUpdate + ' days ago' : 'unknown') + ')</p>';
  html += '<p><b>Anthropic API key:</b> ' + (health.hasApiKey ? '✅ Set' : '❌ Missing') + '</p>';
  html += '<hr>';

  html += '<h3>Data Quality</h3>';
  if (dataChecks.issues.length === 0 && dataChecks.warnings.length === 0) {
    html += '<p style="color:#2e7d32">✅ No issues found</p>';
  }
  if (dataChecks.issues.length > 0) {
    html += '<p><b style="color:#c62828">Errors (' + dataChecks.issues.length + '):</b></p><ul>';
    dataChecks.issues.forEach(function(i) { html += '<li style="color:#c62828">' + i + '</li>'; });
    html += '</ul>';
  }
  if (dataChecks.warnings.length > 0) {
    html += '<p><b style="color:#e65100">Warnings (' + dataChecks.warnings.length + '):</b></p><ul>';
    dataChecks.warnings.forEach(function(w) { html += '<li style="color:#e65100">' + w + '</li>'; });
    html += '</ul>';
  }
  html += '<hr>';

  html += '<h3>AI Commentary</h3>';
  html += '<p style="background:#f9f9f9;padding:12px;border-left:3px solid #bbb;margin:0">' +
    claudeCommentary.replace(/\n/g, '<br>') + '</p>';
  html += '</div>';

  MailApp.sendEmail({
    to:       toEmail,
    subject:  '[' + (overallOk ? '✅ OK' : '❌ ALERT') + '] Omnia Reporting — Daily Health Check',
    htmlBody: html,
  });

  return { ok: overallOk, issues: dataChecks.issues.length, warnings: dataChecks.warnings.length, sentTo: toEmail };
}

// Scheduled function — set a daily time trigger pointing at this (see
// installDailyTrigger() below). Also sends a Slack alert if
// SLACK_WEBHOOK_URL is set in Script Properties.
function runDailyMonitor() {
  // Compute once and pass to both sendMonitorEmail and sendSlackAlert to avoid double sheet reads.
  var health     = runHealthCheck();
  var dataChecks = runDataChecks();
  var result     = sendMonitorEmail('nathan@edgered.com.au', health, dataChecks);
  sendSlackAlert(result.ok, dataChecks, health);
  return result;
}

// One-time setup: in the Apps Script editor, select installDailyTrigger in
// the function dropdown at the top and click Run — this schedules
// runDailyMonitor() to run every morning without needing to click through
// the Triggers UI by hand. Safe to re-run: it clears any existing trigger
// for runDailyMonitor first, so re-running never creates duplicates.
function installDailyTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'runDailyMonitor') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('runDailyMonitor')
    .timeBased()
    .everyDays(1)
    .atHour(7)
    .inTimezone('Australia/Sydney')
    .create();
  Logger.log('Daily trigger installed — runDailyMonitor will run around 7am Australia/Sydney every day.');
}

// Posts a summary to Slack if SLACK_WEBHOOK_URL is set in Script Properties.
function sendSlackAlert(ok, dataChecks, health) {
  var webhookUrl = PropertiesService.getScriptProperties().getProperty('SLACK_WEBHOOK_URL');
  if (!webhookUrl) return;

  var emoji = ok ? ':white_check_mark:' : ':x:';
  var lines = [
    emoji + ' *Omnia Reporting — Daily Health Check*',
    'Status: ' + (ok ? 'All Clear' : 'Issues Detected'),
    'Latest data: ' + (health.latestPeriod || 'unknown') +
      (health.daysSinceUpdate !== null ? ' (' + health.daysSinceUpdate + ' days ago)' : ''),
  ];
  if (dataChecks.issues.length)   lines.push(':red_circle: Errors: ' + dataChecks.issues.join(' | '));
  if (dataChecks.warnings.length) lines.push(':large_orange_circle: Warnings: ' + dataChecks.warnings.join(' | '));

  try {
    UrlFetchApp.fetch(webhookUrl, {
      method:             'post',
      contentType:        'application/json',
      payload:            JSON.stringify({ text: lines.join('\n') }),
      muteHttpExceptions: true,
    });
  } catch(e) {}
}

// ---------- Self Code Review ----------

// Fetches this project's own source via the Apps Script REST API.
// Requires 'https://www.googleapis.com/auth/script.projects.readonly' in appsscript.json oauthScopes.
function fetchOwnSource() {
  var token    = ScriptApp.getOAuthToken();
  var scriptId = ScriptApp.getScriptId();
  var resp = UrlFetchApp.fetch(
    'https://script.googleapis.com/v1/projects/' + scriptId + '/content',
    { headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true }
  );
  if (resp.getResponseCode() !== 200) {
    throw new Error('Source fetch failed (' + resp.getResponseCode() + '): ' + resp.getContentText());
  }
  var files = {};
  // Apps Script names files without extension (e.g. 'Code', 'Index').
  JSON.parse(resp.getContentText()).files.forEach(function(f) { files[f.name] = f.source; });
  return files;
}
// Finds the JSON array/object in Claude's response text, even when Claude
// prefaces it with a sentence or two despite being told to return ONLY JSON
// (e.g. "...uses a [YYYY, MM] tuple format..." before the real answer).
// The old approach — text.match(/\[[\s\S]*\]/) — greedily spans from the
// FIRST openChar to the LAST closeChar in the whole response, so a stray
// bracket mentioned in prose gets glued to the real JSON's closing bracket,
// producing an unparseable mix of both (that's the "Unexpected token 'Y'"
// class of error). This instead tracks bracket depth — skipping brackets
// that appear inside "..." strings, so they don't throw off the count — and
// returns the largest complete balanced span, since the real JSON payload
// (a findings array, a verdict list, ...) is always far bigger than an
// incidental bracket in a sentence. Returns null if no balanced span exists.
function extractJsonSpan(text, openChar, closeChar) {
  var best     = null;
  var inString = false;
  var escaped  = false;
  var depth    = 0;
  var start    = -1;
  for (var i = 0; i < text.length; i++) {
    var ch = text[i];
    if (inString) {
      if (escaped)        { escaped = false; }
      else if (ch === '\\') { escaped = true; }
      else if (ch === '"')  { inString = false; }
      continue;
    }
    if (ch === '"') { inString = true; continue; }
    if (ch === openChar) {
      if (depth === 0) start = i;
      depth++;
      continue;
    }
    if (ch === closeChar && depth > 0) {
      depth--;
      if (depth === 0 && start !== -1) {
        var span = text.slice(start, i + 1);
        if (!best || span.length > best.length) best = span;
        start = -1;
      }
    }
  }
  return best;
}

// Claude's JSON output sometimes isn't valid JSON as-is: bare backslashes in
// string values (\d, \s, \url, etc. — invalid JSON escapes), literal
// newline/tab/CR characters instead of \n/\t/\r escapes (raw control
// characters inside a string are illegal in JSON and make JSON.parse throw
// "Unterminated string"), and — especially in "quote" fields that reproduce
// source code verbatim — literal " characters that Claude forgot to escape
// as \", which prematurely end the JSON string. Walks the text tracking
// whether it's inside a "..." string and fixes all three only there, so
// whitespace used for JSON formatting outside strings is untouched.
function sanitizeClaudeJson(raw) {
  var out = '';
  var inString = false;
  var escaped  = false;
  for (var i = 0; i < raw.length; i++) {
    var ch = raw[i];
    if (!inString) {
      if (ch === '"') inString = true;
      out += ch;
      continue;
    }
    if (escaped) {
      escaped = false;
      if (ch === 'u') {
        var hex = raw.substr(i + 1, 4);
        out += /^[0-9a-fA-F]{4}$/.test(hex) ? '\\u' : '\\\\u';
      } else if ('"\\/bfnrt'.indexOf(ch) !== -1) {
        out += '\\' + ch;
      } else {
        out += '\\\\' + ch; // invalid escape sequence — escape the backslash itself
      }
      continue;
    }
    if (ch === '\\') { escaped = true; continue; }
    if (ch === '"') {
      // Look past trailing whitespace: a real string terminator is followed
      // by a JSON structural character (or end of input). Anything else
      // means this quote is literal content that Claude failed to escape.
      var j = i + 1;
      while (j < raw.length && /\s/.test(raw[j])) j++;
      var next = raw[j];
      if (next === undefined || ',}]:'.indexOf(next) !== -1) {
        inString = false;
        out += ch;
      } else {
        out += '\\"';
      }
      continue;
    }
    if (ch === '\n') { out += '\\n'; continue; }
    if (ch === '\r') { out += '\\r'; continue; }
    if (ch === '\t') { out += '\\t'; continue; }
    out += ch;
  }
  return out;
}

// Sends Code.js (and optionally Index.html) to Claude for a structured review.
// Returns { findings: [...], reviewedAt, filesReviewed }.
function runCodeReview(includeHtml) {
  var src       = fetchCodeFromGitHub(includeHtml);
  var codeJs    = src['Code']  || '';
  var indexHtml = includeHtml ? (src['Index'] || '') : '';
  var seedJs    = src['Seed'] || '';

  var apiKey = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!apiKey) throw new Error('ANTHROPIC_API_KEY not set');

  var content = '=== Code.js ===\n' + codeJs;
  if (indexHtml) content += '\n\n=== Index.html ===\n' + indexHtml;

  if (seedJs) content += '\n\n=== Seed.js (reference only — canonical config keys, not in scope for findings) ===\n' + seedJs;

  // Forces the model to explicitly consider every top-level function instead of
  // letting a single holistic pass over the whole file randomly "notice" only
  // some of them — without this, two real bugs in two different functions can
  // each show up in only one of two otherwise-identical repeat calls, and which
  // ones surface becomes a matter of sampling luck rather than review quality.
  var codeFuncs = findFunctionBlocks(codeJs).blocks.map(function(b) { return 'Code.js:' + b.name; });
  var htmlFuncs = indexHtml ? findFunctionBlocks(indexHtml).blocks.map(function(b) { return 'Index.html:' + b.name; }) : [];
  var allFuncs  = codeFuncs.concat(htmlFuncs);
  var checklist =
    '\nBefore producing the findings array, go through this exact list of functions one by one and explicitly ' +
    'check each of them against every focus area above — do not skip any, and do not let one function\'s ' +
    'issues distract you from checking the rest:\n' +
    allFuncs.map(function(f) { return '- ' + f; }).join('\n') + '\n' +
    '(It is fine — expected, even — for most of these to have no real issue. Only include an entry in the ' +
    'output array for functions where you found an actual, quotable problem.)\n\n';

  var prompt =
    'You are a senior Google Apps Script engineer reviewing the Omnia Reporting web app for Solace Advisory FY26.\n\n' +
    'Review the code for real bugs and security issues only. Focus on:\n' +
    '- Logic bugs: wrong calculations, off-by-one errors, missing return statements\n' +
    '- Date/period bugs: Date object vs "YYYY-MM" string mismatches\n' +
    '- Access control: routes that write or delete data without requireRole()\n' +
    '- Injection risks: unsanitised input written to sheets or email HTML\n' +
    '- Config consistency: cfgMap keys with no fallback, wrong TABS references\n' +
    '- Null dereferences that crash a route\n\n' +
    'Before finalizing findings, apply these checks — the bug classes below are the ones most often missed ' +
    'by a holistic read-through because nothing crashes and nothing looks obviously wrong at a glance:\n' +
    '- Do not lower scrutiny for routes/functions commented as test-only, dev-only, or debug — review them ' +
    'with the same rigor as production code.\n' +
    '- For any function whose comment states an intended behavior (e.g. "anchor to the first day of the ' +
    'following month", a boundary condition, a date rule), explicitly restate what the comment promises, ' +
    'then check the code line-by-line against that restatement — do not just judge whether the code "looks ' +
    'reasonable" on its own.\n' +
    '- Config-key and TABS-reference consistency is a mechanical set-membership check, not a plausibility ' +
    'judgment — do it in two explicit steps, in order:\n' +
    '  1) Extract the canonical list of valid identifiers verbatim from their defining source: every key in ' +
    'the CONFIG_ROWS array (Seed.js) for cfgMap()/cfgNum() lookups, and every key in the TABS object literal ' +
    '(Code.js) for TABS.* references.\n' +
    '  2) Then list every cfgMap()/cfgNum() call and every TABS.* reference in the reviewed code side by side ' +
    'with its literal key string, and check whether that exact string is present in the canonical list from ' +
    'step 1. Do not judge whether a key "looks like" a real one — a key that is plausible but absent from the ' +
    'canonical list is exactly the bug this check exists to catch, and it reads as correct on a normal pass.\n\n' +
    checklist +
    'Return ONLY a JSON array (no other text, no markdown fences) with up to 15 real findings. ' +
    'Only report a finding if you can back it with an exact quote from the code below — never describe ' +
    'code from memory or assumption:\n' +
    '[\n  {\n    "severity": "critical|high|medium|low",\n    "file": "Code.js|Index.html",\n' +
    '    "location": "function name",\n    "title": "short title (max 60 chars)",\n' +
    '    "description": "what is wrong and why it matters",\n' +
    '    "quote": "the exact code (verbatim, character for character, 1-6 lines) this finding is about",\n' +
    '    "suggestion": "how to fix it in plain English — no code snippets or regex patterns"\n  }\n]\n\n' +
    content;

  var resp = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method:             'post',
    contentType:        'application/json',
    headers: { 'x-api-key': apiKey, 'anthropic-version': '2023-06-01' },
    payload:            JSON.stringify({
      model:       CLAUDE_MODEL,
      max_tokens:  8192,
      // Deterministic — without this, two real bugs in two different functions
      // can each surface in only one of two otherwise-identical repeat calls,
      // purely from sampling variance rather than either finding being wrong.
      temperature: 0,
      messages:    [{ role: 'user', content: prompt }],
    }),
    muteHttpExceptions: true,
  });

  var data = JSON.parse(resp.getContentText());
  if (resp.getResponseCode() !== 200) {
    throw new Error(data.error ? data.error.message : 'API error ' + resp.getResponseCode());
  }
  if (data.stop_reason === 'max_tokens') {
    throw new Error('Claude response was cut off by the token limit before finishing — try again with fewer findings or split the review.');
  }

  var text  = (data.content && data.content[0]) ? data.content[0].text : '';
  var match = extractJsonSpan(text, '[', ']');
  if (!match) throw new Error('Could not parse findings from Claude response');

  var findings     = JSON.parse(sanitizeClaudeJson(match));
  var verifyResult = verifyAndFilterFindings(findings, codeJs, indexHtml, apiKey);

  return {
    findings:      verifyResult.findings,
    reviewedAt:    new Date().toISOString(),
    filesReviewed: includeHtml ? ['Code.js', 'Index.html'] : ['Code.js'],
    verification:  verifyResult.report,
  };
}

// Scans fullSource for top-level `function name(` declarations and returns
// [{ name, start, end }] line ranges (end = line before next function, or EOF).
function findFunctionBlocks(fullSource) {
  var lines  = fullSource.split('\n');
  var blocks = [];
  var re     = /^\s*function\s+([A-Za-z0-9_$]+)\s*\(/;
  for (var i = 0; i < lines.length; i++) {
    var m = lines[i].match(re);
    if (m) blocks.push({ name: m[1], start: i, end: lines.length - 1 });
  }
  for (var j = 0; j < blocks.length - 1; j++) blocks[j].end = blocks[j + 1].start - 1;
  return { lines: lines, blocks: blocks };
}

// Contiguous // comment line(s) directly above a function declaration, used
// as a one-line summary of what it does when building the function index.
// Falls back to the function's own signature line if there's no comment.
function leadingCommentFor(lines, startIdx) {
  var comments = [];
  var i = startIdx - 1;
  while (i >= 0) {
    var t = lines[i].trim();
    if (t.indexOf('//') !== 0) break;
    comments.unshift(t.replace(/^\/\/\s?/, ''));
    i--;
  }
  return comments.length ? comments.join(' ') : lines[startIdx].trim();
}

// Builds a lightweight index of every top-level function in a file — just
// its name and a one-line summary, not the full body — cheap enough to hand
// an LLM in full so it can pick relevant functions by meaning rather than by
// keyword/substring matching against the file text.
// Returns { lines, blocks, index: [{ file, name, summary }] }.
function buildFunctionIndex(fileLabel, fullSource) {
  var found = findFunctionBlocks(fullSource);
  var index = found.blocks.map(function(b) {
    return { file: fileLabel, name: b.name, summary: leadingCommentFor(found.lines, b.start).slice(0, 180) };
  });
  return { lines: found.lines, blocks: found.blocks, index: index };
}

// Asks Claude to pick which function(s) — out of the whole-file index — are
// actually relevant to a free-text problem description. This replaces naive
// keyword/substring matching with semantic judgement, the same pattern
// chatProxy used previously (a cheap call to narrow scope before the
// expensive one).
// Returns { matched: [{ file, name }], note: string }.
function selectRelevantFunctionsViaLLM(indexEntries, description, apiKey) {
  if (!indexEntries.length) return { matched: [], note: '' };

  var listing = indexEntries.map(function(e) { return '- ' + e.file + ':' + e.name + ' — ' + e.summary; }).join('\n');
  var prompt =
    'Here is an index of every function in a Google Apps Script web app (name and one-line summary only, not the code):\n\n' +
    listing + '\n\n' +
    'A user described this problem/area to review:\n"' + description + '"\n\n' +
    'Pick ONLY the function(s) from the list above that are actually relevant to that description — ' +
    'do not invent function names that are not in the list. If nothing in the list is genuinely relevant, ' +
    'return an empty "matched" array and use "note" to briefly say what topic/theme the description seems to be ' +
    'about, so a full-file review can at least focus on that theme.\n\n' +
    'Return ONLY a JSON object (no other text, no markdown fences):\n' +
    '{\n  "matched": [ { "file": "Code.js|Index.html", "name": "exact function name from the list" } ],\n' +
    '  "note": "short note — why these were picked, or what theme to focus on if none matched"\n}';

  var resp = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method:             'post',
    contentType:        'application/json',
    headers: { 'x-api-key': apiKey, 'anthropic-version': '2023-06-01' },
    payload:            JSON.stringify({
      model:       CLAUDE_MODEL,
      max_tokens:  700,
      // Deterministic — this pick of relevant functions must stay stable across
      // repeat calls with the same description, otherwise a real bug in a
      // function that's randomly excluded on one run silently disappears.
      temperature: 0,
      messages:    [{ role: 'user', content: prompt }],
    }),
    muteHttpExceptions: true,
  });

  var data = JSON.parse(resp.getContentText());
  if (resp.getResponseCode() !== 200) {
    throw new Error(data.error ? data.error.message : 'API error ' + resp.getResponseCode());
  }
  var text  = (data.content && data.content[0]) ? data.content[0].text : '';
  var match = extractJsonSpan(text, '{', '}');
  if (!match) return { matched: [], note: '' };
  var parsed;
  try { parsed = JSON.parse(sanitizeClaudeJson(match)); } catch (e) { return { matched: [], note: '' }; }

  // Defensive: only trust matches that actually exist in the index we sent —
  // never let a hallucinated function name silently scope the review.
  var validKeys = {};
  indexEntries.forEach(function(e) { validKeys[e.file + ':' + e.name] = true; });
  var matched = Array.isArray(parsed.matched)
    ? parsed.matched.filter(function(m) { return m && validKeys[m.file + ':' + m.name]; })
    : [];
  return { matched: matched, note: String(parsed.note || '') };
}

// Like runCodeReview, but scoped to the function(s) an LLM judges relevant to
// a user-supplied free-text description of the problem area, instead of the
// whole file. Two calls: a cheap one over a name+summary index to pick which
// functions matter (semantic, not keyword-based), then the real review call
// scoped to just those. Falls back to a full-file review (matched:false) if
// nothing in the index was judged relevant.
function runTargetedCodeReview(description, includeHtml) {
  description = String(description || '').trim();
  if (!description) throw new Error('description is required');

  var src       = fetchCodeFromGitHub(includeHtml);
  var codeJs    = src['Code']  || '';
  var indexHtml = includeHtml ? (src['Index'] || '') : '';
  var seedJs    = src['Seed'] || '';

  var apiKey = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!apiKey) throw new Error('ANTHROPIC_API_KEY not set');

  var codeIndex = buildFunctionIndex('Code.js', codeJs);
  var htmlIndex = indexHtml ? buildFunctionIndex('Index.html', indexHtml) : { lines: [], blocks: [], index: [] };
  var combinedIndex = codeIndex.index.concat(htmlIndex.index);

  var selection = { matched: [], note: '' };
  try {
    selection = selectRelevantFunctionsViaLLM(combinedIndex, description, apiKey);
  } catch (e) {
    // Selection step failing shouldn't block the review — just fall back to full-file.
  }

  var matched = selection.matched.length > 0;

  var content, filesReviewed;
  if (matched) {
    var sectionFor = function(entry) {
      var blocks = entry.file === 'Index.html' ? htmlIndex.blocks : codeIndex.blocks;
      var lines  = entry.file === 'Index.html' ? htmlIndex.lines  : codeIndex.lines;
      var b = blocks.filter(function(x) { return x.name === entry.name; })[0];
      if (!b) return null;
      var start = Math.max(0, b.start - 3);
      var end   = Math.min(lines.length - 1, b.end);
      return {
        file: entry.file,
        name: b.name,
        text: '// lines ' + (start + 1) + '-' + (end + 1) + '\n' + lines.slice(start, end + 1).join('\n'),
      };
    };
    var sections = selection.matched.map(sectionFor).filter(Boolean);
    var codeSections = sections.filter(function(s) { return s.file === 'Code.js'; });
    var htmlSections  = sections.filter(function(s) { return s.file === 'Index.html'; });

    content = '';
    if (codeSections.length) {
      content += '=== Code.js (relevant sections only) ===\n' +
        codeSections.map(function(s) { return s.text; }).join('\n\n// ...\n\n');
    }
    if (htmlSections.length) {
      content += '\n\n=== Index.html (relevant sections only) ===\n' +
        htmlSections.map(function(s) { return s.text; }).join('\n\n// ...\n\n');
    }
    filesReviewed = sections.map(function(s) { return s.file + ':' + s.name; });
  } else {
    // Nothing in the index was judged relevant — fall back to full file(s)
    // so the user still gets a review, but the UI is told so it can say so.
    content = '=== Code.js ===\n' + codeJs;
    if (indexHtml) content += '\n\n=== Index.html ===\n' + indexHtml;
    filesReviewed = includeHtml ? ['Code.js', 'Index.html'] : ['Code.js'];
  }
  // Reference only — the canonical config-key source (CONFIG_ROWS in Seed.js) never
  // appears inside a function body, so it would otherwise be invisible to a scoped
  // review even though cfgMap()/cfgNum() calls in the scoped section depend on it.
  if (seedJs) content += '\n\n=== Seed.js (reference only — canonical config keys, not in scope for findings) ===\n' + seedJs;

  // Forces the model to explicitly consider every selected function instead of
  // letting a single holistic pass over the scoped code randomly "notice" only
  // some of them — without this, two real bugs in two different functions can
  // each show up in only one of two otherwise-identical repeat calls.
  var checklist = matched
    ? '\nBefore producing the findings array, go through this exact list of functions one by one and explicitly ' +
      'check each of them against every focus area below — do not skip any, and do not let one function\'s ' +
      'issues distract you from checking the rest:\n' +
      filesReviewed.map(function(f) { return '- ' + f; }).join('\n') + '\n' +
      '(It is fine — expected, even — for most of these to have no real issue. Only include an entry in the ' +
      'output array for functions where you found an actual, quotable problem.)\n\n'
    : '';

  var scopeGuidance;
  if (matched) {
    scopeGuidance = 'Below are the section(s) of the codebase judged most relevant to that description. ' +
      'Only review this scoped code — do not invent issues about code you cannot see.\n\n' + checklist;
  } else {
    scopeGuidance = 'No specific function in this codebase was judged relevant to that description, ' +
      'so the full file(s) are included below instead. ' +
      (selection.note ? 'A first pass suggested this theme to focus on: "' + selection.note + '". ' : '') +
      'Prioritise findings related to that theme/description where the code actually supports it — ' +
      'do not invent anything about code or functions that are not actually present below.\n\n';
  }

  var prompt =
    'You are a senior Google Apps Script engineer reviewing the Omnia Board Reporting web app for Solace Advisory FY26.\n\n' +
    'The user has described a specific problem area to focus on:\n"' + description + '"\n\n' +
    scopeGuidance +
    'Focus on:\n' +
    '- Logic bugs: wrong calculations, off-by-one errors, missing return statements\n' +
    '- Date/period bugs: Date object vs "YYYY-MM" string mismatches\n' +
    '- Access control: routes that write or delete data without requireRole()\n' +
    '- Injection risks: unsanitised input written to sheets or email HTML\n' +
    '- Config consistency: cfgMap keys with no fallback, wrong TABS references\n' +
    '- Null dereferences that crash a route\n\n' +
    'Before finalizing findings, apply these checks — the bug classes below are the ones most often missed ' +
    'by a holistic read-through because nothing crashes and nothing looks obviously wrong at a glance:\n' +
    '- Do not lower scrutiny for routes/functions commented as test-only, dev-only, or debug — review them ' +
    'with the same rigor as production code.\n' +
    '- For any function whose comment states an intended behavior (e.g. "anchor to the first day of the ' +
    'following month", a boundary condition, a date rule), explicitly restate what the comment promises, ' +
    'then check the code line-by-line against that restatement — do not just judge whether the code "looks ' +
    'reasonable" on its own.\n' +
    '- Where a function makes several near-identical calls with different string keys (e.g. multiple ' +
    'cfgNum()/cfgMap() lookups), list every such call and its key side by side, and check each key ' +
    'character-for-character against the config key it should actually reference — a one-character or ' +
    'one-suffix typo (e.g. "_max" vs "_max_pct") is a common injected bug that reads as correct on a normal pass.\n\n' +
    'Return ONLY a JSON array (no other text, no markdown fences) with up to 15 real findings. ' +
    'Only report a finding if you can back it with an exact quote from the code below — never describe ' +
    'code from memory or assumption:\n' +
    '[\n  {\n    "severity": "critical|high|medium|low",\n    "file": "Code.js|Index.html",\n' +
    '    "location": "function name",\n    "title": "short title (max 60 chars)",\n' +
    '    "description": "what is wrong and why it matters",\n' +
    '    "quote": "the exact code (verbatim, character for character, 1-6 lines) this finding is about",\n' +
    '    "suggestion": "how to fix it in plain English — no code snippets or regex patterns"\n  }\n]\n\n' +
    content;

  var resp = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method:             'post',
    contentType:        'application/json',
    headers: { 'x-api-key': apiKey, 'anthropic-version': '2023-06-01' },
    payload:            JSON.stringify({
      model:       CLAUDE_MODEL,
      max_tokens:  8192,
      // Deterministic — without this, two real bugs in two different functions
      // can each surface in only one of two otherwise-identical repeat calls,
      // purely from sampling variance rather than either finding being wrong.
      temperature: 0,
      messages:    [{ role: 'user', content: prompt }],
    }),
    muteHttpExceptions: true,
  });

  var data = JSON.parse(resp.getContentText());
  if (resp.getResponseCode() !== 200) {
    throw new Error(data.error ? data.error.message : 'API error ' + resp.getResponseCode());
  }
  if (data.stop_reason === 'max_tokens') {
    throw new Error('Claude response was cut off by the token limit before finishing — try again with fewer findings or split the review.');
  }

  var text  = (data.content && data.content[0]) ? data.content[0].text : '';
  var match = extractJsonSpan(text, '[', ']');
  if (!match) throw new Error('Could not parse findings from Claude response');

  var findings     = JSON.parse(sanitizeClaudeJson(match));
  var verifyResult = verifyAndFilterFindings(findings, codeJs, indexHtml, apiKey);

  return {
    findings:      verifyResult.findings,
    reviewedAt:    new Date().toISOString(),
    filesReviewed: filesReviewed,
    scope:         { description: description, matched: matched, note: selection.note || '' },
    verification:  verifyResult.report,
  };
}

// ── Code review verification ────────────────────────────────────────────────
// Two cheap layers that run over freshly-generated findings before they reach
// a user, added after a review call fabricated a finding describing a
// duplicate `let` declaration that does not exist in the source (it would
// have thrown a SyntaxError if it did — the file demonstrably still runs).
//
// Layer 1 (staticVerifyFindings) is plain string/regex checks, no API call:
// - every finding must quote real, verbatim code from the file it's about
// - the function name it points to must actually exist
// - any finding claiming a syntax-breaking bug is auto-rejected if the file
//   still compiles cleanly (new Function() over the actual source)
//
// Layer 2 (verifyFindingsWithLLM) is one batched API call — not one per
// finding — that re-reads each survivor against just its own relevant
// snippet and votes CONFIRMED/REFUTED. Cost stays ~2x a single review call
// regardless of how many findings survive layer 1, instead of scaling with
// finding count.

// Pulls the JS out of <script> tags in Index.html (skipping external/src=
// includes and non-JS types like application/ld+json) so it can be
// syntax-checked without the surrounding HTML confusing the parser.
function extractInlineScripts(html) {
  var out = [];
  var re  = /<script\b([^>]*)>([\s\S]*?)<\/script>/gi;
  var m;
  while ((m = re.exec(html)) !== null) {
    var attrs = m[1] || '';
    if (/\bsrc\s*=/i.test(attrs)) continue;
    if (/type\s*=\s*["'](?!text\/javascript)[^"']*["']/i.test(attrs)) continue;
    out.push(m[2]);
  }
  return out.join('\n;\n');
}

// True if jsCode parses as valid JavaScript — doesn't execute it, just checks
// it compiles. This is what actually disproves a "this causes a SyntaxError"
// claim, rather than trusting the model's description of the code.
function checkSyntaxValid(jsCode) {
  try {
    new Function(jsCode);
    return true;
  } catch (e) {
    return false;
  }
}

function escapeRegExp(s) {
  return String(s).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

// Loose existence check for a finding's claimed "location" (function name) —
// covers `function name(`, `const/let/var name =`, and object-method styles
// (`name: function`, `name = function`).
function locationExistsInSource(source, location) {
  var name = String(location || '').trim();
  if (!name) return true; // nothing to check
  var esc = escapeRegExp(name);
  var patterns = [
    new RegExp('function\\s+' + esc + '\\s*\\('),
    new RegExp('(const|let|var)\\s+' + esc + '\\s*='),
    new RegExp(esc + '\\s*[:=]\\s*function'),
    new RegExp(esc + '\\s*=\\s*\\('),
  ];
  return patterns.some(function(re) { return re.test(source); });
}

function normalizeForMatch(s) {
  return String(s || '').replace(/\s+/g, ' ').trim();
}

// A finding's "quote" must be a real, verbatim (whitespace-insensitive)
// substring of the file it claims to be about — otherwise it's describing
// code that isn't actually there.
function quoteExistsInSource(quote, source) {
  var nq = normalizeForMatch(quote);
  if (!nq) return false;
  return normalizeForMatch(source).indexOf(nq) !== -1;
}

var SYNTAX_CLAIM_RE = /syntax\s*error|redeclar|declared\s+twice|duplicate\s+(let|const|var|declaration)|won'?t\s+(compile|parse|load)|throws?\s+(a\s+)?syntaxerror|fail(s)?\s+to\s+(compile|parse|load)/i;

function claimsSyntaxBreak(finding) {
  var text = (finding.title || '') + ' ' + (finding.description || '');
  return SYNTAX_CLAIM_RE.test(text);
}

// Layer 1 — see header comment above. Returns { kept, dropped } where dropped
// entries carry the reason(s) so a caller can audit what got rejected and why.
function staticVerifyFindings(findings, codeJs, indexHtml) {
  var htmlScripts  = indexHtml ? extractInlineScripts(indexHtml) : '';
  var codeSyntaxOk = checkSyntaxValid(codeJs);
  var htmlSyntaxOk = indexHtml ? checkSyntaxValid(htmlScripts) : true;

  var kept = [], dropped = [];
  findings.forEach(function(f) {
    var isHtml = f.file === 'Index.html' || f.file === 'Index';
    var source = isHtml ? indexHtml : codeJs;
    var reasons = [];

    if (!f.quote) {
      reasons.push('no verbatim quote supplied');
    } else if (!quoteExistsInSource(f.quote, source)) {
      reasons.push('quote does not appear verbatim in ' + (f.file || 'the file'));
    }
    if (f.location && !locationExistsInSource(source, f.location)) {
      reasons.push('location "' + f.location + '" was not found in ' + (f.file || 'the file'));
    }
    if (claimsSyntaxBreak(f) && (isHtml ? htmlSyntaxOk : codeSyntaxOk)) {
      reasons.push('claims a syntax-breaking bug but the file compiles cleanly as-is');
    }

    if (reasons.length) dropped.push({ finding: f, reasons: reasons });
    else kept.push(f);
  });
  return { kept: kept, dropped: dropped };
}

// Grabs just the function body a finding's "location" points to (plus a
// couple lines of leading context) so the batched verify call below doesn't
// need to resend the whole file per finding.
function snippetForLocation(source, location, fallback) {
  var found = findFunctionBlocks(source);
  var block = found.blocks.filter(function(b) { return b.name === location; })[0];
  if (!block) return fallback || '';
  var start = Math.max(0, block.start - 2);
  var end   = Math.min(found.lines.length - 1, block.end);
  return found.lines.slice(start, end + 1).join('\n');
}

// Pulls up to maxLines other lines in the file that mention the same
// identifier(s) as the finding's location/quote, outside the primary snippet
// itself — this is what lets the verifier see a design-intent comment or a
// second call site that lives far from the flagged function, without paying
// for the whole file every time. Cheap (plain regex), no extra API cost.
function relatedMentions(source, location, quote, primarySnippet, maxLines) {
  maxLines = maxLines || 12;
  var tokens = {};
  (String(location || '').match(/[A-Za-z_$][A-Za-z0-9_$]{2,}/g) || []).forEach(function(t) { tokens[t] = true; });
  (String(quote || '').match(/[A-Za-z_$][A-Za-z0-9_$]{2,}/g) || []).forEach(function(t) { tokens[t] = true; });
  var names = Object.keys(tokens).filter(function(t) {
    return !/^(function|const|let|var|return|null|undefined|true|false|this)$/.test(t);
  });
  if (!names.length) return '';

  var lines = source.split('\n');
  var res   = [];
  for (var i = 0; i < lines.length && res.length < maxLines; i++) {
    var line = lines[i];
    if (primarySnippet.indexOf(line) !== -1) continue; // already shown in full snippet
    var hit = names.some(function(n) { return line.indexOf(n) !== -1; });
    if (hit) res.push((i + 1) + ': ' + line.trim());
  }
  return res.join('\n');
}

// Layer 2 — see header comment above. One request covering every surviving
// finding at once, each judged against its own snippet PLUS a short list of
// other lines elsewhere in the file that mention the same identifiers — added
// after the verifier confirmed findings that a design-intent comment or a
// second call site (outside the flagged function's own body) would have
// disproved, because it never saw those lines.
function verifyFindingsWithLLM(findings, codeJs, indexHtml, apiKey) {
  if (!findings.length) return [];

  var items = findings.map(function(f, i) {
    var isHtml  = f.file === 'Index.html' || f.file === 'Index';
    var source  = isHtml ? indexHtml : codeJs;
    var snippet = snippetForLocation(source, f.location, f.quote);
    var related = relatedMentions(source, f.location, f.quote, snippet);
    return '--- Finding #' + i + ' ---\n' +
      'File: ' + f.file + '\nLocation: ' + f.location + '\nTitle: ' + f.title + '\n' +
      'Description: ' + f.description + '\n' +
      'Code:\n```\n' + snippet + '\n```\n' +
      (related ? 'Other lines elsewhere in the file mentioning the same identifiers ' +
        '(may confirm, explain, or contradict the finding):\n```\n' + related + '\n```\n' : '');
  }).join('\n');

  var prompt =
    'You are a skeptical senior engineer double-checking a colleague\'s code review findings before they ' +
    'reach a human, specifically to catch cases where the reviewer misremembered or fabricated code structure, ' +
    'or invented a failure scenario that sounds plausible but cannot actually happen.\n\n' +
    'For each finding, judge it against the code snippet AND the "other lines elsewhere" block for that finding, ' +
    'if one is given. Rules:\n' +
    '1. Do not accept the finding\'s own description of what the code does — re-derive it yourself from the ' +
    'snippet. If the finding depends on arithmetic, an index/loop bound, or a formula, recompute it by hand ' +
    'line by line with concrete example values before deciding; if your own recomputation contradicts the ' +
    'finding\'s claimed result, REFUTE and state what you actually got.\n' +
    '2. If the finding claims a failure scenario ("if X happens, then Y breaks"), REFUTE unless the snippet ' +
    'shows X is actually reachable given the real logic — an unverified hypothetical is not a bug.\n' +
    '3. If the finding claims two names/fields/locations are duplicated, orphaned, or inconsistent, check the ' +
    '"other lines elsewhere" block first for a comment or second usage that explains them as intentionally ' +
    'distinct — if one exists, REFUTE.\n' +
    '4. If the finding claims something about ORDER (one line/call happening before or after another), quote ' +
    'the exact two lines in your reason and confirm their relative order as they literally appear — if you ' +
    'cannot see both locations to compare, REFUTE as unverifiable rather than trusting the description.\n' +
    '5. Default to REFUTED whenever you are not certain — an unverified finding reaching the user is worse ' +
    'than a real one being dropped, since real bugs get caught again on the next review pass.\n\n' +
    items + '\n\n' +
    'Return ONLY a JSON array (no other text, no markdown fences), exactly one entry per finding, same order:\n' +
    '[\n  {\n    "index": 0,\n    "verdict": "CONFIRMED|REFUTED",\n    "reason": "one sentence"\n  }\n]';

  var resp = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method:             'post',
    contentType:        'application/json',
    headers: { 'x-api-key': apiKey, 'anthropic-version': '2023-06-01' },
    payload:            JSON.stringify({
      model:      CLAUDE_MODEL,
      max_tokens: 2048,
      messages:   [{ role: 'user', content: prompt }],
    }),
    muteHttpExceptions: true,
  });

  var data     = JSON.parse(resp.getContentText());
  var verdicts = null;
  if (resp.getResponseCode() === 200) {
    var text  = (data.content && data.content[0]) ? data.content[0].text : '';
    var match = extractJsonSpan(text, '[', ']');
    if (match) {
      try { verdicts = JSON.parse(sanitizeClaudeJson(match)); } catch (e) { verdicts = null; }
    }
  }

  // If the verify call itself fails, don't silently drop everything — pass
  // findings through unverified rather than blocking the whole review.
  return findings.map(function(f, i) {
    var v    = verdicts ? verdicts.filter(function(x) { return x && x.index === i; })[0] : null;
    var copy = {};
    for (var k in f) copy[k] = f[k];
    copy.verify = v
      ? { verdict: v.verdict, reason: v.reason }
      : { verdict: 'UNVERIFIED', reason: 'verification call failed or returned no verdict' };
    return copy;
  });
}

// Entry point used by both runCodeReview and runTargetedCodeReview — runs
// both verification layers and returns only the survivors, plus a report of
// what was rejected and why (for auditing the review pipeline itself).
function verifyAndFilterFindings(findings, codeJs, indexHtml, apiKey) {
  var staticResult = staticVerifyFindings(findings, codeJs, indexHtml);
  var verified;
  try {
    verified = verifyFindingsWithLLM(staticResult.kept, codeJs, indexHtml, apiKey);
  } catch (e) {
    verified = staticResult.kept.map(function(f) {
      var copy = {}; for (var k in f) copy[k] = f[k];
      copy.verify = { verdict: 'UNVERIFIED', reason: 'verification call threw: ' + e.message };
      return copy;
    });
  }
  var refuted = verified.filter(function(f) { return f.verify && f.verify.verdict === 'REFUTED'; });
  var final   = verified.filter(function(f) { return !f.verify || f.verify.verdict !== 'REFUTED'; });
  return {
    findings: final,
    report: {
      generated:      findings.length,
      staticRejected: staticResult.dropped.map(function(d) { return { title: d.finding.title, reasons: d.reasons }; }),
      llmRefuted:     refuted.map(function(f) { return { title: f.title, reason: f.verify.reason }; }),
    },
  };
}

// Given a specific finding, asks Claude for an exact old_string → new_string patch.
// Returns { patch: { old_string, new_string, explanation }, file }.
function generateCodeFix(finding) {
  var isHtml     = (finding.file === 'Index.html' || finding.file === 'Index');
  var src        = fetchCodeFromGitHub(isHtml);
  var fileKey    = isHtml ? 'Index' : 'Code';
  var fullSource = src[fileKey] || '';

  // Send only the relevant function + surrounding context (not the full file)
  // to stay well under the Anthropic TPM rate limit.
  var funcName = finding.location || '';
  var lines    = fullSource.split('\n');
  var funcIdx  = -1;
  for (var li = 0; li < lines.length; li++) {
    if (lines[li].indexOf('function ' + funcName) !== -1 ||
        lines[li].indexOf(funcName + ':') !== -1 ||
        lines[li].indexOf(funcName + ' =') !== -1) {
      funcIdx = li;
      break;
    }
  }
  var fileContent;
  if (funcIdx !== -1) {
    var s = Math.max(0, funcIdx - 5);
    var e = Math.min(lines.length - 1, funcIdx + 200);
    fileContent = (s > 0 ? '// [file truncated above]\n' : '') +
                  lines.slice(s, e + 1).join('\n') +
                  (e < lines.length - 1 ? '\n// [file truncated below]' : '');
  } else {
    fileContent = fullSource; // function not found — fall back to full file
  } 

  var apiKey = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!apiKey) throw new Error('ANTHROPIC_API_KEY not set');

  var prompt =
    'You are fixing a specific bug in a Google Apps Script web app.\n\n' +
    'Finding:\n' +
    '- Title: '       + finding.title       + '\n' +
    '- Location: '    + finding.location    + '\n' +
    '- Description: ' + finding.description + '\n' +
    '- Fix: '         + finding.suggestion  + '\n\n' +
    'File (' + finding.file + '):\n```\n' + fileContent + '\n```\n\n' +
    'Return ONLY a JSON object (no other text, no markdown fences):\n' +
    '{\n' +
    '  "old_string": "exact text to replace — include 3-8 surrounding lines so it is unique in the file",\n' +
    '  "new_string": "the replacement text",\n' +
    '  "explanation": "one sentence explaining the change"\n' +
    '}\n\n' +
    'The old_string must be an exact substring of the file above — character for character.';

  var resp = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method:             'post',
    contentType:        'application/json',
    headers: { 'x-api-key': apiKey, 'anthropic-version': '2023-06-01' },
    payload:            JSON.stringify({
      model:      CLAUDE_MODEL,
      max_tokens: 2048,
      messages:   [{ role: 'user', content: prompt }],
    }),
    muteHttpExceptions: true,
  });

  var data = JSON.parse(resp.getContentText());
  if (resp.getResponseCode() !== 200) {
    throw new Error(data.error ? data.error.message : 'API error ' + resp.getResponseCode());
  }

  var text  = (data.content && data.content[0]) ? data.content[0].text : '';
  var match = extractJsonSpan(text, '{', '}');
  if (!match) throw new Error('Could not parse patch from Claude response');

  return { patch: JSON.parse(sanitizeClaudeJson(match)), file: finding.file };
}

// Formats code review findings as an HTML email and sends it.
function emailCodeReviewFindings(toEmail, findings, meta) {
  var SEV_COLOR = { critical: '#c62828', high: '#e65100', medium: '#b45309', low: '#065f46' };
  var SEV_BG    = { critical: '#fee2e2', high: '#fff3e0', medium: '#fffde7', low: '#e8f5e9' };

  var counts = { critical: 0, high: 0, medium: 0, low: 0 };
  findings.forEach(function(f) { if (counts[f.severity] !== undefined) counts[f.severity]++; });

  var files = (meta.filesReviewed || ['Code.js']).join(' + ');
  var ts    = new Date().toLocaleString('en-AU', { timeZone: 'Australia/Sydney' });

  var subject = '[Omnia] Code Review — ' + findings.length + ' finding' + (findings.length !== 1 ? 's' : '');
  if (counts.critical) subject += ' · ' + counts.critical + ' CRITICAL';
  else if (counts.high) subject += ' · ' + counts.high + ' high';

  var summaryChips = ['critical','high','medium','low'].filter(function(s){ return counts[s] > 0; })
    .map(function(s){ return '<span style="background:' + SEV_BG[s] + ';color:' + SEV_COLOR[s] + ';font-size:11px;font-weight:800;padding:3px 10px;border-radius:10px;text-transform:uppercase;margin-right:6px">' + s + ' ' + counts[s] + '</span>'; })
    .join('');

  var rows = findings.map(function(f) {
    var bg    = SEV_BG[f.severity]    || '#f5f5f5';
    var color = SEV_COLOR[f.severity] || '#333';
    return '<tr style="border-bottom:1px solid #e4e2dc;vertical-align:top">' +
      '<td style="padding:10px 8px;white-space:nowrap"><span style="background:' + bg + ';color:' + color + ';font-size:10px;font-weight:800;padding:2px 8px;border-radius:8px;text-transform:uppercase">' + f.severity + '</span></td>' +
      '<td style="padding:10px 8px;font-size:11px;color:#707979;white-space:nowrap">' + (f.file || 'Code.js') + '</td>' +
      '<td style="padding:10px 8px;font-size:11px;font-family:monospace;color:#404849;white-space:nowrap">' + (f.location || '') + '</td>' +
      '<td style="padding:10px 8px">' +
        '<div style="font-size:13px;font-weight:700;color:#002b2f;margin-bottom:4px">' + f.title + '</div>' +
        '<div style="font-size:12px;color:#707979;line-height:1.5;margin-bottom:6px">' + f.description + '</div>' +
        '<div style="font-size:11px;color:#306767;font-style:italic">&#128161; ' + f.suggestion + '</div>' +
      '</td></tr>';
  }).join('');

  var html =
    '<div style="font-family:Manrope,Arial,sans-serif;max-width:820px;margin:0 auto">' +
    '<div style="background:#002b2f;padding:20px 28px;border-radius:8px 8px 0 0">' +
      '<h2 style="color:#c2f0a6;margin:0 0 4px;font-size:18px">Omnia Code Review</h2>' +
      '<p style="color:#7bafb5;margin:0;font-size:12px">' + files + ' &middot; ' + ts + '</p>' +
    '</div>' +
    '<div style="background:#f5f3ec;padding:14px 28px;border-bottom:1px solid #e4e2dc">' + summaryChips + '</div>' +
    '<table style="width:100%;border-collapse:collapse;background:#fff">' +
      '<thead><tr style="border-bottom:2px solid #e4e2dc;background:#fafaf7">' +
        '<th style="padding:8px;text-align:left;font-size:10px;font-weight:700;color:#a0a8a9;text-transform:uppercase">Severity</th>' +
        '<th style="padding:8px;text-align:left;font-size:10px;font-weight:700;color:#a0a8a9;text-transform:uppercase">File</th>' +
        '<th style="padding:8px;text-align:left;font-size:10px;font-weight:700;color:#a0a8a9;text-transform:uppercase">Location</th>' +
        '<th style="padding:8px;text-align:left;font-size:10px;font-weight:700;color:#a0a8a9;text-transform:uppercase">Finding</th>' +
      '</tr></thead>' +
      '<tbody>' + rows + '</tbody>' +
    '</table>' +
    '<div style="background:#f5f3ec;padding:14px 28px;border-radius:0 0 8px 8px;font-size:11px;color:#a0a8a9">' +
      'Apply fixes manually via the Apps Script editor &middot; Omnia Reporting &middot; Solace Advisory FY26' +
    '</div></div>';

  MailApp.sendEmail({ to: toEmail, subject: subject, htmlBody: html, name: 'Omnia Reporting' });
  return { sent: true, sentTo: toEmail, findingsCount: findings.length };
}

// Diagnostic: logs which OAuth scopes the current token actually contains.
// Run this from the editor to check if script.projects.readonly is authorised.
function devCheckToken() {
  var token = ScriptApp.getOAuthToken();
  var resp  = UrlFetchApp.fetch(
    'https://www.googleapis.com/oauth2/v1/tokeninfo?access_token=' + token,
    { muteHttpExceptions: true }
  );
  Logger.log('=== Token info ===');
  Logger.log(resp.getContentText());
}

// ---------- GitHub PR Flow ----------

// Intentionally not a real repo — this demo copy ships with no GITHUB_TOKEN
// in Script Properties, so every call below throws before a request is ever
// made. Point these at a real demo repo (with its own token) if you want the
// "create fix PR" feature to actually work here.
var GITHUB_OWNER = 'your-org';
var GITHUB_REPO  = 'your-demo-repo';
var GITHUB_API   = 'https://api.github.com';

function getGitHubToken() {
  var t = PropertiesService.getScriptProperties().getProperty('GITHUB_TOKEN');
  if (!t) throw new Error('GITHUB_TOKEN not set in Script Properties');
  return t;
}

function githubRequest(method, path, body) {
  var options = {
    method:             method,
    contentType:        'application/json',
    headers: {
      Authorization:         'Bearer ' + getGitHubToken(),
      Accept:                'application/vnd.github+json',
      'X-GitHub-Api-Version': '2022-11-28',
    },
    muteHttpExceptions: true,
  };
  if (body) options.payload = JSON.stringify(body);
  var resp = UrlFetchApp.fetch(GITHUB_API + path, options);
  var code = resp.getResponseCode();
  var data = JSON.parse(resp.getContentText());
  if (code < 200 || code >= 300) {
    throw new Error('GitHub ' + code + ': ' + (data.message || resp.getContentText()));
  }
  return data;
}

// Fetches Code.js (and optionally Index.html) from GitHub main branch.
// Returns { Code, Index?, _codeSha, _htmlSha? } — same shape as fetchOwnSource().
function fetchCodeFromGitHub(includeHtml) {
  var cd = githubRequest('GET', '/repos/' + GITHUB_OWNER + '/' + GITHUB_REPO + '/contents/Code.js?');
  var result  = {
    Code:     Utilities.newBlob(Utilities.base64Decode(cd.content.replace(/\n/g, ''))).getDataAsString(),
    _codeSha: cd.sha,
  };
  if (includeHtml) {
    var hd = githubRequest('GET', '/repos/' + GITHUB_OWNER + '/' + GITHUB_REPO + '/contents/Index.html?');
    result.Index   = Utilities.newBlob(Utilities.base64Decode(hd.content.replace(/\n/g, ''))).getDataAsString();
    result._htmlSha = hd.sha;
  }
  var sd      = githubRequest('GET', '/repos/' + GITHUB_OWNER + '/' + GITHUB_REPO + '/contents/Seed.js');
  result.Seed = Utilities.newBlob(Utilities.base64Decode(sd.content.replace(/\n/g, ''))).getDataAsString();
  return result;
}

function getMainBranchSha() {
  return githubRequest('GET', '/repos/' + GITHUB_OWNER + '/' + GITHUB_REPO + '/git/ref/heads/main').object.sha;
}

function createGitHubBranch(branchName, fromSha) {
  return githubRequest('POST', '/repos/' + GITHUB_OWNER + '/' + GITHUB_REPO + '/git/refs', {
    ref: 'refs/heads/' + branchName,
    sha: fromSha,
  });
}

function commitFileToGitHub(branch, filePath, content, currentSha, commitMessage) {
  return githubRequest('PUT', '/repos/' + GITHUB_OWNER + '/' + GITHUB_REPO + '/contents/' + filePath, {
    message: commitMessage,
    content: Utilities.base64Encode(content, Utilities.Charset.UTF_8),
    sha:     currentSha,
    branch:  branch,
  });
}

function openPullRequest(title, body, headBranch) {
  return githubRequest('POST', '/repos/' + GITHUB_OWNER + '/' + GITHUB_REPO + '/pulls', {
    title: title,
    body:  body,
    head:  headBranch,
    base:  'main',
  });
}

// Creates a branch, commits the patched file, and opens a PR.
// Returns { pr_url, pr_number, branch }.
function createFixPR(finding, patch) {
  var isHtml   = (finding.file === 'Index.html' || finding.file === 'Index');
  var src      = fetchCodeFromGitHub(isHtml);
  var fileKey  = isHtml ? 'Index' : 'Code';
  var filePath = isHtml ? 'Index.html' : 'Code.js';
  var fileSha  = isHtml ? src._htmlSha : src._codeSha;
  var original = src[fileKey] || '';

  if (original.indexOf(patch.old_string) === -1) {
    throw new Error('patch.old_string not found in ' + filePath + ' — patch may be stale, re-run code_fix');
  }
  var patched = original.replace(patch.old_string, patch.new_string);

  var mainSha    = getMainBranchSha();
  var slug       = finding.title.toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '').slice(0, 38);
  var ts         = new Date().getTime().toString().slice(-4);
  var branchName = 'fix/' + slug + '-' + ts;

  createGitHubBranch(branchName, mainSha);
  commitFileToGitHub(branchName, filePath, patched, fileSha, 'fix(' + finding.location + '): ' + finding.title);

  var prBody =
    '## Finding\n\n' +
    '**Severity:** ' + finding.severity + '  \n' +
    '**File:** ' + finding.file + '  \n' +
    '**Location:** `' + finding.location + '`\n\n' +
    '**Issue:** ' + finding.description + '\n\n' +
    '**Suggested fix:** ' + finding.suggestion + '\n\n' +
    '---\n_Auto-generated by Omnia code review · ' + new Date().toISOString() + '_';

  var pr = openPullRequest('[Omnia Fix] ' + finding.title, prBody, branchName);
  return { pr_url: pr.html_url, pr_number: pr.number, branch: branchName };
}

// Lists open PRs on branches starting with fix/ (for the dashboard panel).
function listFixPRs() {
  var prs = githubRequest('GET', '/repos/' + GITHUB_OWNER + '/' + GITHUB_REPO + '/pulls?state=open&per_page=50');
  return prs
    .filter(function(pr) { return pr.head && pr.head.ref && pr.head.ref.indexOf('fix/') === 0; })
    .map(function(pr) {
      return { pr_number: pr.number, pr_url: pr.html_url, title: pr.title, branch: pr.head.ref, created: pr.created_at };
    });
}

// Squash-merges a PR (Approve button).
function mergeFixPR(prNumber, commitTitle) {
  return githubRequest('PUT', '/repos/' + GITHUB_OWNER + '/' + GITHUB_REPO + '/pulls/' + prNumber + '/merge', {
    commit_title: commitTitle || 'fix: merge PR #' + prNumber,
    merge_method: 'squash',
  });
}

// Closes a PR without merging (Reject button).
function closeFixPR(prNumber) {
  return githubRequest('PATCH', '/repos/' + GITHUB_OWNER + '/' + GITHUB_REPO + '/pulls/' + prNumber, {
    state: 'closed',
  });
}

// ---------- Utils ----------
function pick(obj, keys) {
  const out = {};
  keys.forEach(k => { if (obj[k] !== undefined) out[k] = obj[k]; });
  return out;
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}