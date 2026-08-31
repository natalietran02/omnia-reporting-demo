/**
 * Omnia Reporting Demo — one-shot seed script.
 *
 * Run once from the Apps Script editor: select `seedDemoData` in the
 * function dropdown, click Run, then approve the permission prompts.
 * Safe to re-run — every tab is cleared and rewritten from scratch, so
 * re-running just regenerates the same demo dataset (the RNG is seeded,
 * so the numbers come out identical every time).
 *
 * Writes:
 *   - Every data_* tab Code.js reads (fabricated figures, 3 fictional
 *     entities — Solace Advisory, Kestrel Group, Northlight AI — plus an
 *     empty Fernway Data placeholder for structural parity with the real
 *     product's 4-entity model).
 *   - pages / config / editors / access_users, seeded so you can log in
 *     immediately (see ADMIN_EMAIL below).
 *   - memos / memo_versions, empty (matches the real product's Seed.js).
 * activity_log / feedback self-provision on first use — see ensureSheet()
 * in Code.js — so they're not seeded here.
 */

// The account that can log in and drive this demo as an admin from the
// start. Must be on one of Code.js's ALLOWED_DOMAINS or the OTP login will
// reject it — change this to your own address before running if needed.
var ADMIN_EMAIL = 'nathan@edgered.com.au';

function seedDemoData() {
  var raw = buildFakeRawData_();
  var ss = SpreadsheetApp.getActive();

  writeTab_(ss, 'data_key_metrics_monthly', raw.keyMetrics,
    ['period', 'month_date', 'entity', 'division', 'billable_revenue', 'utilisation_pct', 'margin', 'active_clients', 'billable_fte', 'cummulative_billable_revenue']);

  writeTab_(ss, 'data_powerbi_unified_fact', raw.unifiedFact,
    ['month_date', 'fiscal_year', 'year_month', 'month_name', 'calendar_year', 'calendar_quarter', 'entity', 'revenue_status',
     'business_unit', 'role_group', 'location', 'project_code', 'project_name', 'client_name', 'project_type', 'project_stream',
     'client_industry', 'headcount', 'total_days', 'total_revenue', 'total_cost', 'gross_margin', 'margin_pct', 'avg_daily_rate',
     'avg_daily_cost', 'monthly_revenue_target']);

  writeTab_(ss, 'data_hubspot_deals', raw.hubspotDeals,
    ['Record_ID', 'Omnia_Company', 'Company', 'Brand', 'Deal_Name', 'Deal_Owner', 'Deal_Team', 'Deal_Stage', 'Amount', 'Confidence',
     'Close_Date', 'Open_Date', 'Last_Modified_Date', 'Last_Updated', 'Source_System', 'Contributing_OpCo_1', 'Contributing_OpCo_2',
     'Contributing_OpCo_3', 'Contributing_OpCo_4', 'Contributing_OpCo_5', 'Deal_Source', 'Alliance_Partners', 'Logged_In_Partner_Portal',
     'Deal_Probability', 'Deal_Lost_Reason', 'Why_Lost', 'Is_Stalled', 'Stage_Entered_Date', 'Weighted_Open_Pipeline', 'Month']);

  writeTab_(ss, 'data_ref_projects', raw.refProjects,
    ['Project_Code', 'Project_Name', 'Start_Date', 'End_Date', 'Client', 'Industry', 'Omnia_Company']);

  writeTab_(ss, 'data_employee_roster', raw.employeeRoster,
    ['userId', 'Omnia_Company', 'role_title', 'role_code', 'role_group', 'employment_type', 'business_unit', 'location',
     'employment_start_date', 'effective_from', 'termination_date', 'tenure', 'tenure_long', 'FTE', 'actual_weekly_hours',
     'salary_band_year', 'base_salary', 'cost_per_hour']);

  writeTab_(ss, 'data_revenue_target', raw.revenueTarget,
    ['month_date', 'entity', 'total_revenue', 'revenue_target']);

  writeTab_(ss, 'data_utilisation', raw.utilisation,
    ['userId', 'month', 'fiscal_year', 'Omnia_Company', 'business_unit', 'role_group', 'role_code', 'location', 'employment_basis',
     'is_contractor', 'is_billable_staff', 'available_days', 'available_days_full', 'invoiced_days', 'leave_days', 'holiday_days',
     'daily_rate', 'billed_amount', 'fte', 'utilisation_pct', 'utilisation_pct_full']);

  writeTab_(ss, 'data_certifications', raw.certifications,
    ['AADUserId', 'IndividualFirstName', 'IndividualLastName', 'Email', 'CorpEmail', 'PartnerName', 'PartnerCountryLocation',
     'PartnerCityLocation', 'MPNId', 'PGAMpnId', 'TrainingActivityId', 'TrainingType', 'TrainingTitle', 'ActivationStatus',
     'TrainingCompletionDate', 'ExpirationDate', 'Month']);

  writeTab_(ss, 'data_target_daily_rate', raw.targetRates,
    ['FY', 'Grade', 'Utilisation_Target', 'Daily_Rate_Target']);

  writeTab_(ss, 'data_pl', raw.pl,
    ['month', 'Omnia_Company', 'fiscal_year', 'revenue_actual', 'revenue_budget', 'gross_margin_actual', 'gross_margin_forecast',
     'gp_budget', 'ebitda_actual', 'ebitda_budget', 'direct_costs_actual', 'direct_costs_forecast', 'operating_expenses_actual',
     'operating_expenses_forecast', 'wages_salaries_indirect_actual', 'wages_salaries_indirect_forecast', 'employee_benefits_actual',
     'employee_benefits_forecast', 'sales_marketing_actual', 'sales_marketing_forecast', 'travel_expenses_actual',
     'travel_expenses_forecast', 'administrative_expenses_actual', 'administrative_expenses_forecast', 'rent_expense_actual',
     'rent_expense_forecast', 'recruitment_expense_actual', 'recruitment_expense_forecast', 'other_operating_expenses_actual',
     'other_operating_expenses_forecast', 'utilisation_budget', 'headcount_budget', 'daily_rate_budget', 'operating_margin_actual',
     'operating_margin_forecast']);

  // Empty on purpose — Fernway Data (the Elysium-equivalent 4th entity) has
  // no rows in any table, so this tab just needs to exist with the right
  // headers or getAllReportingData() throws "Missing tab" for everyone.
  writeTab_(ss, 'data_elysium_pl_monthly', [],
    ['month_date', 'fiscal_year', 'type', 'Omnia_Company', 'revenue', 'direct_costs', 'gross_profit', 'wages_salaries_indirect',
     'employee_benefits', 'sales_marketing', 'travel_expenses', 'administrative_expenses', 'rent_expense', 'recruitment_expense',
     'other_operating_expenses', 'operating_expenses', 'normalised_ebitda', 'gross_profit_margin_pct', 'normalised_ebitda_margin_pct',
     'billable_fte', 'total_fte', 'average_day_rate', 'billable_employee_utilisation_pct', 'nwd']);

  // ---------- App-owned tabs ----------
  ensureTab_(ss, 'pages', ['slug', 'title', 'order', 'status'], [
    ['takeaways', 'Takeaways', 1, 'active'], ['revenue', 'Revenue', 2, 'active'], ['clients', 'Clients', 3, 'active'],
    ['projects', 'Projects', 4, 'active'], ['deals', 'Deals', 5, 'active'], ['employees', 'Employees', 6, 'active'],
    ['margins', 'Margins', 7, 'active'], ['cost', 'Cost', 8, 'active'], ['utilisation', 'Utilisation', 9, 'active'],
    ['certifications', 'Certifications', 10, 'active'],
  ]);

  ensureTab_(ss, 'config', ['config key', 'default', 'controls'], [
    ['current_period', '2026-08', 'Period the app defaults to — update monthly'],
    ['fy_start_period', '2026-07', 'FY27 starts July 2026'],
    ['entity', 'Solace Advisory', 'Primary entity for health check / AI commentary'],
    ['app_version', '0.1.0-demo', ''],
    ['check_gp_min', 0, 'GP% floor'],
    ['check_gp_max', 80, 'GP% ceiling'],
    ['check_util_max', 110, 'Utilisation % ceiling'],
    ['check_mom_drop_pct', 40, 'MoM revenue drop %'],
  ]);

  ensureTab_(ss, 'editors', ['email', 'role', 'pages', 'status', 'updated_at', 'updated_by'], [
    [ADMIN_EMAIL, 'admin', '*', 'active', new Date(), 'seed'],
  ]);

  ensureTab_(ss, 'access_users', ['email', 'role', 'status', 'added_by', 'added_at', 'updated_at', 'updated_by'], [
    [ADMIN_EMAIL, 'admin', 'active', 'seed', new Date(), '', ''],
  ]);

  ensureTab_(ss, 'memos', ['memo_id', 'page_slug', 'period', 'headline', 'standfirst', 'bullet_1', 'bullet_2', 'bullet_3', 'bullet_4',
    'callout_label', 'callout_text', 'status', 'source', 'model', 'prompt_version', 'published_at', 'published_by', 'updated_at', 'updated_by'], []);
  ensureTab_(ss, 'memo_versions', ['version_id', 'memo_id', 'changed_at', 'changed_by', 'field', 'old_value', 'new_value'], []);

  SpreadsheetApp.getActive().toast('Demo data seeded — log in as ' + ADMIN_EMAIL + ' to see it.');
}

function writeTab_(ss, name, rows, headers) {
  var sh = ss.getSheetByName(name);
  if (sh) sh.clear(); else sh = ss.insertSheet(name);
  var values = [headers].concat(rows.map(function (r) { return headers.map(function (h) { return r[h] != null ? r[h] : ''; }); }));
  sh.getRange(1, 1, values.length, headers.length).setValues(values);
  sh.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  sh.setFrozenRows(1);
}

function ensureTab_(ss, name, headers, rows) {
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sh.setFrozenRows(1);
    if (rows.length) sh.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }
}

// ── Synthetic dataset generator — same logic as the standalone webpage
// demo's fake-data.js, ported to Apps Script (V8 runtime supports the same
// language features). Every figure here is fabricated.
function makeRng_(seed) {
  return function () {
    seed |= 0; seed = (seed + 0x6D2B79F5) | 0;
    var t = Math.imul(seed ^ (seed >>> 15), 1 | seed);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

function buildFakeRawData_() {
  var rng = makeRng_(20260827);
  function rf(min, max) { return min + rng() * (max - min); }
  function ri(min, max) { return Math.floor(rf(min, max + 1)); }
  function pick(arr) { return arr[Math.floor(rng() * arr.length)]; }
  function pad2(n) { return String(n).padStart(2, '0'); }
  function ymd(y, m, d) { return y + '-' + pad2(m) + '-' + pad2(d); }
  function lastDayOfMonth(y, m) { return new Date(y, m, 0).getDate(); }

  var months = [];
  var cursor = { y: 2023, m: 7 };
  var ACTUALS_END = { y: 2026, m: 8 };
  var FORECAST_END = { y: 2027, m: 6 };
  while (cursor.y < FORECAST_END.y || (cursor.y === FORECAST_END.y && cursor.m <= FORECAST_END.m)) {
    var isForecast = (cursor.y > ACTUALS_END.y) || (cursor.y === ACTUALS_END.y && cursor.m > ACTUALS_END.m);
    months.push({ y: cursor.y, m: cursor.m, key: cursor.y + '-' + pad2(cursor.m), isForecast: isForecast });
    cursor.m++; if (cursor.m > 12) { cursor.m = 1; cursor.y++; }
  }
  var actualMonths = months.filter(function (mo) { return !mo.isForecast; });

  var ENTITIES = [
    { name: 'Solace Advisory', size: 1.0, growth: 0.014, bus: ['Data & Analytics', 'AI & Automation', 'Strategy & Advisory'] },
    { name: 'Kestrel Group', size: 0.55, growth: 0.010, bus: ['Managed Services', 'Data & Analytics'] },
    { name: 'Northlight AI', size: 0.30, growth: 0.028, bus: ['AI & Automation', 'Applied Research'] },
  ];
  var GRADES = ['Grade 1', 'Grade 2', 'Grade 3', 'Grade 4', 'Grade 5'];
  var LOCATIONS = ['NSW', 'VIC', 'QLD', 'WA', 'Philippines'];

  var CLIENTS = [
    { name: 'Meridian Bank', industry: 'Financial Services' },
    { name: 'Northfield Retail Group', industry: 'Retail' },
    { name: 'Solstice Health Partners', industry: 'Healthcare' },
    { name: 'Greymoor Government Services', industry: 'Government' },
    { name: 'Bramwell Manufacturing', industry: 'Manufacturing' },
    { name: 'Aurora Cloud Systems', industry: 'Technology' },
    { name: 'Fernbank Insurance', industry: 'Financial Services' },
    { name: 'Halcyon Logistics', industry: 'Logistics' },
    { name: 'Quarry Financial', industry: 'Financial Services' },
    { name: 'Ridgeline Energy', industry: 'Energy' },
    { name: 'Copperleaf Retail', industry: 'Retail' },
    { name: 'Vantage Health Network', industry: 'Healthcare' },
    { name: 'Silverlake Technologies', industry: 'Technology' },
    { name: 'Amberview Insurance', industry: 'Financial Services' },
    { name: 'Thornbury Manufacturing', industry: 'Manufacturing' },
    { name: 'Coastal Grid Utilities', industry: 'Energy' },
    { name: 'Ferngate Financial', industry: 'Financial Services' },
    { name: 'Wrenfield Public Sector', industry: 'Government' },
  ];
  var PROJECT_NOUNS = ['Data Platform', 'Analytics Uplift', 'AI Copilot Rollout', 'Reporting Modernisation', 'Cloud Migration',
    'Customer Insights Program', 'Forecasting Model', 'Automation Pipeline', 'Governance Framework', 'Self-Serve BI Rollout'];
  var PROJECT_TYPES = ['Fixed Fee', 'T&M', 'Retainer'];

  var refProjects = [];
  var projectCodeCounter = 1;
  var projectsByEntity = { 'Solace Advisory': [], 'Kestrel Group': [], 'Northlight AI': [] };
  CLIENTS.forEach(function (client) {
    var entity = pick(ENTITIES).name;
    var nProjects = ri(1, 3);
    for (var i = 0; i < nProjects; i++) {
      var startIdx = ri(0, actualMonths.length - 6);
      var startMo = actualMonths[startIdx];
      var durationMonths = ri(4, 22);
      var stillOpen = rng() < 0.55 || startIdx + durationMonths >= actualMonths.length;
      var endMo = stillOpen ? null : actualMonths[Math.min(startIdx + durationMonths, actualMonths.length - 1)];
      var code = 'PRJ-' + String(projectCodeCounter++).padStart(4, '0');
      var proj = {
        Project_Code: code,
        Project_Name: client.name.split(' ')[0] + ' ' + pick(PROJECT_NOUNS),
        Start_Date: ymd(startMo.y, startMo.m, 1),
        End_Date: endMo ? ymd(endMo.y, endMo.m, lastDayOfMonth(endMo.y, endMo.m)) : '',
        Client: client.name,
        Industry: client.industry,
        Omnia_Company: entity,
      };
      refProjects.push(proj);
      projectsByEntity[entity].push({ proj: proj, startIdx: startIdx, endIdx: endMo ? actualMonths.indexOf(endMo) : (actualMonths.length - 1) });
    }
  });

  var unifiedFact = [];
  var MONTH_NAMES3 = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  actualMonths.forEach(function (mo, idx) {
    ENTITIES.forEach(function (ent) {
      projectsByEntity[ent.name].forEach(function (p) {
        if (idx < p.startIdx || idx > p.endIdx) return;
        var proj = p.proj;
        var bu = pick(ent.bus);
        var monthsIn = idx - p.startIdx;
        var baseRev = 55000 * ent.size * (0.85 + rng() * 0.5) * Math.pow(1 + ent.growth, monthsIn);
        var rev = Math.round(baseRev * (0.9 + rng() * 0.2));
        var marginPct = 0.28 + rng() * 0.22;
        var cost = Math.round(rev * (1 - marginPct));
        var days = Math.round(rev / (700 + rng() * 350));
        var headcount = Math.max(1, Math.round(days / 19));
        var fy = mo.m >= 7 ? mo.y + 1 : mo.y;
        unifiedFact.push({
          month_date: ymd(mo.y, mo.m, 1), fiscal_year: 'FY' + String(fy).slice(-2), year_month: mo.key,
          month_name: MONTH_NAMES3[mo.m - 1], calendar_year: String(mo.y), calendar_quarter: 'Q' + (Math.floor((mo.m - 1) / 3) + 1),
          entity: ent.name, revenue_status: 'Actual', business_unit: bu,
          role_group: pick(['Consulting', 'Engineering', 'Data Science']), location: pick(LOCATIONS),
          project_code: proj.Project_Code, project_name: proj.Project_Name, client_name: proj.Client,
          project_type: pick(PROJECT_TYPES), project_stream: bu, client_industry: proj.Industry,
          headcount: headcount, total_days: days, total_revenue: rev, total_cost: cost, gross_margin: rev - cost,
          margin_pct: marginPct * 100, avg_daily_rate: Math.round(rev / Math.max(days, 1)),
          avg_daily_cost: Math.round(cost / Math.max(days, 1)), monthly_revenue_target: Math.round(rev * (0.92 + rng() * 0.14)),
        });
      });
    });
  });

  var keyMetrics = [];
  var cumByEntity = {};
  actualMonths.forEach(function (mo, idx) {
    ENTITIES.forEach(function (ent) {
      var rows = unifiedFact.filter(function (r) { return r.entity === ent.name && r.month_date === ymd(mo.y, mo.m, 1); });
      var rev = rows.reduce(function (s, r) { return s + r.total_revenue; }, 0);
      var margin = rows.reduce(function (s, r) { return s + r.gross_margin; }, 0);
      var clients = {}; rows.forEach(function (r) { clients[r.client_name] = true; });
      var util = 68 + rng() * 20 - (mo.m === 1 ? 6 : 0);
      cumByEntity[ent.name] = (cumByEntity[ent.name] || 0) + rev;
      keyMetrics.push({
        period: mo.key, month_date: ymd(mo.y, mo.m, 1), entity: ent.name, division: '', billable_revenue: rev,
        utilisation_pct: Math.round(util * 10) / 10, margin: margin, active_clients: Object.keys(clients).length,
        billable_fte: Math.max(1, Math.round(rows.reduce(function (s, r) { return s + r.headcount; }, 0) / 1.15)),
        cummulative_billable_revenue: Math.round(cumByEntity[ent.name]),
      });
    });
  });

  var revenueTarget = [];
  months.forEach(function (mo) {
    ENTITIES.forEach(function (ent) {
      var rows = unifiedFact.filter(function (r) { return r.entity === ent.name && r.month_date === ymd(mo.y, mo.m, 1); });
      var actualRev = rows.reduce(function (s, r) { return s + r.total_revenue; }, 0);
      var target = mo.isForecast
        ? Math.round(220000 * ent.size * Math.pow(1 + ent.growth, months.indexOf(mo)) * (0.95 + rng() * 0.1))
        : Math.round((actualRev || 200000 * ent.size) * (0.9 + rng() * 0.18));
      revenueTarget.push({ month_date: ymd(mo.y, mo.m, 1), entity: ent.name, total_revenue: mo.isForecast ? 0 : actualRev, revenue_target: target });
    });
  });

  var employeeRoster = [];
  var FIRST_NAMES = ['Ava','Noah','Mia','Liam','Zoe','Ethan','Grace','Lucas','Ivy','Owen','Ruby','Jack','Nora','Leo','Isla','Finn','Chloe','Max','Ella','Sam','Priya','Arjun','Wei','Sofia','Marco','Fatima','Tom','Hana','Jules','Kai'];
  var LAST_NAMES = ['Whitfield','Nakamura','Osei','Berg','Calloway','Nguyen','Farrell','Petrov','Alcaraz','Doyle','Marsh','Okafor','Lindqvist','Reyes','Hartley','Sundberg','Delgado','Okonkwo','Vance','Kowalski'];
  var ROLE_TITLES = { 'Grade 1': 'Analyst', 'Grade 2': 'Consultant', 'Grade 3': 'Senior Consultant', 'Grade 4': 'Principal Consultant', 'Grade 5': 'Practice Lead' };
  var empId = 1;
  ENTITIES.forEach(function (ent) {
    var targetHC = Math.round(28 * ent.size);
    for (var i = 0; i < targetHC; i++) {
      var startIdx = ri(0, actualMonths.length - 2);
      var startMo = actualMonths[startIdx];
      var grade = pick(GRADES);
      var isContractor = rng() < 0.16;
      var terminated = rng() < 0.12 && startIdx < actualMonths.length - 4;
      var termMo = terminated ? actualMonths[ri(startIdx + 3, actualMonths.length - 1)] : null;
      employeeRoster.push({
        userId: 'EMP-' + String(empId++).padStart(4, '0'), Omnia_Company: ent.name,
        role_title: (isContractor ? 'Contract ' : '') + ROLE_TITLES[grade], role_code: grade.replace('Grade ', 'G'),
        role_group: pick(['Consulting', 'Engineering', 'Data Science', 'Delivery Leadership']),
        employment_type: isContractor ? 'Contractor' : 'Permanent', business_unit: pick(ent.bus), location: pick(LOCATIONS),
        employment_start_date: ymd(startMo.y, startMo.m, ri(1, 28)), effective_from: ymd(startMo.y, startMo.m, 1),
        termination_date: termMo ? ymd(termMo.y, termMo.m, ri(1, 28)) : '',
        tenure: Math.round(((actualMonths.length - startIdx) / 12) * 10) / 10, tenure_long: '',
        FTE: isContractor ? 1 : (rng() < 0.08 ? 0.8 : 1), actual_weekly_hours: isContractor ? 40 : (rng() < 0.08 ? 30 : 38),
        salary_band_year: 'FY26', base_salary: Math.round((85000 + GRADES.indexOf(grade) * 38000) * (0.92 + rng() * 0.16)),
        cost_per_hour: Math.round(((85000 + GRADES.indexOf(grade) * 38000) / 1650) * (0.92 + rng() * 0.16)),
        _name: FIRST_NAMES[ri(0, FIRST_NAMES.length - 1)] + ' ' + LAST_NAMES[ri(0, LAST_NAMES.length - 1)],
        _grade: grade, _termMoIdx: termMo ? actualMonths.indexOf(termMo) : null, _startIdx: startIdx,
      });
    }
  });

  var utilisation = [];
  actualMonths.forEach(function (mo, idx) {
    var wd = [11, 20, 20, 22, 21, 21, 22, 21, 22, 23, 20, 20][mo.m - 1];
    employeeRoster.forEach(function (emp) {
      if (idx < emp._startIdx) return;
      if (emp._termMoIdx != null && idx > emp._termMoIdx) return;
      var leave = ri(0, 3);
      var holiday = mo.m === 12 || mo.m === 1 ? ri(2, 5) : 0;
      var available = Math.max(1, wd - leave - holiday) * emp.FTE;
      var targetUtil = emp.employment_type === 'Contractor' ? (0.85 + rng() * 0.12) : (0.7 + rng() * 0.2);
      var invoiced = Math.round(available * Math.min(1.08, targetUtil) * 10) / 10;
      var rate = Math.round((650 + GRADES.indexOf(emp._grade) * 220) * (0.92 + rng() * 0.18));
      utilisation.push({
        userId: emp.userId, month: ymd(mo.y, mo.m, 1), fiscal_year: 'FY' + String(mo.m >= 7 ? mo.y + 1 : mo.y).slice(-2),
        Omnia_Company: emp.Omnia_Company, business_unit: emp.business_unit, role_group: emp.role_group, role_code: emp.role_code,
        location: emp.location, employment_basis: emp.employment_type, is_contractor: emp.employment_type === 'Contractor',
        is_billable_staff: true, available_days: available, available_days_full: Math.max(1, wd) * emp.FTE,
        invoiced_days: invoiced, leave_days: leave, holiday_days: holiday, daily_rate: rate,
        billed_amount: Math.round(invoiced * rate), fte: emp.FTE,
        utilisation_pct: Math.round((invoiced / available) * 1000) / 10, utilisation_pct_full: Math.round((invoiced / (Math.max(1, wd) * emp.FTE)) * 1000) / 10,
      });
    });
  });

  var targetRates = [];
  ['FY25', 'FY26', 'FY27'].forEach(function (fy) {
    GRADES.forEach(function (grade, gi) {
      targetRates.push({ FY: fy, Grade: grade, Utilisation_Target: 0.72 + gi * 0.02, Daily_Rate_Target: 680 + gi * 210 });
    });
  });

  var pl = [];
  months.forEach(function (mo) {
    ENTITIES.forEach(function (ent) {
      var rows = unifiedFact.filter(function (r) { return r.entity === ent.name && r.month_date === ymd(mo.y, mo.m, 1); });
      var revActual = rows.reduce(function (s, r) { return s + r.total_revenue; }, 0);
      var idx = months.indexOf(mo);
      var revBudget = Math.round(210000 * ent.size * Math.pow(1 + ent.growth, idx) * (0.95 + rng() * 0.1));
      var directCostsActual = rows.length ? -Math.round(rows.reduce(function (s, r) { return s + r.total_cost; }, 0)) : null;
      var directCostsForecast = -Math.round(revBudget * (0.58 + rng() * 0.06));
      var opexPctActual = 0.16 + rng() * 0.05;
      var wagesActual = rows.length ? -Math.round(revActual * (0.08 + rng() * 0.02)) : null;
      var benefitsActual = rows.length ? -Math.round(revActual * (0.015 + rng() * 0.01)) : null;
      var salesActual = rows.length ? -Math.round(revActual * (0.02 + rng() * 0.01)) : null;
      var travelActual = rows.length ? -Math.round(revActual * (0.005 + rng() * 0.008)) : null;
      var adminActual = rows.length ? -Math.round(revActual * (0.018 + rng() * 0.01)) : null;
      var rentActual = rows.length ? -Math.round(revActual * 0.012) : null;
      var recruitActual = rows.length ? -Math.round(revActual * (0.004 + rng() * 0.006)) : null;
      var otherOpexActual = rows.length ? -Math.round(revActual * (0.006 + rng() * 0.006)) : null;
      var opexActual = rows.length ? [wagesActual, benefitsActual, salesActual, travelActual, adminActual, rentActual, recruitActual, otherOpexActual].reduce(function (s, v) { return s + v; }, 0) : null;
      var gpActual = rows.length ? (revActual + directCostsActual) : null;
      var gpBudget = Math.round(revBudget * (0.36 + rng() * 0.06));
      var ebitdaActual = rows.length ? Math.round(gpActual + opexActual) : null;
      var ebitdaBudget = Math.round(gpBudget * (0.42 + rng() * 0.08));
      pl.push({
        month: ymd(mo.y, mo.m, 1), Omnia_Company: ent.name, fiscal_year: 'FY' + String(mo.m >= 7 ? mo.y + 1 : mo.y).slice(-2),
        revenue_actual: rows.length ? revActual : null, revenue_budget: revBudget,
        gross_margin_actual: gpActual, gross_margin_forecast: rows.length ? null : gpBudget, gp_budget: gpBudget,
        ebitda_actual: ebitdaActual, ebitda_budget: ebitdaBudget,
        direct_costs_actual: directCostsActual, direct_costs_forecast: rows.length ? null : directCostsForecast,
        operating_expenses_actual: opexActual, operating_expenses_forecast: rows.length ? null : -Math.round(revBudget * opexPctActual),
        wages_salaries_indirect_actual: wagesActual, wages_salaries_indirect_forecast: rows.length ? null : -Math.round(revBudget * 0.09),
        employee_benefits_actual: benefitsActual, employee_benefits_forecast: rows.length ? null : -Math.round(revBudget * 0.02),
        sales_marketing_actual: salesActual, sales_marketing_forecast: rows.length ? null : -Math.round(revBudget * 0.025),
        travel_expenses_actual: travelActual, travel_expenses_forecast: rows.length ? null : -Math.round(revBudget * 0.009),
        administrative_expenses_actual: adminActual, administrative_expenses_forecast: rows.length ? null : -Math.round(revBudget * 0.022),
        rent_expense_actual: rentActual, rent_expense_forecast: rows.length ? null : -Math.round(revBudget * 0.012),
        recruitment_expense_actual: recruitActual, recruitment_expense_forecast: rows.length ? null : -Math.round(revBudget * 0.007),
        other_operating_expenses_actual: otherOpexActual, other_operating_expenses_forecast: rows.length ? null : -Math.round(revBudget * 0.009),
        utilisation_budget: 0.76 + rng() * 0.06, headcount_budget: Math.round(28 * ent.size * Math.pow(1 + ent.growth, idx)),
        daily_rate_budget: 780 + Math.round(rng() * 120),
        operating_margin_actual: rows.length ? Math.round(ebitdaActual * 0.82) : null,
        operating_margin_forecast: rows.length ? null : Math.round(ebitdaBudget * 0.82),
      });
    });
  });

  var STAGES = ['Initiated', 'Pain/Gain Defined', 'Qualified', 'Co-Designed', 'Presented', 'Enabled Decision', 'Closed Won', 'Closed Lost'];
  var STAGE_PROB = { 'Initiated': 0.05, 'Pain/Gain Defined': 0.15, 'Qualified': 0.3, 'Co-Designed': 0.45, 'Presented': 0.6, 'Enabled Decision': 0.8, 'Closed Won': 1, 'Closed Lost': 0 };
  var ALLIANCES = ['N/A', 'Not yet known', 'Microsoft', 'Databricks', 'Snowflake', 'Other'];
  var LOST_REASONS = ['Budget deferred', 'Chose incumbent', 'Timing not right', 'Lost to competitor', 'Scope mismatch'];
  var hubspotDeals = [];
  var dealId = 1;
  for (var d = 0; d < 62; d++) {
    var openIdx = ri(0, months.length - 3);
    var openMo = months[openIdx];
    var closeIdx = Math.min(openIdx + ri(1, 5), months.length - 1);
    var closeMo = months[closeIdx];
    var isClosed = closeIdx < actualMonths.length && rng() < 0.62;
    var stage = isClosed ? (rng() < 0.58 ? 'Closed Won' : 'Closed Lost') : pick(STAGES.slice(0, 6));
    var amount = Math.round((28000 + rng() * 420000) / 1000) * 1000;
    var ent = pick(ENTITIES).name;
    var client = pick(CLIENTS);
    var alliance = pick(ALLIANCES);
    hubspotDeals.push({
      Record_ID: 'DEAL-' + String(dealId++).padStart(4, '0'), Omnia_Company: ent, Company: client.name, Brand: ent,
      Deal_Name: client.name.split(' ')[0] + ' — ' + pick(PROJECT_NOUNS),
      Deal_Owner: pick(FIRST_NAMES) + ' ' + pick(LAST_NAMES), Deal_Team: pick(ENTITIES).bus[0],
      Deal_Stage: stage, Amount: amount, Confidence: stage === 'Closed Won' ? 'High' : stage === 'Closed Lost' ? 'Low' : pick(['Low', 'Medium', 'High']),
      Close_Date: ymd(closeMo.y, closeMo.m, ri(1, 28)), Open_Date: ymd(openMo.y, openMo.m, ri(1, 28)),
      Last_Modified_Date: ymd(closeMo.y, closeMo.m, ri(1, 28)), Last_Updated: ymd(closeMo.y, closeMo.m, ri(1, 28)),
      Source_System: 'HubSpot', Contributing_OpCo_1: ent, Contributing_OpCo_2: '', Contributing_OpCo_3: '', Contributing_OpCo_4: '', Contributing_OpCo_5: '',
      Deal_Source: pick(['Referral', 'Outbound', 'Inbound', 'Partner']), Alliance_Partners: alliance,
      Logged_In_Partner_Portal: alliance !== 'N/A' && alliance !== 'Not yet known' ? pick(['Yes', 'No']) : '',
      Deal_Probability: STAGE_PROB[stage] != null ? STAGE_PROB[stage] : 0.3,
      Deal_Lost_Reason: stage === 'Closed Lost' ? pick(LOST_REASONS) : '', Why_Lost: '',
      Is_Stalled: !isClosed && rng() < 0.22 ? 'Yes' : 'No', Stage_Entered_Date: ymd(closeMo.y, closeMo.m, 1),
      Weighted_Open_Pipeline: isClosed ? 0 : Math.round(amount * (STAGE_PROB[stage] || 0.3)),
      Month: ymd(closeMo.y, closeMo.m, 1),
    });
  }

  var CERT_TITLES = ['Azure Data Engineer Associate', 'Power BI Data Analyst Associate', 'Azure AI Engineer Associate', 'Databricks Certified Data Engineer', 'Azure Solutions Architect Expert', 'Microsoft 365 Certified: Fundamentals'];
  var certifications = [];
  var today = new Date(2026, 7, 27);
  employeeRoster.slice(0, 26).forEach(function (emp, i) {
    var completedMo = actualMonths[ri(Math.max(0, actualMonths.length - 20), actualMonths.length - 1)];
    var expiresIn = ri(-2, 18);
    var expDate = new Date(today.getFullYear(), today.getMonth() + expiresIn, ri(1, 28));
    var nameParts = emp._name.split(' ');
    certifications.push({
      AADUserId: emp.userId, IndividualFirstName: nameParts[0], IndividualLastName: nameParts[1],
      Email: (nameParts[0] + '.' + nameParts[1]).toLowerCase() + '@example-demo.com',
      CorpEmail: (nameParts[0] + '.' + nameParts[1]).toLowerCase() + '@example-demo.com',
      PartnerName: emp.Omnia_Company, PartnerCountryLocation: 'Australia', PartnerCityLocation: pick(['Sydney', 'Melbourne', 'Brisbane']),
      MPNId: '4' + String(100000 + i), PGAMpnId: '4' + String(100000 + i),
      TrainingActivityId: 'TRN-' + String(2000 + i), TrainingType: 'Certification', TrainingTitle: pick(CERT_TITLES),
      ActivationStatus: 'Active', TrainingCompletionDate: ymd(completedMo.y, completedMo.m, ri(1, 28)),
      ExpirationDate: ymd(expDate.getFullYear(), expDate.getMonth() + 1, expDate.getDate()),
      Month: ymd(completedMo.y, completedMo.m, 1),
    });
  });

  return {
    keyMetrics: keyMetrics, unifiedFact: unifiedFact, hubspotDeals: hubspotDeals, refProjects: refProjects,
    employeeRoster: employeeRoster, revenueTarget: revenueTarget, utilisation: utilisation,
    certifications: certifications, targetRates: targetRates, pl: pl,
  };
}
