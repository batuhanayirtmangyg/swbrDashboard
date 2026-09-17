/**
 * Supply WBR — Fully Dynamic Dashboard
 * Backend (Google Apps Script, bound to the Supply raw-data spreadsheet).
 *
 * Architecture — server-side aggregation (v3):
 *   The metric sheets are far too large to ship to the browser (7 filter
 *   dimensions → a cross-product of millions of rows). So doGet() inlines only
 *   a tiny SHELL (week list + targets/meta). Each metric card then calls
 *   getSeries(metricKey, filters): the server reads THAT ONE sheet, applies the
 *   (multi-select, combinable) dimension filters, aggregates to a weekly series,
 *   and returns the small {week: value} map. Results are cached per
 *   (metric, filter-signature) so repeat views are instant. A warming trigger
 *   pre-populates the unfiltered ('ALL') case in the background.
 *
 *   Granularity (weekly/monthly/quarterly/yearly) is a CLIENT concern: the
 *   weekly map is granularity-independent and re-bucketed in the browser, so
 *   changing granularity never hits the server.
 *
 *   Local preview: make_snapshot.js inlines a small mock WITH raw rows
 *   (BOOT.metrics present) → the client aggregates in-browser. In production
 *   BOOT.metrics is absent → the client uses getSeries.
 */

var APP_TITLE   = 'Supply WBR';
var CACHE_TTL   = 6 * 60 * 60;   // 6h (CacheService max)
var CACHE_NS    = 'swbr_v3';     // bump to discard any older cached payloads
var WEEK_KEY    = 'week_start';
// Bump on every deploy. Surfaced in the UI badge via BOOT.build so you can tell at a
// glance whether the web app is serving the code you just pasted: the Apps Script editor
// always runs the newest save, but the /exec deployment stays pinned to the version you
// last published — so an un-bumped deployment looks identical to a code bug.
var BUILD       = 'admin-2026-07-29a';

// Spreadsheets holding the metric data.
var SS_DEMAND   = '1sagdK6fAC-yFNQIhWGJ6981jbZutpg7N2TfqJ5fSet4';
var SS_OVERFLOW = '1ZEaWjXpSjWX3H4QR8Y7AQTsXHYOOXglmrOyjr65z0ZE';
var SS_GMV      = '1KzxasSfVV3V4ihN1rQ10IAE7YzpskCAfytG66P8ubIE';
// Supply — the bound spreadsheet, and the home of the Comments / Users / Usage / Feedback
// tabs. Set SS_SUPPLY to its ID (same form as the three above) to stop depending on the
// binding: getActiveSpreadsheet() is null wherever no sheet is attached (standalone copies,
// some trigger and API entry points), which silently takes down every read at once. Leave
// empty to use the bound sheet; run whichSpreadsheet() to find out which one that is.
var SS_SUPPLY   = '16LkrpNJtvZ7Za-ZGmL_OxDVHZynqlTPIUvLbmaps19k';   // "swbr dashboard"

// Single seam for Supply access — every .gs file routes through it (Apps Script merges them
// into one global scope), so hard-coding the ID is a one-line change here.
function _supplySS() {
  if (SS_SUPPLY) return SpreadsheetApp.openById(SS_SUPPLY);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('No bound spreadsheet: set SS_SUPPLY in Code.gs to the Supply spreadsheet ID.');
  return ss;
}

// "Which sheet did that just write to?" — run from the editor; the answer is the URL.
function whichSpreadsheet() {
  var ss = _supplySS();
  var info = { id: ss.getId(), name: ss.getName(), url: ss.getUrl(),
               hardcoded: !!SS_SUPPLY,
               tabs: ss.getSheets().map(function (s) { return s.getName(); }) };
  Logger.log(JSON.stringify(info, null, 2));
  return info;
}

// ── Data source (provider seam) ───────────────────────────────────────────────
// 'snapshot' = read pre-aggregated artifacts from a Shared Drive (the v4 store,
// written nightly by Databricks); 'sheets' = the legacy in-Apps-Script sheet
// aggregation (kept as an instant fallback). Flip at runtime with the Script
// Property 'data_source' — no redeploy. See _snapshotSeries below.
var DEFAULT_DATA_SOURCE = 'snapshot';
var ARTIFACT_FOLDER_ID  = '1Zpbg5SUtyjLAPRXmHg2WcZ3H83j1wV6X';   // "Supply WBR Datasource" (Shared Drive)

// External target spreadsheet (hand-maintained, long-format, year-agnostic):
// Quarter · Calendar Week · Metric · Channel · In Week Value · Cumulative Value.
// Read into the shell as BOOT.targets. NOTE: web app runs as each viewer, so every
// viewer needs READ access to this spreadsheet.
var TARGETS_SS_ID = '1K9YmymwM-gcbOQ4rn1oFuL7tkGUOE8FXjTi49RV-G1Q';
function _dataSource() {
  return PropertiesService.getScriptProperties().getProperty('data_source') || DEFAULT_DATA_SOURCE;
}

// The 7 filter dimension columns (order = dropdown order on the client).
// `source_market` exists only on demand/overflow/gmv sheets; on supply sheets
// its absence makes that filter a silent no-op (correct behaviour).
var FILTER_COLS = [
  'tour_acquisition_channel',
  'supply_geo_area',
  'sales_office',
  'category',
  'destination_corridor',
  'source_corridor',
  'business_model'
];

// ── Metric specs — server mirror of the client METRIC_REGISTRY (data parts only) ─
// `ss` routes each metric to the spreadsheet that holds its authoritative data
// (this also resolves the gwc_nr / no_show_rate / gmv_yoy name collisions across
// spreadsheets — always read the mapped source, never a stale same-named stub).
// kind: 'sum' | 'mean' | 'wrate'. wrate carries {num,den} per row.
var METRIC_SPECS = {
  curation_duration:     { ss:'supply', sheet:'curation_duration',
    series:[ {key:'p90', col:'p90_days_in_curation', kind:'mean'},
             {key:'median', col:'median_days_in_curation', kind:'mean'},
             {key:'submitted', col:'tours_submitted', kind:'sum'} ] },
  curation_throughput:   { ss:'supply', sheet:'curation_throughput',
    series:[ {key:'submitted',       col:'tours_submitted', kind:'sum'},
             {key:'processed',       col:'processed_tours', kind:'sum'},
             {key:'unapproved_rate', kind:'wrate', num:'unapproved_cnt', den:'total_tours'} ] },
  processed_tours:       { ss:'supply', sheet:'processed_tours',
    series:[ {key:'processed', col:'processed_tours', kind:'sum'} ] },
  unprocessed_tours:     { ss:'supply', sheet:'unprocessed_tours', noDims:true,
    series:[ {key:'unprocessed', col:'unprocessed_tours', kind:'sum'},
             {key:'aged', col:'no_decision_14d', kind:'sum'} ] },
  curation_approval:     { ss:'supply', sheet:'curation_approval',
    series:[ {key:'rate', kind:'wrate', num:'approved_tours', den:'total_tours'},
             {key:'approved', col:'approved_tours', kind:'sum'} ] },
  rejection_reasons:     { ss:'supply', sheet:'rejection_reasons',
    series:[ {key:'rate', kind:'wrate', num:'total_rejected', den:'total_tours'} ] },
  tour_creation:         { ss:'supply', sheet:'tour_creation',
    series:[ {key:'cnv', kind:'wrate', num:'tours_win_7_days', den:'tours_created'},
             {key:'created', col:'tours_created', kind:'sum'} ] },
  registrations:         { ss:'supply', sheet:'registrations', noDims:true,
    series:[ {key:'reg', col:'registrations_completed', kind:'sum'} ] },
  activities_acquired:   { ss:'supply', sheet:'activities_acquired',
    series:[ {key:'online', col:'tours_went_online', kind:'sum'} ] },
  tours_by_channel:      { ss:'supply', sheet:'tours_by_channel',
    dim:{ col:'tour_acquisition_channel', kind:'count', valueCol:'tours_went_online' } },
  activities_activated:  { ss:'supply', sheet:'activities_activated',
    series:[ {key:'act', col:'activities_activated', kind:'sum'} ] },
  activation_rate:       { ss:'supply', sheet:'activation_rate',
    series:[ {key:'rate', kind:'wrate', num:'activated_tours', den:'tours_in_activation'} ] },
  '1b30d_by_channel':    { ss:'supply', sheet:'1b30d_by_channel',
    dim:{ col:'channel', kind:'rate', num:'activated_tours', den:'tours_in_activation' } },
  sellout:               { ss:'supply', sheet:'sellout',
    series:[ {key:'total',  kind:'wrate'},
             {key:'top',    kind:'wrate'},
             {key:'nontop', kind:'wrate'} ] },
  // wrate, not mean (Databricks 2026-08-06): the notebook stopped shipping a pre-computed
  // percentage and now emits raw num/den, so `value` is NULL for these two and a `mean` read
  // produced an empty series — the metrics simply vanished from the dashboard. Aggregating
  // Sum(num)/Sum(den) is also the correct answer: averaging per-slice rates is Simpson's paradox
  // and read 77.7% against the memo's 71.9%. num/den names are for the SHEETS fallback only;
  // the artifact path reads the num/den arrays positionally and ignores them.
  uptime:                { ss:'supply', sheet:'uptime',
    series:[ {key:'up', kind:'wrate', num:'uptime_num', den:'uptime_den'} ] },
  price_parity:          { ss:'supply', sheet:'price_parity',
    series:[ {key:'win', kind:'wrate', num:'win_impressions', den:'total_impressions'} ] },
  price_coverage:        { ss:'supply', sheet:'price_coverage',
    series:[ {key:'cov', col:'coverage_pct', kind:'sum'} ] },
  atc_error_rate:        { ss:'supply', sheet:'atc_error_rate',
    series:[ {key:'overall',      kind:'wrate'},
             {key:'connected',    kind:'wrate'},
             {key:'notconnected', kind:'wrate'} ] },
  gwc_nr:                { ss:'demand', sheet:'gwc_nr',
    series:[ {key:'rate', kind:'wrate', num:'total_gwc', den:'travel_nr'} ] },
  supplier_cancellation: { ss:'demand', sheet:'supplier_cancellation',
    series:[ {key:'nonfm', kind:'wrate', num:'non_fm_cancellations', den:'travel_bookings'},
             {key:'fm', kind:'wrate', num:'all_fm_cancellations', den:'travel_bookings'} ] },
  no_show_rate:          { ss:'overflow', sheet:'no_show_rate',
    series:[ {key:'rate', kind:'wrate', num:'no_show_bookings', den:'bookings'} ] },
  reseller_rate:         { ss:'overflow', sheet:'reseller_rate',
    series:[ {key:'rate', kind:'wrate', num:'reseller_bookings', den:'total_bookings'} ] },
  gmv_yoy:               { ss:'gmv', sheet:'gmv_yoy',
    series:[ {key:'total', col:'gmv_total', kind:'sum'},
             {key:'tp', col:'gmv_tp', kind:'sum'},
             {key:'dg', col:'gmv_dg', kind:'sum'} ] },
  net_revenue:           { ss:'demand', sheet:'net_revenue',
    series:[ {key:'total', col:'total', kind:'sum'} ] },
  bookings:              { ss:'demand', sheet:'bookings',
    series:[ {key:'total', col:'total', kind:'sum'} ] },
  alv:                   { ss:'demand', sheet:'alv',
    series:[ {key:'total', col:'total', kind:'sum'} ] }
};

// Non-metric tabs in the Supply spreadsheet: never metric data, and never inlined into
// BOOT.aux by _buildShell (Feedback and Usage hold every user's email — admin-gated RPCs
// are the only read path).
var CMT_TABS = ['Comments', 'Replies', 'Users', 'Usage', 'Feedback', 'Snapshot_Diff'];

// ── Web app entry ──────────────────────────────────────────────────────────
// doGet inlines only the SHELL (small). Metric series are fetched per card.
function doGet(e) {
  var t = HtmlService.createTemplateFromFile('index');
  var shell = null;
  try { shell = getShell(); } catch (err) { shell = null; }
  t.bootstrapJson = shell ? JSON.stringify(shell) : 'null';
  t.appTitle = APP_TITLE;
  var p = (e && e.parameter) || {};
  // Deep link. commentId/view come from @mention emails; t/gran/week/f/s come from a tile's
  // Share button and pin the whole view (target tile, granularity, as-of week, filters, series
  // chip). One JSON blob so index.html needs no change when params are added.
  t.deepLinkJson = JSON.stringify({
    commentId: p.commentId || '', view: p.view || '',
    t: p.t || '', gran: p.gran || '', week: p.week || '', f: p.f || '', s: p.s || '',
    c: p.c || '', o: p.o || ''   // compare-tab state / per-tile ⊞ toggles
  });
  var userEmail = '';
  try { userEmail = Session.getActiveUser().getEmail() || ''; } catch(e1) {}
  // Compatibility shim only — the client now reads getAdminState() instead. Keep this until
  // the index.html that dropped `<?!= isAdminFlag ?>` is confirmed deployed: an unset
  // template variable makes t.evaluate() throw and takes the whole dashboard down.
  t.isAdminFlag = isAdmin(userEmail);
  try { logVisit('page_load', p.view || '', '', ''); } catch(e2) {}
  return t.evaluate()
    .setTitle(APP_TITLE)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}

// ── Analytics / admin ──────────────────────────────────────────────────────
var USAGE_SHEET   = 'Usage';
var USAGE_HEADERS = ['Timestamp','Email','DisplayName','Action','Tab','Week','Granularity'];

function _adminEmails() {
  return (PropertiesService.getScriptProperties().getProperty('ADMIN_EMAILS') || '')
    .split(',').map(function(e){ return e.trim().toLowerCase(); }).filter(Boolean);
}
function isAdmin(email) {
  return !!email && _adminEmails().indexOf(email.toLowerCase()) !== -1;
}

// Why the admin gate silently fails, made visible. Run from the editor to prove the
// ADMIN_EMAILS property exists on THIS project; run from the deployed page's devtools
// (google.script.run.withSuccessHandler(console.log).diagnoseAdmin()) to see the email the
// web app actually sees. activeEmail:'' with a non-empty effectiveEmail = the deployment is
// "Execute as: Me", so no viewer can ever be an admin. Never fall back to the effective
// user — under execute-as-me that would make every viewer an admin of the Usage log.
// Returns the caller's own identity only, never the roster: every top-level function here
// is callable by any viewer.
function diagnoseAdmin() {
  var active = '', effective = '';
  try { active    = Session.getActiveUser().getEmail()    || ''; } catch (e) {}
  try { effective = Session.getEffectiveUser().getEmail() || ''; } catch (e) {}
  return { activeEmail: active, effectiveEmail: effective,
           adminCount: _adminEmails().length, isAdmin: isAdmin(active), build: BUILD };
}

// Who is viewing, fetched after first paint (see applyAdminState in javascript.html). This
// replaced a template-substituted flag: the flag and the button it controlled lived in
// index.html while the code using them lived in javascript.html, so deploying one file
// without the other hid the whole feature with no error anywhere.
// Per-viewer, so it must NEVER be folded into the shell — getShell() caches under one shared
// key (CACHE_NS + ':shell'), which would serve the first viewer's identity to everyone.
function getAdminState() {
  var email = '';
  try { email = Session.getActiveUser().getEmail() || ''; } catch (e) {}
  // url: the /exec address, needed client-side to build tile share links. The page itself runs in a
  // sandboxed googleusercontent.com iframe, so location.href there is NOT the app URL — the server
  // is the only place that knows it. Rides on this call because it is already made after first
  // paint; it must not move into getShell(), which caches under one shared key.
  var url = '';
  try { url = ScriptApp.getService().getUrl() || ''; } catch (e2) {}
  return { email: email, isAdmin: isAdmin(email),
           adminCount: _adminEmails().length, build: BUILD, url: url };
}

function setupUsageSheet() {
  var ss = _supplySS();
  var sh = ss.getSheetByName(USAGE_SHEET) || ss.insertSheet(USAGE_SHEET);
  if (sh.getLastRow() < 1) {
    sh.getRange(1, 1, 1, USAGE_HEADERS.length).setValues([USAGE_HEADERS]);
    sh.setFrozenRows(1);
  }
  Logger.log('setupUsageSheet: done — ' + ss.getName() + ' ' + ss.getUrl());
  return 'ok';
}

function logVisit(action, tab, week, gran) {
  var u = Session.getActiveUser();
  var sh = _supplySS().getSheetByName(USAGE_SHEET);
  if (!sh) return;
  sh.appendRow([new Date().toISOString(), u.getEmail() || '', u.getName() || '',
                action || 'page_load', tab || '', week || '', gran || '']);
}

function getUsageLog() {
  var email = '';
  try { email = Session.getActiveUser().getEmail() || ''; } catch(e) {}
  if (!isAdmin(email)) return { error: 'unauthorized' };
  var sh = _supplySS().getSheetByName(USAGE_SHEET);
  if (!sh || sh.getLastRow() < 2) return { rows: [], headers: USAGE_HEADERS };
  var data = sh.getRange(2, 1, sh.getLastRow() - 1, USAGE_HEADERS.length).getValues();
  return { rows: data, headers: USAGE_HEADERS };
}

// ── Spreadsheet routing ──────────────────────────────────────────────────────
function _ssFor(tag) {
  if (tag === 'demand')   return SpreadsheetApp.openById(SS_DEMAND);
  if (tag === 'overflow') return SpreadsheetApp.openById(SS_OVERFLOW);
  if (tag === 'gmv')      return SpreadsheetApp.openById(SS_GMV);
  return _supplySS();   // supply (bound)
}
function _tz() {
  return _supplySS().getSpreadsheetTimeZone() || 'UTC';
}

// ── Shell: week list + aux (targets/meta). Small; inlined by doGet, cached. ───
function getShell() {
  var cached = _cacheGet(CACHE_NS + ':shell');
  if (cached) return cached;
  var shell = _buildShell();
  _cachePut(CACHE_NS + ':shell', shell);
  return shell;
}
function refreshShell() { _cacheBust(CACHE_NS + ':shell'); return getShell(); }

function _buildShell() {
  var tz = _tz();
  var activeSS = _supplySS();
  var aux = {};
  var weekSet = {};

  // AUX = every non-metric, non-comment tab in the Supply sheet (targets_*, meta).
  var metricSheetNames = {};
  Object.keys(METRIC_SPECS).forEach(function (k) { metricSheetNames[METRIC_SPECS[k].sheet] = true; });

  activeSS.getSheets().forEach(function (sh) {
    var name = sh.getName();
    if (sh.isSheetHidden()) return;
    if (CMT_TABS.indexOf(name) >= 0) return;
    if (metricSheetNames[name]) return;               // metric tabs fetched on demand
    var data = _readSheet(sh, tz, 4000);              // aux tabs are tiny
    if (data && data.headers.indexOf(WEEK_KEY) < 0) aux[name] = data;
  });

  // Canonical week list. In snapshot mode it comes from the artifacts (the sheets
  // may lag or eventually be archived); otherwise read the week_start column of a
  // small metric sheet.
  var weeks;
  if (_dataSource() === 'snapshot') {
    weeks = _snapshotWeeks();
  } else {
    ['registrations', 'unprocessed_tours', 'processed_tours'].some(function (nm) {
      var sh = activeSS.getSheetByName(nm);
      if (!sh) return false;
      var last = sh.getLastRow();
      if (last < 2) return false;
      var col = _readSheet(sh, tz, 4000);
      var wi = col.headers.indexOf(WEEK_KEY);
      if (wi < 0) return false;
      col.rows.forEach(function (r) { if (r[wi]) weekSet[r[wi]] = true; });
      return Object.keys(weekSet).length > 0;
    });
    weeks = Object.keys(weekSet).sort().reverse();
  }
  return { generatedAt: new Date().toISOString(), build: BUILD, salt: _salt(),
           weeks: weeks, aux: aux, targets: _readTargets() };
}

// ── Targets: read the external long-format target sheet into a compact lookup ──
// BOOT.targets = { <metricNameNorm>: { _name, _pct, <channelNorm>: { 'Q#|cw':{w,c},
// 'cw<cw>':{w,c} } } }.  w = In Week Value, c = Cumulative Value (numbers; % parsed
// to their numeric value, e.g. "14.70%" -> 14.7). Boundary weeks (which appear in
// two quarters) are keyed per quarter, with a quarter-agnostic 'cw<cw>' fallback.
function _tnorm(s) { return String(s == null ? '' : s).toLowerCase().replace(/[^a-z0-9]/g, ''); }
function _tParseNum(v) {
  if (v === '' || v === null || v === undefined) return null;
  if (typeof v === 'number') return v;
  var n = parseFloat(String(v).replace(/[,%\s]/g, ''));
  return isNaN(n) ? null : n;
}
function _readTargets() {
  try {
    var ss = SpreadsheetApp.openById(TARGETS_SS_ID);
    var sh = ss.getSheets()[0]; if (!sh) return null;
    var vals = sh.getDataRange().getDisplayValues(); if (vals.length < 2) return null;
    var hi = -1;
    for (var i = 0; i < Math.min(vals.length, 6); i++) {
      var rl = vals[i].map(function (c) { return String(c).toLowerCase(); }).join('|');
      if (rl.indexOf('metric') >= 0 && rl.indexOf('calendar week') >= 0) { hi = i; break; }
    }
    if (hi < 0) return null;
    var H = vals[hi].map(function (c) { return String(c).trim().toLowerCase(); });
    var qi = H.indexOf('quarter'), cwi = H.indexOf('calendar week'), mi = H.indexOf('metric'),
        chi = H.indexOf('channel'), wi = H.indexOf('in week value'), cvi = H.indexOf('cumulative value');
    if (mi < 0 || cwi < 0 || wi < 0) return null;
    var out = {};
    for (var r = hi + 1; r < vals.length; r++) {
      var row = vals[r];
      var name = String(row[mi] || '').trim(); if (!name) continue;
      var cw = parseInt(row[cwi], 10); if (isNaN(cw)) continue;
      var q = qi >= 0 ? String(row[qi] || '').trim().toUpperCase() : '';
      var ch = (chi >= 0 ? String(row[chi] || '').trim() : '') || 'Total';
      var wCell = row[wi], cCell = cvi >= 0 ? row[cvi] : '';
      var nk = _tnorm(name);
      var byName = out[nk] || (out[nk] = { _name: name, _pct: false });
      if ((typeof wCell === 'string' && wCell.indexOf('%') >= 0) ||
          (typeof cCell === 'string' && cCell.indexOf('%') >= 0)) byName._pct = true;
      var ck = _tnorm(ch);
      var byCh = byName[ck] || (byName[ck] = {});
      var cell = { w: _tParseNum(wCell), c: _tParseNum(cCell) };
      if (q) byCh[q + '|' + cw] = cell;
      byCh['cw' + cw] = cell;                       // quarter-agnostic fallback (last row wins)
    }
    return out;
  } catch (e) { Logger.log('targets read failed: ' + e); return null; }
}
// Editor diagnostic: list the distinct target Metric names + channels + unit, so the
// dashboard TARGET_MAP can be completed with the exact sheet strings.
function diagnoseTargets() {
  var t = _readTargets();
  if (!t) { Logger.log('targets: sheet unreadable/empty (check sharing + structure)'); return; }
  Object.keys(t).forEach(function (nk) {
    var nm = t[nk];
    var chans = Object.keys(nm).filter(function (k) { return k.charAt(0) !== '_'; });
    Logger.log('"' + nm._name + '"' + (nm._pct ? ' [%]' : ' [count]') + ' — channels: ' + chans.join(', '));
  });
}

// ── Per-metric series — the workhorse RPC ─────────────────────────────────────
// filters: { colName: [selectedValues...] }  (multi-select, combinable, AND).
// Returns { series:[{key,label?,kind,isDim,member?,weekly}], domains:{col:[vals]} }.
// Cached per (metric, filter-signature).
function getSeries(metricKey, filters) {
  var spec = METRIC_SPECS[metricKey];
  if (!spec) return { series: [], domains: {} };
  var rawFilters = filters;  // capture before normalization for recursive calls
  filters = _normFilters(filters);
  var sig = _filterSig(filters);
  var src = _dataSource();
  var ckey = CACHE_NS + ':m:' + metricKey + ':' + src + ':' + _salt() + ':' + sig;

  var cached = _cacheGet(ckey);
  if (cached) return cached;

  var out;
  if (metricKey === 'curation_throughput') {
    // Prefer a real artifact when Databricks has written one; fall back to the
    // virtual stitch (curation_duration.submitted + processed_tours.processed)
    // only when no artifact exists in the manifest.
    var man = _manifest(), manEntry = man.metrics && man.metrics[metricKey];
    if (src === 'snapshot' && manEntry) {
      out = _snapshotSeries(metricKey, spec, filters);
    } else if (src !== 'snapshot') {
      var raw = _readMetricSheet(spec);
      out = _aggregate(spec, raw.headers, raw.rows, filters);
    } else {
      // Virtual fallback: artifact not in manifest → stitch from sibling metrics.
      var subOut = getSeries('curation_duration', rawFilters);
      var procOut = getSeries('processed_tours', rawFilters);
      var subSer = (subOut.series || []).filter(function(s){ return s.key === 'submitted'; });
      var procSer = (procOut.series || []).filter(function(s){ return s.key === 'processed'; });
      out = { series: subSer.concat(procSer), domains: subOut.domains || {} };
    }
  } else if (src === 'snapshot') {
    out = _snapshotSeries(metricKey, spec, filters);   // Drive artifacts (v4)
  } else {
    var raw = _readMetricSheet(spec);                  // legacy: one sheet, no row cap
    out = _aggregate(spec, raw.headers, raw.rows, filters);
  }
  // One guard for every provider (snapshot / sheets / virtual stitch): without the
  // quarter-split grain, periodSet falls back to s.weekly and drops the whole week that
  // straddles the quarter boundary — ~20% of a QTD figure in the first weeks of a quarter.
  // The client surfaces this as a banner; silence here is what let it ship unnoticed.
  if (!(out.series || []).some(function (s) { return !!s.qtd_weekly; })) out.qtdFallback = true;
  _cachePut(ckey, out);
  return out;
}

// Batched convenience for the client (still one sheet read per metric).
function getSeriesBatch(metricKeys, filters) {
  var res = {};
  (metricKeys || []).forEach(function (k) { res[k] = getSeries(k, filters); });
  return res;
}

// Read a metric's sheet in full (no artificial row cap — the whole point is to
// stop the old MAX_ROWS truncation that was silently dropping data).
var MAX_COLS = 60;
function _readMetricSheet(spec) {
  var sh = _ssFor(spec.ss).getSheetByName(spec.sheet);
  if (!sh) return { headers: [], rows: [] };
  return _readSheet(sh, _tz(), 0) || { headers: [], rows: [] };
}

// Generic sheet reader. rowCap 0 = no cap (read to real last row).
// Date cells are formatted to yyyy-MM-dd, but MEMOIZED by timestamp: a 300k-row
// sheet has only ~25-105 distinct week values, so this collapses ~300k expensive
// Utilities.formatDate() calls down to a few dozen (the single biggest speedup
// vs the original per-cell formatting). Rows are also mutated in place, avoiding
// a per-row array allocation.
function _readSheet(sh, tz, rowCap) {
  var lastRow = sh.getLastRow();
  var lastCol = Math.min(sh.getLastColumn(), MAX_COLS);
  if (rowCap && lastRow > rowCap) lastRow = rowCap;
  if (lastRow < 2 || lastCol < 1) return { headers: (lastRow >= 1 ? sh.getRange(1,1,1,lastCol).getValues()[0].map(String) : []), rows: [] };
  var values = sh.getRange(1, 1, lastRow, lastCol).getValues();
  var headers = values[0].map(function (h) { return String(h).trim(); });
  var rows = [];
  var dcache = {};                                   // timestamp -> yyyy-MM-dd
  for (var i = 1; i < values.length; i++) {
    var row = values[i], empty = true;
    for (var c = 0; c < row.length; c++) {
      var v = row[c];
      if (v instanceof Date) {
        var t = v.getTime(), s = dcache[t];
        if (s === undefined) { s = dcache[t] = Utilities.formatDate(v, tz, 'yyyy-MM-dd'); }
        row[c] = s; empty = false;
      } else if (v !== '' && v !== null) { empty = false; }
    }
    if (!empty) rows.push(row);
  }
  return { headers: headers, rows: rows };
}

// ── Aggregation — mirrors the client computeMetric() exactly ──────────────────
function _num(v) {
  if (v === null || v === undefined || v === '') return null;
  var n = Number(v); return isNaN(n) ? null : n;
}
// Aggregate a metric's rows to weekly series under `filters`. All column indices
// are resolved ONCE up front — the hot loop over (up to 300k) rows touches only
// array indices, never headers.indexOf(), which is what made warmAll blow the
// 6-min limit. Filter sets become hash maps for O(1) membership.
function _aggregate(spec, headers, rows, filters) {
  var wi = headers.indexOf(WEEK_KEY);
  var series = [];
  var domains = {};

  // domain columns present in this sheet → [{ci, set}]
  var domCols = [];
  for (var fi = 0; fi < FILTER_COLS.length; fi++) {
    var dci = headers.indexOf(FILTER_COLS[fi]);
    if (dci >= 0) domCols.push({ col: FILTER_COLS[fi], ci: dci, set: {} });
  }
  var nDom = domCols.length;

  // active filters present in this sheet → [{ci, map}] (absent cols are no-ops)
  var act = [];
  for (var ai = 0; ai < filters.active.length; ai++) {
    var aci = headers.indexOf(filters.active[ai].col);
    if (aci < 0) continue;
    var map = {}; filters.active[ai].set.forEach(function (v) { map[v] = true; });
    // col carried so the CELL value can be grouped before it is tested — 'Inbound' is the
    // complement of the named sales channels, not a literal the sheet necessarily holds.
    act.push({ ci: aci, map: map, col: filters.active[ai].col });
  }
  var nAct = act.length;

  if (spec.dim) {
    var di = headers.indexOf(spec.dim.col);
    var vc = spec.dim.kind === 'count' ? headers.indexOf(spec.dim.valueCol) : -1;
    var ni = spec.dim.kind === 'rate' ? headers.indexOf(spec.dim.num) : -1;
    var dgi = spec.dim.kind === 'rate' ? headers.indexOf(spec.dim.den) : -1;
    var members = [], totalWeekly = {}, memberWeekly = {};
    for (var r = 0; r < rows.length; r++) {
      var row = rows[r], wk = row[wi]; if (!wk) continue;
      for (var dq = 0; dq < nDom; dq++) { var dv = row[domCols[dq].ci]; if (dv !== '' && dv !== null) domCols[dq].set[_normDim(domCols[dq].col, dv)] = true; }
      var skip = false;
      for (var aq = 0; aq < nAct; aq++) { if (!act[aq].map[String(_normDim(act[aq].col, row[act[aq].ci]))]) { skip = true; break; } }
      if (skip) continue;
      // Grouped here so two raw channels meaning Inbound merge into ONE member series.
      var mem = String(_normDim(spec.dim.col, row[di]));
      var mw = memberWeekly[mem]; if (!mw) { mw = memberWeekly[mem] = {}; members.push(mem); }
      if (spec.dim.kind === 'count') {
        var v = _num(row[vc]) || 0;
        mw[wk] = (mw[wk] || 0) + v;
        totalWeekly[wk] = (totalWeekly[wk] || 0) + v;
      } else {
        var nn = _num(row[ni]) || 0, dd = _num(row[dgi]) || 0;
        var prev = mw[wk] || { num:0, den:0 }; prev.num += nn; prev.den += dd; mw[wk] = prev;
        var t = totalWeekly[wk] || { num:0, den:0 }; t.num += nn; t.den += dd; totalWeekly[wk] = t;
      }
    }
    var dk = spec.dim.kind === 'count' ? 'sum' : 'wrate';
    series.push({ key:'total', label:'Total', kind:dk, isDim:false, weekly:totalWeekly });
    members.sort().forEach(function (mem) {
      series.push({ key:'dim__' + mem, label:mem, kind:dk, isDim:true, member:mem, weekly:memberWeekly[mem] });
    });
  } else {
    var defs = spec.series.map(function (s) {
      return { s:s, ci: s.col ? headers.indexOf(s.col) : -1,
               ni: s.num ? headers.indexOf(s.num) : -1,
               di: s.den ? headers.indexOf(s.den) : -1, wacc:{} };
    });
    var nDefs = defs.length;
    for (var r2 = 0; r2 < rows.length; r2++) {
      var row2 = rows[r2], wk2 = row2[wi]; if (!wk2) continue;
      for (var dq2 = 0; dq2 < nDom; dq2++) { var dv2 = row2[domCols[dq2].ci]; if (dv2 !== '' && dv2 !== null) domCols[dq2].set[_normDim(domCols[dq2].col, dv2)] = true; }
      var skip2 = false;
      for (var aq2 = 0; aq2 < nAct; aq2++) { if (!act[aq2].map[String(_normDim(act[aq2].col, row2[act[aq2].ci]))]) { skip2 = true; break; } }
      if (skip2) continue;
      for (var d = 0; d < nDefs; d++) {
        var def = defs[d], a = def.wacc[wk2];
        if (!a) a = def.wacc[wk2] = { sum:0, n:0, num:0, den:0 };
        if (def.s.kind === 'wrate') { a.num += _num(row2[def.ni]) || 0; a.den += _num(row2[def.di]) || 0; }
        else { var vv = _num(row2[def.ci]); if (vv !== null) { a.sum += vv; a.n++; } }
      }
    }
    defs.forEach(function (def) {
      var weekly = {};
      Object.keys(def.wacc).forEach(function (wk) {
        var a = def.wacc[wk];
        weekly[wk] = (def.s.kind === 'wrate') ? (a.den ? { num:a.num, den:a.den } : null)
                   : (def.s.kind === 'mean')  ? (a.n ? a.sum / a.n : null)
                   : a.sum;
      });
      series.push({ key:def.s.key, kind:def.s.kind, isDim:false, weekly:weekly });
    });
  }

  domCols.forEach(function (dc) {
    domains[dc.col] = Object.keys(dc.set).map(String).sort();
  });
  return { series: series, domains: domains };
}

// ── Channel grouping (FPOPS-75). MUST mirror normalizeDimMember() in javascript.html ──
// "Inbound" is a COMPLEMENT, not a value: it is every acquisition channel that is not one of
// the two named sales channels. Enumerating the aliases the source happens to use today
// ('Marketing' / 'Organic' / 'Unknown') silently drops any new value into nothing — and that
// is precisely how this broke: the client displayed and requested 'Inbound' while the artifact
// dict held the raw names, so EVERY filtered fetch on that column matched zero rows and
// returned '—' with no error anywhere. The free-dim breakdowns looked fine because those
// members were relabelled after the fact, which is what made it look like a display bug.
// The two maps are function-LOCAL on purpose: the test harnesses grab() individual function
// declarations out of this file, so a module-scope var beside them reads undefined there.
function _normDim(col, v) {
  if (col !== 'tour_acquisition_channel' && col !== 'channel') return v;
  return ({ 'inhousesales': 'In House Sales', 'inhouse': 'In House Sales',
            'insidesales': 'Inside Sales' })[_tnorm(v)] || 'Inbound';
}

// filters {col:[vals]} → { active:[{col,set}], sig } normaliser.
// The requested values are grouped here, at the single chokepoint every provider goes
// through, so a link carrying a raw legacy value ('Marketing') resolves to the same rows —
// and the same 6h cache entry — as the 'Inbound' the UI offers.
function _normFilters(filters) {
  var active = [];
  if (filters) {
    FILTER_COLS.forEach(function (col) {
      var vals = filters[col];
      if (vals && vals.length) {
        var seen = {}, set = [];
        vals.map(String).forEach(function (v) {
          var g = _normDim(col, v);
          if (!seen[g]) { seen[g] = true; set.push(g); }
        });
        active.push({ col: col, set: set });
      }
    });
  }
  return { active: active };
}
function _filterSig(f) {
  if (!f.active.length) return 'ALL';
  return f.active.map(function (a) {
    return a.col + '=' + a.set.slice().sort().join(',');
  }).sort().join('|');
}

// ══ Snapshot provider (v4): pre-aggregated artifacts from a Shared Drive ══════
// The nightly Databricks job writes one gzipped columnar JSON per metric to
// ARTIFACT_FOLDER_ID, plus manifest.json. getSeries routes here when
// data_source='snapshot'. Output is byte-for-byte the same {series,domains}
// shape as the legacy _aggregate, so nothing downstream changes.
//
// Two artifact layouts (self-describing via the top-level `layout` field):
//   'full'      — the whole week x dim cross-product; any filter combo aggregable.
//   'marginals' — unfiltered (__ALL__) + per-single-dimension aggregates only, for
//                 the 5 mega metrics too large to parse in Apps Script. Answers
//                 unfiltered and single-dimension filters (incl. multi-select)
//                 EXACTLY (sum -> Σvalue, wrate -> Σnum/Σden). A 2+ dimension combo
//                 is not representable and returns { unsupported:true } (that combo
//                 is served by the live-SQL tier once provisioned).

// The gold table's series_key names were built independently of the dashboard
// registry keys, so they don't all match. Map registryKey -> gold series_key here
// (only where they differ) so _aggFull/_aggMarginals match the right rows; this
// map is also the reference for the future live-SQL tier. NOTE: curation_approval's
// 'approved' series is simply absent from the gold table (add it upstream); its
// primary 'rate' series works regardless.
var ART_SERIES_ALIAS = {
  activities_acquired:  { online: 'total' },
  activities_activated: { act: 'total' },
  processed_tours:      { processed: 'total' },
  registrations:        { reg: 'total' },
  unprocessed_tours:    { aged: 'no_decision_14d' },
  tour_creation:        { cnv: 'rate', created: 'tours_created' },
  // sellout series names match registry directly (top, nontop) — no alias needed
  uptime:               { up: 'rate' },
  price_parity:         { win: 'rate' },
  price_coverage:       { cov: 'total' },
  atc_error_rate:       { notconnected: 'not_connected' }
};

function _artNum(v) { return (v === null || v === undefined) ? 0 : (Number(v) || 0); }

// name -> Drive fileId for everything in the artifact folder (one listing, cached).
// getFileById is proven to work on this Shared Drive; a folder index avoids a
// per-file getFilesByName and tolerates stale extra files sitting in the folder.
// Pass fresh=true to bypass the cached listing (a pipeline run that ADDS a filename is
// otherwise invisible until the salt is bumped or CACHE_TTL lapses).
function _folderIndex(fresh) {
  var ck = CACHE_NS + ':folderidx:' + _salt();
  if (!fresh) { var cached = _cacheGet(ck); if (cached) return cached; }
  var idx = {};
  var files = DriveApp.getFolderById(ARTIFACT_FOLDER_ID).getFiles();
  while (files.hasNext()) { var f = files.next(); idx[f.getName()] = f.getId(); }
  _cachePut(ck, idx);
  return idx;
}
function _driveText(name) {
  var id = _folderIndex()[name];
  if (!id) throw new Error('file not in artifact folder: ' + name);
  return DriveApp.getFileById(id).getBlob().getDataAsString();
}
function _fetchArtifact(name) {
  var id = _folderIndex()[name];
  if (!id) throw new Error('artifact not in folder: ' + name);
  var blob = DriveApp.getFileById(id).getBlob().setContentType('application/x-gzip');
  return JSON.parse(Utilities.ungzip(blob).getDataAsString());
}
function _manifest() {
  var ck = CACHE_NS + ':manifest:' + _salt();
  var cached = _cacheGet(ck);
  if (cached) return cached;
  var man = JSON.parse(_driveText('manifest.json'));
  _cachePut(ck, man);
  return man;
}
// Week list for the shell — from a tiny artifact (registrations ~1KB), else the first metric.
function _snapshotWeeks() {
  var ck = CACHE_NS + ':snapweeks:' + _salt();
  var cached = _cacheGet(ck);
  if (cached) return cached;
  var man = _manifest(), metrics = man.metrics || {};
  var pick = metrics.registrations ? 'registrations' : Object.keys(metrics)[0];
  var weeks = [];
  if (pick && metrics[pick].files && metrics[pick].files.length) {
    var art = _fetchArtifact(metrics[pick].files[0]);
    weeks = (art.weeks || []).slice().sort().reverse();
  }
  _cachePut(ck, weeks);
  return weeks;
}

// QTD artifact filenames for a metric, resolved from the Drive folder (the manifest does
// not register _qtd keys). Prefers the single-file artifact and only falls back to _part
// shards — never mixes the two. A stale _part set left behind when a metric moved to the
// compact marginals layout would otherwise also match the prefix, dragging tens of MB of
// dead weight into every cache miss and leaving which artifact wins up to sort order.
var _QTD_RELISTED = false;   // at most one cache-bypassing Drive listing per execution
function _qtdFiles(metricKey) {
  var qtdKey = metricKey + '_qtd', single = qtdKey + '.json.gz', idx = _folderIndex();
  // A cached listing that predates the artifact upload makes the single file look absent,
  // which is indistinguishable from "no QTD artifact exists" AND lets stale _part shards
  // win. Re-list once (per execution, not per metric) before believing the miss.
  if (!idx[single] && !_QTD_RELISTED) { _QTD_RELISTED = true; idx = _folderIndex(true); }
  if (idx[single]) return [single];
  return Object.keys(idx).filter(function (f) {
    return f.indexOf(qtdKey + '_part') === 0 && /\.json\.gz$/.test(f);
  }).sort();
}
function _snapshotSeries(metricKey, spec, filters) {
  var man = _manifest();
  var entry = man.metrics && man.metrics[metricKey];
  if (!entry || !entry.files || !entry.files.length) return { series: [], domains: {} };
  var arts = entry.files.map(_fetchArtifact);
  var layout = arts[0].layout || 'full';
  var out = (layout === 'marginals')
    ? _aggMarginals(arts[0], spec, filters, metricKey)
    : _aggFull(arts, spec, filters, metricKey);
  var qtdKey = metricKey + '_qtd';
  var qtdFiles = _qtdFiles(metricKey);
  if (!qtdFiles.length) {
    // No QTD artifact -> periodSet falls back to s.weekly, which under-counts any
    // quarter-straddling week. Never silent: getSeries flags it on the payload.
    Logger.log('QTD MISS ' + metricKey + ': no files matching "' + qtdKey + '*.json.gz"');
  } else {
    try {
      var qtdArts = qtdFiles.map(_fetchArtifact);
      var qtdLayout = qtdArts[0].layout || 'full';  // use QTD artifact's own layout, not main's
      var qtdOut = (qtdLayout === 'marginals')
        ? _aggMarginals(qtdArts[0], spec, filters, metricKey)
        : _aggFull(qtdArts, spec, filters, metricKey);
      var hit = 0;
      qtdOut.series.forEach(function(qs) {
        var match = out.series.filter(function(s){ return s.key === qs.key; })[0];
        if (match) { match.qtd_weekly = qs.weekly; hit++; }
      });
      if (!hit) Logger.log('QTD MISS ' + metricKey + ': no series key matched. main=[' +
        out.series.map(function(s){return s.key;}) + '] qtd=[' + qtdOut.series.map(function(s){return s.key;}) + ']');
    } catch(e) {
      // Most likely cause: the QTD artifact is full-layout and too large to ungzip/parse
      // in Apps Script (see diagnoseQtd). Fall back rather than break the whole metric.
      Logger.log('QTD FAIL ' + metricKey + ' (' + qtdFiles.join(', ') + '): ' + e);
    }
  }
  return out;
}

// Aggregate FULL-layout artifact shards to weekly series. Mirrors _aggregate:
// sum -> Σvalue, mean -> Σvalue/n (mean-of-rows), wrate -> {Σnum,Σden}; dim
// metrics -> total + one dim__<member> series per breakdown value.
function _aggFull(arts, spec, filters, metricKey) {
  var domains = {};
  var isDimMetric = !!spec.dim;
  var defs, byKey, totalWeekly, memberWeekly, members, dk;

  if (isDimMetric) {
    totalWeekly = {}; memberWeekly = {}; members = {};
    dk = (spec.dim.kind === 'count') ? 'sum' : 'wrate';
  } else {
    var alias = ART_SERIES_ALIAS[metricKey] || {};
    defs = spec.series.map(function (s) { return { key: s.key, kind: s.kind, art: alias[s.key] || s.key }; });
    byKey = {};
    // Register BOTH the registry key and the gold-table key -> same slot, so the
    // provider matches whether the artifact uses canonical names or the legacy ones.
    defs.forEach(function (d) { var slot = { kind: d.kind, key: d.key, acc: {} }; byKey[d.key] = slot; byKey[d.art] = slot; });
  }

  arts.forEach(function (A) {
    var weeks = A.weeks || [], dimNames = A.dim_names || [], dimDict = A.dim_dict || {}, R = A.rows || {};
    var wIdx = R.week_idx || [], sKey = R.series_key || [], dimsArr = R.dims || [],
        valArr = R.value || [], numArr = R.num || [], denArr = R.den || [];
    var n = wIdx.length;

    // Positional fallback: if the artifact's series_key names don't match the registry
    // (or the alias), map the artifact's own `series[]` to spec.series by order. Fixes
    // any naming drift for same-arity metrics without needing an exact alias.
    if (!isDimMetric && A.series && A.series.length === defs.length) {
      A.series.forEach(function (sk, i) { if (byKey[sk] === undefined) byKey[sk] = byKey[defs[i].key]; });
    }

    if (A.domains) Object.keys(A.domains).forEach(function (c) {
      var set = domains[c] || (domains[c] = {});
      (A.domains[c] || []).forEach(function (v) { set[_normDim(c, v)] = true; });
    });

    // active filters present in this artifact -> [{di, allowed:{idx:true}}]; absent = no-op.
    // The DICT value is grouped before it is tested, so a filter on 'Inbound' selects every raw
    // channel that is not one of the named sales channels — the complement rule, not a list.
    var fm = [];
    filters.active.forEach(function (f) {
      var di = dimNames.indexOf(f.col);
      if (di < 0) return;
      var dict = dimDict[f.col] || [], sel = {};
      f.set.forEach(function (v) { sel[v] = true; });
      var allowed = {};
      for (var i = 0; i < dict.length; i++) if (sel[_normDim(f.col, dict[i])]) allowed[i] = true;
      fm.push({ di: di, allowed: allowed });
    });
    var nf = fm.length;

    // breakdown column for dim metrics; artifacts unified 1b30d's 'channel' into
    // tour_acquisition_channel, so fall back to that if spec.dim.col is absent.
    var bIdx = -1;
    if (isDimMetric) {
      bIdx = dimNames.indexOf(spec.dim.col);
      if (bIdx < 0) bIdx = dimNames.indexOf('tour_acquisition_channel');
    }

    for (var r = 0; r < n; r++) {
      var drow = dimsArr[r], skip = false;
      for (var q = 0; q < nf; q++) { if (!fm[q].allowed[drow[fm[q].di]]) { skip = true; break; } }
      if (skip) continue;
      var wk = weeks[wIdx[r]];
      if (isDimMetric) {
        // Grouped HERE, not on the client: two raw channels that both mean Inbound must merge
        // into ONE member series. Relabelling them downstream would emit two rows called
        // 'Inbound' whose sum is the real figure.
        var mem = (bIdx >= 0 && dimDict[dimNames[bIdx]])
          ? String(_normDim(dimNames[bIdx], dimDict[dimNames[bIdx]][drow[bIdx]])) : 'All';
        var mw = memberWeekly[mem]; if (!mw) { mw = memberWeekly[mem] = {}; members[mem] = true; }
        if (spec.dim.kind === 'count') {
          var v = _artNum(valArr[r]);
          mw[wk] = (mw[wk] || 0) + v;
          totalWeekly[wk] = (totalWeekly[wk] || 0) + v;
        } else {
          var pm = mw[wk] || { num: 0, den: 0 }; pm.num += _artNum(numArr[r]); pm.den += _artNum(denArr[r]); mw[wk] = pm;
          var pt = totalWeekly[wk] || { num: 0, den: 0 }; pt.num += _artNum(numArr[r]); pt.den += _artNum(denArr[r]); totalWeekly[wk] = pt;
        }
      } else {
        var slot = byKey[sKey[r]]; if (!slot) continue;
        var a = slot.acc[wk]; if (!a) a = slot.acc[wk] = { sum: 0, n: 0, num: 0, den: 0 };
        if (slot.kind === 'wrate') { a.num += _artNum(numArr[r]); a.den += _artNum(denArr[r]); }
        else { var vv = valArr[r]; if (vv !== null && vv !== undefined) { a.sum += vv; a.n++; } }
      }
    }
  });

  var series = [];
  if (isDimMetric) {
    series.push({ key: 'total', label: 'Total', kind: dk, isDim: false, weekly: totalWeekly });
    Object.keys(members).sort().forEach(function (mem) {
      series.push({ key: 'dim__' + mem, label: mem, kind: dk, isDim: true, member: mem, weekly: memberWeekly[mem] });
    });
  } else {
    defs.forEach(function (d) {
      var slot = byKey[d.key], weekly = {};
      Object.keys(slot.acc).forEach(function (wk) {
        var a = slot.acc[wk];
        weekly[wk] = (d.kind === 'wrate') ? (a.den ? { num: a.num, den: a.den } : null)
                   : (d.kind === 'mean')  ? (a.n ? a.sum / a.n : null)
                   : a.sum;
      });
      series.push({ key: d.key, kind: d.kind, isDim: false, weekly: weekly });
    });
  }
  var domOut = {};
  Object.keys(domains).forEach(function (c) { domOut[c] = Object.keys(domains[c]).sort(); });
  return { series: series, domains: domOut };
}

function _isArr(x) { return Object.prototype.toString.call(x) === '[object Array]'; }

// Aggregate MARGINALS-layout artifact. Unfiltered -> __ALL__; single dimension -> that
// dim's block filtered to the selected value(s); two dimensions -> the named pair block
// if the export carries one; anything else -> unsupported (with a reason).
function _aggMarginals(A, spec, filters, metricKey) {
  var weeks = A.weeks || [], marg = A.marginals || {}, domOut = {};
  // Grouped, so the dropdown never offers a raw channel the filter path would then group away.
  if (A.domains) Object.keys(A.domains).forEach(function (c) {
    var seen = {};
    (A.domains[c] || []).forEach(function (v) { seen[_normDim(c, v)] = true; });
    domOut[c] = Object.keys(seen).sort();
  });

  // Arity is decided on the filters this metric actually carries — NOT on which of
  // them happen to have a marginals block. A filter on a column the metric has no
  // data for at all is a genuine no-op (this mirrors _aggFull, which skips dim
  // columns absent from the artifact). But a column the metric DOES carry and has
  // no block for is NOT a no-op: dropping it silently returns a lower-dimensional
  // total dressed up as the requested cut, and getSeries pins that for the full 6h
  // TTL. Deciding arity off `activeDims` was exactly that bug.
  var relevant = filters.active.filter(function (f) { return !!(A.domains && A.domains[f.col]); });
  var activeDims = relevant.filter(function (f) { return marg[f.col]; });
  function _sel(f) { var s = {}; f.set.forEach(function (v) { s[v] = true; }); return { col: f.col, set: s }; }
  var block, selDims = [];
  if (relevant.length === 0) {
    block = marg.__ALL__;
  } else if (relevant.length === 1 && activeDims.length === 1) {
    block = marg[activeDims[0].col];
    selDims = [_sel(activeDims[0])];
  } else if (relevant.length === 2) {
    // Named 2-dim PAIR blocks (DATABRICKS_side_spec.md §3f). The notebook exports a
    // small number of pairs it was asked for — not the cross-product, which is why the
    // artifact is marginals in the first place. Key is the two column names sorted
    // alphabetically and joined with '|', matching the exporter's _DIM_PAIRS.
    var pk = relevant.map(function (f) { return f.col; }).sort().join('|');
    block = marg[pk];
    if (block) selDims = relevant.map(_sel);
  }
  if (!block) {
    // Not a defect and not a live-query problem: this metric's artifact is exported
    // as __ALL__ + one block per SINGLE dimension (plus any named pairs) because its
    // full cross-product blew the ~17 MB / 5s Apps Script fetch budget. The rows for
    // this combination are simply not in the file. Fixed upstream by adding the pair
    // to the export, not by a new data source.
    if (relevant.length === 0) return { series: [], domains: domOut };
    return { series: [], domains: domOut, unsupported: true,
             reason: relevant.length > 1
               ? 'This metric cannot be cut by these ' + relevant.length + ' dimensions together.'
               : 'This metric cannot be cut by "' + relevant[0].col + '".' };
  }

  var alias = ART_SERIES_ALIAS[metricKey] || {};
  var defs = (spec.series || []).map(function (s) { return { key: s.key, kind: s.kind, art: alias[s.key] || s.key }; });
  if (!defs.length && A.series) defs = A.series.map(function (k) { return { key: k, kind: (A.kind && A.kind[k]) || 'sum', art: k }; });
  var byKey = {}; defs.forEach(function (d) { var slot = { kind: d.kind, key: d.key, acc: {} }; byKey[d.key] = slot; byKey[d.art] = slot; });

  var sKey = block.series_key || [], wIdx = block.week_idx || [],
      valArr = block.value || [], numArr = block.num || [], denArr = block.den || [];
  // A single-dim block carries value_dict/value_idx as flat arrays; a pair block keys
  // both by column name. Normalise to one list of {dict, idx, set} so the row loop is
  // identical for 0, 1 or 2 constrained dimensions.
  var _vd = block.value_dict, _vi = block.value_idx;
  var _keyed = _vd && !_isArr(_vd);
  var checks = selDims.map(function (d) {
    return { set: d.set, col: d.col,
             dict: (_keyed ? (_vd[d.col] || []) : (_vd || [])),
             idx:  (_keyed ? (_vi[d.col] || []) : (_vi || [])) };
  });
  var nChk = checks.length;
  var n = sKey.length;
  for (var r = 0; r < n; r++) {
    var skip = false;
    for (var q = 0; q < nChk; q++) {
      // Same grouping rule as _aggFull: test the GROUPED dict value, so 'Inbound' is the
      // complement of the named sales channels rather than a literal the artifact may not hold.
      if (!checks[q].set[_normDim(checks[q].col, checks[q].dict[checks[q].idx[r]])]) { skip = true; break; }
    }
    if (skip) continue;
    var slot = byKey[sKey[r]]; if (!slot) continue;
    var wk = weeks[wIdx[r]];
    var a = slot.acc[wk]; if (!a) a = slot.acc[wk] = { sum: 0, n: 0, num: 0, den: 0 };
    if (slot.kind === 'wrate') { a.num += _artNum(numArr[r]); a.den += _artNum(denArr[r]); }
    else { var x = valArr[r]; if (x !== null && x !== undefined) { a.sum += x; a.n++; } }
  }

  var series = defs.map(function (d) {
    var slot = byKey[d.key], weekly = {};
    Object.keys(slot.acc).forEach(function (wk) {
      var a = slot.acc[wk];
      weekly[wk] = (d.kind === 'wrate') ? (a.den ? { num: a.num, den: a.den } : null)
                 : (d.kind === 'mean')  ? (a.n ? a.sum / a.n : null)
                 : a.sum;
    });
    return { key: d.key, kind: d.kind, isDim: false, weekly: weekly };
  });
  return { series: series, domains: domOut };
}

// ── Background warming — pre-populate the unfiltered ('ALL') case ─────────────
// Reading all sheets in one execution would blow the 6-min limit, so warmCache
// is TIME-BOXED: it warms metrics starting from a stored cursor until the time
// budget is hit, then saves the cursor and stops. Install it on a short trigger
// (e.g. every 5 min via installWarmTrigger); over a few ticks the whole set is
// warm, and it keeps cycling to stay fresh. Metrics that are already cached are
// skipped instantly, so most ticks are cheap.
var WARM_BUDGET_MS = 240000;   // ~4 min — safe margin under the 6-min limit
function warmCache() {
  var keys = Object.keys(METRIC_SPECS);
  var props = PropertiesService.getScriptProperties();
  var cur = parseInt(props.getProperty('warm_cursor') || '0', 10) || 0;
  refreshShell();
  var t0 = new Date().getTime(), done = [], i = 0;
  for (; i < keys.length; i++) {
    if (new Date().getTime() - t0 > WARM_BUDGET_MS) break;   // out of budget → resume next tick
    var k = keys[(cur + i) % keys.length];
    try { getSeries(k, null); done.push(k); } catch (e) { Logger.log('warm ' + k + ' failed: ' + e); }
  }
  props.setProperty('warm_cursor', String((cur + i) % keys.length));
  Logger.log('warmCache: warmed ' + done.length + ' [' + done.join(', ') + ']');
  return done;
}

// Manual full warm — may exceed the 6-min limit on the full dataset (that's why
// warmCache is time-boxed for triggers). Kept for small datasets / local use.
function warmAll() {
  refreshShell();
  var t0 = new Date().getTime();
  Object.keys(METRIC_SPECS).forEach(function (k) {
    try { getSeries(k, null); } catch (e) { Logger.log('warmAll ' + k + ' failed: ' + e); }
  });
  Logger.log('warmAll done in ' + ((new Date().getTime() - t0) / 1000) + 's');
  return 'ok';
}

// One-time: install a 5-minute warmCache trigger (removes any prior warm trigger
// first). Time-boxed warmCache advances its cursor each tick until all warm.
function installWarmTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'warmAll' || t.getHandlerFunction() === 'warmCache') {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('warmCache').timeBased().everyMinutes(5).create();
  return 'warmCache trigger installed (every 5 min)';
}

// Cache-version salt — bumping it invalidates every cached series at once
// (we can't enumerate the per-signature keys). Used by the Refresh button.
function _salt() {
  return PropertiesService.getScriptProperties().getProperty('cache_salt') || '0';
}
// Refresh button target: invalidate all caches and return the fresh shell so the
// client can re-init. Series recompute lazily on the next render (fresh salt).
function refreshBootstrap() {
  var props = PropertiesService.getScriptProperties();
  var next = (parseInt(_salt(), 10) || 0) + 1;
  props.setProperty('cache_salt', String(next));
  return refreshShell();
}
// Same effect as the Refresh button, runnable from the editor. Note a BROWSER refresh
// does NOT do this — the series cache key embeds _salt(), so without a bump you keep
// being served pre-fix payloads for the full 6h TTL.
function bustCache() {
  var next = (parseInt(_salt(), 10) || 0) + 1;
  PropertiesService.getScriptProperties().setProperty('cache_salt', String(next));
  _cacheBust(CACHE_NS + ':shell');
  Logger.log('cache_salt -> ' + next + ', shell busted, build=' + BUILD);
  return next;
}

// ── Chunked CacheService helpers (values can exceed the 100 KB per-key cap) ──
function _cachePut(key, obj) {
  try {
    var cache = CacheService.getScriptCache();
    var str = JSON.stringify(obj);
    var size = 90000;
    var parts = Math.ceil(str.length / size) || 1;
    var map = {};
    map[key + '_meta'] = String(parts);
    for (var i = 0; i < parts; i++) map[key + '_' + i] = str.substr(i * size, size);
    cache.putAll(map, CACHE_TTL);
  } catch (err) { /* best-effort */ }
}
function _cacheGet(key) {
  try {
    var cache = CacheService.getScriptCache();
    var meta = cache.get(key + '_meta');
    if (!meta) return null;
    var parts = parseInt(meta, 10);
    var keys = [];
    for (var i = 0; i < parts; i++) keys.push(key + '_' + i);
    var got = cache.getAll(keys);
    var str = '';
    for (var j = 0; j < parts; j++) { var p = got[key + '_' + j]; if (p == null) return null; str += p; }
    return JSON.parse(str);
  } catch (err) { return null; }
}
function _cacheBust(key) {
  try { CacheService.getScriptCache().remove(key + '_meta'); } catch (e) {}
}

// ── Diagnostics (run from the editor) ─────────────────────────────────────────
function diagnoseSeries() {
  var k = 'gmv_yoy';
  var spec = METRIC_SPECS[k];
  var t0 = new Date().getTime();
  var raw = _readMetricSheet(spec);
  var tRead = new Date().getTime() - t0;
  var t1 = new Date().getTime();
  var out = _aggregate(spec, raw.headers, raw.rows, _normFilters(null));
  var tAgg = new Date().getTime() - t1;
  Logger.log(k + ': ' + raw.rows.length + ' rows -> ' + out.series.length + ' series, ' +
    Object.keys(out.series[0].weekly).length + ' weeks');
  Logger.log('READ ' + tRead + 'ms   AGG ' + tAgg + 'ms   TOTAL ' + (tRead + tAgg) + 'ms');
  Logger.log('domains: ' + JSON.stringify(Object.keys(out.domains).map(function (c) {
    return c + '(' + out.domains[c].length + ')';
  })));
}

// Snapshot-mode smoke test (run from the editor). Fetches artifacts from Drive,
// aggregates, and logs timing + the latest weekly value so you can eyeball
// correctness against the gold table before flipping the dashboard over.
function diagnoseSnapshot() {
  var bad = [];
  Object.keys(METRIC_SPECS).forEach(function (k) {
    var t0 = new Date().getTime();
    try {
      var out = _snapshotSeries(k, METRIC_SPECS[k], _normFilters(null));
      var ms = new Date().getTime() - t0, s0 = out.series[0];
      var wks = s0 ? Object.keys(s0.weekly || {}) : [];
      var flag = (!out.series.length || !wks.length) ? '  <-- EMPTY' : '';
      if (flag) bad.push(k);
      Logger.log(k + ': ' + ms + 'ms | series=' + out.series.length + ' weeks=' + wks.length +
        (out.unsupported ? ' UNSUPPORTED' : '') + ' | domains=' + Object.keys(out.domains).length + flag);
    } catch (e) {
      bad.push(k + '(ERR)');
      Logger.log(k + ': ERROR ' + e.message);
    }
  });
  Logger.log(bad.length ? ('>>> PROBLEM METRICS: ' + bad.join(', ')) : '>>> ALL ' + Object.keys(METRIC_SPECS).length + ' METRICS OK');
}
// Dump a metric's artifact series_key names vs the registry keys + resolved weekly
// counts — reveals any naming drift the positional fallback had to bridge.
function diagnoseMetric(metricKey) {
  var spec = METRIC_SPECS[metricKey];
  var man = _manifest(), entry = man.metrics && man.metrics[metricKey];
  if (!entry) { Logger.log(metricKey + ': not in manifest'); return; }
  var arts = entry.files.map(_fetchArtifact);
  Logger.log(metricKey + ' artifact layout=' + (arts[0].layout || 'full') +
    ' | artifact series=' + JSON.stringify(arts[0].series) +
    ' | registry series=' + JSON.stringify((spec.series || []).map(function (s) { return s.key; })) +
    ' | alias=' + JSON.stringify(ART_SERIES_ALIAS[metricKey] || {}));
  var out = _snapshotSeries(metricKey, spec, _normFilters(null));
  out.series.forEach(function (s) {
    var wks = Object.keys(s.weekly || {});
    Logger.log('  -> series ' + s.key + ' (' + s.kind + '): ' + wks.length + ' weeks' +
      (wks.length ? ', latest ' + wks.sort().slice(-1)[0] + '=' + JSON.stringify(s.weekly[wks.sort().slice(-1)[0]]) : ''));
  });
}
// Why is QTD wrong? Run this. For each metric it reports whether a _qtd artifact was
// found, whether it loaded, and whether qtd_weekly actually landed on the series —
// plus the quarter-split buckets it produced, which is what QTD sums over.
// Call with no arg to sweep every metric, or diagnoseQtd('gmv_yoy') for one.
// What the DASHBOARD will actually render. Goes through getSeries, so it hits the same
// cache the client does — passing null gives filter sig 'ALL', the identical cache key
// the browser's unfiltered call uses. diagnoseQtd below calls _snapshotSeries directly
// and so reports OK off fresh artifact reads even while the browser is served a stale
// payload; that gap is what hid this bug. Use this one to confirm a fix landed.
function verifyQtd(metricKey) {
  var mk = metricKey || 'gmv_yoy';
  var out = getSeries(mk, null);
  var now = new Date();
  var qs = new Date(Date.UTC(now.getUTCFullYear(), Math.floor(now.getUTCMonth() / 3) * 3, 1))
             .toISOString().slice(0, 10);
  Logger.log('build=' + BUILD + ' salt=' + _salt() + ' src=' + _dataSource() +
             ' | ' + mk + ' | quarter from ' + qs);
  (out.series || []).forEach(function (s) {
    var wm = s.qtd_weekly || s.weekly || {};
    var sum = 0, num = 0, den = 0, n = 0;
    Object.keys(wm).forEach(function (w) {
      if (w < qs) return;
      var v = wm[w]; if (v === null || v === undefined) return;
      n++;
      if (typeof v === 'object') { num += (v.num || 0); den += (v.den || 0); }
      else sum += Number(v);
    });
    Logger.log('  ' + s.key + ': ' + (s.qtd_weekly ? 'QTD-grain OK' : 'WEEKLY-FALLBACK (wrong)') +
      ' | ' + n + ' buckets | QTD=' + (den ? (num / den * 100).toFixed(2) + '%' : Math.round(sum)));
  });
}
function diagnoseQtd(only) {
  var folderIdx = _folderIndex();
  var keys = only ? [only] : Object.keys(METRIC_SPECS);
  keys.forEach(function (mk) {
    var spec = METRIC_SPECS[mk];
    if (!spec) { Logger.log(mk + ': not in METRIC_SPECS'); return; }
    var qtdKey = mk + '_qtd';
    var files = _qtdFiles(mk);
    // Flag stale shards the pipeline left behind — they are dead weight, not data.
    var ignored = Object.keys(folderIdx).filter(function (f) {
      return f.indexOf(qtdKey) === 0 && /\.json\.gz$/.test(f) && files.indexOf(f) === -1;
    }).sort();
    if (ignored.length) Logger.log('STALE ' + mk + ': delete from Drive -> ' + ignored.join(', '));
    if (!files.length) {
      Logger.log('MISS ' + mk + ': no artifact matching "' + qtdKey + '*.json.gz"');
      return;
    }
    var loaded = null, err = null;
    try { loaded = files.map(_fetchArtifact); } catch (e) { err = e; }
    if (err) { Logger.log('FAIL ' + mk + ' [' + files.join(', ') + ']: fetch/parse threw: ' + err); return; }
    var out;
    try { out = _snapshotSeries(mk, spec, _normFilters(null)); }
    catch (e) { Logger.log('FAIL ' + mk + ': _snapshotSeries threw: ' + e); return; }
    var withQtd = out.series.filter(function (s) { return !!s.qtd_weekly; });
    Logger.log((withQtd.length ? 'OK   ' : 'MISS ') + mk +
      ' [' + files.join(', ') + '] layout=' + (loaded[0].layout || 'full') +
      ' | ' + withQtd.length + '/' + out.series.length + ' series got qtd_weekly');
    out.series.forEach(function (s) {
      if (!s.qtd_weekly) { Logger.log('     ' + s.key + ': NO qtd_weekly (QTD will use s.weekly)'); return; }
      // Non-Monday keys are the quarter-boundary splits — the whole point of this grain.
      var splits = Object.keys(s.qtd_weekly).filter(function (w) {
        var d = new Date(w + 'T00:00:00Z'); return d.getUTCDay() !== 1;
      }).sort();
      Logger.log('     ' + s.key + ': ' + Object.keys(s.qtd_weekly).length + ' qtd weeks, ' +
        splits.length + ' boundary splits' + (splits.length ? ' (' + splits.slice(-4).join(', ') + ')' : ''));
    });
  });
}
// Trash orphaned `<metric>_qtd_part*.json.gz` shards left behind when a metric moved from
// the `full` to the `marginals` layout: the single file supersedes them, but the shards keep
// matching the prefix, so any folder-index miss on the single file drags tens of MB into
// Utilities.ungzip and throws. DRY RUN by default — call cleanStaleQtdArtifacts(true) to
// actually trash. Only ever touches metrics that HAVE a single-file artifact, so legitimate
// multi-part sets (net_revenue, bookings) are never candidates.
function cleanStaleQtdArtifacts(reallyDelete) {
  var idx = _folderIndex(true), n = 0;
  Object.keys(METRIC_SPECS).forEach(function (mk) {
    var qtdKey = mk + '_qtd';
    if (!idx[qtdKey + '.json.gz']) return;              // no single file -> parts are the real artifact
    Object.keys(idx).filter(function (f) {
      return f.indexOf(qtdKey + '_part') === 0 && /\.json\.gz$/.test(f);
    }).sort().forEach(function (f) {
      n++;
      if (reallyDelete) { DriveApp.getFileById(idx[f]).setTrashed(true); Logger.log('TRASHED ' + f); }
      else Logger.log('WOULD TRASH ' + f + '  (superseded by ' + qtdKey + '.json.gz)');
    });
  });
  Logger.log(n ? (reallyDelete ? n + ' file(s) trashed. Click Refresh to re-list the folder.'
                              : n + ' stale file(s). Re-run as cleanStaleQtdArtifacts(true) to trash them.')
              : 'No stale QTD shards.');
}
function countRows() {
  Object.keys(METRIC_SPECS).forEach(function (k) {
    var spec = METRIC_SPECS[k];
    var sh = _ssFor(spec.ss).getSheetByName(spec.sheet);
    Logger.log(k + ' [' + spec.ss + '/' + spec.sheet + '] : ' + (sh ? sh.getLastRow() : 'MISSING') + ' rows');
  });
}

function checkDriveAccess() {
  try {
    DriveApp.getFolderById(ARTIFACT_FOLDER_ID);
    return { hasAccess: true };
  } catch (e) {
    return { hasAccess: false, folderUrl: 'https://drive.google.com/drive/folders/' + ARTIFACT_FOLDER_ID };
  }
}

// Comment backend is in comment_swbr_trading.gs
