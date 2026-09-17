/**
 * Snapshot history — weekly frozen copies of what the dashboard served, plus an
 * append-only log of every restatement.
 *
 * WHY: the nightly Databricks job overwrites <metric>.json.gz in place and rewrites the
 * 4 data spreadsheets, so nothing anywhere records what the dashboard said last week.
 * When a PAST week's number silently moves (the curation_duration COALESCE partial
 * rebuild, the 1b30d numerator fix in v38) there is no before-picture to diff against.
 * DATABRICKS_side_spec.md §3c specified "keep the last ~14 versions" — never implemented.
 *
 * WHAT: snapshotHistory() runs weekly, calls getSeries(metric, null) for all 25 metrics
 * (the unfiltered case warmCache already keeps warm, so most calls are cache hits),
 * and writes ONE small gzipped file per week to a history/ subfolder of
 * ARTIFACT_FOLDER_ID. It then diffs against the previous snapshot and APPENDS every
 * change to the Snapshot_Diff tab.
 *
 * The log is append-only on purpose. A compare-on-request report is exactly what fails
 * when nobody thinks to look for three weeks: a restatement from any past week has to
 * still be findable, as a timestamped row, months later. Same reason the email fires
 * itself rather than waiting to be asked.
 *
 * The history/ SUBFOLDER matters: _folderIndex() lists direct children only, so these
 * files are invisible to _manifest / _qtdFiles / the _qtd prefix scan. Never write them
 * into the artifact folder root.
 *
 * Grain = whatever getSeries returns unfiltered: each series' weekly map plus the
 * metric's designated spec.dim member breakdown. Capturing all 7 FILTER_COLS would be
 * ~175 RPCs per run (past the 6-min limit) and the 5 marginals megas cannot supply it at
 * all — full dim grain belongs in the Databricks follow-up.
 *
 * Entry points (all editor-run except the trigger):
 *   setupSnapshotSheet()                  — once, creates the Snapshot_Diff tab
 *   installSnapshotTrigger()              — once, weekly Monday ~09:00
 *   snapshotHistory()                     — write today's snapshot + diff vs previous
 *   diffSnapshots(dateA, dateB[, tolPct]) — compare ANY two snapshots, no write
 *   weekDrift(metric, series, weekStart)  — one week's value across every snapshot ever
 *   listSnapshots()                       — the asof dates on file
 */

var HIST_SUBFOLDER = 'history';
// Escape hatch: creating the subfolder needs Content Manager on the artifact Shared Drive. If the
// account running the trigger only has read there, put a folder id here instead — a permissions
// wall must not be the thing that stops history from ever starting.
var HIST_FOLDER_ID = '';
var SNAP_SHEET     = 'Snapshot_Diff';
// 'week_start' is load-bearing, not cosmetic: _buildShell inlines every non-metric tab
// into BOOT.aux unless it is in CMT_TABS *or* carries a week_start header. The tab is in
// CMT_TABS; this header is the belt to that braces, so a future CMT_TABS edit cannot
// silently start shipping the whole log to every viewer.
var SNAP_HEADERS   = ['logged_at','asof_new','asof_old','metric','series','member',
                      'week_start','old','new','delta','delta_pct','flag'];
var SNAP_FLAG_COL  = 11;      // 0-based index of 'flag' in SNAP_HEADERS
var SNAP_TOL_PCT   = 0.5;     // relative move that counts as a restatement
var SNAP_TOL_ABS   = 1e-9;    // float-noise floor (accumulated means never land exact)
var SNAP_MAX_ROWS  = 5000;    // per run — one bad pipeline day must not fill the tab

// ── Drive plumbing ───────────────────────────────────────────────────────────
function _histFolder() {
  if (HIST_FOLDER_ID) return DriveApp.getFolderById(HIST_FOLDER_ID);
  var root = DriveApp.getFolderById(ARTIFACT_FOLDER_ID);
  var it = root.getFoldersByName(HIST_SUBFOLDER);
  return it.hasNext() ? it.next() : root.createFolder(HIST_SUBFOLDER);
}
function _snapName(asof) { return 'asof_' + asof + '.json.gz'; }

// The asof dates on file, oldest first.
function listSnapshots() {
  var out = [], files = _histFolder().getFiles();
  while (files.hasNext()) {
    var m = /^asof_(\d{4}-\d{2}-\d{2})\.json\.gz$/.exec(files.next().getName());
    if (m) out.push(m[1]);
  }
  out.sort();
  Logger.log('snapshots on file (' + out.length + '): ' + out.join(', '));
  return out;
}

function _readSnapshot(asof) {
  var it = _histFolder().getFilesByName(_snapName(asof));
  if (!it.hasNext()) throw new Error('no snapshot for ' + asof + ' (have: ' + listSnapshots().join(', ') + ')');
  var blob = it.next().getBlob().setContentType('application/x-gzip');
  return JSON.parse(Utilities.ungzip(blob).getDataAsString());
}

// ── The weekly writer ────────────────────────────────────────────────────────
function snapshotHistory() {
  var t0 = new Date().getTime();
  var asof = Utilities.formatDate(new Date(), _tz(), 'yyyy-MM-dd');
  var dv = '';
  try { dv = _manifest().data_version || ''; } catch (e) { Logger.log('manifest unreadable: ' + e); }

  var snap = { asof: asof, generated_at: new Date().toISOString(), data_version: dv,
               build: BUILD, data_source: _dataSource(), salt: _salt(),
               failed: [], metrics: {} };

  Object.keys(METRIC_SPECS).forEach(function (k) {
    try {
      var out = getSeries(k, null);
      snap.metrics[k] = {
        series: (out.series || []).map(function (s) {
          // Explicit field list, same discipline as enrichSeries: qtd_weekly is a second
          // grain that restates independently of the Monday grain, so it is captured too.
          return { key: s.key, kind: s.kind, isDim: !!s.isDim, member: s.member || '',
                   weekly: s.weekly || {}, qtd_weekly: s.qtd_weekly || null };
        }),
        domains: out.domains || {}
      };
    } catch (e) {
      // A partial snapshot that SAYS it is partial is fine; a silently short one is not.
      snap.failed.push(k + ': ' + e);
      Logger.log('snapshot ' + k + ' failed: ' + e);
    }
  });

  var folder = _histFolder(), name = _snapName(asof);
  var stale = folder.getFilesByName(name);            // re-run on the same day replaces
  while (stale.hasNext()) stale.next().setTrashed(true);
  var file = folder.createFile(
    Utilities.gzip(Utilities.newBlob(JSON.stringify(snap), 'application/json'), name).setName(name));

  var secs = ((new Date().getTime() - t0) / 1000).toFixed(1);
  Logger.log('snapshot ' + asof + ': ' + Object.keys(snap.metrics).length + ' metrics, ' +
             snap.failed.length + ' failed, ' + Math.round(file.getSize() / 1024) + ' KB, ' +
             secs + 's, data_version=' + (dv || '(none)'));

  var prev = listSnapshots().filter(function (d) { return d < asof; }).pop();
  if (!prev) {
    Logger.log('no previous snapshot — nothing to diff. History starts here.');
    return { asof: asof, prev: null, rows: 0 };
  }

  var prevSnap = _readSnapshot(prev);
  var rows = _diffObjects(prevSnap, snap, SNAP_TOL_PCT);
  var n = _snapAppend(rows);
  // Two snapshots sharing a data_version came from the SAME pipeline run, so any RESTATED
  // row between them is a cache/aggregation bug on our side, not an upstream restatement.
  var same = prevSnap.data_version && prevSnap.data_version === snap.data_version;
  if (same) Logger.log('NOTE: ' + prev + ' and ' + asof + ' share data_version ' + snap.data_version +
                       ' — differences here are OUR bug, not a pipeline restatement.');
  Logger.log('diff ' + prev + ' -> ' + asof + ': ' + n + ' rows logged (' + _snapCounts(rows) + ')');
  _snapNotify(asof, prev, rows, same);
  return { asof: asof, prev: prev, rows: n, counts: _snapCounts(rows) };
}

// ── Diff ─────────────────────────────────────────────────────────────────────
// Public: compare ANY two snapshots on file. Logs a summary, returns the rows, writes
// nothing — so "compare today against three months ago" is one editor call.
function diffSnapshots(dateA, dateB, tolPct) {
  var a = _readSnapshot(dateA), b = _readSnapshot(dateB);
  var rows = _diffObjects(a, b, tolPct == null ? SNAP_TOL_PCT : tolPct);
  Logger.log('diff ' + dateA + ' -> ' + dateB + ': ' + rows.length + ' rows (' + _snapCounts(rows) + ')');
  _snapTop(rows, 20).forEach(function (r) {
    Logger.log('  ' + r[SNAP_FLAG_COL] + '  ' + r[3] + '.' + r[4] + (r[5] ? ' [' + r[5] + ']' : '') +
               '  ' + r[6] + '  ' + r[7] + ' -> ' + r[8] +
               (r[10] === '' ? '' : '  (' + r[10] + '%)'));
  });
  return rows;
}

// A weekly map value is either a number or {num,den} (wrate). null when absent.
function _snapVal(w) {
  if (w === null || w === undefined) return null;
  if (typeof w === 'object') return w.den ? (Number(w.num) || 0) / Number(w.den) : 0;
  return Number(w) || 0;
}
function _snapMoved(a, b, tolPct) {
  if (a === b) return false;
  var d = Math.abs(b - a);
  if (d <= SNAP_TOL_ABS) return false;
  var base = Math.max(Math.abs(a), Math.abs(b));
  if (!base) return true;                         // 0 -> non-zero is always a change
  return (d / base) * 100 > tolPct;
}

// Pure over two snapshot objects — this is the unit the harness tests. Row shape is
// SNAP_HEADERS; flags are RESTATED / NEW_WEEK / DISAPPEARED / DOMAIN_CHANGE.
function _diffObjects(oldSnap, newSnap, tolPct) {
  if (tolPct == null) tolPct = SNAP_TOL_PCT;
  var stamp = new Date().toISOString();
  var aAs = oldSnap.asof || '', bAs = newSnap.asof || '';
  var rows = [];
  function push(metric, series, member, week, o, n, flag) {
    var num = (typeof o === 'number' && typeof n === 'number');
    var d = num ? n - o : '';
    var base = num ? Math.max(Math.abs(o), Math.abs(n)) : 0;
    rows.push([stamp, bAs, aAs, metric, series, member || '', week || '',
               o === null ? '' : o, n === null ? '' : n,
               d === '' ? '' : d,
               (num && base) ? Number(((d / base) * 100).toFixed(3)) : '',
               flag]);
  }

  var oM = oldSnap.metrics || {}, nM = newSnap.metrics || {};
  var metrics = Object.keys(oM);
  Object.keys(nM).forEach(function (k) { if (metrics.indexOf(k) < 0) metrics.push(k); });

  metrics.forEach(function (mk) {
    var o = oM[mk], n = nM[mk];
    // A whole metric vanishing is one loud row, not a wall of per-week rows. Note the
    // asymmetry: a metric APPEARING is not worth logging (it is how a new metric ships).
    if (o && !n) { push(mk, '(metric)', '', '', 'present', 'absent', 'DISAPPEARED'); return; }
    if (!o || !n) return;

    var oS = {}, nS = {};
    (o.series || []).forEach(function (s) { oS[s.key] = s; });
    (n.series || []).forEach(function (s) { nS[s.key] = s; });

    Object.keys(oS).forEach(function (sk) {
      if (!nS[sk]) push(mk, sk, oS[sk].member || '', '', 'present', 'absent', 'DISAPPEARED');
    });

    Object.keys(nS).forEach(function (sk) {
      var os = oS[sk], ns = nS[sk];
      if (!os) return;                              // new series — see the note above
      _snapDiffGrain(push, mk, sk, ns.member, os.weekly, ns.weekly, ns.kind, '', tolPct);
      // The QTD grain splits the quarter-straddling week, so it can be restated on its own.
      if (os.qtd_weekly || ns.qtd_weekly) {
        _snapDiffGrain(push, mk, sk, ns.member, os.qtd_weekly || {}, ns.qtd_weekly || {},
                       ns.kind, '@qtd', tolPct);
      }
    });

    // A dimension value appearing or vanishing is itself a bug signal (a renamed channel,
    // a dropped join) and costs nothing to check.
    var oD = o.domains || {}, nD = n.domains || {};
    var cols = Object.keys(oD);
    Object.keys(nD).forEach(function (c) { if (cols.indexOf(c) < 0) cols.push(c); });
    cols.forEach(function (c) {
      var ov = oD[c] || [], nv = nD[c] || [];
      ov.forEach(function (v) { if (nv.indexOf(v) < 0) push(mk, c, String(v), '', 'present', 'absent', 'DOMAIN_CHANGE'); });
      nv.forEach(function (v) { if (ov.indexOf(v) < 0) push(mk, c, String(v), '', 'absent', 'present', 'DOMAIN_CHANGE'); });
    });
  });

  return rows;
}

// One grain of one series. For a wrate the NUM and DEN are compared separately as well as
// the ratio: a compensating num+den move leaves the rate identical, which is precisely the
// silent restatement this whole file exists to catch.
function _snapDiffGrain(push, mk, sk, member, ow, nw, kind, suffix, tolPct) {
  ow = ow || {}; nw = nw || {};
  var weeks = Object.keys(ow);
  Object.keys(nw).forEach(function (w) { if (weeks.indexOf(w) < 0) weeks.push(w); });
  weeks.sort();

  weeks.forEach(function (w) {
    var a = ow[w], b = nw[w];
    var name = sk + suffix;
    if (a !== undefined && b === undefined) { push(mk, name, member, w, _snapVal(a), null, 'DISAPPEARED'); return; }
    if (a === undefined) { push(mk, name, member, w, null, _snapVal(b), 'NEW_WEEK'); return; }

    if (kind === 'wrate' && typeof a === 'object' && typeof b === 'object') {
      var an = Number(a.num) || 0, ad = Number(a.den) || 0;
      var bn = Number(b.num) || 0, bd = Number(b.den) || 0;
      if (_snapMoved(an, bn, tolPct)) push(mk, name + '.num', member, w, an, bn, 'RESTATED');
      if (_snapMoved(ad, bd, tolPct)) push(mk, name + '.den', member, w, ad, bd, 'RESTATED');
    }
    var av = _snapVal(a), bv = _snapVal(b);
    if (_snapMoved(av, bv, tolPct)) push(mk, name, member, w, av, bv, 'RESTATED');
  });
}

// ── Sheet + notification ─────────────────────────────────────────────────────
function setupSnapshotSheet() {
  var ss = _supplySS();
  var sh = ss.getSheetByName(SNAP_SHEET) || ss.insertSheet(SNAP_SHEET);
  if (sh.getLastRow() < 1) {
    sh.getRange(1, 1, 1, SNAP_HEADERS.length).setValues([SNAP_HEADERS]);
    sh.setFrozenRows(1);
  }
  Logger.log('setupSnapshotSheet: done — ' + ss.getName() + ' ' + ss.getUrl());
  return 'ok';
}

function _snapCounts(rows) {
  var c = {};
  rows.forEach(function (r) { var f = r[SNAP_FLAG_COL]; c[f] = (c[f] || 0) + 1; });
  return Object.keys(c).sort().map(function (k) { return k + ':' + c[k]; }).join(' ') || 'no changes';
}
// Signal only, worst first. NEW_WEEK is dropped: every series gets one every single week
// (~300 rows) and "a new week arrived" is never the bug — it would bury the rows that are.
// It still shows in _snapCounts, so the summary stays honest. A vanished level is louder
// than any delta, so flag priority outranks magnitude.
var SNAP_FLAG_RANK = { DISAPPEARED: 0, DOMAIN_CHANGE: 1, RESTATED: 2, NEW_WEEK: 9 };
function _snapTop(rows, n) {
  return rows.filter(function (r) { return r[SNAP_FLAG_COL] !== 'NEW_WEEK'; })
             .sort(function (a, b) {
               var ra = SNAP_FLAG_RANK[a[SNAP_FLAG_COL]], rb = SNAP_FLAG_RANK[b[SNAP_FLAG_COL]];
               if (ra !== rb) return ra - rb;
               return Math.abs(Number(b[9]) || 0) - Math.abs(Number(a[9]) || 0);
             })
             .slice(0, n);
}

function _snapAppend(rows) {
  var sh = _supplySS().getSheetByName(SNAP_SHEET);
  if (!sh) throw new Error('no "' + SNAP_SHEET + '" tab — run setupSnapshotSheet() once');
  var out = _snapTop(rows, SNAP_MAX_ROWS);
  if (!out.length) return 0;
  var kept = _snapTop(rows, Infinity).length;
  if (kept > SNAP_MAX_ROWS) {
    // Never truncate silently — a capped run that looks complete reads as "all clear".
    out = out.slice(0, SNAP_MAX_ROWS - 1);
    out.push([new Date().toISOString(), rows[0][1], rows[0][2], '(truncated)', '', '', '', '', '',
              kept - out.length, '', 'TRUNCATED']);
  }
  sh.getRange(sh.getLastRow() + 1, 1, out.length, SNAP_HEADERS.length).setValues(out);
  return out.length;
}

function _snapNotify(asof, prev, rows, sameDataVersion) {
  var bad = rows.filter(function (r) {
    var f = r[SNAP_FLAG_COL];
    return f === 'RESTATED' || f === 'DISAPPEARED' || f === 'DOMAIN_CHANGE';
  });
  if (!bad.length) return;
  var to = _adminEmails().join(',');
  if (!to) { Logger.log(bad.length + ' restatements but ADMIN_EMAILS is unset — no email sent'); return; }
  var url = '';
  try { url = _supplySS().getUrl() + '#gid=' + _supplySS().getSheetByName(SNAP_SHEET).getSheetId(); } catch (e) {}
  var body = 'Supply WBR data changed between the ' + prev + ' and ' + asof + ' snapshots.\n\n' +
    _snapCounts(rows) + '\n\n' +
    (sameDataVersion ? 'NOTE: both snapshots came from the SAME pipeline run (data_version ' +
      'unchanged), so these differences are a dashboard-side bug, not an upstream restatement.\n\n' : '') +
    'Biggest movers:\n' +
    _snapTop(bad, 10).map(function (r) {
      return '  ' + r[SNAP_FLAG_COL] + '  ' + r[3] + '.' + r[4] + (r[5] ? ' [' + r[5] + ']' : '') +
             '  ' + r[6] + '  ' + r[7] + ' -> ' + r[8] + (r[10] === '' ? '' : '  (' + r[10] + '%)');
    }).join('\n') +
    '\n\nFull log: ' + url;
  GmailApp.sendEmail(to, '[Supply WBR] ' + bad.length + ' data changes since ' + prev, body);
  Logger.log('notified ' + to);
}

// ── One week, every snapshot ──────────────────────────────────────────────────
// "When did CW28 change, and how many times?" — the question a single pairwise diff
// cannot answer. Pass the series key as it appears in the log (e.g. 'total', 'dim__In House').
function weekDrift(metricKey, seriesKey, weekStart) {
  var out = [], prev = null;
  listSnapshots().forEach(function (d) {
    var snap;
    try { snap = _readSnapshot(d); } catch (e) { Logger.log('  ' + d + '  unreadable: ' + e); return; }
    var m = snap.metrics && snap.metrics[metricKey];
    var s = m && (m.series || []).filter(function (x) { return x.key === seriesKey; })[0];
    var v = s ? _snapVal(s.weekly[weekStart]) : null;
    out.push({ asof: d, value: v, delta: (prev === null || v === null) ? '' : v - prev,
               data_version: snap.data_version || '' });
    if (v !== null) prev = v;
  });
  Logger.log(metricKey + '.' + seriesKey + ' @ ' + weekStart);
  out.forEach(function (r) {
    Logger.log('  ' + r.asof + '  ' + (r.value === null ? '(absent)' : r.value) +
               (r.delta === '' || r.delta === 0 ? '' : '   Δ ' + r.delta) + '   ' + r.data_version);
  });
  return out;
}

// ── Trigger ───────────────────────────────────────────────────────────────────
// Monday 09:00, after the nightly pipeline. Apps Script honours the hour ±15 min.
function installSnapshotTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'snapshotHistory') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('snapshotHistory').timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(9).create();
  return 'snapshotHistory trigger installed (Mondays ~09:00)';
}
