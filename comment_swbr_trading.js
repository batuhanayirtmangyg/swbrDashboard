/* ===========================================================================
   COMMENTING SYSTEM — backend module for the GYG SWBR Dashboard
   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
   Add this as a NEW script file in the Apps Script project:
     Apps Script editor > File > New script > name it "CommentBackend"
     Paste this entire file in.  Do not merge into the existing Code.gs.

   All private helpers are prefixed _cmt so they never clash with existing
   functions (the project already has _ensureSheet, _logEvent, etc.).

   ONE-TIME SETUP (run from the Apps Script editor after pasting):
     1. Run setupCommentDb()      — creates Comments / Replies / Users tabs
     2. Redeploy the web app as "Execute as: User accessing the web app"
        (Deploy > Manage deployments > Edit > change Execute as)
        This is required for Session.getActiveUser() to return each viewer's
        real email rather than the script owner's.  Each viewer will be asked
        to authorise spreadsheets + gmail.send on their first open.

   FRONTEND WIRING (after adding comments.html — sent separately):
     - Add  <?!= include('comments'); ?>  just before </body> in index.html
     - On renderable elements add  data-comment-target="metric:gmv_cw"  etc.
     - At the end of your main render function add:
         if (window.afterDashboardRender) window.afterDashboardRender();
   =========================================================================== */

// ── Constants ────────────────────────────────────────────────────────────────

const CMT_SHEET     = 'Comments';
const REPLY_SHEET   = 'Replies';
const USER_SHEET    = 'Users';

const CMT_HEADERS   = ['CommentID','AuthorEmail','TargetElementID','Timestamp',
                        'TextContent','IsResolved','ViewKey','AnchorQuote'];
const REPLY_HEADERS = ['ReplyID','ParentCommentID','AuthorEmail','Timestamp','TextContent'];
const USER_HEADERS  = ['Name','Email'];

const THREADS_CACHE_KEY = 'gyg_td_threads_v1';
const THREADS_CACHE_TTL = 30;    // seconds — short TTL, busted on every write
const MAX_CMT_LEN       = 4000;

// ── One-time setup ───────────────────────────────────────────────────────────

// Creates Comments, Replies, and Users tabs with the correct headers.
// Idempotent — safe to run more than once.
function setupCommentDb() {
  const ss = _supplySS();
  _cmtEnsureSheet(ss, CMT_SHEET,   CMT_HEADERS);
  _cmtEnsureSheet(ss, REPLY_SHEET, REPLY_HEADERS);
  const u = _cmtEnsureSheet(ss, USER_SHEET, USER_HEADERS);
  if (u.getLastRow() < 2) {
    u.appendRow(['Me', Session.getActiveUser().getEmail() || '']);
  }
  Logger.log('setupCommentDb: done');
  return 'ok';
}

// ── Session ──────────────────────────────────────────────────────────────────

function getSessionUser() {
  return Session.getActiveUser().getEmail() || '';
}

// ── Auto-registration ────────────────────────────────────────────────────────
// Called from comments.html on every page open.  Returns { email, name, isNew }.
// If the viewer is not in the Users sheet yet, derives their display name from
// the email local-part ("john.smith" -> "John Smith"), appends them, and calls
// warmCache() so colleagues can immediately @mention the new user.

function ensureUserRegistered() {
  const me = getSessionUser();
  if (!me) return { email: '', name: '', isNew: false };

  const ss = _supplySS();
  let sh = ss.getSheetByName(USER_SHEET);
  if (!sh) { setupCommentDb(); sh = ss.getSheetByName(USER_SHEET); }

  const users = _cmtRowsToObjects(sh);
  const found = users.find(u => String(u.Email).toLowerCase() === me.toLowerCase());
  if (found) return { email: me, name: found.Name || me, isNew: false };

  const name = me.split('@')[0]
    .split('.')
    .map(p => p.charAt(0).toUpperCase() + p.slice(1))
    .join(' ');

  _cmtAppendRow(USER_SHEET, USER_HEADERS, { Name: name, Email: me });
  _bustCmtCache();
  try { warmCache(); } catch (e_) { /* non-fatal if warmCache isn't installed yet */ }
  return { email: me, name, isNew: true };
}

// Add a colleague who hasn't visited yet to the Users directory.
// Called from the @mention "Add to team" button.
// Returns { email, name, alreadyExists }.

function addUserToDirectory(payload) {
  payload = payload || {};
  const email = String(payload.email || '').trim();
  if (!email || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email))
    throw new Error('Invalid email address.');

  const name = String(payload.name || '').trim() ||
    email.split('@')[0].split('.').map(p => p.charAt(0).toUpperCase() + p.slice(1)).join(' ');

  const ss = _supplySS();
  const sh = ss.getSheetByName(USER_SHEET) || _cmtEnsureSheet(ss, USER_SHEET, USER_HEADERS);
  const exists = _cmtRowsToObjects(sh)
    .some(u => String(u.Email).toLowerCase() === email.toLowerCase());
  if (exists) return { email, name, alreadyExists: true };

  _cmtAppendRow(USER_SHEET, USER_HEADERS, { Name: name, Email: email });
  _bustCmtCache();
  return { email, name, alreadyExists: false };
}

// ── Read ─────────────────────────────────────────────────────────────────────
// Lightweight RPC fetched after first paint — NOT part of the heavy data
// bootstrap so the dashboard's fast load is unaffected.
// Returns { me, users, comments, replies }.  Cached briefly; busted on writes.

function getThreads() {
  const cache  = CacheService.getScriptCache();
  const cached = cache.get(THREADS_CACHE_KEY);
  if (cached) {
    const payload = JSON.parse(cached);
    payload.me = getSessionUser();   // always live — user-agnostic cache
    return payload;
  }
  const ss = _supplySS();
  const payload = {
    users:    _cmtRowsToObjects(ss.getSheetByName(USER_SHEET)),
    comments: _cmtRowsToObjects(ss.getSheetByName(CMT_SHEET)),
    replies:  _cmtRowsToObjects(ss.getSheetByName(REPLY_SHEET)),
  };
  try { cache.put(THREADS_CACHE_KEY, JSON.stringify(payload), THREADS_CACHE_TTL); } catch (e_) {}
  payload.me = getSessionUser();
  return payload;
}

// ── Write ────────────────────────────────────────────────────────────────────

function addComment(payload) {
  payload = payload || {};
  const text = _cmtClean(payload.text);
  if (!text) throw new Error('Comment text is required.');
  const me = getSessionUser();

  const row = {
    CommentID:       Utilities.getUuid(),
    AuthorEmail:     me,
    TargetElementID: String(payload.targetElementId || '').slice(0, 200),
    Timestamp:       new Date().toISOString(),
    TextContent:     text,
    IsResolved:      false,
    ViewKey:         String(payload.viewKey || '').slice(0, 60),
    AnchorQuote:     String(payload.anchorQuote || '').slice(0, 300),
  };

  _cmtAppendRow(CMT_SHEET, CMT_HEADERS, row);
  _bustCmtCache();
  _cmtNotifyMentions(text, {
    commentId: row.CommentID,
    view:      row.ViewKey,
    target:    row.TargetElementID,
    quote:     row.AnchorQuote,
    author:    me,
  });
  return row;
}

function addReply(payload) {
  payload = payload || {};
  const text   = _cmtClean(payload.text);
  const parent = String(payload.parentCommentId || '');
  if (!text)   throw new Error('Reply text is required.');
  if (!parent) throw new Error('parentCommentId is required.');
  const me = getSessionUser();

  const row = {
    ReplyID:         Utilities.getUuid(),
    ParentCommentID: parent,
    AuthorEmail:     me,
    Timestamp:       new Date().toISOString(),
    TextContent:     text,
  };

  _cmtAppendRow(REPLY_SHEET, REPLY_HEADERS, row);
  _bustCmtCache();
  _cmtNotifyMentions(text, {
    commentId: parent,
    view:      String(payload.viewKey || ''),
    target:    '',
    quote:     '',
    author:    me,
  });
  return row;
}

function setResolved(payload, isResolved) {
  // Accept either two separate args (legacy) or a single {id, val} object (current client)
  var commentId = (payload && typeof payload === 'object') ? payload.id : payload;
  var resolved   = (payload && typeof payload === 'object') ? payload.val : isResolved;
  _cmtUpdateCell(CMT_SHEET, 'CommentID', String(commentId), 'IsResolved', !!resolved);
  _bustCmtCache();
  return true;
}

function deleteComment(commentId) {
  _cmtDeleteAuthorOnly(CMT_SHEET, 'CommentID', String(commentId));
  _cmtDeleteMatching(
    _supplySS().getSheetByName(REPLY_SHEET),
    'ParentCommentID',
    String(commentId)
  );
  _bustCmtCache();
  return true;
}

function deleteReply(replyId) {
  _cmtDeleteAuthorOnly(REPLY_SHEET, 'ReplyID', String(replyId));
  _bustCmtCache();
  return true;
}

// ── Mention notifications ────────────────────────────────────────────────────
// Re-derives mentioned emails server-side against the Users sheet (authoritative).
// Client input is ignored — no one can email arbitrary addresses via the composer.

function _cmtNotifyMentions(text, ctx) {
  const ss    = _supplySS();
  const users = _cmtRowsToObjects(ss.getSheetByName(USER_SHEET));
  const known = {};
  users.forEach(u => { if (u.Email) known[String(u.Email).toLowerCase()] = String(u.Email); });

  const mentioned = {};
  const re = /@([A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,})/g;
  let m;
  while ((m = re.exec(text)) !== null) {
    const key = m[1].toLowerCase();
    if (known[key] && known[key].toLowerCase() !== String(ctx.author).toLowerCase()) {
      mentioned[known[key]] = true;
    }
  }

  const to = Object.keys(mentioned);
  if (!to.length) return;

  let url = '';
  try { url = ScriptApp.getService().getUrl(); } catch (e_) {}
  const link = url
    ? `${url}?commentId=${encodeURIComponent(ctx.commentId)}${ctx.view ? '&view=' + encodeURIComponent(ctx.view) : ''}`
    : '';

  const subject = `${ctx.author || 'Someone'} mentioned you in the GYG SWBR Dashboard`;
  const safe    = _cmtEscHtml(text);
  const where   = ctx.target
    ? `<div style="color:#666;font-size:12px;margin:6px 0">on <b>${_cmtEscHtml(ctx.target)}</b>${ctx.quote ? ' — "' + _cmtEscHtml(ctx.quote) + '"' : ''}</div>`
    : '';

  const html = `<div style="font-family:Arial,sans-serif;max-width:520px">
    <p><b>${_cmtEscHtml(ctx.author)}</b> mentioned you in a comment:</p>
    ${where}
    <blockquote style="border-left:3px solid #FA4F2D;margin:8px 0;padding:6px 12px;color:#222">${safe}</blockquote>
    ${link ? `<p><a href="${link}" style="background:#FA4F2D;color:#fff;padding:8px 16px;border-radius:6px;text-decoration:none">Open the thread</a></p>` : ''}
  </div>`;

  try {
    GmailApp.sendEmail(to.join(','), subject, text, {
      name:     'GYG SWBR Dashboard',
      htmlBody: html,
    });
  } catch (e_) { /* quota / scope — non-fatal; comment is already saved */ }
}

// ── Sheet helpers (LockService-guarded writes) ────────────────────────────────
// All prefixed _cmt to avoid clashing with any existing helpers in Code.gs.

function _cmtClean(t) {
  return String(t == null ? '' : t).trim().slice(0, MAX_CMT_LEN);
}

function _bustCmtCache() {
  try { CacheService.getScriptCache().remove(THREADS_CACHE_KEY); } catch (e_) {}
}

function _cmtEnsureSheet(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() < 1) sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  return sh;
}

function _cmtRowsToObjects(sh) {
  if (!sh || sh.getLastRow() < 2) return [];
  const values  = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
  const headers = values[0].map(h => String(h).trim());
  const out = [];
  for (let i = 1; i < values.length; i++) {
    const o = {}; let blank = true;
    headers.forEach((h, c) => {
      let v = values[i][c];
      if (v instanceof Date) v = v.toISOString();
      if (v !== '' && v !== null) blank = false;
      o[h] = v;
    });
    if (!blank) out.push(o);
  }
  return out;
}

function _cmtAppendRow(sheetName, headers, obj) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ss = _supplySS();
    const sh = ss.getSheetByName(sheetName) || _cmtEnsureSheet(ss, sheetName, headers);
    sh.appendRow(headers.map(h => (obj[h] === undefined ? '' : obj[h])));
  } finally { lock.releaseLock(); }
}

function _cmtFindRow(sh, keyCol, keyVal) {
  const values = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
  const ci = values[0].indexOf(keyCol);
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][ci]) === keyVal) {
      return { rowIndex: i + 1, headers: values[0], row: values[i] };
    }
  }
  return null;
}

function _cmtUpdateCell(sheetName, keyCol, keyVal, setCol, setVal) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sh = _supplySS().getSheetByName(sheetName);
    const f  = _cmtFindRow(sh, keyCol, keyVal);
    if (!f) throw new Error('Row not found.');
    const c = f.headers.indexOf(setCol);
    if (c < 0) throw new Error('Column not found.');
    sh.getRange(f.rowIndex, c + 1).setValue(setVal);
  } finally { lock.releaseLock(); }
}

function _cmtDeleteAuthorOnly(sheetName, keyCol, keyVal) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sh = _supplySS().getSheetByName(sheetName);
    const f  = _cmtFindRow(sh, keyCol, keyVal);
    if (!f) return;
    const ae = f.headers.indexOf('AuthorEmail');
    if (ae >= 0 && String(f.row[ae]) !== getSessionUser()) {
      throw new Error('Only the author can delete this.');
    }
    sh.deleteRow(f.rowIndex);
  } finally { lock.releaseLock(); }
}

function _cmtDeleteMatching(sh, keyCol, keyVal) {
  if (!sh || sh.getLastRow() < 2) return;
  const values = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
  const ci = values[0].indexOf(keyCol);
  for (let i = values.length - 1; i >= 1; i--) {
    if (String(values[i][ci]) === keyVal) sh.deleteRow(i + 1);
  }
}

function _cmtEscHtml(s) {
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

// ── FPOPS-56: Auto-expire unresolved comments after 7 days ───────────────────
// Auto-resolves (does not delete) and emails tagged users.
// Schedule by running installExpireTrigger() once from the script editor.

function expireOldComments() {
  const ss = _supplySS();
  const sh = ss.getSheetByName(CMT_SHEET);
  if (!sh || sh.getLastRow() < 2) return;
  const rows = sh.getDataRange().getValues();
  const H = rows[0];
  const iRes = H.indexOf('IsResolved'), iTs = H.indexOf('Timestamp'),
        iTxt = H.indexOf('TextContent'), iId = H.indexOf('CommentID');
  const ONE_WEEK_MS = 7 * 24 * 60 * 60 * 1000;
  const now = new Date();
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][iRes]) continue;
    var ts = new Date(rows[i][iTs]);
    if (isNaN(ts) || now - ts < ONE_WEEK_MS) continue;
    sh.getRange(i + 1, iRes + 1).setValue(true);
    var mentions = String(rows[i][iTxt]).match(/@([A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,})/g) || [];
    mentions.forEach(function(m) {
      try {
        GmailApp.sendEmail(
          m.slice(1),
          'GYG SWBR Dashboard: comment auto-resolved after 7 days',
          'A comment you were tagged in was auto-resolved because it had no reply for 7 days.\n\nComment ID: ' + rows[i][iId] + '\n\nOpen the dashboard to view or reopen it.'
        );
      } catch(e) { Logger.log('expireOldComments: email failed for ' + m + ': ' + e); }
    });
  }
}

// Run once from the Apps Script editor to register the daily trigger:
function installExpireTrigger() {
  // Idempotent: remove any existing trigger first
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'expireOldComments') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('expireOldComments').timeBased().everyDays(1).atHour(6).create();
  Logger.log('installExpireTrigger: done');
}
