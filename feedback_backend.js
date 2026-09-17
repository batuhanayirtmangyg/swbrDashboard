/* ===========================================================================
   FEEDBACK SYSTEM — backend for the Supply WBR Dashboard
   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
   ONE-TIME SETUP (run once from the Apps Script editor after pasting):
     Run setupFeedbackSheet() — creates the Feedback tab.
   =========================================================================== */

const FB_SHEET = 'Feedback';
const FB_HEADERS = [
  'Timestamp', 'Email', 'View', 'Granularity', 'WeekLabel',
  'Filters', 'CacheState', 'Browser', 'ScreenWidth',
  'Q1_Usability', 'Q2_Type', 'Q3_Area', 'Q4_Specifics',
  'Q5_Metric', 'Q6_Notes', 'Q7_Blocks', 'Q8_Screenshot', 'Q9_FollowupOK'
];

function setupFeedbackSheet() {
  const ss = _supplySS();
  let sh = ss.getSheetByName(FB_SHEET);
  if (!sh) sh = ss.insertSheet(FB_SHEET);
  if (sh.getLastRow() < 1) {
    sh.getRange(1, 1, 1, FB_HEADERS.length).setValues([FB_HEADERS]);
    sh.setFrozenRows(1);
  }
  Logger.log('setupFeedbackSheet: done — ' + ss.getName() + ' ' + ss.getUrl());
  return 'ok';
}

// Admin-only inbox. Same gate and same {rows,headers} / {error} shape as getUsageLog
// (Code.gs) so the client renders both with one function. isAdmin lives in Code.gs — Apps
// Script merges all .gs into one global scope, so don't redeclare it here.
function getFeedbackLog() {
  var email = '';
  try { email = Session.getActiveUser().getEmail() || ''; } catch (e) {}
  if (!isAdmin(email)) return { error: 'unauthorized' };
  var sh = _supplySS().getSheetByName(FB_SHEET);
  if (!sh || sh.getLastRow() < 2) return { rows: [], headers: FB_HEADERS };
  return { rows: sh.getRange(2, 1, sh.getLastRow() - 1, FB_HEADERS.length).getValues(),
           headers: FB_HEADERS };
}

function submitFeedback(payload) {
  payload = payload || {};
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ss = _supplySS();
    let sh = ss.getSheetByName(FB_SHEET);
    if (!sh) { setupFeedbackSheet(); sh = ss.getSheetByName(FB_SHEET); }
    const me = Session.getActiveUser().getEmail() || '';
    sh.appendRow(FB_HEADERS.map(function(h) {
      if (h === 'Timestamp') return new Date().toISOString();
      if (h === 'Email') return me;
      return payload[h] !== undefined ? String(payload[h]) : '';
    }));
  } finally { lock.releaseLock(); }
  return true;
}
