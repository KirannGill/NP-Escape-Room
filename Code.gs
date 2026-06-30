/**
 * TRU MN-NP Nursing Escape Room — Google Sheets Logger v2
 * ---------------------------------------------------------
 * Tabs created automatically on first run:
 *   Logins   — one row every time a student enters their name
 *   Anna     — one row per Case 1 attempt (fixed columns)
 *   Tyler    — one row per Case 2 attempt (fixed columns)
 *   Hannah   — one row per Case 3 attempt (fixed columns)
 *   Mistakes — one row per wrong/skipped answer across all cases
 *   Summary  — one row per attempt-set, cross-case rollup
 *
 * Deploy as Web App: Execute as Me, Anyone can access.
 * Paste the /exec URL into home.html (SHEET_WEBHOOK_URL)
 * and nursing_escape_room_updated.html (SHEETS_URL).
 */

const APPS_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbyfSGaKDGSFBZ3dCqqgsaDi68JhgGNPZue3GWIZGbwSLshRFyP-cTUTAnfs7sfZiKXWiQ/exec';

const LOGIN_HEADERS   = ['Student Name', 'Date', 'Time'];
const MISTAKES_HEADERS = ['Student Name', 'Date', 'Case', 'Question', 'Answer Given (Full Text)', 'Correct Answer'];
const CASE_HEADERS    = [
  'Student Name', 'Date', 'Login Timestamp', 'Completion Timestamp',
  'Score (Correct/Total)', 'Percentage', 'Pass/Fail', 'Time Taken (mm:ss)',
  'Unscored Decision Point Answers'
];
const SUMMARY_HEADERS = [
  'Student Name', 'Date', 'Time (first login)',
  'Case 1 Score', 'Case 1 Pass/Fail', 'Case 1 Time',
  'Case 2 Score', 'Case 2 Pass/Fail', 'Case 2 Time',
  'Case 3 Score', 'Case 3 Pass/Fail', 'Case 3 Time',
  'Total Score', 'Total Pass/Fail', 'Total Time'
];

// ---------------------------------------------------------------------------
// HTTP ENTRY POINTS
// ---------------------------------------------------------------------------

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    if      (body.action === 'logLogin')      logLogin(body);
    else if (body.action === 'logCaseResult') logCaseResult(body);
    return jsonResponse({ status: 'ok' });
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.message });
  }
}

function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || '';
  if (action === 'getAllData') {
    return jsonResponse(getAllData());
  }
  return jsonResponse({ status: 'ok' });
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---------------------------------------------------------------------------
// LOGIN LOGGING
// ---------------------------------------------------------------------------

function logLogin(body) {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const sh  = getOrCreateSheet(ss, 'Logins', LOGIN_HEADERS);
  const now = new Date();
  sh.appendRow([
    body.studentName || '',
    fmt(now, 'yyyy-MM-dd'),
    fmt(now, 'HH:mm:ss')
  ]);
}

// ---------------------------------------------------------------------------
// CASE RESULT LOGGING
// ---------------------------------------------------------------------------

function logCaseResult(body) {
  const ss            = SpreadsheetApp.getActiveSpreadsheet();
  const caseId        = body.caseId;
  const caseSheetName = { anna: 'Anna', tyler: 'Tyler', hannah: 'Hannah' }[caseId];
  if (!caseSheetName) return;

  const studentName          = body.studentName          || 'Unknown Student';
  const loginTimestamp       = body.loginTimestamp       || '';
  const completionTimestamp  = body.completionTimestamp  || new Date().toISOString();
  const score                = body.score                || { correct: 0, total: 0 };
  const pct                  = score.total > 0 ? Math.round((score.correct / score.total) * 100) : 0;
  const passed               = pct >= 70;
  const timeTaken            = fmtSecs(body.caseTime || 0);
  const mistakes             = body.skippedOrIncorrect   || [];
  const unscored             = body.unscoredAnswers      || [];
  const now                  = new Date();

  // 1. Write to per-case tab (fixed columns only — no dynamic question columns)
  const caseSh = getOrCreateSheet(ss, caseSheetName, CASE_HEADERS);
  caseSh.appendRow([
    studentName,
    fmt(now, 'yyyy-MM-dd'),
    loginTimestamp,
    completionTimestamp,
    `${score.correct}/${score.total}`,
    `${pct}%`,
    passed ? 'Pass' : 'Fail',
    timeTaken,
    unscored.map(u => `${u.question}: ${u.answer}`).join(' | ')
  ]);

  // 2. Write one row per mistake to the Mistakes tab
  if (mistakes.length > 0) {
    const mistakeSh  = getOrCreateSheet(ss, 'Mistakes', MISTAKES_HEADERS);
    const dateStr    = fmt(now, 'yyyy-MM-dd');
    const caseLabel  = { anna: 'Case 1 — Anna Jacobs', tyler: 'Case 2 — Tyler Haley', hannah: 'Case 3 — Hannah Howard' }[caseId];
    const mistakeRows = mistakes.map(m => [
      studentName,
      dateStr,
      caseLabel,
      m.question  || '',
      m.given     || '(skipped — no answer given)',
      m.correct   || ''
    ]);
    if (mistakeRows.length > 0) {
      mistakeSh.getRange(
        mistakeSh.getLastRow() + 1, 1,
        mistakeRows.length, MISTAKES_HEADERS.length
      ).setValues(mistakeRows);
    }
  }

  // 3. Update Summary tab
  updateSummaryRow(ss, studentName, loginTimestamp, caseId, {
    score:    `${score.correct}/${score.total}`,
    passFail: passed ? 'Pass' : 'Fail',
    time:     timeTaken,
    correct:  score.correct,
    total:    score.total,
    seconds:  body.caseTime || 0
  });
}

// ---------------------------------------------------------------------------
// SUMMARY TAB
// ---------------------------------------------------------------------------

function updateSummaryRow(ss, studentName, loginTimestamp, caseId, result) {
  const sh     = getOrCreateSheet(ss, 'Summary', SUMMARY_HEADERS);
  const colMap = {
    anna:   { score: 4, passFail: 5, time: 6 },
    tyler:  { score: 7, passFail: 8, time: 9 },
    hannah: { score: 10, passFail: 11, time: 12 }
  };
  const cols = colMap[caseId];
  if (!cols) return;

  const data = sh.getDataRange().getValues();
  let targetRow = -1;
  for (let r = data.length - 1; r >= 1; r--) {
    if (data[r][0] === studentName && !data[r][cols.score - 1]) {
      targetRow = r + 1;
      break;
    }
  }

  if (targetRow === -1) {
    const now    = new Date();
    const newRow = new Array(SUMMARY_HEADERS.length).fill('');
    newRow[0] = studentName;
    newRow[1] = fmt(now, 'yyyy-MM-dd');
    newRow[2] = loginTimestamp || fmt(now, 'HH:mm:ss');
    sh.appendRow(newRow);
    targetRow = sh.getLastRow();
  }

  sh.getRange(targetRow, cols.score).setValue(result.score);
  sh.getRange(targetRow, cols.passFail).setValue(result.passFail);
  sh.getRange(targetRow, cols.time).setValue(result.time);
  recomputeTotals(sh, targetRow);
}

function recomputeTotals(sh, row) {
  const v             = sh.getRange(row, 1, 1, SUMMARY_HEADERS.length).getValues()[0];
  const scoreIdxs     = [3, 6, 9];   // 0-based col indices for Case 1/2/3 Score
  const passIdxs      = [4, 7, 10];
  const timeIdxs      = [5, 8, 11];
  let totalC = 0, totalT = 0, totalSecs = 0, allPassed = true, completedCount = 0;

  scoreIdxs.forEach((si, i) => {
    const s = v[si];
    if (s && typeof s === 'string' && s.includes('/')) {
      const [c, t] = s.split('/').map(Number);
      if (!isNaN(c) && !isNaN(t)) { totalC += c; totalT += t; completedCount++; }
    }
    if (!v[passIdxs[i]] || v[passIdxs[i]] === 'Fail') allPassed = false;
    const ts = v[timeIdxs[i]];
    if (ts && typeof ts === 'string' && ts.includes(':')) {
      const [m, s2] = ts.split(':').map(Number);
      if (!isNaN(m) && !isNaN(s2)) totalSecs += (m * 60 + s2);
    }
  });

  if (completedCount > 0) sh.getRange(row, 13).setValue(`${totalC}/${totalT}`);
  if (completedCount === 3) sh.getRange(row, 14).setValue(allPassed ? 'Pass' : 'Fail');
  if (totalSecs > 0) sh.getRange(row, 15).setValue(fmtSecs(totalSecs));
}

// ---------------------------------------------------------------------------
// getAllData — called by dashboard.html via ?action=getAllData
// Returns structured JSON the dashboard can render directly.
// ---------------------------------------------------------------------------

function getAllData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // --- Summary rows ---
  const summarySh   = ss.getSheetByName('Summary');
  const summaryRows = summarySh ? sheetToObjects(summarySh) : [];

  // --- Per-case rows ---
  const caseData = {};
  ['Anna', 'Tyler', 'Hannah'].forEach(name => {
    const sh = ss.getSheetByName(name);
    caseData[name.toLowerCase()] = sh ? sheetToObjects(sh) : [];
  });

  // --- Mistakes rows ---
  const mistakesSh   = ss.getSheetByName('Mistakes');
  const mistakeRows  = mistakesSh ? sheetToObjects(mistakesSh) : [];

  // --- Logins rows ---
  const loginsSh   = ss.getSheetByName('Logins');
  const loginRows  = loginsSh ? sheetToObjects(loginsSh) : [];

  return {
    summary:  summaryRows,
    cases:    caseData,
    mistakes: mistakeRows,
    logins:   loginRows
  };
}

// ---------------------------------------------------------------------------
// HELPERS
// ---------------------------------------------------------------------------

function getOrCreateSheet(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(headers);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#1a365d').setFontColor('#ffffff');
  }
  return sh;
}

/** Convert a sheet's data to an array of plain objects keyed by header row. */
function sheetToObjects(sh) {
  const data = sh.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i]; });
    return obj;
  });
}

function fmt(date, pattern) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), pattern);
}

function fmtSecs(totalSeconds) {
  const m = Math.floor(totalSeconds / 60);
  const s = Math.floor(totalSeconds % 60);
  return `${String(m).padStart(2, '0')}:${String(s).padStart(2, '0')}`;
}
