/**
 * TRU MN-NP Nursing Escape Room — Google Sheets Logger
 * Tabs created automatically on first run:
 *   - Anna     (one row per attempt at Case 1)
 *   - Tyler    (one row per attempt at Case 2)
 *   - Hannah   (one row per attempt at Case 3)
 *   - Summary  (one row per attempt-set, cross-case rollup)
 *   - Logins   (one row every time a student enters their name on the home page)
 */

const CASE_SHEETS = ['Anna', 'Tyler', 'Hannah'];
const LOGIN_SHEET = 'Logins';
const SUMMARY_SHEET = 'Summary';

// Column headers for each per-case tab.
// "Q: ..." columns are added dynamically per-case the first time a question is seen,
// so the header row grows over time as new question types are encountered.
const CASE_FIXED_HEADERS = [
  'Student Name', 'Date', 'Login Timestamp', 'Completion Timestamp',
  'Score (Correct/Total)', 'Percentage', 'Pass/Fail', 'Time Taken (mm:ss)'
];
const UNSCORED_HEADER = 'Unscored Decision Point Answers';

const SUMMARY_HEADERS = [
  'Student Name', 'Date', 'Time (first login)',
  'Case 1 Score', 'Case 1 Pass/Fail', 'Case 1 Time',
  'Case 2 Score', 'Case 2 Pass/Fail', 'Case 2 Time',
  'Case 3 Score', 'Case 3 Pass/Fail', 'Case 3 Time',
  'Total Score', 'Total Pass/Fail', 'Total Time'
];

const LOGIN_HEADERS = ['Student Name', 'Date', 'Time'];

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;

    if (action === 'logLogin') {
      logLogin(body);
    } else if (action === 'logCaseResult') {
      logCaseResult(body);
    }

    return ContentService.createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  // Kept for compatibility with any existing dashboard fetches; returns a simple ok.
  return ContentService.createTextOutput(JSON.stringify({ status: 'ok' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---------------------------------------------------------------------------
// LOGIN LOGGING — called from home.html the moment a student enters their name
// ---------------------------------------------------------------------------
function logLogin(body) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateSheet(ss, LOGIN_SHEET, LOGIN_HEADERS);
  const now = new Date();
  sheet.appendRow([
    body.studentName || '',
    Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
    Utilities.formatDate(now, Session.getScriptTimeZone(), 'HH:mm:ss')
  ]);
}

// ---------------------------------------------------------------------------
// CASE RESULT LOGGING — called from nursing_escape_room_updated.html when a
// case is completed (markCaseCompleted). Writes one row to the matching
// per-case tab, and updates (or creates) that student's row on Summary.
// ---------------------------------------------------------------------------
function logCaseResult(body) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const caseId = body.caseId; // 'anna' | 'tyler' | 'hannah'
  const caseSheetName = { anna: 'Anna', tyler: 'Tyler', hannah: 'Hannah' }[caseId];
  if (!caseSheetName) return;

  const studentName = body.studentName || 'Unknown Student';
  const loginTimestamp = body.loginTimestamp || '';
  const completionTimestamp = body.completionTimestamp || new Date().toISOString();
  const score = body.score || { correct: 0, total: 0 };
  const pct = score.total > 0 ? Math.round((score.correct / score.total) * 100) : 0;
  const passed = pct >= 70;
  const timeTaken = formatSeconds(body.caseTime || 0);

  // skippedOrIncorrect: array of { question, given, correct } for every question
  // the student either skipped or answered incorrectly across the whole case.
  const mistakes = body.skippedOrIncorrect || [];
  // unscoredAnswers: array of { question, answer } for free-text / no-point
  // decision-point questions (e.g. DP1 free response, Hannah's checkbox list).
  const unscored = body.unscoredAnswers || [];

  const sheet = getOrCreateCaseSheet(ss, caseSheetName, mistakes);

  // Build the row in the same column order as the header row currently on the sheet.
  const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const row = headerRow.map(header => {
    switch (header) {
      case 'Student Name': return studentName;
      case 'Date': return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
      case 'Login Timestamp': return loginTimestamp;
      case 'Completion Timestamp': return completionTimestamp;
      case 'Score (Correct/Total)': return `${score.correct}/${score.total}`;
      case 'Percentage': return `${pct}%`;
      case 'Pass/Fail': return passed ? 'Pass' : 'Fail';
      case 'Time Taken (mm:ss)': return timeTaken;
      case UNSCORED_HEADER: return unscored.map(u => `${u.question}: ${u.answer}`).join(' | ');
      default: {
        // This header is a per-question column (added dynamically). Find a
        // matching mistake entry; if found, the student got it wrong/skipped.
        const match = mistakes.find(m => m.question === header);
        if (match) return match.given ? `Skipped/Incorrect — answered: ${match.given}` : 'Skipped';
        return ''; // got it right, or question not part of this case
      }
    }
  });
  sheet.appendRow(row);

  updateSummaryRow(ss, studentName, loginTimestamp, caseId, {
    score: `${score.correct}/${score.total}`,
    passFail: passed ? 'Pass' : 'Fail',
    time: timeTaken,
    correct: score.correct,
    total: score.total,
    seconds: body.caseTime || 0
  });
}

// Ensures the per-case tab exists with the fixed headers, the unscored column,
// and a column for every distinct question seen so far (existing + new ones
// from this submission). New question columns are appended to the right.
function getOrCreateCaseSheet(ss, sheetName, mistakes) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow([...CASE_FIXED_HEADERS, UNSCORED_HEADER]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, sheet.getLastColumn()).setFontWeight('bold');
  }

  const headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
  const existingHeaders = headerRange.getValues()[0];
  const existingSet = new Set(existingHeaders);

  const newQuestionHeaders = mistakes
    .map(m => m.question)
    .filter(q => q && !existingSet.has(q));

  // De-duplicate within this batch too
  const uniqueNew = [...new Set(newQuestionHeaders)];

  if (uniqueNew.length > 0) {
    const startCol = sheet.getLastColumn() + 1;
    sheet.getRange(1, startCol, 1, uniqueNew.length).setValues([uniqueNew]);
    sheet.getRange(1, startCol, 1, uniqueNew.length).setFontWeight('bold');
  }

  return sheet;
}

function getOrCreateSheet(ss, sheetName, headers) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }
  return sheet;
}

// ---------------------------------------------------------------------------
// SUMMARY TAB — one row per student per attempt-set. Since each case can be
// completed independently (possibly across different days), we find the most
// recent incomplete row for this student (missing this case's columns) and
// fill it in; if none exists, we start a new row.
// ---------------------------------------------------------------------------
function updateSummaryRow(ss, studentName, loginTimestamp, caseId, result) {
  const sheet = getOrCreateSheet(ss, SUMMARY_SHEET, SUMMARY_HEADERS);
  const colMap = {
    anna:   { score: 4, passFail: 5, time: 6 },
    tyler:  { score: 7, passFail: 8, time: 9 },
    hannah: { score: 10, passFail: 11, time: 12 }
  };
  const cols = colMap[caseId];
  if (!cols) return;

  const data = sheet.getDataRange().getValues();
  let targetRow = -1;

  // Look for the most recent row for this student where this case's column is still empty.
  for (let r = data.length - 1; r >= 1; r--) {
    if (data[r][0] === studentName && !data[r][cols.score - 1]) {
      targetRow = r + 1; // 1-indexed sheet row
      break;
    }
  }

  if (targetRow === -1) {
    // No existing open row — start a new one.
    const now = new Date();
    const newRow = new Array(SUMMARY_HEADERS.length).fill('');
    newRow[0] = studentName;
    newRow[1] = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    newRow[2] = loginTimestamp || Utilities.formatDate(now, Session.getScriptTimeZone(), 'HH:mm:ss');
    sheet.appendRow(newRow);
    targetRow = sheet.getLastRow();
  }

  sheet.getRange(targetRow, cols.score).setValue(result.score);
  sheet.getRange(targetRow, cols.passFail).setValue(result.passFail);
  sheet.getRange(targetRow, cols.time).setValue(result.time);

  recomputeTotals(sheet, targetRow);
}

// Recomputes the Total Score / Total Pass-Fail / Total Time columns for a
// summary row once any of the three per-case results have been filled in.
function recomputeTotals(sheet, row) {
  const values = sheet.getRange(row, 1, 1, SUMMARY_HEADERS.length).getValues()[0];
  const caseScoreCols = [3, 6, 9]; // 0-indexed: Case 1/2/3 Score columns
  const casePassCols = [4, 7, 10];
  const caseTimeCols = [5, 8, 11];

  let totalCorrect = 0, totalQuestions = 0, totalSeconds = 0;
  let anyCompleted = false, allPassed = true, anyFailed = false;

  caseScoreCols.forEach((scoreCol, i) => {
    const scoreStr = values[scoreCol]; // e.g. "8/10"
    if (scoreStr && typeof scoreStr === 'string' && scoreStr.includes('/')) {
      const [c, t] = scoreStr.split('/').map(Number);
      if (!isNaN(c) && !isNaN(t)) {
        totalCorrect += c;
        totalQuestions += t;
        anyCompleted = true;
      }
    }
    const passFail = values[casePassCols[i]];
    if (passFail === 'Fail') { allPassed = false; anyFailed = true; }
    if (!passFail) allPassed = false;

    const timeStr = values[caseTimeCols[i]]; // "mm:ss"
    if (timeStr && typeof timeStr === 'string' && timeStr.includes(':')) {
      const [m, s] = timeStr.split(':').map(Number);
      if (!isNaN(m) && !isNaN(s)) totalSeconds += (m * 60 + s);
    }
  });

  if (anyCompleted) {
    sheet.getRange(row, 13).setValue(`${totalCorrect}/${totalQuestions}`); // Total Score
  }
  // Only mark overall Pass/Fail once all three cases have a result.
  const allThreeDone = caseScoreCols.every(c => values[c]) || (totalQuestions > 0 && [values[3], values[6], values[9]].filter(Boolean).length === 3);
  if ([values[3], values[6], values[9]].filter(Boolean).length === 3) {
    sheet.getRange(row, 14).setValue(allPassed ? 'Pass' : 'Fail'); // Total Pass/Fail
  }
  if (totalSeconds > 0) {
    sheet.getRange(row, 15).setValue(formatSeconds(totalSeconds)); // Total Time
  }
}

function formatSeconds(totalSeconds) {
  const m = Math.floor(totalSeconds / 60);
  const s = Math.floor(totalSeconds % 60);
  return `${String(m).padStart(2, '0')}:${String(s).padStart(2, '0')}`;
}
