// ============================================================
//  TUTOR TRACKER — PWA + LIVE LOG + IMMUTABLE ARCHIVE
// ============================================================

const LIVE_SHEET_NAME = 'Sheet1';
const ARCHIVE_SHEET_NAME = 'Sheet2';
const LIVE_HEADERS = ['Timestamp', 'Date', 'Student', 'Complete', 'Paid', 'Amount', 'Notes', 'Session ID'];
const ARCHIVE_HEADERS = [...LIVE_HEADERS, 'Session Backup'];

function doGet(e) {
  return handleRequest(e);
}

function doPost(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  try {
    const payload = parsePayload(e);
    if (!payload) {
      return makeResponse({ status: 'ok', message: 'Tutor Tracker webhook active ✓' });
    }

    const action = payload.action || 'log';
    const session = payload.session;

    if (!session) {
      return makeResponse({ status: 'error', message: 'No session data received' });
    }

    if (action === 'log') return makeResponse(addSession(session));
    if (action === 'update_paid') return makeResponse(updatePaid(session));

    return makeResponse({ status: 'error', message: `Unknown action: ${action}` });
  } catch (err) {
    return makeResponse({ status: 'error', message: err.toString() });
  }
}

function parsePayload(e) {
  if (e && e.postData && e.postData.contents) return JSON.parse(e.postData.contents);
  if (e && e.parameter && e.parameter.payload) return JSON.parse(decodeURIComponent(e.parameter.payload));
  return null;
}

function addSession(session) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const liveSheet = getOrCreateSheet(ss, LIVE_SHEET_NAME, LIVE_HEADERS);
  const archiveSheet = getOrCreateSheet(ss, ARCHIVE_SHEET_NAME, ARCHIVE_HEADERS);

  const sessionId = String(session.id || '');
  if (!sessionId) return { status: 'error', message: 'Missing session id' };

  if (findRowBySessionId(liveSheet, sessionId)) {
    return { status: 'duplicate', message: 'Session already exists in Sheet1' };
  }

  const liveRow = buildLiveRow(session);
  liveSheet.appendRow(liveRow);

  if (session.complete) {
    const archiveRow = [...liveRow, buildBackupString(session)];
    archiveSheet.appendRow(archiveRow);
  }

  formatSheet(liveSheet, LIVE_HEADERS.length);
  formatSheet(archiveSheet, ARCHIVE_HEADERS.length);

  return { status: 'ok', message: 'Session logged to Sheet1 and archived to Sheet2 when complete ✓' };
}

function updatePaid(session) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const liveSheet = getOrCreateSheet(ss, LIVE_SHEET_NAME, LIVE_HEADERS);
  const rowIndex = findRowBySessionId(liveSheet, String(session.id || ''));

  if (!rowIndex) {
    return { status: 'not_found', message: 'Session ID not found in Sheet1' };
  }

  liveSheet.getRange(rowIndex, 5).setValue(session.paid ? 'Yes' : 'No');
  liveSheet.getRange(rowIndex, 1).setValue(new Date().toLocaleString());

  return { status: 'ok', message: 'Sheet1 paid status updated ✓' };
}

function buildLiveRow(session) {
  return [
    new Date().toLocaleString(),
    session.date || '',
    session.studentName || session.studentId || 'Unknown',
    session.complete ? 'Yes' : 'No',
    session.paid ? 'Yes' : 'No',
    Number(session.amount || 0),
    session.notes || '',
    String(session.id || ''),
  ];
}

function buildBackupString(session) {
  return JSON.stringify({
    id: String(session.id || ''),
    date: session.date || '',
    studentName: session.studentName || '',
    studentId: session.studentId || '',
    complete: Boolean(session.complete),
    paid: Boolean(session.paid),
    amount: Number(session.amount || 0),
    notes: session.notes || '',
  });
}

function getOrCreateSheet(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  ensureHeaders(sheet, headers);
  return sheet;
}

function ensureHeaders(sheet, headers) {
  const existingHeader = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  const matches = existingHeader.every((v, i) => String(v) === String(headers[i]));

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    styleHeader(sheet, headers.length);
    return;
  }

  if (!matches) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    styleHeader(sheet, headers.length);
  }
}

function styleHeader(sheet, width) {
  sheet.getRange(1, 1, 1, width)
    .setFontWeight('bold')
    .setBackground('#1c2128')
    .setFontColor('#adbac7');
  sheet.setFrozenRows(1);
}

function findRowBySessionId(sheet, id) {
  if (!id) return null;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;

  const ids = sheet.getRange(2, 8, lastRow - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) return i + 2;
  }
  return null;
}

function formatSheet(sheet, columnCount) {
  for (let i = 1; i <= columnCount; i++) sheet.autoResizeColumn(i);
}

function makeResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
