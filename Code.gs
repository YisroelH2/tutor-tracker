// ============================================================
// TUTOR TRACKER — PWA Webhook Backend (Sheet1 + Sheet2)
// ============================================================

const LIVE_SHEET_NAME = "Sheet1";      // Live Log (rewritable)
const ARCHIVE_SHEET_NAME = "Sheet2";   // Session Archive (append-only)
const HEADERS = [
  "Timestamp",
  "Date",
  "Student",
  "Complete",
  "Paid",
  "Amount",
  "Notes",
  "Session ID",
  "Backup Payload"
];

function doGet(e) { return handleRequest(e); }
function doPost(e) { return handleRequest(e); }

function handleRequest(e) {
  try {
    const data = parseRequestBody(e);
    if (!data) return makeResponse({ status: "ok", message: "Tutor Tracker Webhook Active ✓" });

    const action = data.action || "log";
    const session = data.session;
    if (!session) return makeResponse({ status: "error", message: "No session data received" });

    if (action === "log") return makeResponse(addSession(session));
    if (action === "update_paid") return makeResponse(updatePaid(session));

    return makeResponse({ status: "error", message: "Unknown action: " + action });
  } catch (err) {
    return makeResponse({ status: "error", message: String(err) });
  }
}

function parseRequestBody(e) {
  if (e && e.postData && e.postData.contents) return JSON.parse(e.postData.contents);
  if (e && e.parameter && e.parameter.payload) return JSON.parse(decodeURIComponent(e.parameter.payload));
  return null;
}

function addSession(session) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const liveSheet = getOrCreate(ss, LIVE_SHEET_NAME);
  const archiveSheet = getOrCreate(ss, ARCHIVE_SHEET_NAME);

  ensureHeaders(liveSheet);
  ensureHeaders(archiveSheet);

  if (findRowBySessionId(liveSheet, session.id)) {
    return { status: "duplicate", message: "Session already exists in Sheet1" };
  }

  const row = buildRow(session);
  liveSheet.appendRow(row);

  if (session.complete) {
    archiveSheet.appendRow(row);
  }

  formatSheet(liveSheet);
  return { status: "ok", message: "Session saved to Sheet1 and archived in Sheet2 when complete ✓" };
}

function updatePaid(session) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const liveSheet = getOrCreate(ss, LIVE_SHEET_NAME);
  ensureHeaders(liveSheet);

  const rowIndex = findRowBySessionId(liveSheet, session.id);
  if (!rowIndex) return { status: "not_found", message: "Session ID not found in Sheet1" };

  liveSheet.getRange(rowIndex, 5).setValue(session.paid ? "Yes" : "No");
  liveSheet.getRange(rowIndex, 1).setValue(new Date().toLocaleString());

  return { status: "ok", message: "Sheet1 paid status updated ✓" };
}

function buildRow(s) {
  const normalized = {
    id: s.id || "",
    date: s.date || "",
    studentName: s.studentName || s.studentId || "Unknown",
    complete: !!s.complete,
    paid: !!s.paid,
    amount: Number(s.amount) || 0,
    notes: s.notes || ""
  };

  const backupPayload = JSON.stringify(normalized);

  return [
    new Date().toLocaleString(),
    normalized.date,
    normalized.studentName,
    normalized.complete ? "Yes" : "No",
    normalized.paid ? "Yes" : "No",
    normalized.amount,
    normalized.notes,
    normalized.id,
    backupPayload
  ];
}

function getOrCreate(ss, name) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    ensureHeaders(sheet);
  }
  return sheet;
}

function ensureHeaders(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.getRange(1, 1, 1, HEADERS.length)
      .setFontWeight("bold")
      .setBackground("#1c2128")
      .setFontColor("#adbac7");
    sheet.setFrozenRows(1);
  }
}

function findRowBySessionId(sheet, id) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;

  const ids = sheet.getRange(2, 8, lastRow - 1, 1).getValues();
  const target = String(id || "");
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === target) return i + 2;
  }
  return null;
}

function formatSheet(sheet) {
  for (var c = 1; c <= HEADERS.length; c++) sheet.autoResizeColumn(c);
}

function makeResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
