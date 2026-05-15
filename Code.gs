// ============================================================
// TUTOR TRACKER - Apps Script backend
// Sheet 1: Live_Sessions (rewritten on Sync)
// Sheet 2: Archive (append-only, never cleared)
// ============================================================

const LIVE_SHEET_NAME = "Live_Sessions";
const ARCHIVE_SHEET_NAME = "Archive";

const LIVE_HEADERS = [
  "Timestamp",
  "Date",
  "Student",
  "Complete",
  "Paid",
  "Amount",
  "Notes",
  "Session ID"
];

const ARCHIVE_HEADERS = [
  ...LIVE_HEADERS,
  "Bulk Import"
];

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
      return jsonResponse({ status: "ok", message: "Tutor Tracker webhook active" });
    }

    if (payload.action === "sync_all") {
      return jsonResponse(syncAll(payload.sessions || [], payload.students || []));
    }

    if (payload.action === "log" && payload.session) {
      return jsonResponse(addOrUpdateLiveSession(payload.session, payload.students || []));
    }

    if (payload.action === "update_paid" && payload.session) {
      return jsonResponse(updateLivePaid(payload.session));
    }

    return jsonResponse({ status: "error", message: "Unknown or missing action" });
  } catch (err) {
    return jsonResponse({ status: "error", message: String(err) });
  }
}

function parsePayload(e) {
  if (e && e.postData && e.postData.contents) {
    return JSON.parse(e.postData.contents);
  }
  if (e && e.parameter && e.parameter.payload) {
    return JSON.parse(decodeURIComponent(e.parameter.payload));
  }
  return null;
}

function syncAll(sessions, students) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const liveSheet = getOrCreateSheet(ss, LIVE_SHEET_NAME);
  const archiveSheet = getOrCreateSheet(ss, ARCHIVE_SHEET_NAME);

  resetLiveSheet(liveSheet);
  ensureArchiveHeaders(archiveSheet);

  const liveRows = sessions.map((session) => buildLiveRow(session, students));
  if (liveRows.length) {
    liveSheet.getRange(2, 1, liveRows.length, LIVE_HEADERS.length).setValues(liveRows);
  }

  const existingArchiveIds = getExistingArchiveIds(archiveSheet);
  const archiveRows = [];
  sessions.forEach((session) => {
    const id = String(session.id || "");
    if (!id || existingArchiveIds[id]) return;
    archiveRows.push(buildArchiveRow(session, students));
    existingArchiveIds[id] = true;
  });

  if (archiveRows.length) {
    archiveSheet
      .getRange(archiveSheet.getLastRow() + 1, 1, archiveRows.length, ARCHIVE_HEADERS.length)
      .setValues(archiveRows);
  }

  formatSheet(liveSheet, LIVE_HEADERS.length);
  formatSheet(archiveSheet, ARCHIVE_HEADERS.length);

  return {
    status: "ok",
    message: "Live_Sessions rewritten and Archive appended",
    liveRows: liveRows.length,
    archiveRowsAdded: archiveRows.length
  };
}

function addOrUpdateLiveSession(session, students) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const liveSheet = getOrCreateSheet(ss, LIVE_SHEET_NAME);
  const archiveSheet = getOrCreateSheet(ss, ARCHIVE_SHEET_NAME);

  if (liveSheet.getLastRow() === 0) resetLiveSheet(liveSheet);
  ensureArchiveHeaders(archiveSheet);

  const row = buildLiveRow(session, students);
  const rowIndex = findLiveRowBySessionId(liveSheet, session.id);
  if (rowIndex) {
    liveSheet.getRange(rowIndex, 1, 1, LIVE_HEADERS.length).setValues([row]);
  } else {
    liveSheet.appendRow(row);
  }

  const id = String(session.id || "");
  const existingArchiveIds = getExistingArchiveIds(archiveSheet);
  if (id && !existingArchiveIds[id]) {
    archiveSheet.appendRow(buildArchiveRow(session, students));
  }

  formatSheet(liveSheet, LIVE_HEADERS.length);
  formatSheet(archiveSheet, ARCHIVE_HEADERS.length);

  return { status: "ok", message: "Session saved" };
}

function updateLivePaid(session) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const liveSheet = getOrCreateSheet(ss, LIVE_SHEET_NAME);
  if (liveSheet.getLastRow() === 0) resetLiveSheet(liveSheet);

  const rowIndex = findLiveRowBySessionId(liveSheet, session.id);
  if (!rowIndex) return { status: "not_found", message: "Session ID not found in Live_Sessions" };

  liveSheet.getRange(rowIndex, 1).setValue(new Date());
  liveSheet.getRange(rowIndex, 5).setValue(truthy(session.paid) ? "Yes" : "No");
  return { status: "ok", message: "Paid status updated in Live_Sessions" };
}

function resetLiveSheet(sheet) {
  sheet.clearContents();
  sheet.getRange(1, 1, 1, LIVE_HEADERS.length).setValues([LIVE_HEADERS]);
  styleHeader(sheet, LIVE_HEADERS.length);
}

function ensureArchiveHeaders(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, ARCHIVE_HEADERS.length).setValues([ARCHIVE_HEADERS]);
    styleHeader(sheet, ARCHIVE_HEADERS.length);
    return;
  }

  const current = sheet.getRange(1, 1, 1, ARCHIVE_HEADERS.length).getValues()[0];
  const matches = ARCHIVE_HEADERS.every((header, index) => String(current[index]) === header);
  if (!matches) {
    sheet.getRange(1, 1, 1, ARCHIVE_HEADERS.length).setValues([ARCHIVE_HEADERS]);
    styleHeader(sheet, ARCHIVE_HEADERS.length);
  }
}

function buildLiveRow(session, students) {
  const normalized = normalizeSession(session, students);
  return [
    new Date(),
    normalized.date,
    normalized.student,
    normalized.complete,
    normalized.paid,
    normalized.amount,
    normalized.notes,
    normalized.id
  ];
}

function buildArchiveRow(session, students) {
  const liveRow = buildLiveRow(session, students);
  const normalized = normalizeSession(session, students);
  return [...liveRow, buildBulkImportString(normalized)];
}

function normalizeSession(session, students) {
  const student =
    session.studentName ||
    findStudentName(session.studentId, students) ||
    session.studentId ||
    "Unknown";

  return {
    id: String(session.id || ""),
    date: session.date || session.dateStr || "",
    student,
    complete: truthy(session.complete) ? "Yes" : "No",
    paid: truthy(session.paid) ? "Yes" : "No",
    amount: Number(session.amount || 0),
    notes: session.notes || ""
  };
}

function buildBulkImportString(session) {
  return [
    session.date,
    session.student,
    session.complete,
    session.paid,
    session.amount,
    session.notes
  ].join("\t");
}

function findStudentName(studentId, students) {
  if (!studentId || !students || !students.length) return "";
  const match = students.find((student) => String(student.id) === String(studentId));
  return match ? match.name : "";
}

function truthy(value) {
  return value === true || value === "true" || value === "Yes" || value === "yes" || value === 1 || value === "1";
}

function getExistingArchiveIds(sheet) {
  const ids = {};
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return ids;

  const values = sheet.getRange(2, 8, lastRow - 1, 1).getValues();
  values.forEach((row) => {
    const id = String(row[0] || "");
    if (id) ids[id] = true;
  });
  return ids;
}

function findLiveRowBySessionId(sheet, id) {
  const target = String(id || "");
  if (!target || sheet.getLastRow() < 2) return null;

  const values = sheet.getRange(2, 8, sheet.getLastRow() - 1, 1).getValues();
  for (let index = 0; index < values.length; index++) {
    if (String(values[index][0] || "") === target) return index + 2;
  }
  return null;
}

function getOrCreateSheet(ss, name) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  return sheet;
}

function styleHeader(sheet, width) {
  sheet.getRange(1, 1, 1, width)
    .setFontWeight("bold")
    .setBackground("#1c2128")
    .setFontColor("#adbac7");
  sheet.setFrozenRows(1);
}

function formatSheet(sheet, width) {
  for (let column = 1; column <= width; column++) {
    sheet.autoResizeColumn(column);
  }
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
