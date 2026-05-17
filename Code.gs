// ============================================================
// TUTOR TRACKER - Apps Script backend
// Live_Sessions is rewritten on full sync.
// Archive is append-only and never cleared.
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

    const action = payload.action || "";

    if (action === "sync_all" || action === "syncAllSessions") {
      return jsonResponse(syncAllSessions(payload.sessions || [], payload.students || []));
    }

    if (action === "log" && payload.session) {
      return jsonResponse(logSingleSession(payload.session, payload.students || []));
    }

    if (action === "update_paid" && payload.session) {
      return jsonResponse(updateLivePaid(payload.session));
    }

    return jsonResponse({ status: "error", message: "Unknown or missing action: " + action });
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

function syncAllSessions(sessions, students) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const liveSheet = getOrCreateSheet(ss, LIVE_SHEET_NAME);
  const archiveSheet = getOrCreateSheet(ss, ARCHIVE_SHEET_NAME);

  // This is the key behavior: full sync replaces Sheet 1 with the phone's current state.
  rewriteLiveSheet(liveSheet, sessions, students);

  // Archive remains a permanent ledger: append only session IDs that are not already archived.
  appendNewArchiveRows(archiveSheet, sessions, students);

  formatSheet(liveSheet, LIVE_HEADERS.length);
  formatSheet(archiveSheet, ARCHIVE_HEADERS.length);

  return {
    status: "ok",
    message: "Live_Sessions rewritten; Archive appended only",
    liveRows: sessions.length
  };
}

function rewriteLiveSheet(sheet, sessions, students) {
  sheet.clearContents();
  sheet.getRange(1, 1, 1, LIVE_HEADERS.length).setValues([LIVE_HEADERS]);
  styleHeader(sheet, LIVE_HEADERS.length);

  const rows = sessions.map((session) => buildLiveRow(session, students));
  if (rows.length) {
    sheet.getRange(2, 1, rows.length, LIVE_HEADERS.length).setValues(rows);
  }
}

function logSingleSession(session, students) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const liveSheet = getOrCreateSheet(ss, LIVE_SHEET_NAME);
  const archiveSheet = getOrCreateSheet(ss, ARCHIVE_SHEET_NAME);

  ensureLiveHeaders(liveSheet);
  ensureArchiveHeaders(archiveSheet);

  const row = buildLiveRow(session, students);
  const rowIndex = findRowBySessionId(liveSheet, session.id);

  if (rowIndex) {
    liveSheet.getRange(rowIndex, 1, 1, LIVE_HEADERS.length).setValues([row]);
  } else {
    liveSheet.appendRow(row);
  }

  appendNewArchiveRows(archiveSheet, [session], students);

  formatSheet(liveSheet, LIVE_HEADERS.length);
  formatSheet(archiveSheet, ARCHIVE_HEADERS.length);

  return { status: "ok", message: "Session logged" };
}

function updateLivePaid(session) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const liveSheet = getOrCreateSheet(ss, LIVE_SHEET_NAME);
  ensureLiveHeaders(liveSheet);

  const rowIndex = findRowBySessionId(liveSheet, session.id);
  if (!rowIndex) {
    return { status: "not_found", message: "Session ID not found in Live_Sessions" };
  }

  liveSheet.getRange(rowIndex, 1).setValue(new Date());
  liveSheet.getRange(rowIndex, 5).setValue(isTruthy(session.paid) ? "Yes" : "No");

  return { status: "ok", message: "Paid status updated in Live_Sessions" };
}

function appendNewArchiveRows(sheet, sessions, students) {
  ensureArchiveHeaders(sheet);

  const existingIds = getExistingArchiveIds(sheet);
  const rows = [];

  sessions.forEach((session) => {
    const id = String(session.id || "");
    if (!id || existingIds[id]) return;

    rows.push(buildArchiveRow(session, students));
    existingIds[id] = true;
  });

  if (rows.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, ARCHIVE_HEADERS.length).setValues(rows);
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
    complete: isTruthy(session.complete) ? "Yes" : "No",
    paid: isTruthy(session.paid) ? "Yes" : "No",
    amount: Number(session.amount || 0),
    notes: session.notes || ""
  };
}
