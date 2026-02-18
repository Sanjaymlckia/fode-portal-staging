/******************** UTIL ********************/

function mustGetSheet_(ss, name) {
  var sh = ss.getSheetByName(name);
  if (!sh) throw new Error("Missing sheet: " + name);
  return sh;
}

function clean_(v) {
  return (v === null || v === undefined) ? "" : String(v).trim();
}

function normalize_(v) {
  if (v === null || v === undefined) return "";
  if (typeof v === "object") return JSON.stringify(v);
  return String(v);
}

function slug_(s) {
  return clean_(s).toLowerCase().replace(/[^a-z0-9]+/g, "_");
}

function log_(sheet, label, msg) {
  sheet.appendRow([new Date(), label, msg || ""]);
}

function payloadSummary_(p) {
  return JSON.stringify({
    First_Name: p.First_Name,
    Last_Name: p.Last_Name,
    Grade: p.Grade_Applying_For,
    Intake: p.Intake_Year || p["Intake Year"] || ""
  });
}

function esc_(s) {
  return String(s === null || s === undefined ? "" : s)
    .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;").replace(/'/g, "&#039;");
}

function shallowCopy_(obj) {
  var out = {};
  for (var k in obj) out[k] = obj[k];
  return out;
}

function isValidEmail_(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(clean_(email));
}

// IMPORTANT: regex literal must be /^https?:\/\//i (no extra escaping)
function isHttpUrl_(s) {
  return /^https?:\/\//i.test(String(s || "").trim());
}

function getPortalEditableFields_() {
  if (CONFIG.PORTAL_EDIT_MODE === "ALL_VISIBLE_EXCEPT_NON_EDIT") {
    var nonEdit = new Set(CONFIG.PORTAL_NON_EDIT_FIELDS || []);
    var exclude = new Set(CONFIG.PORTAL_EDIT_EXCLUDE_FIELDS || []);
    return (CONFIG.PORTAL_VISIBLE_FIELDS || []).filter(function (f) {
      return !nonEdit.has(f) && !exclude.has(f);
    });
  }
  return CONFIG.PORTAL_EDIT_FIELDS || [];
}

/******************** PHASE 0 — SAFETY RAILS ********************/

function assertDriveId_(id, label) {
  if (!id || typeof id !== "string" || !/^[a-zA-Z0-9_-]{20,}$/.test(id)) {
    throw new Error(label + " missing/invalid: [" + id + "]");
  }
}

/**
 * Phase 0 — dependency smoke test (read-only).
 * Run this after every change before doing any other testing.
 */
function test_Smoke() {
  Logger.log("===== SMOKE TEST START =====");
  Logger.log("CONFIG.VERSION = " + (CONFIG.VERSION || "MISSING"));

  // Validate IDs
  assertDriveId_(CONFIG.SHEET_ID, "CONFIG.SHEET_ID");
  assertDriveId_(CONFIG.LOG_SHEET_ID, "CONFIG.LOG_SHEET_ID");
  assertDriveId_(CONFIG.ROOT_FOLDER_ID, "CONFIG.ROOT_FOLDER_ID");

  // Open main spreadsheet + data sheet
  var ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  var dataSh = ss.getSheetByName(CONFIG.DATA_SHEET);
  if (!dataSh) throw new Error("Missing DATA_SHEET tab: " + CONFIG.DATA_SHEET);
  Logger.log("DATA_SHEET OK: " + CONFIG.DATA_SHEET);

  // Open portal log spreadsheet + sheet
  var logSS = SpreadsheetApp.openById(CONFIG.LOG_SHEET_ID);
  var portalLogSh = logSS.getSheetByName(CONFIG.LOG_SHEET_NAME);
  if (!portalLogSh) throw new Error("Missing LOG_SHEET_NAME tab: " + CONFIG.LOG_SHEET_NAME);
  Logger.log("LOG_SHEET_NAME OK: " + CONFIG.LOG_SHEET_NAME);

  // Open root Drive folder
  var root = DriveApp.getFolderById(CONFIG.ROOT_FOLDER_ID);
  Logger.log("ROOT_FOLDER OK: " + root.getName());

  // Optional: Exam sites tab check (non-fatal warning)
  var exam = ss.getSheetByName(CONFIG.EXAM_SITES_SHEET);
  if (!exam) Logger.log("WARN: Missing EXAM_SITES_SHEET tab: " + CONFIG.EXAM_SITES_SHEET);
  else Logger.log("EXAM_SITES_SHEET OK: " + CONFIG.EXAM_SITES_SHEET);

  Logger.log("===== SMOKE TEST PASSED =====");
}

/******************** PHASE 1 — CENTRALIZED SHEET IO ********************/

/**
 * Builds a 1-based column index map from the header row.
 */
function getHeaderIndexMap_(sheet) {
  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) {
    return String(h).trim();
  });
  var map = {};
  headers.forEach(function (h, i) {
    if (h) map[h] = i + 1;
  });
  return map;
}

/**
 * Returns an object mapping header -> value for the given row.
 */
function getRowObject_(sheet, rowIndex) {
  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var values = sheet.getRange(rowIndex, 1, 1, lastCol).getValues()[0];
  var obj = {};
  for (var i = 0; i < headers.length; i++) {
    var key = String(headers[i]).trim();
    if (key) obj[key] = values[i];
  }
  return obj;
}

/**
 * Finds applicant row by ApplicantID first, then Parent_Email.
 * Returns the 1-based row number, or null if not found.
 */
function findApplicantRow_(sheet, id, email) {
  var headerMap = getHeaderIndexMap_(sheet);

  var idHeader = SCHEMA.APPLICANT_ID;
  var emailHeader = SCHEMA.PARENT_EMAIL;

  if (!headerMap[idHeader]) throw new Error("Missing header: " + idHeader);
  if (!headerMap[emailHeader]) throw new Error("Missing header: " + emailHeader);

  var idCol = headerMap[idHeader];
  var emailCol = headerMap[emailHeader];

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;

  var idNorm = String(id || "").trim();
  var emailNorm = String(email || "").toLowerCase().trim();

  if (idNorm) {
    var idVals = sheet.getRange(2, idCol, lastRow - 1, 1).getValues();
    for (var r1 = 0; r1 < idVals.length; r1++) {
      if (String(idVals[r1][0]).trim() === idNorm) return r1 + 2;
    }
  }

  if (emailNorm) {
    var emailVals = sheet.getRange(2, emailCol, lastRow - 1, 1).getValues();
    for (var r2 = 0; r2 < emailVals.length; r2++) {
      if (String(emailVals[r2][0] || "").toLowerCase().trim() === emailNorm) return r2 + 2;
    }
  }

  return null;
}

/**
 * Applies a header->value patch to the given row (writes values).
 */
function applyPatch_(sheet, rowIndex, patchObj) {
  var headerMap = getHeaderIndexMap_(sheet);

  Object.keys(patchObj).forEach(function (header) {
    if (!headerMap[header]) throw new Error("Missing header (patch): " + header);
  });

  Object.keys(patchObj).forEach(function (header) {
    sheet.getRange(rowIndex, headerMap[header]).setValue(patchObj[header]);
  });
}

/******************** PHASE 2 — PORTAL LOGGING (NON-FATAL) ********************/

function appendPortalLog_(eventObj) {
  // Log into the main staging spreadsheet tab for visibility during testing.
  // (The older LOG_SHEET_ID / LOG_SHEET_NAME config is kept for reference but not used here.)
  assertDriveId_(CONFIG.SHEET_ID, "CONFIG.SHEET_ID");
  var ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  var sh = ss.getSheetByName(CONFIG.LOG_SHEET);
  if (!sh) throw new Error("Missing log sheet tab in main sheet: " + CONFIG.LOG_SHEET);

  var row = [
    new Date(),
    CONFIG.VERSION || "",
    eventObj.route || "",
    eventObj.applicantId || "",
    eventObj.email || "",
    eventObj.status || "",
    eventObj.message || ""
  ];
  sh.appendRow(row);
}

function safePortalLog_(eventObj, throwOnFail) {
  try {
    appendPortalLog_(eventObj);
  } catch (err) {
    Logger.log("PORTAL LOG FAILURE (non-fatal): " + err.message);
    if (throwOnFail) throw err;
  }
}

function test_PortalLogWrite() {
  safePortalLog_({
    route: "test_PortalLogWrite",
    applicantId: "TEST",
    email: "test@example.com",
    status: "OK",
    message: "Portal log smoke write"
  }, true);
  Logger.log("Done");
}
