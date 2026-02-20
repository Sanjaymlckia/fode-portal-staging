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

function newPortalSecret_() {
  // 64 hex chars from two UUIDs; take first 32 for compact URL-safe token.
  var bytes = Utilities.getUuid().replace(/-/g, "") + Utilities.getUuid().replace(/-/g, "");
  return bytes.slice(0, 32);
}

function hashPortalSecret_(secret) {
  var bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(secret || ""),
    Utilities.Charset.UTF_8
  );
  var out = [];
  for (var i = 0; i < bytes.length; i++) {
    var b = (bytes[i] + 256) % 256;
    var h = b.toString(16);
    out.push(h.length === 1 ? "0" + h : h);
  }
  return out.join("");
}

function openPortalSecrets_() {
  assertDriveId_(CONFIG.PORTAL_SECRETS_SHEET_ID, "CONFIG.PORTAL_SECRETS_SHEET_ID");
  var ss = SpreadsheetApp.openById(CONFIG.PORTAL_SECRETS_SHEET_ID);
  var tabName = clean_(CONFIG.PORTAL_SECRETS_TAB || "PortalSecrets");
  var sh = ss.getSheetByName(tabName);
  if (!sh) sh = ss.insertSheet(tabName);
  ensurePortalSecretsHeaders_(sh);
  return sh;
}

function ensurePortalSecretsHeaders_(sheet) {
  var expected = [
    "ApplicantID",
    "Email",
    "Full_Name",
    "Secret_Plain",
    "Secret_Hash",
    "Created_At",
    "Last_Rotated_At",
    "Status"
  ];
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, expected.length).setValues([expected]);
    return;
  }
  var lastCol = Math.max(sheet.getLastColumn(), expected.length);
  var current = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function (v) {
    return clean_(v);
  });
  var changed = false;
  for (var i = 0; i < expected.length; i++) {
    if (current.indexOf(expected[i]) === -1) {
      current.push(expected[i]);
      changed = true;
    }
  }
  if (changed) {
    sheet.getRange(1, 1, 1, current.length).setValues([current]);
  }
}

function findPortalSecretsRowByApplicantId_(sheet, applicantId) {
  var id = clean_(applicantId);
  if (!id) return null;
  var idx = getHeaderIndexMap_(sheet);
  if (!idx.ApplicantID) throw new Error("PortalSecrets missing header: ApplicantID");
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  var vals = sheet.getRange(2, idx.ApplicantID, lastRow - 1, 1).getValues();
  for (var i = 0; i < vals.length; i++) {
    if (clean_(vals[i][0]) === id) return i + 2;
  }
  return null;
}

function readPortalSecretsRecord_(sheet, rowIndex) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var row = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
  var out = {};
  for (var i = 0; i < headers.length; i++) {
    var h = clean_(headers[i]);
    if (h) out[h] = row[i];
  }
  return out;
}

function makePortalSecretPlain_() {
  var p1 = Utilities.getUuid().replace(/-/g, "");
  var p2 = Utilities.getUuid().replace(/-/g, "");
  return (p1 + p2).slice(0, 32);
}

function getOrCreateActivePortalSecret_(applicantId, email, fullName, admissionsSheet, rowNumber, opts) {
  opts = opts || {};
  var dryRun = opts.dryRun === true;
  var id = clean_(applicantId);
  if (!id) throw new Error("Missing ApplicantID for PortalSecrets");

  var secretsSheet = opts.secretsSheet || openPortalSecrets_();
  var idx = getHeaderIndexMap_(secretsSheet);
  var rowIndex = findPortalSecretsRowByApplicantId_(secretsSheet, id);

  if (rowIndex) {
    var rec = readPortalSecretsRecord_(secretsSheet, rowIndex);
    var status = clean_(rec.Status);
    var existingPlain = clean_(rec.Secret_Plain);
    var existingHash = clean_(rec.Secret_Hash);
    var createdAt = rec.Created_At || "";
    if (status === "Active" && existingPlain && existingHash) {
      if (!dryRun && admissionsSheet && rowNumber && !clean_(opts.skipAdmissionsHashWrite || "")) {
        var rowObj = getRowObject_(admissionsSheet, Number(rowNumber));
        var currentHash = clean_(rowObj[SCHEMA.PORTAL_TOKEN_HASH] || "");
        if (!currentHash || opts.forceHashWrite === true) {
          setPortalTokenHashForRow_(admissionsSheet, Number(rowNumber), existingHash);
        }
      }
      return {
        secretPlain: existingPlain,
        secretHash: existingHash,
        createdAt: createdAt,
        existed: true,
        created: false,
        rowIndex: rowIndex
      };
    }
  }

  var generatedPlain = makePortalSecretPlain_();
  var generatedHash = hashPortalSecret_(generatedPlain);
  var now = new Date();
  if (!dryRun) {
    secretsSheet.appendRow([
      id,
      clean_(email),
      clean_(fullName),
      generatedPlain,
      generatedHash,
      now.toISOString(),
      "",
      "Active"
    ]);
    if (admissionsSheet && rowNumber && !clean_(opts.skipAdmissionsHashWrite || "")) {
      setPortalTokenHashForRow_(admissionsSheet, Number(rowNumber), generatedHash);
    }
  }
  return {
    secretPlain: generatedPlain,
    secretHash: generatedHash,
    createdAt: now.toISOString(),
    existed: false,
    created: true,
    rowIndex: null,
    dryRun: dryRun
  };
}

function withSpreadsheetRetry_(fn) {
  var delays = [200, 600, 1400];
  var lastErr = null;
  for (var i = 0; i < delays.length + 1; i++) {
    try {
      return fn();
    } catch (e) {
      lastErr = e;
      var msg = e && e.message ? String(e.message) : String(e);
      var retriable = msg.indexOf("Service Spreadsheets failed") >= 0;
      if (!retriable || i >= delays.length) throw e;
      Utilities.sleep(delays[i]);
    }
  }
  if (lastErr) throw lastErr;
  throw new Error("Spreadsheet retry failed");
}

function buildPortalSecretsIndex_(secretsSheet) {
  var idx = getHeaderIndexMap_(secretsSheet);
  if (!idx.ApplicantID) throw new Error("PortalSecrets missing header: ApplicantID");
  var data = withSpreadsheetRetry_(function () {
    return secretsSheet.getDataRange().getValues();
  });
  var byApplicantId = {};
  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    var applicantId = clean_(row[idx.ApplicantID - 1]);
    if (!applicantId) continue;
    byApplicantId[applicantId] = {
      rowIndex: r + 1,
      status: clean_(idx.Status ? row[idx.Status - 1] : ""),
      secretHash: clean_(idx.Secret_Hash ? row[idx.Secret_Hash - 1] : "")
    };
  }
  return {
    byApplicantId: byApplicantId,
    lastRow: withSpreadsheetRetry_(function () { return secretsSheet.getLastRow(); })
  };
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

function normalizeToUrlList_(cellValue) {
  var raw = clean_(cellValue);
  if (!raw) return [];
  var parts = raw.split(/\r?\n|,/).map(function (s) {
    return clean_(s);
  }).filter(function (s) {
    return !!s;
  });
  var out = [];
  var seen = {};
  for (var i = 0; i < parts.length; i++) {
    var p = parts[i];
    if (!/^https?:\/\//i.test(p)) continue;
    if (seen[p]) continue;
    seen[p] = true;
    out.push(p);
  }
  return out;
}

function appendUrlToCell_(existingValue, newUrl) {
  var url = clean_(newUrl);
  if (!url) return clean_(existingValue);
  var urls = normalizeToUrlList_(existingValue);
  if (urls.indexOf(url) === -1) urls.push(url);
  return urls.join("\n");
}
