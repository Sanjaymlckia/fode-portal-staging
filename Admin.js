/******************** ADMIN APP ********************/

function renderAdminApp_(e) {
  var email = getActiveUserEmail_();
  if (!isAdmin_(email)) {
    return HtmlService.createHtmlOutput("<h3>Access denied</h3><p>Not authorized.</p>")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  var t = HtmlService.createTemplateFromFile("AdminUI");
  t.BRAND = CONFIG.BRAND || {};
  t.USER_EMAIL = email;
  t.WEBAPP_URL = CONFIG.WEBAPP_URL;
  return t.evaluate()
    .setTitle((CONFIG.BRAND && CONFIG.BRAND.name ? CONFIG.BRAND.name : "FODE Admin") + " - Document Verification")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function isAdmin_(email) {
  email = String(email || "").toLowerCase().trim();
  return (CONFIG.ADMIN_EMAILS || []).map(function(e){
    return String(e).toLowerCase().trim();
  }).indexOf(email) >= 0;
}

function admin_searchApplicants(payload) {
  var adminEmail = getActiveUserEmail_();
  if (!isAdmin_(adminEmail)) throw new Error("Access denied");

  payload = payload || {};
  var applicantId = clean_(payload.applicantId || "");
  var email = clean_(payload.email || "").toLowerCase();

  if (!applicantId && !email) return { ok: true, rows: [] };

  var sh = openDataSheet_();
  var values = sh.getDataRange().getValues();
  if (values.length < 2) return { ok: true, rows: [] };

  var headers = values[0];
  var idx = headerIndex_(headers);
  requireHeaders_(idx, [
    "ApplicantID","First_Name","Last_Name","Parent_Email_Corrected",
    "Doc_Verification_Status","Payment_Verified","Portal_Access_Status"
  ]);

  var out = [];
  for (var r = 1; r < values.length; r++) {
    var row = values[r];
    var rid = clean_(row[idx.ApplicantID - 1]);
    var remail = clean_(row[idx.Parent_Email_Corrected - 1]).toLowerCase();
    var match = (applicantId && rid === applicantId) || (email && remail === email);
    if (!match) continue;

    out.push({
      rowNumber: r + 1,
      applicantId: rid,
      name: (clean_(row[idx.First_Name - 1]) + " " + clean_(row[idx.Last_Name - 1])).trim(),
      docStatus: clean_(row[idx.Doc_Verification_Status - 1]) || "Pending",
      paymentVerified: clean_(row[idx.Payment_Verified - 1]),
      portalAccess: clean_(row[idx.Portal_Access_Status - 1]) || "Open"
    });
  }

  return { ok: true, rows: out };
}

function admin_getApplicantDetail(payload) {
  var adminEmail = getActiveUserEmail_();
  if (!isAdmin_(adminEmail)) throw new Error("Access denied");

  payload = payload || {};
  var rowNumber = Number(payload.rowNumber || 0);
  if (!rowNumber || rowNumber < 2) throw new Error("Invalid rowNumber");

  var sh = openDataSheet_();
  var lastCol = sh.getLastColumn();
  var headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var idx = headerIndex_(headers);
  requireHeaders_(idx, [
    "ApplicantID","First_Name","Last_Name","Parent_Email_Corrected","Payment_Verified",
    "Portal_Access_Status","Doc_Verification_Status","Doc_Last_Verified_At","Doc_Last_Verified_By",
    "Birth_ID_Passport_File","Latest_School_Report_File","Transfer_Certificate_File","Passport_Photo_File","Fee_Receipt_File",
    "Birth_ID_Status","Birth_ID_Comment","Report_Status","Report_Comment","Transfer_Status","Transfer_Comment","Photo_Status","Photo_Comment","Receipt_Status","Receipt_Comment"
  ]);

  var row = sh.getRange(rowNumber, 1, 1, lastCol).getValues()[0];
  var detail = {
    _rowNumber: rowNumber,
    ApplicantID: clean_(row[idx.ApplicantID - 1]),
    First_Name: clean_(row[idx.First_Name - 1]),
    Last_Name: clean_(row[idx.Last_Name - 1]),
    Parent_Email_Corrected: clean_(row[idx.Parent_Email_Corrected - 1]),
    Payment_Verified: clean_(row[idx.Payment_Verified - 1]),
    Portal_Access_Status: clean_(row[idx.Portal_Access_Status - 1]) || "Open",
    Doc_Verification_Status: clean_(row[idx.Doc_Verification_Status - 1]) || "Pending",
    Doc_Last_Verified_At: row[idx.Doc_Last_Verified_At - 1],
    Doc_Last_Verified_By: clean_(row[idx.Doc_Last_Verified_By - 1])
  };

  var map = CONFIG.DOC_FIELDS || [];
  detail._docs = map.map(function (m) {
    var url = clean_(row[idx[m.file] - 1]);
    return {
      label: m.label,
      file: m.file,
      statusField: m.status,
      commentField: m.comment,
      required: m.required !== false,
      url: url,
      hasFile: /^https?:\/\/?/i.test(url),
      status: normalizeDocStatus_(clean_(row[idx[m.status] - 1]) || "Pending"),
      comment: clean_(row[idx[m.comment] - 1])
    };
  });

  return { ok: true, detail: detail };
}

function admin_updateDocStatuses(payload) {
  var adminEmail = getActiveUserEmail_();
  if (!isAdmin_(adminEmail)) throw new Error("Access denied");

  payload = payload || {};
  var rowNumber = Number(payload.rowNumber || 0);
  var docs = payload.docs || [];
  if (!rowNumber || rowNumber < 2) throw new Error("Invalid rowNumber");
  if (!Array.isArray(docs)) throw new Error("Invalid docs payload");

  var sh = openDataSheet_();
  var idx = headerIndex_(sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]);
  requireHeaders_(idx, ["ApplicantID","Doc_Verification_Status","Doc_Last_Verified_At","Doc_Last_Verified_By","Portal_Access_Status"]);
  var applicantId = clean_(sh.getRange(rowNumber, idx.ApplicantID).getValue());
  if (!applicantId) throw new Error("Missing ApplicantID in target row.");

  var docMap = CONFIG.DOC_FIELDS || [];
  for (var i = 0; i < docs.length; i++) {
    var d = docs[i] || {};
    var file = clean_(d.file || "");
    var mapping = findDocMapping_(file, d.statusField, d.commentField, docMap);
    if (!mapping) throw new Error("Invalid document mapping.");
    var status = normalizeDocStatus_(d.status);
    var comment = clean_(d.comment || "");
    adminVerifyDocument(applicantId, mapping.file, toRouteStatusKey_(status), adminEmail || "admin", comment);
  }

  var overall = recomputeOverallDocStatus_(sh, rowNumber, idx, docMap);
  setCell_(sh, rowNumber, idx, "Doc_Verification_Status", overall);
  if (overall === "Fraudulent") setCell_(sh, rowNumber, idx, "Portal_Access_Status", "Locked");
  setCell_(sh, rowNumber, idx, "Doc_Last_Verified_At", new Date());
  setCell_(sh, rowNumber, idx, "Doc_Last_Verified_By", adminEmail || "admin");

  log_(openLogSheet_(), "ADMIN_DOC_UPDATE", "row=" + rowNumber + " by=" + (adminEmail || "admin") + " overall=" + overall);
  return { ok: true, overallStatus: overall };
}

function admin_setOverallStatus(payload) {
  var adminEmail = getActiveUserEmail_();
  if (!isAdmin_(adminEmail)) throw new Error("Access denied");

  payload = payload || {};
  var rowNumber = Number(payload.rowNumber || 0);
  var action = normalizeDocStatus_(payload.action);
  var reason = clean_(payload.reason || "");
  if (!rowNumber || rowNumber < 2) throw new Error("Invalid rowNumber");
  if (["Pending","Verified","Rejected","Fraudulent"].indexOf(action) === -1) throw new Error("Invalid action");
  if ((action === "Rejected" || action === "Fraudulent") && !reason) throw new Error("Reason required");

  var sh = openDataSheet_();
  var idx = headerIndex_(sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]);
  requireHeaders_(idx, ["Doc_Verification_Status","Portal_Access_Status","Doc_Last_Verified_At","Doc_Last_Verified_By"]);
  var patch = {};
  patch[SCHEMA.DOC_VERIFICATION_STATUS] = action;
  if (action === "Fraudulent") {
    patch[SCHEMA.PORTAL_ACCESS_STATUS] = "Locked";
  }
  patch[SCHEMA.DOC_LAST_VERIFIED_AT] = new Date();
  patch[SCHEMA.DOC_LAST_VERIFIED_BY] = adminEmail || "admin";
  applyPatch_(sh, rowNumber, patch);

  log_(openLogSheet_(), "ADMIN_OVERALL_STATUS", "row=" + rowNumber + " action=" + action + " by=" + (adminEmail || "admin") + " reason=" + (reason || "-"));
  return { ok: true };
}

function admin_setPortalAccess(payload) {
  var adminEmail = getActiveUserEmail_();
  if (!isAdmin_(adminEmail)) throw new Error("Access denied");

  payload = payload || {};
  var rowNumber = Number(payload.rowNumber || 0);
  var status = clean_(payload.status || "");
  if (!rowNumber || rowNumber < 2) throw new Error("Invalid rowNumber");
  if (status !== "Open" && status !== "Locked") throw new Error("Invalid status");

  var sh = openDataSheet_();
  var idx = headerIndex_(sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]);
  requireHeaders_(idx, ["Portal_Access_Status","Doc_Last_Verified_At","Doc_Last_Verified_By"]);

  setCell_(sh, rowNumber, idx, "Portal_Access_Status", status);
  setCell_(sh, rowNumber, idx, "Doc_Last_Verified_At", new Date());
  setCell_(sh, rowNumber, idx, "Doc_Last_Verified_By", adminEmail || "admin");

  log_(openLogSheet_(), "ADMIN_PORTAL_ACCESS", "row=" + rowNumber + " status=" + status + " by=" + (adminEmail || "admin"));
  return { ok: true };
}

/******************** ADMIN HELPERS ********************/

function openDataSheet_() {
  var ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  return mustGetSheet_(ss, CONFIG.DATA_SHEET);
}

function openLogSheet_() {
  var ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  return mustGetSheet_(ss, CONFIG.LOG_SHEET);
}

function headerIndex_(headersRow) {
  var out = {};
  for (var i = 0; i < headersRow.length; i++) {
    var h = clean_(headersRow[i]);
    if (h) out[h] = i + 1;
  }
  return out;
}

function requireHeaders_(idx, required) {
  for (var i = 0; i < required.length; i++) {
    if (!idx[required[i]]) throw new Error("Missing required header: " + required[i]);
  }
}

function setCell_(sh, rowNumber, idx, header, value) {
  sh.getRange(rowNumber, idx[header]).setValue(value);
}

function getActiveUserEmail_() {
  try { return clean_(Session.getActiveUser().getEmail()); } catch (e) { return ""; }
}

function findDocMapping_(file, statusField, commentField, docMap) {
  var i;
  if (file) {
    for (i = 0; i < docMap.length; i++) if (docMap[i].file === file) return docMap[i];
  }
  if (statusField && commentField) {
    for (i = 0; i < docMap.length; i++) {
      if (docMap[i].status === statusField && docMap[i].comment === commentField) return docMap[i];
    }
  }
  return null;
}

function normalizeDocStatus_(s) {
  var v = clean_(s).toLowerCase();
  if (v === "verified") return "Verified";
  if (v === "rejected") return "Rejected";
  if (v === "fraudulent") return "Fraudulent";
  return "Pending";
}

function toRouteStatusKey_(status) {
  if (status === "Verified") return "VERIFIED";
  if (status === "Rejected") return "REJECTED";
  if (status === "Fraudulent") return "FRAUDULENT";
  return "PENDING_REVIEW";
}

function recomputeOverallDocStatus_(sh, rowNumber, idx, docMap) {
  var row = sh.getRange(rowNumber, 1, 1, sh.getLastColumn()).getValues()[0];
  var requiredVerified = true;
  var hasFraudulent = false;
  var hasRejected = false;

  for (var i = 0; i < docMap.length; i++) {
    var m = docMap[i];
    var st = normalizeDocStatus_(row[idx[m.status] - 1]);
    if (st === "Fraudulent") hasFraudulent = true;
    if (st === "Rejected") hasRejected = true;
    if (m.required !== false && st !== "Verified") requiredVerified = false;
  }

  if (hasFraudulent) return "Fraudulent";
  if (hasRejected) return "Rejected";
  if (requiredVerified) return "Verified";
  return "Pending";
}

function test_AdminAuth() {
  Logger.log("Active user: " + Session.getActiveUser().getEmail());
}
