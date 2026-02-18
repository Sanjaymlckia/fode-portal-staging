/************************************************************
ADMIN DOCUMENT VERIFICATION (MINIMAL AUDIT MODEL)
- One admin verifies docs via an admin UI (to be added next)
- We store: per-doc Status + Comment, and a single "last verification" stamp
- Roll-up: Docs_Verified becomes "Yes" only when all REQUIRED docs are VERIFIED
  (Transfer_Certificate is optional by default)
************************************************************/

function adminVerifyDocument(applicantId, fieldName, newStatus, adminUser, comment) {
  applicantId = clean_(applicantId);
  fieldName = clean_(fieldName);
  newStatus = clean_(newStatus);
  adminUser = clean_(adminUser) || "ADMIN";
  comment = clean_(comment);

  if (!applicantId) throw new Error("Missing ApplicantID.");
  if (!fieldName) throw new Error("Missing fieldName.");
  if (!newStatus) throw new Error("Missing newStatus.");

  if (!CONFIG.DOC_STATUS[newStatus]) {
    throw new Error("Invalid status. Allowed: " + Object.keys(CONFIG.DOC_STATUS).join(", "));
  }

  var ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  var sheet = mustGetSheet_(ss, CONFIG.DATA_SHEET);
  var logSheet = mustGetSheet_(ss, CONFIG.LOG_SHEET);

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var idCol = headers.indexOf(CONFIG.APPLICANT_ID_HEADER);
  if (idCol === -1) throw new Error("ApplicantID column missing.");

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) throw new Error("No records found.");

  var data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();

  var rowIndex = -1;
  var record = null;

  for (var i = 0; i < data.length; i++) {
    if (clean_(data[i][idCol]) === applicantId) {
      rowIndex = i + 2;
      record = {};
      for (var c = 0; c < headers.length; c++) record[headers[c]] = data[i][c];
      break;
    }
  }

  if (rowIndex === -1) throw new Error("Applicant not found: " + applicantId);

  var docMeta = docMetaByField_(fieldName);
  if (!docMeta) throw new Error("Invalid document field: " + fieldName);

  if (!hasHeader_(sheet, docMeta.status)) throw new Error("Status column missing: " + docMeta.status);

  // Build updates
  var updates = {};
  updates[docMeta.status] = newStatus;

  if (docMeta.comment && hasHeader_(sheet, docMeta.comment)) {
    updates[docMeta.comment] = comment || "";
  }

  // Minimal audit (last verification only)
  if (hasHeader_(sheet, SCHEMA.DOC_LAST_VERIFIED_AT)) updates[SCHEMA.DOC_LAST_VERIFIED_AT] = new Date().toISOString();
  if (hasHeader_(sheet, SCHEMA.DOC_LAST_VERIFIED_BY)) updates[SCHEMA.DOC_LAST_VERIFIED_BY] = adminUser;

  // Append File_Log
  var line = new Date().toISOString()
    + " | ADMIN_VERIFY | " + fieldName
    + " | status=" + newStatus
    + " | by=" + adminUser
    + " | comment=" + (comment || "-");
  if (hasHeader_(sheet, SCHEMA.FILE_LOG)) {
    updates[SCHEMA.FILE_LOG] = appendLog_(clean_(record[SCHEMA.FILE_LOG] || ""), line);
  }

  // Roll-up Docs_Verified (required docs must be VERIFIED)
  var requiredFields = [];
  for (var d = 0; d < CONFIG.DOCS.length; d++) {
    var f = CONFIG.DOCS[d].field;
    var isOptional = (CONFIG.OPTIONAL_DOC_FIELDS || []).indexOf(f) >= 0;
    if (!isOptional) requiredFields.push(f);
  }

  var allOk = true;
  for (var r = 0; r < requiredFields.length; r++) {
    var f2 = requiredFields[r];
    var meta2 = docMetaByField_(f2);
    var st = (f2 === fieldName) ? newStatus : clean_(record[meta2.status] || "");
    if (st !== "VERIFIED") { allOk = false; break; }
  }

  if (hasHeader_(sheet, SCHEMA.DOCS_VERIFIED)) {
    updates[SCHEMA.DOCS_VERIFIED] = allOk ? "Yes" : "";
  }

  // Optional roll-up stamps (if you keep these headers)
  if (allOk) {
    if (hasHeader_(sheet, SCHEMA.VERIFIED_BY)) updates[SCHEMA.VERIFIED_BY] = adminUser;
    if (hasHeader_(sheet, SCHEMA.VERIFIED_AT)) updates[SCHEMA.VERIFIED_AT] = new Date().toISOString();
  }

  writeBack_(sheet, rowIndex, updates);

  log_(logSheet, "ADMIN VERIFY", applicantId + " | " + fieldName + " | " + newStatus);

  return { ok: true, applicantId: applicantId, field: fieldName, status: newStatus, docsVerified: allOk ? "Yes" : "" };
}
