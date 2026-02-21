/************************************************************
ADMIN DOCUMENT VERIFICATION (MINIMAL AUDIT MODEL)
- One admin verifies docs via an admin UI (to be added next)
- We store: per-doc Status + Comment, and a single "last verification" stamp
- Roll-up: Docs_Verified becomes "Yes" only when all REQUIRED docs are VERIFIED
  (Transfer_Certificate is optional by default)
************************************************************/

function respondDiag_(e) {
  return jsonOut_(diagStatus_());
}

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

  var rowIndex = findApplicantRow_(sheet, applicantId, null);
  if (!rowIndex) throw new Error("Applicant not found: " + applicantId);

  var rowObj = getRowObject_(sheet, rowIndex);
  var headerMap = getHeaderIndexMap_(sheet);

  var docMeta = docMetaByField_(fieldName);
  if (!docMeta) throw new Error("Invalid document field: " + fieldName);
  if (!headerMap[docMeta.status]) throw new Error("Status column missing: " + docMeta.status);

  // Build patch
  var patch = {};
  patch[docMeta.status] = newStatus;

  if (docMeta.comment && headerMap[docMeta.comment]) {
    patch[docMeta.comment] = comment || "";
  }

  if (headerMap[SCHEMA.DOC_LAST_VERIFIED_AT]) patch[SCHEMA.DOC_LAST_VERIFIED_AT] = new Date().toISOString();
  if (headerMap[SCHEMA.DOC_LAST_VERIFIED_BY]) patch[SCHEMA.DOC_LAST_VERIFIED_BY] = adminUser;

  // Roll-up Docs_Verified (required docs must be VERIFIED)
  var requiredFields = [];
  for (var d = 0; d < CONFIG.DOC_FIELDS.length; d++) {
    var f = CONFIG.DOC_FIELDS[d].file;
    var isOptional = (CONFIG.OPTIONAL_DOC_FIELDS || []).indexOf(f) >= 0;
    if (!isOptional) requiredFields.push(f);
  }

  var allOk = true;
  for (var r = 0; r < requiredFields.length; r++) {
    var f2 = requiredFields[r];
    var meta2 = docMetaByField_(f2);
    var st = (f2 === fieldName) ? newStatus : clean_(rowObj[meta2.status] || "");
    if (st !== "VERIFIED") { allOk = false; break; }
  }

  if (headerMap[SCHEMA.DOCS_VERIFIED]) {
    patch[SCHEMA.DOCS_VERIFIED] = allOk ? "Yes" : "";
  }

  // Optional roll-up stamps (if you keep these headers)
  if (allOk) {
    if (headerMap[SCHEMA.VERIFIED_BY]) patch[SCHEMA.VERIFIED_BY] = adminUser;
    if (headerMap[SCHEMA.VERIFIED_AT]) patch[SCHEMA.VERIFIED_AT] = new Date().toISOString();
  }

  applyPatch_(sheet, rowIndex, patch);

  log_(logSheet, "ADMIN VERIFY", applicantId + " | " + fieldName + " | " + newStatus);

  return { ok: true, applicantId: applicantId, field: fieldName, status: newStatus, docsVerified: allOk ? "Yes" : "" };
}
