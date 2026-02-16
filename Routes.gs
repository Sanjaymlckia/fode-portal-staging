/************************************************************
ADMIN DOCUMENT VERIFICATION
************************************************************/

function adminVerifyDocument(applicantId, fieldName, newStatus, adminUser, comment) {
  applicantId = clean_(applicantId);
  fieldName = clean_(fieldName);
  newStatus = clean_(newStatus);
  adminUser = clean_(adminUser) || "ADMIN";
  comment = clean_(comment);

  var ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  var sheet = mustGetSheet_(ss, CONFIG.DATA_SHEET);
  var logSheet = mustGetSheet_(ss, CONFIG.LOG_SHEET);

  var headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  var idCol = headers.indexOf(CONFIG.APPLICANT_ID_HEADER);
  if (idCol === -1) throw new Error("ApplicantID column missing.");

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) throw new Error("No records found.");

  var data = sheet.getRange(2,1,lastRow-1,sheet.getLastColumn()).getValues();

  var rowIndex = -1;
  var record = null;

  for (var i = 0; i < data.length; i++) {
    if (clean_(data[i][idCol]) === applicantId) {
      rowIndex = i + 2;
      record = {};
      for (var c = 0; c < headers.length; c++) {
        record[headers[c]] = data[i][c];
      }
      break;
    }
  }

  if (rowIndex === -1) throw new Error("Applicant not found.");

  var docMeta = docMetaByField_(fieldName);
  if (!docMeta) throw new Error("Invalid document field.");

  if (!hasHeader_(sheet, docMeta.status)) throw new Error("Status column missing.");

  var updates = {};

  updates[docMeta.status] = newStatus;

  if (hasHeader_(sheet, docMeta.comment)) {
    updates[docMeta.comment] = comment;
  }

  updates["Doc_Last_Verified_At"] = new Date().toISOString();
  updates["Doc_Last_Verified_By"] = adminUser;

  // If fraud suspected → lock portal immediately
  if (newStatus === CONFIG.DOC_STATUS.FRAUD) {
    updates["Payment_Verified"] = "LOCKED_FRAUD";
  }

  var line = new Date().toISOString()
    + " | ADMIN_VERIFY | " + fieldName
    + " | status=" + newStatus
    + " | by=" + adminUser
    + " | comment=" + (comment || "-");

  updates.File_Log = appendLog_(clean_(record.File_Log || ""), line);

  writeBack_(sheet, rowIndex, updates);

  log_(logSheet, "ADMIN VERIFY", applicantId + " | " + fieldName + " | " + newStatus);

  return { ok: true, applicantId: applicantId, field: fieldName, status: newStatus };
}
