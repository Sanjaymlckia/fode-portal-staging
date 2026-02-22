/******************** ADMIN APP (REWORKED FOR HASHED PORTAL TOKENS) ********************/
var ADMIN_DETAIL_SIG = "ADMIN_DETAIL_SIG_20260220_v1";

function renderAdminApp_(e) {
  var email = getActiveUserEmail_();
  if (!isAdmin_(email)) {
    return HtmlService.createHtmlOutput("<h3>Access denied</h3><p>Not authorized.</p>")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  var t = HtmlService.createTemplateFromFile("AdminUI");
  t.BRAND = CONFIG.BRAND || {};
  t.USER_EMAIL = email;
  t.WEBAPP_URL = CONFIG.WEBAPP_URL_ADMIN || CONFIG.WEBAPP_URL;
  t.ADMIN_ROLE = getAdminRole_(email);
  t.IS_SUPER = getAdminRole_(email) === "SUPER";
  t.CAN_OVERRIDE = canOverrideOverall_(email);
  t.PAYMENT_BADGE = "Payment Pending";
  t.PAYMENT_VERIFIED_BOOL = false;
  t.OVERALL_DOC_STATUS = "Pending";
  t.BUILD_VERSION = CONFIG.VERSION;
  t.BUILD_RENDERED_AT = new Date().toISOString();
  t.BUILD_SCRIPT_ID = ScriptApp.getScriptId();
  t.STUDENT_URL_READY = isStudentUrlConfigured_();
  t.STUDENT_URL_WARNING = getStudentUrlWarning_();
  return t.evaluate()
    .setTitle((CONFIG.BRAND && CONFIG.BRAND.name ? CONFIG.BRAND.name : "FODE Admin") + " - Document Verification")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function isAdmin_(email) {
  email = String(email || "").toLowerCase().trim();
  return (CONFIG.ADMIN_EMAILS || []).map(function (e) {
    return String(e).toLowerCase().trim();
  }).indexOf(email) >= 0;
}

function getAdminRole_(email) {
  var e = String(email || "").toLowerCase().trim();
  var roles = CONFIG.ADMIN_ROLES || {};
  var role = String(roles[e] || "").toUpperCase();
  if (role === "SUPER") return "SUPER";
  return "VERIFIER";
}

function requireSuperAdmin_(email) {
  if (getAdminRole_(email) !== "SUPER") {
    throw new Error("Access denied: SUPER admin required");
  }
}

function isStudentUrlConfigured_() {
  var studentBase = clean_(CONFIG.WEBAPP_URL_STUDENT || "");
  if (!studentBase) return false;
  if (/[<>]/.test(studentBase)) return false;
  if (studentBase.indexOf("STUDENT_DEPLOYMENT_ID") >= 0) return false;
  return /^https?:\/\//i.test(studentBase);
}

function getStudentUrlWarning_() {
  return isStudentUrlConfigured_() ? "" : "Student URL not configured. Set CONFIG.WEBAPP_URL_STUDENT to the student deployment /exec URL.";
}

/******************** ADMIN API ********************/

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
    "ApplicantID", "First_Name", "Last_Name",
    "Doc_Verification_Status", "Receipt_Status", "Portal_Access_Status"
  ]);
  var hasParentEmail = !!idx.Parent_Email;
  var hasParentEmailCorrected = !!idx.Parent_Email_Corrected;

  var out = [];
  for (var r = 1; r < values.length; r++) {
    var row = values[r];
    var rid = clean_(row[idx.ApplicantID - 1]);
    var parentEmail = hasParentEmail ? clean_(row[idx.Parent_Email - 1]).toLowerCase() : "";
    var correctedEmail = hasParentEmailCorrected ? clean_(row[idx.Parent_Email_Corrected - 1]).toLowerCase() : "";
    var effectiveEmail = correctedEmail || parentEmail;
    var match = (applicantId && rid === applicantId) || (email && effectiveEmail === email);
    if (!match) continue;
    var statusRow = {
      Birth_ID_Status: idx.Birth_ID_Status ? clean_(row[idx.Birth_ID_Status - 1]) : "",
      Birth_Status: idx.Birth_Status ? clean_(row[idx.Birth_Status - 1]) : "",
      Report_Status: idx.Report_Status ? clean_(row[idx.Report_Status - 1]) : "",
      Photo_Status: idx.Photo_Status ? clean_(row[idx.Photo_Status - 1]) : "",
      Transfer_Status: idx.Transfer_Status ? clean_(row[idx.Transfer_Status - 1]) : "",
      Receipt_Status: idx.Receipt_Status ? clean_(row[idx.Receipt_Status - 1]) : ""
    };
    var paymentBadge = derivePaymentBadge_(statusRow);
    var docStage = computeDocVerificationStatus_(statusRow);

    out.push({
      rowNumber: r + 1,
      applicantId: rid,
      name: (clean_(row[idx.First_Name - 1]) + " " + clean_(row[idx.Last_Name - 1])).trim(),
      email: effectiveEmail,
      docStatus: docStage,
      paymentVerified: paymentBadge === "Verified" ? "Payment Verified" : (paymentBadge === "Rejected" ? "Payment Rejected" : "Payment Pending"),
      portalAccess: clean_(row[idx.Portal_Access_Status - 1]) || "Open"
    });
  }

  return { ok: true, rows: out };
}

function admin_getApplicantDetail(payload) {
  try {
    Logger.log("SIG admin_getApplicantDetail: %s row=%s id=%s", ADMIN_DETAIL_SIG, payload && payload.rowNumber, payload && payload.applicantId);
    var adminEmail = getActiveUserEmail_();
    if (!isAdmin_(adminEmail)) {
      return { ok: false, error: "Access denied" };
    }

    if (!payload) {
      return { ok: false, error: "Missing payload" };
    }

    var rowNumber = Number(payload.rowNumber);
    var applicantId = clean_(payload.applicantId || "");
    if (!rowNumber || rowNumber < 2 || Math.floor(rowNumber) !== rowNumber) {
      return { ok: false, error: "Missing/invalid RowNumber" };
    }

    var sheet = SpreadsheetApp.openById(CONFIG.SHEET_ID).getSheetByName(CONFIG.DATA_SHEET);
    if (!sheet) {
      return { ok: false, error: "Data sheet not found" };
    }

    var lastRow = sheet.getLastRow();
    if (rowNumber > lastRow) {
      return { ok: false, error: "Row out of range: " + rowNumber };
    }

    var lastCol = sheet.getLastColumn();
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var idx = headerIndex_(headers);
    requireHeaders_(idx, [
      "ApplicantID", "First_Name", "Last_Name", "Parent_Email_Corrected",
      "Portal_Access_Status", "Doc_Verification_Status", "Doc_Last_Verified_At", "Doc_Last_Verified_By",
      "PortalTokenIssuedAt",
      "Birth_ID_Passport_File", "Latest_School_Report_File", "Transfer_Certificate_File", "Passport_Photo_File", "Fee_Receipt_File",
      "Birth_ID_Status", "Birth_ID_Comment", "Report_Status", "Report_Comment", "Transfer_Status", "Transfer_Comment",
      "Photo_Status", "Photo_Comment", "Receipt_Status", "Receipt_Comment"
    ]);

    var values = sheet.getRange(rowNumber, 1, 1, lastCol).getValues();
    var displayRow = sheet.getRange(rowNumber, 1, 1, lastCol).getDisplayValues()[0];
    if (!values || !values.length) {
      return { ok: false, error: "Row empty for RowNumber=" + rowNumber };
    }

    var row = values[0];
    var rowApplicantId = clean_(row[idx.ApplicantID - 1]);
    if (!rowApplicantId) {
      if (applicantId) {
        return { ok: false, error: "Row not found for ApplicantID=" + applicantId + " RowNumber=" + rowNumber };
      }
      return { ok: false, error: "Row empty for RowNumber=" + rowNumber };
    }

    var issuedAtRaw = row[idx.PortalTokenIssuedAt - 1];
    var issuedAtDate = issuedAtRaw ? new Date(issuedAtRaw) : null;
    var tokenAgeDays = null;
    if (issuedAtDate && !isNaN(issuedAtDate.getTime())) {
      tokenAgeDays = Math.floor((new Date().getTime() - issuedAtDate.getTime()) / (24 * 60 * 60 * 1000));
    }
    var tokenExpired = tokenAgeDays !== null && tokenAgeDays > Number(CONFIG.PORTAL_TOKEN_MAX_AGE_DAYS || 0);

    var detailObj = {
      _rowNumber: rowNumber,
      ApplicantID: rowApplicantId,
      First_Name: clean_(row[idx.First_Name - 1]),
      Last_Name: clean_(row[idx.Last_Name - 1]),
      Parent_Email_Corrected: clean_(row[idx.Parent_Email_Corrected - 1]),
      Birth_ID_Status: idx.Birth_ID_Status ? clean_(row[idx.Birth_ID_Status - 1]) : "",
      Birth_Status: idx.Birth_Status ? clean_(row[idx.Birth_Status - 1]) : "",
      Report_Status: clean_(row[idx.Report_Status - 1]),
      Photo_Status: clean_(row[idx.Photo_Status - 1]),
      Transfer_Status: clean_(row[idx.Transfer_Status - 1]),
      Receipt_Status: clean_(row[idx.Receipt_Status - 1]),
      Payment_Verified: idx.Payment_Verified ? clean_(row[idx.Payment_Verified - 1]) : "",
      Portal_Access_Status: clean_(row[idx.Portal_Access_Status - 1]) || "Open",
      Doc_Verification_Status: clean_(row[idx.Doc_Verification_Status - 1]) || "Pending",
      Doc_Last_Verified_At: row[idx.Doc_Last_Verified_At - 1],
      Doc_Last_Verified_By: clean_(row[idx.Doc_Last_Verified_By - 1]),
      PortalTokenIssuedAt: issuedAtDate && !isNaN(issuedAtDate.getTime()) ? issuedAtDate.toISOString() : "",
      PortalTokenAgeDays: tokenAgeDays,
      PortalTokenExpired: tokenExpired,
      PortalTokenMaxAgeDays: Number(CONFIG.PORTAL_TOKEN_MAX_AGE_DAYS || 0)
    };

    var map = CONFIG.DOC_FIELDS || [];
    detailObj._docs = map.map(function (m) {
      var url = clean_(displayRow[idx[m.file] - 1]);
      return {
        label: m.label,
        file: m.file,
        statusField: m.status,
        commentField: m.comment,
        required: m.required !== false,
        url: url,
        hasFile: /^https?:\/\//i.test(url),
        status: normalizeDocStatus_(clean_(row[idx[m.status] - 1]) || "Pending"),
        comment: clean_(row[idx[m.comment] - 1])
      };
    });

    detailObj.Parent_Email_Corrected = String(detailObj.Parent_Email_Corrected || "");
    var docStageComputed = computeDocVerificationStatus_(detailObj);
    var paymentBadge = derivePaymentBadge_(detailObj);
    var overallComputed = computeOverallStatus_(detailObj);
    var paymentVerifiedBool = paymentBadge === "Verified";
    var overallStored = idx.Overall_Status ? clean_(row[idx.Overall_Status - 1]) : "";
    var canOverride = canOverrideOverall_(adminEmail);
    var isOverridden = !!(canOverride && overallStored && overallStored !== overallComputed);
    detailObj.Payment_Verified = paymentVerifiedBool ? "Yes" : "";
    detailObj.Payment_Verified_Bool = paymentVerifiedBool;
    detailObj.Payment_Badge = paymentBadge;
    detailObj.Doc_Verification_Status_Computed = docStageComputed;
    detailObj.Overall_Status_Computed = overallComputed;
    detailObj.Portal_Locked_Computed = isPortalLocked_(detailObj);
    detailObj.Portal_Lock_Reason = getPortalLockReason_(detailObj);
    detailObj.Overall_Status_Stored = overallStored;
    detailObj.Overall_IsOverridden = isOverridden;
    detailObj.Overall_OverrideValue = isOverridden ? overallStored : "";
    detailObj.overallComputed = overallComputed;
    detailObj.overallStored = overallStored;
    detailObj.isOverridden = isOverridden;
    detailObj.Doc_Verification_Status = docStageComputed;
    detailObj.Portal_Access_Status = String(detailObj.Portal_Access_Status || "");
    detailObj.Doc_Verification_Status = String(detailObj.Doc_Verification_Status || "Pending");
    detailObj._docs = (detailObj._docs || []).map(function (d) {
      d.url = asStringUrl_(d.url);
      d.hasFile = /^https?:\/\//i.test(d.url);
      return d;
    });

    if (!detailObj) {
      return { ok: false, error: "Failed to build detail object" };
    }

    Logger.log("DOC_URL_SAMPLE: %s", JSON.stringify(detailObj._docs.map(function (d) {
      return { file: d.file, url: d.url, t: typeof d.url };
    })));
    var resultObject = { ok: true, detail: detailObj };
    Logger.log("DETAIL RETURN SHAPE: %s", JSON.stringify(resultObject));
    return resultObject;
  } catch (e) {
    return { ok: false, error: "admin_getApplicantDetail failed: " + (e && e.message ? e.message : String(e)) };
  }
}

function safeJson_(obj) {
  return JSON.stringify(obj, function (key, val) {
    if (val === undefined) return null;
    if (val instanceof Date) return val.toISOString();
    return val;
  });
}

function admin_getApplicantDetail_json(payload) {
  var SIG = "DETAIL_JSON_V1_20260220";
  Logger.log("SIG admin_getApplicantDetail_json: %s row=%s id=%s",
    SIG,
    payload && payload.rowNumber,
    payload && payload.applicantId
  );

  var res = admin_getApplicantDetail(payload);

  var json = JSON.stringify(res, function (k, v) {
    if (v === undefined) return null;
    if (v instanceof Date) return v.toISOString();
    return v;
  });

  Logger.log("SIG admin_getApplicantDetail_json returning length=%s",
    json ? json.length : 0
  );

  return json;
}

/**
 * Alias maintained for UI: "Generate Portal Link"
 */
function admin_generatePortalLink(payload) {
  return admin_resetPortalLink(payload);
}

function admin_resetPortalSecret(payload) {
  return admin_resetPortalLink(payload);
}

/**
 * REWORKED:
 * - Generates a new plain secret (used in the portal link as `s=...`)
 * - Stores ONLY:
 *   - PortalTokenHash
 *   - PortalTokenIssuedAt
 * - Updates Doc_Last_Verified_At / By
 */
function admin_resetPortalLink(payload) {
  var adminEmail = getActiveUserEmail_();
  if (!isAdmin_(adminEmail)) throw new Error("Access denied");
  requireSuperAdmin_(adminEmail);

  payload = payload || {};
  var rowNumber = Number(payload.rowNumber || 0);
  if (!rowNumber || rowNumber < 2) throw new Error("Invalid rowNumber");
  if (!isStudentUrlConfigured_()) throw new Error(getStudentUrlWarning_());

  var sh = openDataSheet_();
  ensureHeadersExist_(sh, ["PortalTokenHash", "PortalTokenIssuedAt", "Portal_Access_Status"]);
  var idx = headerIndex_(sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]);

  // Require hashed-token columns (NOT PortalSecret)
  requireHeaders_(idx, [
    "ApplicantID",
    "PortalTokenHash",
    "PortalTokenIssuedAt",
    "Doc_Last_Verified_At",
    "Doc_Last_Verified_By"
  ]);

  var applicantId = clean_(sh.getRange(rowNumber, idx["ApplicantID"]).getValue());
  if (!applicantId) throw new Error("ApplicantID missing");
  var rowObj = getRowObject_(sh, rowNumber);
  var parentEmail = clean_(rowObj[SCHEMA.PARENT_EMAIL] || "");
  var parentEmailCorrected = clean_(rowObj[SCHEMA.PARENT_EMAIL_CORRECTED] || "");
  var fullName = (clean_(rowObj.First_Name) + " " + clean_(rowObj.Last_Name)).trim();
  var emailForSecret = parentEmailCorrected || parentEmail;

  // Generate new secret and store hash + issue time
  var secret = newPortalSecret_();
  var secretHash = hashPortalSecret_(secret);
  var issuedAt = new Date();
  var patch = {};
  patch["PortalTokenHash"] = secretHash;
  patch["PortalTokenIssuedAt"] = issuedAt;
  patch["Doc_Last_Verified_At"] = new Date();
  patch["Doc_Last_Verified_By"] = adminEmail || "admin";
  applyPatch_(sh, rowNumber, patch);
  syncPortalSecretsActive_(applicantId, emailForSecret, fullName, secret, secretHash);

  var refreshedRow = getRowObject_(sh, rowNumber);
  var issuedAtRef = refreshedRow.PortalTokenIssuedAt ? new Date(refreshedRow.PortalTokenIssuedAt) : issuedAt;
  var tokenAgeDays = 0;
  if (issuedAtRef && !isNaN(issuedAtRef.getTime())) {
    tokenAgeDays = Math.floor((new Date().getTime() - issuedAtRef.getTime()) / (24 * 60 * 60 * 1000));
  }
  var link = buildPortalLink_(applicantId, secret);
  log_(openLogSheet_(), "ADMIN_PORTAL_LINK_RESET",
    "row=" + rowNumber + " applicantId=" + applicantId + " by=" + (adminEmail || "admin"));

  return {
    ok: true,
    link: link,
    portalUrl: link,
    newPortalUrl: link,
    newSecret: secret,
    applicantId: applicantId,
    issuedAt: issuedAtRef && !isNaN(issuedAtRef.getTime()) ? issuedAtRef.toISOString() : issuedAt.toISOString(),
    secretIssuedAt: issuedAtRef && !isNaN(issuedAtRef.getTime()) ? issuedAtRef.toISOString() : issuedAt.toISOString(),
    tokenAgeDays: tokenAgeDays,
    warning: ""
  };

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
  var lastCol = sh.getLastColumn();
  var headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var idx = headerIndex_(headers);
  var cols = resolveStatusCols_(idx);
  var displayRow = sh.getRange(rowNumber, 1, 1, lastCol).getDisplayValues()[0];
  requireHeaders_(idx, ["ApplicantID", "Doc_Verification_Status", "Doc_Last_Verified_At", "Doc_Last_Verified_By", "Portal_Access_Status"]);
  var applicantId = clean_(sh.getRange(rowNumber, idx.ApplicantID).getValue());
  if (!applicantId) throw new Error("Missing ApplicantID in target row.");
  if (!cols.receipt) {
    var missingReceiptDbg = newDebugId_();
    throw new Error("Missing Receipt_Status column mapping. Debug: " + missingReceiptDbg);
  }
  var dbgId = newDebugId_();
  function hasUploadedFileForMapping_(mapping) {
    if (!mapping || !mapping.file) return false;
    if (!idx[mapping.file]) return false;
    var url = clean_(displayRow[idx[mapping.file] - 1]);
    return /^https?:\/\//i.test(url);
  }

  var docMap = CONFIG.DOC_FIELDS || [];
  var prepared = [];
  for (var i = 0; i < docs.length; i++) {
    var d = docs[i] || {};
    var file = clean_(d.file || "");
    var mapping = findDocMapping_(file, d.statusField, d.commentField, docMap);
    if (!mapping) throw new Error("Invalid document mapping.");
    var status = normalizeDocStatus_(d.status);
    if (status === "Verified" && !hasUploadedFileForMapping_(mapping)) {
      throw new Error("Cannot set Verified: " + (mapping.label || mapping.file) + " has no uploaded file.");
    }
    var comment = clean_(d.comment || "");
    prepared.push({
      mapping: mapping,
      status: status,
      comment: comment
    });
  }

  for (var p = 0; p < prepared.length; p++) {
    var item = prepared[p];
    adminVerifyDocument(applicantId, item.mapping.file, toRouteStatusKey_(item.status), adminEmail || "admin", item.comment);
  }

  var refreshedRow = getRowObject_(sh, rowNumber);
  var docStage = computeDocVerificationStatus_(refreshedRow);
  var paymentBadge = derivePaymentBadge_(refreshedRow);
  var paymentVerified = paymentBadge === "Verified";
  var overallComputed = computeOverallStatus_(refreshedRow);
  if (cols.paymentCompat) setCell_(sh, rowNumber, idx, cols.paymentCompat, paymentVerified ? "Yes" : "");
  if (cols.docStage) setCell_(sh, rowNumber, idx, cols.docStage, docStage);
  if (cols.overall) setCell_(sh, rowNumber, idx, cols.overall, overallComputed);
  setCell_(sh, rowNumber, idx, "Doc_Last_Verified_At", new Date());
  setCell_(sh, rowNumber, idx, "Doc_Last_Verified_By", adminEmail || "admin");

  var savedSnapshot = readRowSnapshot_(applicantId);
  var freshDetailRes = admin_getApplicantDetail({ rowNumber: rowNumber, applicantId: applicantId });
  var freshDetail = (freshDetailRes && freshDetailRes.ok === true && freshDetailRes.detail) ? freshDetailRes.detail : null;

  log_(openLogSheet_(), "ADMIN_DOC_UPDATE", "row=" + rowNumber + " by=" + (adminEmail || "admin") + " docStage=" + docStage + " payment=" + paymentBadge + " overall=" + overallComputed);
  log_(openLogSheet_(), "DOC_STATUS_SAVE", JSON.stringify({
    applicantId: applicantId,
    receiptStatus: clean_(refreshedRow.Receipt_Status || ""),
    docStageComputed: docStage,
    overallComputed: overallComputed,
    dbgId: dbgId
  }));
  return {
    ok: true,
    debugId: dbgId,
    docVerificationStatusComputed: docStage,
    paymentBadge: paymentBadge,
    paymentVerified: paymentVerified ? "Yes" : "",
    overallStatusComputed: overallComputed,
    overallStatus: overallComputed,
    detail: freshDetail,
    savedSnapshot: savedSnapshot
  };
}

function admin_setOverallStatus(payload) {
  var adminEmail = getActiveUserEmail_();
  if (!isAdmin_(adminEmail)) throw new Error("Access denied");

  payload = payload || {};
  var rowNumber = Number(payload.rowNumber || 0);
  var requested = clean_(payload.action || "");
  var reason = clean_(payload.reason || "");
  if (!rowNumber || rowNumber < 2) throw new Error("Invalid rowNumber");
  if (["Pending", "Docs_Verified", "Verified", "Rejected", "Fraudulent"].indexOf(requested) === -1) throw new Error("Invalid action");
  if ((requested === "Rejected" || requested === "Fraudulent") && !reason) throw new Error("Reason required");

  var sh = openDataSheet_();
  var idx = headerIndex_(sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]);
  requireHeaders_(idx, ["Doc_Verification_Status", "Portal_Access_Status", "Doc_Last_Verified_At", "Doc_Last_Verified_By"]);
  var cols = resolveStatusCols_(idx);
  var rowObj = getRowObject_(sh, rowNumber);
  var docStage = computeDocVerificationStatus_(rowObj);
  var paymentBadge = derivePaymentBadge_(rowObj);
  var paymentVerified = paymentBadge === "Verified";
  var computed = computeOverallStatus_(rowObj);
  var canOverride = canOverrideOverall_(adminEmail);
  var finalStatus = canOverride ? requested : computed;

  if (canOverride && requested !== computed) {
    logAudit_("OVERRIDE_OVERALL", {
      user: adminEmail,
      rowNumber: rowNumber,
      computed: computed,
      forced: requested
    });
  }

  var patch = {};
  if (cols.paymentCompat) patch[cols.paymentCompat] = paymentVerified ? "Yes" : "";
  if (cols.docStage) patch[cols.docStage] = docStage;
  if (cols.overall) patch[cols.overall] = finalStatus;
  if (finalStatus === "Fraudulent") {
    patch[SCHEMA.PORTAL_ACCESS_STATUS] = "Locked";
  }
  patch[SCHEMA.DOC_LAST_VERIFIED_AT] = new Date();
  patch[SCHEMA.DOC_LAST_VERIFIED_BY] = adminEmail || "admin";
  applyPatch_(sh, rowNumber, patch);

  log_(openLogSheet_(), "ADMIN_OVERALL_STATUS", "row=" + rowNumber + " action=" + finalStatus + " requested=" + requested + " by=" + (adminEmail || "admin") + " reason=" + (reason || "-"));
  return {
    ok: true,
    overallStatus: finalStatus,
    overallStatusComputed: computed,
    overallComputed: computed,
    overallStored: finalStatus,
    docVerificationStatusComputed: docStage,
    paymentBadge: paymentBadge,
    computed: computed,
    overridden: !!(canOverride && requested !== computed),
    overallIsOverridden: !!(canOverride && requested !== computed),
    isOverridden: !!(canOverride && requested !== computed),
    overallOverrideValue: !!(canOverride && requested !== computed) ? finalStatus : "",
    paymentVerified: paymentVerified ? "Yes" : ""
  };
}

function admin_setPortalAccess(payload) {
  var adminEmail = getActiveUserEmail_();
  if (!isAdmin_(adminEmail)) throw new Error("Access denied");
  requireSuperAdmin_(adminEmail);

  payload = payload || {};
  var rowNumber = Number(payload.rowNumber || 0);
  var status = clean_(payload.status || "");
  if (!rowNumber || rowNumber < 2) throw new Error("Invalid rowNumber");
  if (status !== "Open" && status !== "Locked") throw new Error("Invalid status");

  var sh = openDataSheet_();
  var idx = headerIndex_(sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]);
  requireHeaders_(idx, ["Portal_Access_Status", "Doc_Last_Verified_At", "Doc_Last_Verified_By"]);
  var rowObj = getRowObject_(sh, rowNumber);
  var paymentBadge = derivePaymentBadge_(rowObj);
  if (status === "Open" && paymentBadge === "Verified") {
    throw new Error("Cannot unlock after payment verification.");
  }

  setCell_(sh, rowNumber, idx, "Portal_Access_Status", status);
  setCell_(sh, rowNumber, idx, "Doc_Last_Verified_At", new Date());
  setCell_(sh, rowNumber, idx, "Doc_Last_Verified_By", adminEmail || "admin");

  log_(openLogSheet_(), "ADMIN_PORTAL_ACCESS", "row=" + rowNumber + " status=" + status + " by=" + (adminEmail || "admin"));
  var refreshed = getRowObject_(sh, rowNumber);
  var applicantId = clean_(refreshed.ApplicantID || "");
  var detailRes = applicantId ? admin_getApplicantDetail({ rowNumber: rowNumber, applicantId: applicantId }) : null;
  return {
    ok: true,
    paymentBadge: derivePaymentBadge_(refreshed),
    portalAccessStatus: clean_(refreshed.Portal_Access_Status || status),
    detail: (detailRes && detailRes.ok === true) ? detailRes.detail : null
  };
}

function admin_verifyPayment(payload) {
  return admin_setPaymentVerified(payload);
}

function admin_setPaymentVerified(payload) {
  var adminEmail = getActiveUserEmail_();
  if (!isAdmin_(adminEmail)) throw new Error("Access denied");
  requireSuperAdmin_(adminEmail);

  payload = payload || {};
  var rowNumber = Number(payload.rowNumber || 0);
  if (!rowNumber || rowNumber < 2) throw new Error("Invalid rowNumber");
  var note = clean_(payload.comment || payload.reason || "");

  var sh = openDataSheet_();
  var lastCol = sh.getLastColumn();
  var headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var idx = headerIndex_(headers);
  requireHeaders_(idx, ["ApplicantID", "Receipt_Status", "Doc_Last_Verified_At", "Doc_Last_Verified_By"]);
  var cols = resolveStatusCols_(idx);
  if (!cols.receipt) {
    var missingReceiptDbg = newDebugId_();
    throw new Error("Missing Receipt_Status column mapping. Debug: " + missingReceiptDbg);
  }

  var beforeRow = getRowObject_(sh, rowNumber);
  var wasPaymentVerified = derivePaymentBadge_(beforeRow) === "Verified";

  // Legacy endpoint is mapped to receipt verification only.
  setCell_(sh, rowNumber, idx, "Receipt_Status", "Verified");
  if (idx.Receipt_Comment && note) setCell_(sh, rowNumber, idx, "Receipt_Comment", note);

  var refreshedRow = getRowObject_(sh, rowNumber);
  var docStage = computeDocVerificationStatus_(refreshedRow);
  var paymentBadge = derivePaymentBadge_(refreshedRow);
  var paymentVerified = paymentBadge === "Verified";
  var computedOverall = computeOverallStatus_(refreshedRow);
  if (cols.paymentCompat) setCell_(sh, rowNumber, idx, cols.paymentCompat, paymentVerified ? "Yes" : "");
  if (cols.docStage) setCell_(sh, rowNumber, idx, cols.docStage, docStage);
  if (cols.overall) setCell_(sh, rowNumber, idx, cols.overall, computedOverall);
  setCell_(sh, rowNumber, idx, "Doc_Last_Verified_At", new Date());
  setCell_(sh, rowNumber, idx, "Doc_Last_Verified_By", adminEmail || "admin");

  var transitionToYes = !wasPaymentVerified && paymentVerified;
  var crm = { attempted: false, ok: true, debugId: "" };
  if (transitionToYes) {
    crm = crm_syncOnPaymentVerified_(rowNumber, sh, idx);
  }

  log_(openLogSheet_(), "ADMIN_PAYMENT_VERIFIED", "row=" + rowNumber + " by=" + (adminEmail || "admin") + " via=Receipt_Status transitionToYes=" + (transitionToYes ? "1" : "0"));
  var applicantId = clean_(refreshedRow.ApplicantID || "");
  var savedSnapshot = applicantId ? readRowSnapshot_(applicantId) : null;
  var freshDetailRes = applicantId ? admin_getApplicantDetail({ rowNumber: rowNumber, applicantId: applicantId }) : null;
  var freshDetail = (freshDetailRes && freshDetailRes.ok === true && freshDetailRes.detail) ? freshDetailRes.detail : null;
  return {
    ok: true,
    paymentVerified: paymentVerified ? "Yes" : "",
    paymentBadge: paymentBadge,
    docVerificationStatusComputed: docStage,
    overallStatusComputed: computedOverall,
    overallStatus: computedOverall,
    detail: freshDetail,
    savedSnapshot: savedSnapshot,
    crm: crm
  };
}

function crm_syncOnPaymentVerified_(rowNumber, sh, idx) {
  var dbgId = newDebugId_();
  var sheet = sh || openDataSheet_();
  var map = idx || headerIndex_(sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]);
  var rowObj = getRowObject_(sheet, rowNumber);
  var applicantId = clean_(rowObj.ApplicantID || "");
  try {
    var stableFormId = ensureStableFormId_(rowObj, sheet, rowNumber, map);
    if (stableFormId) rowObj.FormID = stableFormId;
    var payload = buildCrmPayloadFromRow_(rowObj);
    payload.formId = clean_(stableFormId || payload.formId || "");

    var token = getZohoToken_();
    var contactRes = upsertZohoContact_(token, payload);
    var dealRes = upsertZohoDeal_(token, payload, payload.folderUrl || "", contactRes.id);

    var patch = {};
    if (map.Contact_ID) patch.Contact_ID = clean_(contactRes.id || "");
    if (map.Deal_ID) patch.Deal_ID = clean_(dealRes.id || "");
    var crmResponseObj = {
      ok: true,
      dbgId: dbgId,
      formId: payload.formId || "",
      contactId: clean_(contactRes.id || ""),
      dealId: clean_(dealRes.id || "")
    };
    if (map.CRM_Response) patch.CRM_Response = JSON.stringify(crmResponseObj);
    if (Object.keys(patch).length) applyPatch_(sheet, rowNumber, patch);

    log_(openLogSheet_(), "ZOHO_OK", JSON.stringify({
      dbgId: dbgId,
      applicantId: applicantId,
      rowNumber: rowNumber,
      contactId: clean_(contactRes.id || ""),
      dealId: clean_(dealRes.id || "")
    }));
    return { attempted: true, ok: true, debugId: dbgId, contactId: clean_(contactRes.id || ""), dealId: clean_(dealRes.id || "") };
  } catch (e) {
    var errMsg = "ERROR: " + String(e && e.message ? e.message : e);
    try {
      if (map.CRM_Response) {
        applyPatch_(sheet, rowNumber, { CRM_Response: errMsg });
      }
    } catch (writeErr) {}
    log_(openLogSheet_(), "ZOHO_ERROR", JSON.stringify({
      dbgId: dbgId,
      applicantId: applicantId,
      rowNumber: rowNumber,
      message: String(e && e.message ? e.message : e)
    }));
    return { attempted: true, ok: false, debugId: dbgId, error: errMsg };
  }
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

function logAudit_(label, payload) {
  log_(openLogSheet_(), clean_(label || "AUDIT"), JSON.stringify(payload || {}));
}

function getCol_(idx, candidates) {
  var names = Array.isArray(candidates) ? candidates : [candidates];
  for (var i = 0; i < names.length; i++) {
    var k = clean_(names[i]);
    if (k && idx[k]) return k;
  }
  return "";
}

function resolveStatusCols_(idx) {
  return {
    docStage: getCol_(idx, ["Doc_Verification_Status"]),
    overall: getCol_(idx, ["Overall_Status"]),
    paymentCompat: getCol_(idx, ["Payment_Verified"]),
    receipt: getCol_(idx, ["Receipt_Status"])
  };
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

function ensureHeadersExist_(sh, headers) {
  var current = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var changed = false;
  for (var i = 0; i < headers.length; i++) {
    if (current.indexOf(headers[i]) === -1) {
      current.push(headers[i]);
      changed = true;
    }
  }
  if (changed) sh.getRange(1, 1, 1, current.length).setValues([current]);
}

function getActiveUserEmail_() {
  try { return clean_(Session.getActiveUser().getEmail()); } catch (e) { return ""; }
}

function buildPortalLink_(applicantId, secret) {
  var base = clean_(CONFIG.WEBAPP_URL_STUDENT || "");
  if (!isStudentUrlConfigured_()) throw new Error(getStudentUrlWarning_());
  return base + "?view=portal&id=" + encodeURIComponent(applicantId) + "&s=" + encodeURIComponent(secret);
}

function buildPortalLinkFromBase_(base, applicantId, secret) {
  return base + "?view=portal&id=" + encodeURIComponent(applicantId) + "&s=" + encodeURIComponent(secret);
}

function buildCsvLine_(cells) {
  return cells.map(function (cell) {
    var v = String(cell === null || cell === undefined ? "" : cell);
    if (/[",\n]/.test(v)) v = '"' + v.replace(/"/g, '""') + '"';
    return v;
  }).join(",");
}

function resolveExportRowNumbers_(payload, lastRow) {
  payload = payload || {};
  var scope = clean_(payload.scope || "");
  if (scope === "auto") scope = "search_first";
  if (scope === "search") scope = "search_only";
  var requested = Array.isArray(payload.currentSearchRowNumbers)
    ? payload.currentSearchRowNumbers
    : (Array.isArray(payload.rowNumbers) ? payload.rowNumbers : []);
  var out = [];
  var emptySearchAllowed = scope === "search_only";
  var startRow = Math.max(2, Number(payload.startRow || 2));
  var batchSize = Math.max(1, Number(payload.batchSize || payload.maxRows || payload.limit || 200));
  var maxRows = Math.max(0, Number(payload.maxRows || payload.batchSize || payload.limit || 0));

  if (scope === "search_only" || scope === "search_first") {
    var seenSearch = {};
    for (var s = 0; s < requested.length; s++) {
      var ns = Number(requested[s] || 0);
      if (!ns || ns < 2 || ns > lastRow || seenSearch[ns]) continue;
      seenSearch[ns] = true;
      out.push(ns);
      if (scope === "search_first" && out.length >= batchSize) break;
    }
    if (out.length || emptySearchAllowed) return out;
    // search_first fallback when no current search results
  }

  if (requested.length && scope !== "all") {
    var seen = {};
    for (var i = 0; i < requested.length; i++) {
      var n = Number(requested[i] || 0);
      if (!n || n < 2 || n > lastRow || seen[n]) continue;
      seen[n] = true;
      out.push(n);
      if (maxRows > 0 && out.length >= maxRows) break;
    }
    return out;
  }
  var endRow = Math.min(lastRow, startRow + batchSize - 1);
  for (var row = startRow; row <= endRow; row++) out.push(row);
  return out;
}

function syncPortalSecretsActive_(applicantId, email, fullName, secretPlain, secretHash) {
  var sh = openPortalSecrets_();
  var idx = getHeaderIndexMap_(sh);
  var rowIndex = findPortalSecretsRowByApplicantId_(sh, applicantId);
  var nowIso = new Date().toISOString();
  var patch = {
    ApplicantID: clean_(applicantId),
    Email: clean_(email),
    Full_Name: clean_(fullName),
    Secret_Plain: clean_(secretPlain),
    Secret_Hash: clean_(secretHash),
    Last_Rotated_At: nowIso,
    Status: "Active"
  };
  if (rowIndex) {
    if (!idx.Created_At || !idx.Last_Rotated_At) throw new Error("PortalSecrets schema missing required headers");
    var existingCreatedAt = sh.getRange(rowIndex, idx.Created_At).getValue();
    if (!clean_(existingCreatedAt)) patch.Created_At = nowIso;
    applyPatch_(sh, rowIndex, patch);
    return rowIndex;
  }
  sh.appendRow([
    clean_(applicantId),
    clean_(email),
    clean_(fullName),
    clean_(secretPlain),
    clean_(secretHash),
    nowIso,
    nowIso,
    "Active"
  ]);
  return sh.getLastRow();
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

function toPlainString_(v) {
  if (v === null || v === undefined) return "";
  if (Array.isArray(v)) {
    return v.filter(function (x) { return !!x; }).map(function (x) {
      return String(x).trim();
    }).filter(function (x) { return !!x; }).join(", ");
  }
  if (typeof v === "string") return v.trim();
  if (typeof v === "number" || typeof v === "boolean") return String(v);
  if (v && typeof v.getUrl === "function") {
    try { return String(v.getUrl() || "").trim(); } catch (e) {}
  }
  return String(v).trim();
}

function asStringUrl_(v) {
  var out = "";
  if (v === null || v === undefined) {
    out = "";
  } else if (Array.isArray(v)) {
    var first = "";
    for (var i = 0; i < v.length; i++) {
      var candidate = String(v[i] === null || v[i] === undefined ? "" : v[i]).trim();
      if (candidate) {
        first = candidate;
        break;
      }
    }
    if (first) out = first;
    else {
      out = v.map(function (x) {
        return String(x === null || x === undefined ? "" : x).trim();
      }).filter(function (x) { return !!x; }).join(", ");
    }
  } else if (typeof v === "string") {
    out = v.trim();
  } else {
    out = String(v).trim();
  }
  if (out.indexOf("[Ljava.lang.Object;") >= 0) return "";
  if (out === "undefined" || out === "null") return "";
  return out;
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

function admin_backfillPortalTokens(payload) {
  var adminEmail = getActiveUserEmail_();
  if (!isAdmin_(adminEmail)) throw new Error("Access denied");
  requireSuperAdmin_(adminEmail);

  payload = payload || {};
  var dryRun = payload.dryRun !== false;
  var startRow = Math.max(2, Number(payload.startRow || 2));
  var limit = Number(payload.limit);
  if (!limit || limit < 1) limit = 200;

  var sh = openDataSheet_();
  ensureHeadersExist_(sh, ["PortalTokenHash", "PortalTokenIssuedAt", "Portal_Access_Status"]);
  var lastCol = withSpreadsheetRetry_(function () { return sh.getLastColumn(); });
  var headers = withSpreadsheetRetry_(function () { return sh.getRange(1, 1, 1, lastCol).getValues()[0]; });
  var idx = headerIndex_(headers);
  requireHeaders_(idx, ["ApplicantID", "PortalTokenHash", "PortalTokenIssuedAt"]);

  var lastRow = withSpreadsheetRetry_(function () { return sh.getLastRow(); });
  if (startRow > lastRow) {
    return {
      ok: true,
      dryRun: dryRun,
      startRow: startRow,
      limit: limit,
      endRow: startRow - 1,
      lastRow: lastRow,
      nextStartRow: "",
      checked: 0,
      updated: 0,
      skipped: 0,
      generatedCount: 0,
      rekeyedCount: 0,
      createdSecretsRows: 0
    };
  }
  var endRow = Math.min(lastRow, startRow + limit - 1);
  var batchSize = endRow - startRow + 1;
  var batchValues = withSpreadsheetRetry_(function () {
    return sh.getRange(startRow, 1, batchSize, lastCol).getValues();
  });

  var secretsSheet = openPortalSecrets_();
  var secretsIndex = buildPortalSecretsIndex_(secretsSheet);

  var checked = 0;
  var updated = 0;
  var skipped = 0;
  var generatedCount = 0;
  var rekeyedCount = 0;
  var createdSecretsRows = 0;
  var admissionsTouched = false;
  var secretsAppendRows = [];
  var now = new Date();
  var nowIso = now.toISOString();

  for (var i = 0; i < batchValues.length; i++) {
    var rowNumber = startRow + i;
    var row = batchValues[i];
    var applicantId = clean_(idx.ApplicantID ? row[idx.ApplicantID - 1] : "");
    if (!applicantId) continue;
    checked++;
    var admissionsHash = clean_(idx.PortalTokenHash ? row[idx.PortalTokenHash - 1] : "");
    var emailCorrected = clean_(idx.Parent_Email_Corrected ? row[idx.Parent_Email_Corrected - 1] : "");
    var emailRaw = clean_(idx.Parent_Email ? row[idx.Parent_Email - 1] : "");
    var emailForSecret = emailCorrected || emailRaw;
    var firstName = clean_(idx.First_Name ? row[idx.First_Name - 1] : "");
    var lastName = clean_(idx.Last_Name ? row[idx.Last_Name - 1] : "");
    var fullName = (firstName + " " + lastName).trim();
    var portalRec = secretsIndex.byApplicantId[applicantId] || null;
    var hasSecretRecord = !!portalRec;
    var hasActiveSecret = !!(portalRec && portalRec.status === "Active" && portalRec.secretHash);

    if (!admissionsHash) {
      var secretPlain1 = newPortalSecret_();
      var secretHash1 = hashPortalSecret_(secretPlain1);
      updated++;
      generatedCount++;
      createdSecretsRows++;
      secretsAppendRows.push([
        applicantId,
        emailForSecret,
        fullName,
        secretPlain1,
        secretHash1,
        nowIso,
        nowIso,
        "Active"
      ]);
      secretsIndex.byApplicantId[applicantId] = {
        rowIndex: (secretsIndex.lastRow || 1) + secretsAppendRows.length,
        status: "Active",
        secretHash: secretHash1
      };
      if (!dryRun) {
        if (idx.PortalTokenHash) row[idx.PortalTokenHash - 1] = secretHash1;
        if (idx.PortalTokenIssuedAt) row[idx.PortalTokenIssuedAt - 1] = now;
        admissionsTouched = true;
      }
      continue;
    }

    if (!hasSecretRecord) {
      var secretPlain2 = newPortalSecret_();
      var secretHash2 = hashPortalSecret_(secretPlain2);
      updated++;
      rekeyedCount++;
      createdSecretsRows++;
      secretsAppendRows.push([
        applicantId,
        emailForSecret,
        fullName,
        secretPlain2,
        secretHash2,
        nowIso,
        nowIso,
        "Active"
      ]);
      secretsIndex.byApplicantId[applicantId] = {
        rowIndex: (secretsIndex.lastRow || 1) + secretsAppendRows.length,
        status: "Active",
        secretHash: secretHash2
      };
      continue;
    }

    if (hasActiveSecret || hasSecretRecord) {
      skipped++;
    } else {
      skipped++;
    }
  }

  if (!dryRun) {
    if (admissionsTouched) {
      withSpreadsheetRetry_(function () {
        sh.getRange(startRow, 1, batchSize, lastCol).setValues(batchValues);
      });
    }
    if (secretsAppendRows.length) {
      var startAppendRow = withSpreadsheetRetry_(function () { return secretsSheet.getLastRow(); }) + 1;
      withSpreadsheetRetry_(function () {
        secretsSheet.getRange(startAppendRow, 1, secretsAppendRows.length, 8).setValues(secretsAppendRows);
      });
    }
  }

  var nextStartRow = endRow < lastRow ? (endRow + 1) : "";
  Logger.log(
    "ADMIN_TOKEN_BACKFILL batch dryRun=%s start=%s end=%s limit=%s lastRow=%s nextStart=%s checked=%s updated=%s skipped=%s generated=%s rekeyed=%s createdSecretsRows=%s",
    dryRun, startRow, endRow, limit, lastRow, nextStartRow, checked, updated, skipped, generatedCount, rekeyedCount, createdSecretsRows
  );
  log_(openLogSheet_(), "ADMIN_TOKEN_BACKFILL",
    "dryRun=" + dryRun
    + " startRow=" + startRow
    + " endRow=" + endRow
    + " limit=" + limit
    + " lastRow=" + lastRow
    + " nextStartRow=" + nextStartRow
    + " checked=" + checked
    + " updated=" + updated
    + " skipped=" + skipped
    + " generated=" + generatedCount
    + " rekeyed=" + rekeyedCount
    + " createdSecretsRows=" + createdSecretsRows
    + " by=" + (adminEmail || "admin"));

  return {
    ok: true,
    dryRun: dryRun,
    startRow: startRow,
    limit: limit,
    endRow: endRow,
    lastRow: lastRow,
    nextStartRow: nextStartRow,
    checked: checked,
    updated: updated,
    skipped: skipped,
    generatedCount: generatedCount,
    rekeyedCount: rekeyedCount,
    createdSecretsRows: createdSecretsRows
  };
}

function admin_backfillPortalTokensDryRun(payload) {
  payload = payload || {};
  payload.dryRun = true;
  return admin_backfillPortalTokens(payload);
}

function admin_backfillPortalTokensApply(payload) {
  payload = payload || {};
  payload.dryRun = false;
  return admin_backfillPortalTokens(payload);
}

function admin_exportPortalLinksCsv(payload) {
  var adminEmail = getActiveUserEmail_();
  if (!isAdmin_(adminEmail)) throw new Error("Access denied");
  requireSuperAdmin_(adminEmail);
  if (!isStudentUrlConfigured_()) throw new Error(getStudentUrlWarning_());

  payload = payload || {};
  var sh = openDataSheet_();
  var secretsSheet = openPortalSecrets_();
  var lastRow = sh.getLastRow();
  var rows = resolveExportRowNumbers_(payload, lastRow);
  var out = [["ApplicantID", "PortalUrl"]];
  var exportedCount = 0;
  var generatedCount = 0;
  var rekeyedCount = 0;

  for (var i = 0; i < rows.length; i++) {
    var rowNumber = rows[i];
    var rowObj = getRowObject_(sh, rowNumber);
    var applicantId = clean_(rowObj.ApplicantID || "");
    if (!applicantId) continue;

    var emailCorrected = clean_(rowObj.Parent_Email_Corrected || "");
    var emailRaw = clean_(rowObj.Parent_Email || "");
    var email = emailCorrected || emailRaw || "";
    var fullName = (clean_(rowObj.First_Name || "") + " " + clean_(rowObj.Last_Name || "")).trim();
    var admissionsHash = clean_(rowObj.PortalTokenHash || "");
    var hasSecretRecord = !!findPortalSecretsRowByApplicantId_(secretsSheet, applicantId);

    var secretInfo = getOrCreateActivePortalSecret_(applicantId, email, fullName, sh, rowNumber, {
      secretsSheet: secretsSheet
    });
    if (secretInfo.created) generatedCount++;
    if (admissionsHash && !hasSecretRecord) rekeyedCount++;

    var link = buildPortalLinkFromBase_(clean_(CONFIG.WEBAPP_URL_STUDENT || ""), applicantId, secretInfo.secretPlain);
    out.push([applicantId, link]);
    exportedCount++;
  }

  var lines = out.map(function (row) { return buildCsvLine_(row); });
  var csv = lines.join("\n");
  var fileStamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "GMT", "yyyyMMdd_HHmmss");
  var filename = "portal-links-" + fileStamp + ".csv";

  log_(openLogSheet_(), "ADMIN_EXPORT_PORTAL_LINKS",
    "exported=" + exportedCount + " generated=" + generatedCount + " rekeyed=" + rekeyedCount + " by=" + (adminEmail || "admin"));

  return {
    ok: true,
    detail: {
      csv: csv,
      filename: filename,
      exportedCount: exportedCount,
      generatedCount: generatedCount,
      rekeyedCount: rekeyedCount
    }
  };
}

function test_AdminAuth() {
  Logger.log("Active user: " + Session.getActiveUser().getEmail());
}

function test_AdminResetPortalLink() {
  var rowNumber = 2;
  var res = admin_resetPortalLink({ rowNumber: rowNumber });
  Logger.log("admin_resetPortalLink.link = " + (res.link || ""));

  var sh = openDataSheet_();
  var idx = headerIndex_(sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]);
  requireHeaders_(idx, ["PortalTokenHash", "PortalTokenIssuedAt"]);
  var tokenHash = clean_(sh.getRange(rowNumber, idx["PortalTokenHash"]).getValue());
  var issuedAt = sh.getRange(rowNumber, idx["PortalTokenIssuedAt"]).getValue();
  if (!tokenHash || !issuedAt) throw new Error("Portal token fields not updated.");
  Logger.log("admin_resetPortalLink token fields updated for row " + rowNumber);
}

function test_BackfillPortalTokens_DryRun() {
  var res = admin_backfillPortalTokens({ dryRun: true, limit: 0 });
  Logger.log("backfill dryRun checked=" + res.checked + " updated=" + res.updated);
}

function audit_NoHardcodedRowDefaults() {
  Logger.log("Run: rg -n \"\\|\\|\\s*17|payload\\.rowNumber\\s*\\|\\|\\s*[0-9]+|selectedRow\\s*=\\s*[0-9]+\" Admin.js AdminUI.html Code.js");
}
