/******************** ADMIN APP (REWORKED FOR HASHED PORTAL TOKENS) ********************/
var ADMIN_DETAIL_SIG = "ADMIN_DETAIL_SIG_20260220_v1";

function makeDebugId_() {
  return adminDebugId_();
}

function adminDebugId_() {
  try {
    if (typeof newDebugId_ === "function") return newDebugId_();
  } catch (_e) {}
  return "ADM-" + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd-HHmmss") + "-" + Math.floor(Math.random() * 100000);
}

function ok_(data, debugId) {
  var out = { ok: true, debugId: String(debugId || adminDebugId_()) };
  var src = (data && typeof data === "object") ? data : {};
  for (var k in src) {
    if (Object.prototype.hasOwnProperty.call(src, k)) out[k] = src[k];
  }
  return out;
}

function err_(code, message, debugId, extra) {
  var out = {
    ok: false,
    code: clean_(code || "ERROR"),
    message: clean_(message || "Server returned an error"),
    debugId: String(debugId || adminDebugId_())
  };
  var src = (extra && typeof extra === "object") ? extra : {};
  for (var k in src) {
    if (Object.prototype.hasOwnProperty.call(src, k)) out[k] = src[k];
  }
  if (!out.error) out.error = out.message;
  return out;
}

function parseOverrideFlag_(payload, key) {
  var p = payload || {};
  var v = p[key];
  return v === true || v === 1 || String(v || "").toLowerCase() === "true";
}

function logAdminApiException_(fnName, debugId, e) {
  try {
    logAudit_("ADMIN_API_EXCEPTION", {
      endpoint: String(fnName || ""),
      debugId: String(debugId || adminDebugId_()),
      message: String(e && e.message ? e.message : e),
      stack: String(e && e.stack ? e.stack : "")
    });
  } catch (_logErr) {
    try {
      Logger.log("ADMIN_API_EXCEPTION %s %s", String(debugId || ""), String(e && e.message ? e.message : e));
    } catch (_logErr2) {}
  }
}

function withEnvelope_(fnName, fn) {
  var debugId = makeDebugId_();
  try {
    var out = fn(debugId);
    if (out && typeof out === "object" && typeof out.ok === "boolean") {
      if (!out.debugId) out.debugId = debugId;
      return out;
    }
    if (out && typeof out === "object") return ok_(out, debugId);
    return ok_({ value: out }, debugId);
  } catch (e) {
    logAdminApiException_(fnName, debugId, e);
    return err_("EXCEPTION", String(e && e.message ? e.message : e), debugId);
  }
}

function renderAdminApp_(e) {
  var email = getActiveUserEmail_();
  if (!isAdmin_(email)) {
    var debugId = makeDebugId_();
    var activeUserEmail = "";
    var effectiveUserEmail = "";
    var safeUrl = "";
    try { activeUserEmail = String(Session.getActiveUser().getEmail() || ""); } catch (_au) {}
    try { effectiveUserEmail = String(Session.getEffectiveUser().getEmail() || ""); } catch (_eu) {}
    try {
      safeUrl = String(ScriptApp.getService().getUrl() || "").replace(/[?#].*$/, "");
    } catch (_urlErr) {}
    try {
      logAudit_("ADMIN_ACCESS_DENIED", {
        activeUserEmail: activeUserEmail,
        effectiveUserEmail: effectiveUserEmail,
        url: safeUrl,
        debugId: debugId
      });
    } catch (_logErr) {}
    return HtmlService.createHtmlOutput("<h3>Access denied</h3><p>Not authorized. Debug ID: " + debugId + "</p>")
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
    var rowObj = {};
    for (var c = 0; c < headers.length; c++) {
      var hk = clean_(headers[c]);
      if (hk) rowObj[hk] = row[c];
    }
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
    var docsFollowupSentAt = getDocsFollowupSentAt_(rowObj);
    var eligibleDocsFollowUp = computeEligibleDocsFollowUp_(rowObj, docsFollowupSentAt);

    out.push({
      rowNumber: r + 1,
      applicantId: rid,
      name: (clean_(row[idx.First_Name - 1]) + " " + clean_(row[idx.Last_Name - 1])).trim(),
      email: effectiveEmail,
      docStatus: docStage,
      paymentVerified: paymentBadge === "Verified" ? "Payment Verified" : (paymentBadge === "Rejected" ? "Payment Rejected" : "Payment Pending"),
      portalAccess: clean_(row[idx.Portal_Access_Status - 1]) || "Open",
      eligibleDocsFollowUp: !!eligibleDocsFollowUp,
      docsFollowupSentAt: safeStr_(docsFollowupSentAt || "")
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

    var sheet = openDataSheet_();
    if (!sheet) {
      return { ok: false, code: "DATA_SHEET_NOT_FOUND", debugId: newDebugId_(), error: "Data sheet not found" };
    }

    var lastRow = sheet.getLastRow();
    var rowNumberValid = !!(rowNumber && rowNumber >= 2 && Math.floor(rowNumber) === rowNumber && rowNumber <= lastRow);
    if (!rowNumberValid) {
      if (applicantId) {
        rowNumber = findRowByApplicantId_(sheet, applicantId);
        rowNumberValid = !!(rowNumber && rowNumber >= 2);
      }
    }
    if (!rowNumberValid) {
      var dbgMissing = newDebugId_();
      if (!applicantId) {
        return {
          ok: false,
          code: "MISSING_ROWNUMBER_AND_ID",
          message: "Cannot review record: missing rowNumber and ApplicantID.",
          debugId: dbgMissing,
          error: "Cannot review record: missing rowNumber and ApplicantID. Debug ID: " + dbgMissing
        };
      }
      return {
        ok: false,
        code: "DETAIL_ROW_NOT_FOUND",
        message: "Could not locate applicant record for review.",
        debugId: dbgMissing,
        error: "Could not locate applicant record for review. Debug ID: " + dbgMissing
      };
    }

    var lastCol = sheet.getLastColumn();
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var idx = headerIndex_(headers);
    requireHeaders_(idx, [
      "ApplicantID", "First_Name", "Last_Name", "Parent_Email", "Parent_Email_Corrected",
      "Portal_Access_Status", "Doc_Verification_Status", "Doc_Last_Verified_At", "Doc_Last_Verified_By",
      "PortalTokenIssuedAt",
      "Birth_ID_Passport_File", "Latest_School_Report_File", "Transfer_Certificate_File", "Passport_Photo_File", "Fee_Receipt_File",
      "Birth_ID_Status", "Birth_ID_Comment", "Report_Status", "Report_Comment", "Transfer_Status", "Transfer_Comment",
      "Photo_Status", "Photo_Comment", "Receipt_Status", "Receipt_Comment"
    ]);

    var values = sheet.getRange(rowNumber, 1, 1, lastCol).getValues();
    var displayRow = sheet.getRange(rowNumber, 1, 1, lastCol).getDisplayValues()[0];
    if (!values || !values.length) {
      return { ok: false, code: "ROW_EMPTY", debugId: newDebugId_(), error: "Row empty for RowNumber=" + rowNumber };
    }

    var row = values[0];
    var rowApplicantId = clean_(row[idx.ApplicantID - 1]);
    if (!rowApplicantId) {
      if (applicantId) {
        return { ok: false, error: "Row not found for ApplicantID=" + applicantId + " RowNumber=" + rowNumber };
      }
      return { ok: false, code: "ROW_EMPTY", debugId: newDebugId_(), error: "Row empty for RowNumber=" + rowNumber };
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
      Parent_Email: clean_(row[idx.Parent_Email - 1]),
      Parent_Email_Corrected: clean_(row[idx.Parent_Email_Corrected - 1]),
      Birth_ID_Status: idx.Birth_ID_Status ? clean_(row[idx.Birth_ID_Status - 1]) : "",
      Birth_Status: idx.Birth_Status ? clean_(row[idx.Birth_Status - 1]) : "",
      Report_Status: clean_(row[idx.Report_Status - 1]),
      Photo_Status: clean_(row[idx.Photo_Status - 1]),
      Transfer_Status: clean_(row[idx.Transfer_Status - 1]),
      Receipt_Status: clean_(row[idx.Receipt_Status - 1]),
      Portal_Submitted: idx.Portal_Submitted ? clean_(row[idx.Portal_Submitted - 1]) : "",
      Docs_Verified: idx.Docs_Verified ? clean_(row[idx.Docs_Verified - 1]) : "",
      Payment_Verified: idx.Payment_Verified ? clean_(row[idx.Payment_Verified - 1]) : "",
      Registration_Complete: idx.Registration_Complete ? clean_(row[idx.Registration_Complete - 1]) : "",
      Fee_Receipt_File: idx.Fee_Receipt_File ? clean_(row[idx.Fee_Receipt_File - 1]) : "",
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
      var rawValue = displayRow[idx[m.file] - 1];
      var resolvedUrls = normalizeToUrlList_(rawValue, m.file);
      var url = resolvedUrls.length ? clean_(resolvedUrls[0]) : clean_(rawValue);
      return {
        label: m.label,
        file: m.file,
        statusField: m.status,
        commentField: m.comment,
        required: m.required !== false,
        url: url,
        hasFile: resolvedUrls.length > 0 || /^https?:\/\//i.test(url),
        status: normalizeDocStatus_(clean_(row[idx[m.status] - 1]) || "Pending"),
        comment: clean_(row[idx[m.comment] - 1])
      };
    });

    detailObj.Effective_Email = clean_(detailObj.Parent_Email_Corrected || detailObj.Parent_Email || "");
    detailObj.Parent_Email_Corrected = String(detailObj.Parent_Email_Corrected || "");
    var docStageComputed = computeDocVerificationStatus_(detailObj);
    var paymentBadge = derivePaymentBadge_(detailObj);
    var overallComputed = computeOverallStatus_(detailObj);
    var paymentVerifiedBool = paymentBadge === "Verified";
    var overallStored = idx.Overall_Status ? clean_(row[idx.Overall_Status - 1]) : "";
    var canOverride = canOverrideOverall_(adminEmail);
    var isSuperAdminCaller = canBypassPaymentFreeze_(adminEmail);
    var isOverridden = !!(canOverride && overallStored && overallStored !== overallComputed);
    detailObj.Payment_Received = (nonEmpty_(clean_(detailObj.Fee_Receipt_File || "")) || nonEmpty_(clean_(detailObj.Receipt_Status || ""))) ? "Yes" : "No";
    detailObj.Docs_Verified = clean_(detailObj.Docs_Verified || "") === "Yes" || docStageComputed === "Verified" ? "Yes" : "No";
    detailObj.Portal_Submitted = clean_(detailObj.Portal_Submitted || "") === "Yes" ? "Yes" : "No";
    detailObj.Payment_Verified = paymentVerifiedBool ? "Yes" : "No";
    detailObj.Enrolled_Confirmed = paymentVerifiedBool ? "Yes" : "No";
    detailObj.Payment_Verified_Bool = paymentVerifiedBool;
    detailObj.paymentVerified = paymentVerifiedBool;
    detailObj.isPaymentVerified = paymentVerifiedBool;
    detailObj.isSuperAdmin = !!isSuperAdminCaller;
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
  return admin_getPortalLink(payload);
}

function admin_resetPortalSecret(payload) {
  return admin_resetPortalLink(payload);
}

function admin_getPortalLink(payload) {
  payload = payload || {};
  var debugId = clean_(payload.debugId || "") || adminDebugId_();
  var applicantIdForLog = clean_(payload.applicantId || payload.id || "");
  var callerEmail = "";
  try { callerEmail = String(Session.getActiveUser().getEmail() || ""); } catch (_callerErr) {}
  Logger.log("PORTAL_LINK_START " + JSON.stringify({
    debugId: debugId,
    applicantId: applicantIdForLog,
    caller: callerEmail
  }));
  try {
    var adminEmail = getActiveUserEmail_();
    if (!isAdmin_(adminEmail)) return { ok: false, code: "PORTAL_LINK_ERROR", debugId: debugId, message: "Link generation failed" };
    var rowNumber = Number(payload.rowNumber || 0);
    if (!rowNumber || rowNumber < 2) return { ok: false, code: "PORTAL_LINK_ERROR", debugId: debugId, message: "Link generation failed" };

    var sh = openDataSheet_();
    ensureHeadersExist_(sh, ["ApplicantID"]);
    var rowObj = getRowObject_(sh, rowNumber);
    var applicantId = clean_(rowObj.ApplicantID || "");
    if (!applicantId) return { ok: false, code: "PORTAL_LINK_ERROR", debugId: debugId, message: "Link generation failed" };

    var secretRes = getPortalSecretForApplicant_(applicantId);
    if (!secretRes || secretRes.ok !== true) return { ok: false, code: "PORTAL_LINK_ERROR", debugId: debugId, message: "Link generation failed" };
    var portalUrl = buildStudentPortalUrl_(applicantId, secretRes.secret);
    logAdminEvent_("PORTAL_URL_GENERATED", {
      operatorEmail: adminEmail || "",
      applicantId: applicantId,
      rowNumber: rowNumber,
      debugId: debugId
    });
    return {
      ok: true,
      link: portalUrl,
      portalUrl: portalUrl,
      applicantId: applicantId,
      debugId: debugId
    };
  } catch (e) {
    Logger.log("PORTAL_LINK_THROW " + JSON.stringify({
      debugId: debugId,
      message: String(e),
      stack: String((e && e.stack) || "")
    }));
    return { ok: false, code: "PORTAL_LINK_ERROR", debugId: debugId, message: "Link generation failed" };
  }
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
  payload = payload || {};
  var debugId = clean_(payload.debugId || "") || adminDebugId_();
  var applicantIdForLog = clean_(payload.applicantId || payload.id || "");
  var callerEmail = "";
  try { callerEmail = String(Session.getActiveUser().getEmail() || ""); } catch (_callerErr) {}
  Logger.log("PORTAL_RESET_START " + JSON.stringify({
    debugId: debugId,
    applicantId: applicantIdForLog,
    caller: callerEmail
  }));
  try {
    var adminEmail = getActiveUserEmail_();
    if (!isAdmin_(adminEmail)) return { ok: false, code: "PORTAL_RESET_ERROR", debugId: debugId, message: "Link generation failed" };
    var rowNumber = Number(payload.rowNumber || 0);
    if (!rowNumber || rowNumber < 2) return { ok: false, code: "PORTAL_RESET_ERROR", debugId: debugId, message: "Link generation failed" };

    var sh = openDataSheet_();
    ensureHeadersExist_(sh, ["ApplicantID"]);
    var rowObj = getRowObject_(sh, rowNumber);
    var applicantId = clean_(rowObj.ApplicantID || "");
    if (!applicantId) return { ok: false, code: "PORTAL_RESET_ERROR", debugId: debugId, message: "Link generation failed" };

    var newSecret = Utilities.getUuid();
    var setRes = setPortalSecretForApplicant_(applicantId, newSecret);
    if (!setRes || setRes.ok !== true) return { ok: false, code: "PORTAL_RESET_ERROR", debugId: debugId, message: "Link generation failed" };

    var portalUrl = buildStudentPortalUrl_(applicantId, newSecret);
    logAdminEvent_("PORTAL_URL_RESET", {
      operatorEmail: adminEmail || "",
      applicantId: applicantId,
      rowNumber: rowNumber,
      debugId: debugId
    });
    logAdminEvent_("PORTAL_URL_GENERATED", {
      operatorEmail: adminEmail || "",
      applicantId: applicantId,
      rowNumber: rowNumber,
      debugId: debugId
    });
    return {
      ok: true,
      link: portalUrl,
      portalUrl: portalUrl,
      debugId: debugId
    };
  } catch (e) {
    Logger.log("PORTAL_RESET_THROW " + JSON.stringify({
      debugId: debugId,
      message: String(e),
      stack: String((e && e.stack) || "")
    }));
    return { ok: false, code: "PORTAL_RESET_ERROR", debugId: debugId, message: "Link generation failed" };
  }

}

function admin_updateDocStatuses(payload) {
  return withEnvelope_("admin_updateDocStatuses", function(dbgId) {
    try {
      Logger.log("SAVE_CALLED " + JSON.stringify({
        debugId: String(dbgId || ""),
        keys: payload && typeof payload === "object" ? Object.keys(payload) : null,
        applicantId: payload && payload.applicantId ? String(payload.applicantId) : "",
        rowNumber: payload && payload.rowNumber ? Number(payload.rowNumber) : null,
        actor: String((Session.getActiveUser && Session.getActiveUser().getEmail && Session.getActiveUser().getEmail()) || "")
      }));
    } catch (_logErr) {}
    return admin_updateDocStatuses_impl_(payload, dbgId);
  });
}

function admin_updateDocStatuses_impl_(payload, dbgId) {
  dbgId = String(dbgId || adminDebugId_());
  try {
  var adminEmail = getActiveUserEmail_();
  if (!isAdmin_(adminEmail)) return err_("ACCESS_DENIED", "Access denied", dbgId);

  payload = payload || {};
  var rowNumber = Number(payload.rowNumber || 0);
  var docs = payload.docs || [];
  if (!rowNumber || rowNumber < 2) return err_("VALIDATION", "Invalid rowNumber", dbgId);
  if (!Array.isArray(docs)) return err_("VALIDATION", "Invalid docs payload", dbgId);

  var sh = openDataSheet_();
  var lastCol = sh.getLastColumn();
  var headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var idx = headerIndex_(headers);
  var cols = resolveStatusCols_(idx);
  var overridePaymentBeforeDocs = parseOverrideFlag_(payload, "overridePaymentBeforeDocs");
  var bypassPaymentFreeze = parseOverrideFlag_(payload, "bypassPaymentFreeze") || parseOverrideFlag_(payload, "overridePaymentFreeze");
  var displayRow = sh.getRange(rowNumber, 1, 1, lastCol).getDisplayValues()[0];
  requireHeaders_(idx, ["ApplicantID", "Doc_Verification_Status", "Doc_Last_Verified_At", "Doc_Last_Verified_By", "Portal_Access_Status"]);
  var applicantId = clean_(sh.getRange(rowNumber, idx.ApplicantID).getValue());
  if (!applicantId) return err_("VALIDATION", "Missing ApplicantID in target row.", dbgId);
  if (!cols.receipt) return err_("VALIDATION", "Missing Receipt_Status column mapping.", dbgId);
  var currentRowObj = getRowObject_(sh, rowNumber);
  var priorRowObj = {};
  for (var bk in currentRowObj) {
    if (Object.prototype.hasOwnProperty.call(currentRowObj, bk)) priorRowObj[bk] = currentRowObj[bk];
  }
  var paymentFreezeActive = isPaymentFreezeActive_(currentRowObj);
  var canBypassFreeze = canBypassPaymentFreeze_(adminEmail);
  if (paymentFreezeActive) {
    if (!canBypassFreeze) {
      return err_("PAYMENT_FREEZE", "Payment is verified. Only Super Admin can unlock this record for editing.", dbgId);
    }
    if (!bypassPaymentFreeze) {
      return err_("PAYMENT_FREEZE_REQUIRES_BYPASS", "Payment is verified. Use Unlock to enable Super Admin override before saving.", dbgId);
    }
  }
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

  var wantsReceiptVerified = false;
  var beforePaymentVerified = isPaymentVerifiedDerived_(currentRowObj) === true;
  var prospectiveRow = {};
  for (var key in currentRowObj) {
    if (Object.prototype.hasOwnProperty.call(currentRowObj, key)) prospectiveRow[key] = currentRowObj[key];
  }
  for (var p0 = 0; p0 < prepared.length; p0++) {
    var prep = prepared[p0];
    if (prep.mapping && prep.mapping.status) prospectiveRow[prep.mapping.status] = prep.status;
    if (prep.mapping && prep.mapping.comment) prospectiveRow[prep.mapping.comment] = prep.comment;
    if (prep.mapping && prep.mapping.status === cols.receipt && prep.status === "Verified") wantsReceiptVerified = true;
  }
  if (wantsReceiptVerified) {
    var prospectivePaymentVerified = derivePaymentBadge_(prospectiveRow) === "Verified" || isPaymentVerifiedDerived_(prospectiveRow) === true;
    if (!beforePaymentVerified && prospectivePaymentVerified && !canBypassPaymentFreeze_(adminEmail)) {
      logAdminEvent_("PAYVER_NOT_ALLOWED_BLOCK", {
        applicantId: applicantId,
        rowNumber: rowNumber,
        actor: adminEmail || "",
        dbg: dbgId
      });
      return err_("PAYVER_NOT_ALLOWED", "Only Super Admin can verify payments.", dbgId);
    }
    var prospectiveDocStage = computeDocVerificationStatus_(prospectiveRow);
    var docsVerifiedAfterSave = (clean_(prospectiveRow.Docs_Verified || "") === "Yes") || prospectiveDocStage === "Verified";
    if (!docsVerifiedAfterSave) {
      if (!overridePaymentBeforeDocs) {
        logAudit_("PAYMENT_BEFORE_DOCS_BLOCK", {
          applicantId: applicantId,
          actor: adminEmail || "",
          rowNumber: rowNumber,
          debugId: dbgId
        });
        return err_("PAYMENT_BEFORE_DOCS_REQUIRES_OVERRIDE", "Docs not verified. Confirm override to verify payment.", dbgId);
      }
      logAudit_("PAYMENT_BEFORE_DOCS_OVERRIDE", {
        applicantId: applicantId,
        rowNumber: rowNumber,
        actor: adminEmail || "",
        debugId: dbgId
      });
    }
  }

  for (var p = 0; p < prepared.length; p++) {
    var item = prepared[p];
    adminVerifyDocument(applicantId, item.mapping.file, toRouteStatusKey_(item.status), adminEmail || "admin", item.comment);
  }
  if (paymentFreezeActive && canBypassFreeze && bypassPaymentFreeze) {
    logAudit_("PAYMENT_FREEZE_BYPASS", {
      endpoint: "admin_updateDocStatuses",
      applicantId: applicantId,
      rowNumber: rowNumber,
      actor: adminEmail || "",
      debugId: dbgId,
      changedFields: prepared.map(function(item) {
        return item && item.mapping ? String(item.mapping.status || "") : "";
      }).filter(function(v){ return !!v; })
    });
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
  var finalRowObj = getRowObject_(sh, rowNumber);
  var actions = runVerificationAutomations_(sh, rowNumber, idx, priorRowObj, finalRowObj, dbgId);
  var beforeVerified = (clean_(priorRowObj["Payment_Verified"]) === "Yes");
  var afterVerified = (clean_(finalRowObj["Payment_Verified"]) === "Yes");
  var transition = (!beforeVerified && afterVerified);
  var workflowWarnings = [];
  if (transition && CONFIG.EMAIL_ENABLE_PAYMENT_VERIFIED_TRIGGERS === true) {
    var wf = handlePaymentVerifiedTrigger_(finalRowObj, dbgId);
    actions.payverWorkflow = safeStr_(wf && wf.code || "");
    if (wf && Array.isArray(wf.warnings) && wf.warnings.length) {
      workflowWarnings = workflowWarnings.concat(wf.warnings);
    }
  }

  log_(openLogSheet_(), "ADMIN_DOC_UPDATE", "row=" + rowNumber + " by=" + (adminEmail || "admin") + " docStage=" + docStage + " payment=" + paymentBadge + " overall=" + overallComputed);
  log_(openLogSheet_(), "DOC_STATUS_SAVE", JSON.stringify({
    applicantId: applicantId,
    receiptStatus: clean_(refreshedRow.Receipt_Status || ""),
    docStageComputed: docStage,
    overallComputed: overallComputed,
    dbgId: dbgId
  }));
  return ok_({
    rowNumber: rowNumber,
    applicantId: applicantId,
    changedCount: prepared.length,
    docVerificationStatusComputed: docStage,
    paymentBadge: paymentBadge,
    paymentVerified: paymentVerified ? "Yes" : "",
    overallStatusComputed: overallComputed,
    overallStatus: overallComputed,
    actions: actions,
    emailTriggered: !!(actions && actions.emailTriggered),
    warnings: ((actions && Array.isArray(actions.warnings)) ? actions.warnings : []).concat(workflowWarnings),
    dbg: dbgId
  }, dbgId);
  } catch (e) {
    logAdminApiException_("admin_updateDocStatuses", dbgId, e);
    return err_("EXCEPTION", String(e && e.message ? e.message : e), dbgId);
  }
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
  return withEnvelope_("admin_setPaymentVerified", function(dbgId) {
    return admin_setPaymentVerified_impl_(payload, dbgId);
  });
}

function admin_setPaymentVerified_impl_(payload, dbgId) {
  dbgId = String(dbgId || adminDebugId_());
  try {
  var adminEmail = getActiveUserEmail_();
  if (!isAdmin_(adminEmail)) return err_("ACCESS_DENIED", "Access denied", dbgId);
  try { requireSuperAdmin_(adminEmail); } catch (_superErr) { return err_("ACCESS_DENIED", "Access denied: SUPER admin required", dbgId); }

  payload = payload || {};
  var rowNumber = Number(payload.rowNumber || 0);
  var overridePaymentBeforeDocs = parseOverrideFlag_(payload, "overridePaymentBeforeDocs");
  var bypassPaymentFreeze = parseOverrideFlag_(payload, "bypassPaymentFreeze") || parseOverrideFlag_(payload, "overridePaymentFreeze");
  if (!rowNumber || rowNumber < 2) return err_("VALIDATION", "Invalid rowNumber", dbgId);
  var note = clean_(payload.comment || payload.reason || "");

  var sh = openDataSheet_();
  var lastCol = sh.getLastColumn();
  var headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var idx = headerIndex_(headers);
  requireHeaders_(idx, ["ApplicantID", "Receipt_Status", "Doc_Last_Verified_At", "Doc_Last_Verified_By"]);
  var cols = resolveStatusCols_(idx);
  if (!cols.receipt) return err_("VALIDATION", "Missing Receipt_Status column mapping.", dbgId);

  var beforeRow = getRowObject_(sh, rowNumber);
  var priorRowObj = {};
  for (var bk in beforeRow) {
    if (Object.prototype.hasOwnProperty.call(beforeRow, bk)) priorRowObj[bk] = beforeRow[bk];
  }
  var paymentFreezeActive = isPaymentFreezeActive_(beforeRow);
  if (paymentFreezeActive) {
    if (!bypassPaymentFreeze) {
      return err_("PAYMENT_FREEZE_REQUIRES_BYPASS", "Payment is already verified. Use Unlock to enable Super Admin override before saving.", dbgId);
    }
    logAudit_("PAYMENT_FREEZE_BYPASS", {
      endpoint: "admin_setPaymentVerified",
      applicantId: clean_(beforeRow.ApplicantID || ""),
      rowNumber: rowNumber,
      actor: adminEmail || "",
      debugId: dbgId,
      changedFields: ["Receipt_Status"]
    });
  }
  var docsVerifiedNow = (clean_(beforeRow.Docs_Verified || "") === "Yes") || computeDocVerificationStatus_(beforeRow) === "Verified";
  if (!docsVerifiedNow) {
    if (!overridePaymentBeforeDocs) {
      logAudit_("PAYMENT_BEFORE_DOCS_BLOCK", {
        applicantId: clean_(beforeRow.ApplicantID || ""),
        actor: adminEmail || "",
        rowNumber: rowNumber,
        debugId: dbgId
      });
      return err_("PAYMENT_BEFORE_DOCS_REQUIRES_OVERRIDE", "Docs not verified. Confirm override to verify payment.", dbgId);
    }
    logAudit_("PAYMENT_BEFORE_DOCS_OVERRIDE", {
      applicantId: clean_(beforeRow.ApplicantID || ""),
      rowNumber: rowNumber,
      actor: adminEmail || "",
      debugId: dbgId
    });
  }
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
  var finalRowObj = getRowObject_(sh, rowNumber);
  var actions = runVerificationAutomations_(sh, rowNumber, idx, priorRowObj, finalRowObj, dbgId);

  var beforeVerified = (clean_(priorRowObj["Payment_Verified"]) === "Yes");
  var afterVerified = (clean_(finalRowObj["Payment_Verified"]) === "Yes");
  var transitionToYes = (!beforeVerified && afterVerified);
  var workflowWarnings = [];
  var crm = { attempted: false, ok: true, debugId: "", dryRun: CONFIG.CRM_PUSH_DRY_RUN === true };
  if (transitionToYes && CONFIG.EMAIL_ENABLE_PAYMENT_VERIFIED_TRIGGERS === true) {
    var wf = handlePaymentVerifiedTrigger_(finalRowObj, dbgId);
    actions.payverWorkflow = safeStr_(wf && wf.code || "");
    if (wf && Array.isArray(wf.warnings) && wf.warnings.length) {
      workflowWarnings = workflowWarnings.concat(wf.warnings);
    }
  }

  log_(openLogSheet_(), "ADMIN_PAYMENT_VERIFIED", "row=" + rowNumber + " by=" + (adminEmail || "admin") + " via=Receipt_Status transitionToYes=" + (transitionToYes ? "1" : "0"));
  var applicantId = clean_(refreshedRow.ApplicantID || "");
  return ok_({
    rowNumber: rowNumber,
    applicantId: applicantId,
    paymentVerified: paymentVerified ? "Yes" : "",
    paymentBadge: paymentBadge,
    docVerificationStatusComputed: docStage,
    overallStatusComputed: computedOverall,
    overallStatus: computedOverall,
    actions: actions,
    emailTriggered: !!(actions && actions.emailTriggered),
    warnings: ((actions && Array.isArray(actions.warnings)) ? actions.warnings : []).concat(workflowWarnings),
    dbg: dbgId,
    crm: crm
  }, dbgId);
  } catch (e) {
    logAdminApiException_("admin_setPaymentVerified", dbgId, e);
    return err_("EXCEPTION", String(e && e.message ? e.message : e), dbgId);
  }
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
  return getWorkingSheet_();
}

function openLogSheet_() {
  var ss = getWorkingSpreadsheet_();
  return mustGetSheet_(ss, CONFIG.LOG_SHEET);
}

function logAudit_(label, payload) {
  log_(openLogSheet_(), clean_(label || "AUDIT"), JSON.stringify(payload || {}));
}

function patchIfHeadersPresent_(sh, rowNumber, idx, patchObj) {
  var patch = {};
  var src = patchObj || {};
  for (var k in src) {
    if (!Object.prototype.hasOwnProperty.call(src, k)) continue;
    if (idx && idx[k]) patch[k] = src[k];
  }
  if (!Object.keys(patch).length) return false;
  applyPatch_(sh, rowNumber, patch);
  return true;
}

function joinEmails_(arr) {
  var list = Array.isArray(arr) ? arr : [];
  return list.map(function(v){ return safeStr_(v); }).filter(function(v){ return !!v; }).join(",");
}

function rowStudentName_(row) {
  var r = row || {};
  return (safeStr_(r.First_Name || "") + " " + safeStr_(r.Last_Name || "")).trim();
}

function rowProgramSummary_(row) {
  var r = row || {};
  return safeStr_(r.Program || r.Program_Applied_For || r.Intake || r.Intake_Name || "");
}

function rowSubjectsSummary_(row) {
  var r = row || {};
  return safeStr_(r.Subjects_Selected_Canonical || r.Subjects_Selected || "");
}

function sendQuoteEmail_(rowObj, debugId) {
  var row = rowObj || {};
  var to = pickParentEmail_(row);
  if (!to) {
    logAdminEvent_("QUOTE_EMAIL_MISSING_RECIPIENT", { applicantId: safeStr_(row.ApplicantID), debugId: debugId });
    return { ok: true, status: "skipped", reason: "missing_recipient" };
  }
  var applicantId = safeStr_(row.ApplicantID);
  var studentName = rowStudentName_(row) || "Student";
  var program = rowProgramSummary_(row) || "(program pending)";
  var subjects = rowSubjectsSummary_(row) || "(subjects not listed)";
  var subject = "FODE Application Docs Verified - Next Steps (" + applicantId + ")";
  var body = [
    "Dear Parent/Guardian,",
    "",
    "Your student's FODE application documents have been verified.",
    "",
    "Student: " + studentName,
    "Applicant ID: " + applicantId,
    "Program/Intake: " + program,
    "Subjects: " + subjects,
    "",
    "Next steps:",
    "1. Please proceed with payment.",
    "2. Upload the payment receipt in the student portal.",
    "3. Wait for payment verification and enrollment processing.",
    "",
    "Payment instructions:",
    safeStr_(CONFIG.PAYMENT_INSTRUCTIONS_TEXT || "Please contact the office for payment instructions."),
    "",
    "Support: WhatsApp +675 7860 4013 | Email: mlc@minervacenters.com",
    "",
    "Regards,",
    "Minerva Learning Centers Ltd"
  ].join("\n");
  var cc = safeStr_(CONFIG.EMAIL_ADMIN_ALERTS_TO || "");
  var sent = adminSendEmail_(to, subject, body, { cc: cc });
  if (!sent.ok) {
    logAdminEvent_("QUOTE_EMAIL_FAILED", {
      applicantId: applicantId,
      to: to,
      cc: cc,
      from: safeStr_(sent.from || CONFIG.EMAIL_FROM_ADDRESS || ""),
      replyTo: safeStr_(sent.replyTo || CONFIG.EMAIL_REPLY_TO || ""),
      debugId: debugId,
      error: safeStr_(sent.error || "Quote email failed")
    });
    return { ok: false, status: "failed", code: "EMAIL_SEND_FAILED", message: safeStr_(sent.error || "Quote email failed") };
  }
  logAdminEvent_("QUOTE_EMAIL_SENT", {
    applicantId: applicantId,
    to: to,
    cc: safeStr_(sent.cc || cc),
    from: safeStr_(sent.from || CONFIG.EMAIL_FROM_ADDRESS || ""),
    replyTo: safeStr_(sent.replyTo || CONFIG.EMAIL_REPLY_TO || ""),
    debugId: debugId
  });
  return { ok: true, status: "sent" };
}

function sendPaymentEmail_(rowObj, debugId) {
  var row = rowObj || {};
  var to = pickParentEmail_(row);
  if (!to) {
    logAdminEvent_("PAYMENT_EMAIL_MISSING_RECIPIENT", { applicantId: safeStr_(row.ApplicantID), debugId: debugId });
    return { ok: true, status: "skipped", reason: "missing_recipient" };
  }
  var applicantId = safeStr_(row.ApplicantID);
  var studentName = rowStudentName_(row) || "Student";
  var subject = "FODE Payment Verified - Enrollment Processing (" + applicantId + ")";
  var body = [
    "Dear Parent/Guardian,",
    "",
    "We confirm that payment for the FODE application has been verified.",
    "",
    "Student: " + studentName,
    "Applicant ID: " + applicantId,
    "",
    "Next steps:",
    "Your enrollment and study access will now be processed. We will contact you shortly with the next instructions.",
    "",
    "Support: WhatsApp +675 7860 4013 | Email: mlc@minervacenters.com",
    "",
    "Regards,",
    "Minerva Learning Centers Ltd"
  ].join("\n");
  var cc = joinEmails_(CONFIG.INTERNAL_FINANCE_EMAILS || []);
  var sent = adminSendEmail_(to, subject, body, { cc: cc });
  if (!sent.ok) return { ok: false, status: "failed", code: "EMAIL_SEND_FAILED", message: safeStr_(sent.error || "Payment email failed") };
  logAdminEvent_("PAYMENT_CONFIRM_EMAIL_SENT", { applicantId: applicantId, to: to, cc: cc, debugId: debugId });
  return { ok: true, status: "sent" };
}

function triggerInvoiceWebhook_(rowObj, debugId) {
  var row = rowObj || {};
  var mode = safeStr_(CONFIG.INVOICE_TRIGGER_MODE || "LOG_ONLY") || "LOG_ONLY";
  if (mode === "LOG_ONLY") {
    logAdminEvent_("INVOICE_TRIGGER_LOG_ONLY", { applicantId: safeStr_(row.ApplicantID), debugId: debugId });
    return { ok: true, mode: mode, httpStatus: 0 };
  }
  if (mode !== "WEBHOOK") {
    return { ok: false, code: "INVOICE_TRIGGER_MODE_INVALID", message: "Invalid invoice trigger mode: " + mode };
  }
  var url = safeStr_(CONFIG.INVOICE_WEBHOOK_URL || "");
  if (!url) return { ok: false, code: "INVOICE_WEBHOOK_URL_MISSING", message: "Invoice webhook URL is not configured" };
  var payload = {
    applicantId: safeStr_(row.ApplicantID),
    firstName: safeStr_(row.First_Name),
    lastName: safeStr_(row.Last_Name),
    name: rowStudentName_(row),
    email: pickParentEmail_(row),
    program: rowProgramSummary_(row),
    subjects: rowSubjectsSummary_(row),
    amountK: safeStr_(row.Fee_Total_Kina || row.Total_Fee_Kina || row.Total_Fee || ""),
    debugId: String(debugId || "")
  };
  var res = UrlFetchApp.fetch(url, {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  var code = Number(res && res.getResponseCode ? res.getResponseCode() : 0);
  if (code >= 200 && code < 300) return { ok: true, mode: mode, httpStatus: code };
  return { ok: false, code: "INVOICE_WEBHOOK_HTTP_" + String(code || 0), message: "Invoice webhook failed with HTTP " + String(code || 0), httpStatus: code };
}

function getDocsFollowupSentAt_(rowObj) {
  var row = rowObj || {};
  var applicantId = clean_(row.ApplicantID || row.applicantId || "");
  var key = buildDocsFollowupKey_(CONFIG.DATA_MODE, applicantId);
  try {
    return safeStr_(PropertiesService.getScriptProperties().getProperty(key) || "");
  } catch (_e) {
    return "";
  }
}

function computeEligibleDocsFollowUp_(rowObj, sentAtOpt) {
  if (CONFIG.DOCS_FOLLOWUP_ENABLE !== true) return false;
  var row = rowObj || {};
  var docsVerified = isYes_(row.Docs_Verified) || computeDocVerificationStatus_(row) === "Verified";
  if (!docsVerified) return false;
  if (!getRowEmailForStudent_(row)) return false;
  var sentAt = safeStr_(sentAtOpt || getDocsFollowupSentAt_(row));
  if (sentAt) return false;
  return true;
}

function composeDocsFollowupBody_(rowObj, portalUrl) {
  var row = rowObj || {};
  var applicantId = safeStr_(row.ApplicantID || "");
  var studentName = rowStudentName_(row) || "Student";
  var subjectCount = countSubjectsFromRow_(row);
  var baseK = Number(CONFIG.FODE_FEE_BASE_K || 600);
  var perSubjectK = Number(CONFIG.FODE_FEE_PER_SUBJECT_K || 450);
  var totalK = baseK + (subjectCount * perSubjectK);
  var url = safeStr_(portalUrl || "");
  return [
    "Dear Parent/Guardian,",
    "",
    "Your student's FODE application documents have been verified.",
    "",
    "Student: " + studentName,
    "Applicant ID: " + applicantId,
    "Program/Intake: " + safeStr_(row.Program_Applied_For || row.Program || row.Intake || ""),
    "Subjects: " + safeStr_(row.Subjects_Selected_Canonical || row.Subjects_Selected || ""),
    "",
    "Quote summary:",
    "- Base fee: K" + String(baseK),
    "- Per subject fee: K" + String(perSubjectK),
    "- Subject count: " + String(subjectCount),
    "- Estimated total: K" + String(totalK),
    "",
    "Student Portal Link:",
    url,
    "",
    "Bank Details:",
    "Kundu International Academy",
    "Account Number: 7027138796",
    "BSP BANK, BSP Haus, Konedobu, Port Moresby",
    "BSB No: 088950",
    "",
    "Next steps:",
    "1. Pay the total fee shown in your quote by bank deposit/transfer.",
    "2. In the payment reference, write: Applicant ID + Student Name.",
    "3. After payment, reopen your student portal link above and upload your payment receipt.",
    "4. Once receipt is verified, we will confirm enrolment and release program access.",
    "5. Keep this email and your portal link safe for re-uploads/updates.",
    "",
    "Support: fode@kundu.ac",
    "",
    "Regards,",
    "FODE Admissions"
  ].join("\n");
}

function admin_sendDocsFollowupEmails(payload) {
  return withEnvelope_("admin_sendDocsFollowupEmails", function(dbgId) {
    var adminEmail = getActiveUserEmail_();
    if (!isAdmin_(adminEmail)) return err_("ACCESS_DENIED", "Access denied", dbgId);
    if (CONFIG.DOCS_FOLLOWUP_ENABLE !== true) return ok_({
      summary: { sentCount: 0, failedCount: 0 },
      results: [],
      dbg: dbgId
    }, dbgId);

    payload = payload || {};
    var rowNumbers = Array.isArray(payload.rowNumbers) ? payload.rowNumbers : [];
    var normalized = [];
    var seen = {};
    for (var i = 0; i < rowNumbers.length; i++) {
      var n = Number(rowNumbers[i] || 0);
      if (!Number.isFinite(n) || n < 2) continue;
      n = Math.floor(n);
      if (seen[n]) continue;
      seen[n] = true;
      normalized.push(n);
    }
    if (!normalized.length) {
      return err_("VALIDATION", "rowNumbers is required.", dbgId);
    }

    var sh = getWorkingSheet_();
    var results = [];
    var sentCount = 0;
    var failedCount = 0;
    for (var ri = 0; ri < normalized.length; ri++) {
      var rowNumber = normalized[ri];
      var rowObj = getRowObject_(sh, rowNumber);
      rowObj._rowNumber = rowNumber;
      var applicantId = safeStr_(rowObj.ApplicantID || ("ROW-" + rowNumber));
      var sentAt = getDocsFollowupSentAt_(rowObj);
      var eligible = computeEligibleDocsFollowUp_(rowObj, sentAt);
      var baseAudit = {
        operator: adminEmail || "",
        applicantId: applicantId,
        rowNumber: rowNumber,
        debugId: dbgId
      };
      if (!eligible) {
        logAdminEvent_("DOCS_FOLLOWUP_CLICK", {
          operator: baseAudit.operator,
          applicantId: applicantId,
          rowNumber: rowNumber,
          outcome: "NOT_ELIGIBLE",
          debugId: dbgId
        });
        results.push({ ok: false, code: "NOT_ELIGIBLE", message: "Not eligible for docs follow-up.", applicantId: applicantId, ApplicantID: applicantId, rowNumber: rowNumber });
        failedCount++;
        continue;
      }
      var to = getRowEmailForStudent_(rowObj);
      if (!to) {
        logAdminEvent_("DOCS_FOLLOWUP_CLICK", {
          operator: baseAudit.operator,
          applicantId: applicantId,
          rowNumber: rowNumber,
          outcome: "NO_PARENT_EMAIL",
          debugId: dbgId
        });
        results.push({ ok: false, code: "NO_PARENT_EMAIL", message: "Parent/guardian email is missing or invalid.", applicantId: applicantId, ApplicantID: applicantId, rowNumber: rowNumber });
        failedCount++;
        continue;
      }

      var secretRes = getPortalSecretForApplicant_(applicantId);
      if (!secretRes || secretRes.ok !== true) {
        logAdminEvent_("DOCS_FOLLOWUP_CLICK", {
          operator: baseAudit.operator,
          applicantId: applicantId,
          rowNumber: rowNumber,
          outcome: "PORTAL_LINK_ERROR",
          debugId: dbgId
        });
        results.push({ ok: false, code: "PORTAL_LINK_ERROR", message: "Portal link error.", applicantId: applicantId, ApplicantID: applicantId, rowNumber: rowNumber });
        failedCount++;
        continue;
      }
      var portalUrl = buildStudentPortalUrl_(applicantId, secretRes.secret);
      var subject = safeStr_(CONFIG.DOCS_FOLLOWUP_EMAIL_SUBJECT || "FODE Application - Documents Verified | Quote, Payment Instructions & Next Steps");
      var body = composeDocsFollowupBody_(rowObj, portalUrl);
      var sendOpts = {
        cc: safeStr_(CONFIG.DOCS_FOLLOWUP_CC || ""),
        replyTo: safeStr_(CONFIG.DOCS_FOLLOWUP_REPLY_TO || CONFIG.EMAIL_REPLY_TO || ""),
        name: safeStr_(CONFIG.EMAIL_FROM_NAME || "FODE Admissions"),
        senderMode: safeStr_(CONFIG.DOCS_FOLLOWUP_SENDER_MODE || CONFIG.EMAIL_SENDER_MODE || "DEFAULT")
      };
      var sent = adminSendEmail_(to, subject, body, sendOpts);
      if (!sent.ok) {
        logAdminEvent_("DOCS_FOLLOWUP_CLICK", {
          operator: baseAudit.operator,
          applicantId: applicantId,
          rowNumber: rowNumber,
          outcome: "EMAIL_SEND_FAILED",
          error: safeStr_(sent.error || ""),
          debugId: dbgId
        });
        results.push({ ok: false, code: "EMAIL_SEND_FAILED", message: safeStr_(sent.error || "Email send failed"), applicantId: applicantId, ApplicantID: applicantId, rowNumber: rowNumber });
        failedCount++;
        continue;
      }

      var key = buildDocsFollowupKey_(CONFIG.DATA_MODE, applicantId);
      var ts = nowIso_();
      PropertiesService.getScriptProperties().setProperty(key, ts);
      logAdminEvent_("DOCS_FOLLOWUP_CLICK", {
        operator: baseAudit.operator,
        applicantId: applicantId,
        rowNumber: rowNumber,
        outcome: "SENT",
        debugId: dbgId
      });
      logAdminEvent_("DOCS_FOLLOWUP_EMAIL_SENT", {
        operator: baseAudit.operator,
        applicantId: applicantId,
        rowNumber: rowNumber,
        to: to,
        portalUrl: portalUrl,
        cc: safeStr_(CONFIG.DOCS_FOLLOWUP_CC || ""),
        replyTo: safeStr_(CONFIG.DOCS_FOLLOWUP_REPLY_TO || CONFIG.EMAIL_REPLY_TO || ""),
        sentKey: key,
        sentAt: ts,
        debugId: dbgId
      });
      results.push({ ok: true, code: "SENT", message: "Docs follow-up sent.", applicantId: applicantId, ApplicantID: applicantId, rowNumber: rowNumber, sentAt: ts });
      sentCount++;
    }

    return ok_({
      summary: { sentCount: sentCount, failedCount: failedCount },
      results: results,
      dbg: dbgId
    }, dbgId);
  });
}

function admin_updateParentEmailCorrected(payload) {
  return withEnvelope_("admin_updateParentEmailCorrected", function (dbgId) {
    var operatorEmail = getActiveUserEmail_();
    if (!isAdmin_(operatorEmail)) return err_("ACCESS_DENIED", "Access denied", dbgId);
    requireSuperAdmin_(operatorEmail);
    if (!(CONFIG && CONFIG.SUPERADMIN_ALLOW_EMAIL_OVERRIDE_POST_DOCS_VERIFIED === true)) {
      return err_("FEATURE_DISABLED", "Email override is disabled by config.", dbgId);
    }

    payload = payload || {};
    var rowNumber = Number(payload.rowNumber || 0);
    var applicantIdInput = clean_(payload.applicantId || "");
    var newEmail = clean_(payload.newEmail || "").toLowerCase();
    var reason = clean_(payload.reason || "");
    if (!rowNumber || rowNumber < 2) return err_("VALIDATION", "rowNumber is required.", dbgId);
    if (!applicantIdInput) return err_("VALIDATION", "applicantId is required.", dbgId);
    if (!newEmail) return err_("VALIDATION", "newEmail is required.", dbgId);
    if (!reason) return err_("VALIDATION", "reason is required.", dbgId);
    var emailOk = false;
    if (typeof isValidEmail_ === "function") emailOk = !!isValidEmail_(newEmail);
    else emailOk = /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(newEmail);
    if (!emailOk) return err_("VALIDATION", "Invalid email format.", dbgId);

    var sh = openDataSheet_();
    var idx = headerIndex_(sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]);
    requireHeaders_(idx, ["ApplicantID", "Parent_Email", "Parent_Email_Corrected", "Docs_Verified", "Payment_Verified"]);
    var rowObj = getRowObject_(sh, rowNumber);
    var applicantId = clean_(rowObj.ApplicantID || "");
    if (!applicantId) return err_("ROW_NOT_FOUND", "Applicant row not found.", dbgId);
    if (applicantId !== applicantIdInput) {
      return err_("VALIDATION", "ApplicantID mismatch for row.", dbgId);
    }

    var beforeEmail = clean_(rowObj.Parent_Email_Corrected || rowObj.Parent_Email || "");
    var beforeDocsVerified = clean_(rowObj.Docs_Verified || "");
    var beforePaymentVerified = clean_(rowObj.Payment_Verified || "");
    applyPatch_(sh, rowNumber, { Parent_Email_Corrected: newEmail });

    var deletedKey = deleteDocsFollowupKey_(CONFIG.DATA_MODE, applicantId);
    logAdminEvent_("EMAIL_OVERRIDE_SUPERADMIN", {
      operatorEmail: operatorEmail || "",
      applicantId: applicantId,
      rowNumber: rowNumber,
      beforeEmail: beforeEmail,
      afterEmail: newEmail,
      beforeDocsVerified: beforeDocsVerified,
      beforePaymentVerified: beforePaymentVerified,
      reason: reason,
      debugId: dbgId
    });
    logAdminEvent_("DOCS_FOLLOWUP_RESET_EMAIL_OVERRIDE", {
      operatorEmail: operatorEmail || "",
      applicantId: applicantId,
      rowNumber: rowNumber,
      key: deletedKey,
      debugId: dbgId
    });

    return ok_({
      applicantId: applicantId,
      newEmail: newEmail,
      docsFollowupReset: true,
      dbg: dbgId
    }, dbgId);
  });
}

function countSubjectsFromRow_(rowObj) {
  var row = rowObj || {};
  var csv = safeStr_(row.Subjects_Selected_Canonical || row.Subjects_Selected || "");
  if (!csv) return 0;
  return csv.split(",").map(function(v){ return safeStr_(v); }).filter(function(v){ return !!v; }).length;
}

function buildPaymentVerifiedEmailOptions_() {
  var opts = {};
  var senderMode = safeStr_(CONFIG.EMAIL_SENDER_MODE || "DEFAULT").toUpperCase();
  if (senderMode === "ALIAS" && safeStr_(CONFIG.EMAIL_FROM_ADDRESS || "")) {
    opts.from = safeStr_(CONFIG.EMAIL_FROM_ADDRESS || "");
  }
  if (safeStr_(CONFIG.EMAIL_REPLY_TO || "")) opts.replyTo = safeStr_(CONFIG.EMAIL_REPLY_TO || "");
  return opts;
}

function sendPaymentVerifiedStudentQuoteEmail_(rowObj, debugId) {
  var row = rowObj || {};
  var to = getRowEmailForStudent_(row);
  if (!to) {
    logAdminEvent_("EMAIL_STUDENT_SKIPPED_NO_EMAIL", {
      applicantId: safeStr_(row.ApplicantID || ""),
      debugId: debugId
    });
    return { ok: true, status: "skipped", warning: "Student/parent email missing or invalid" };
  }
  var applicantId = safeStr_(row.ApplicantID || "");
  var studentName = rowStudentName_(row) || "Student";
  var subjectCount = countSubjectsFromRow_(row);
  var baseK = Number(CONFIG.FODE_FEE_BASE_K || 600);
  var perSubjectK = Number(CONFIG.FODE_FEE_PER_SUBJECT_K || 450);
  var totalK = baseK + (perSubjectK * subjectCount);
  var subject = safeStr_(CONFIG.EMAIL_STUDENT_SUBJECT_PAYMENT_VERIFIED || "FODE Application - Payment Verified | Next Steps & Bank Details");
  var body = [
    "Dear Parent/Guardian,",
    "",
    "Payment for the FODE application has been verified.",
    "",
    "Student: " + studentName,
    "Applicant ID: " + applicantId,
    "Subjects: " + safeStr_(row.Subjects_Selected_Canonical || row.Subjects_Selected || "(not listed)"),
    "",
    "Fee summary:",
    "- Base fee: K" + String(baseK),
    "- Per subject: K" + String(perSubjectK),
    "- Subject count: " + String(subjectCount),
    "- Estimated total: K" + String(totalK),
    "",
    String(CONFIG.FODE_BANK_DETAILS_TEXT || ""),
    "",
    String(CONFIG.FODE_NEXT_STEPS_TEXT || ""),
    "",
    "For assistance, please reply to " + safeStr_(CONFIG.EMAIL_REPLY_TO || "fode@kundu.ac") + ".",
    "",
    "Regards,",
    "FODE Admissions"
  ].join("\n");
  var sendOpts = buildPaymentVerifiedEmailOptions_();
  var sent = adminSendEmail_(to, subject, body, sendOpts);
  if (!sent.ok) {
    return { ok: false, code: "EMAIL_SEND_FAILED", message: safeStr_(sent.error || "Failed to send student payment verified email") };
  }
  return { ok: true, status: "sent", to: to };
}

function sendPaymentVerifiedAdminReleaseEmail_(rowObj, debugId) {
  var row = rowObj || {};
  var adminTo = parseCsvEmails_(CONFIG.EMAIL_RELEASE_ADMIN_TO || "");
  if (!adminTo) {
    return { ok: false, code: "EMAIL_CONFIG_MISSING", message: "EMAIL_RELEASE_ADMIN_TO is empty or invalid" };
  }
  var applicantId = safeStr_(row.ApplicantID || "");
  var subjectTpl = safeStr_(CONFIG.EMAIL_ADMIN_SUBJECT_PAYMENT_VERIFIED || "FODE - Payment Verified | Release Access for ApplicantID");
  var subject = subjectTpl.indexOf("ApplicantID") >= 0 ? subjectTpl.replace("ApplicantID", applicantId || "UNKNOWN") : (subjectTpl + " " + applicantId);
  var body = [
    "Payment Verified - Release Access Required",
    "",
    "Applicant ID: " + applicantId,
    "Student Name: " + (rowStudentName_(row) || "(unknown)"),
    "Program/Grade: " + safeStr_(row.Program_Applied_For || row.Grade_Applying_For || row.Program || ""),
    "Subjects: " + safeStr_(row.Subjects_Selected_Canonical || row.Subjects_Selected || ""),
    "Parent Name: " + safeStr_(row.Parent_Name || ""),
    "Parent Phone: " + safeStr_(row.Parent_Phone || row.Phone_Number || ""),
    "Parent Email: " + getRowEmailForStudent_(row),
    "",
    "Action: Release access / enable program / add to LMS / allow next stage.",
    "Debug ID: " + String(debugId || "")
  ].join("\n");
  var sendOpts = buildPaymentVerifiedEmailOptions_();
  var sent = adminSendEmail_(adminTo, subject, body, sendOpts);
  if (!sent.ok) {
    return { ok: false, code: "EMAIL_SEND_FAILED", message: safeStr_(sent.error || "Failed to send admin release email") };
  }
  return { ok: true, status: "sent", to: adminTo };
}

function handlePaymentVerifiedEmailTriggers_(rowObj, debugId) {
  var row = rowObj || {};
  var warnings = [];
  if (CONFIG.EMAIL_ENABLE_PAYMENT_VERIFIED_TRIGGERS !== true) {
    return { ok: true, status: "disabled", warnings: warnings };
  }
  var applicantId = safeStr_(row.ApplicantID || "");
  var mode = safeStr_(CONFIG.DATA_MODE || "UNKNOWN") || "UNKNOWN";
  var key = "PAYVER_SENT::" + mode + "::" + (applicantId || ("ROW-" + safeStr_(row._rowNumber || "")));
  var props = PropertiesService.getScriptProperties();
  var already = safeStr_(props.getProperty(key) || "");
  if (already) {
    warnings.push("Payment verified email already sent");
    logAdminEvent_("PAYVER_EMAIL_SKIPPED_ALREADY_SENT", { applicantId: applicantId, sentKey: key, debugId: debugId });
    return { ok: true, status: "skipped", warnings: warnings, sentKey: key, alreadySentAt: already };
  }

  var studentRes = sendPaymentVerifiedStudentQuoteEmail_(row, debugId);
  if (!studentRes.ok) {
    logAdminEvent_("PAYMENT_EMAIL_SEND_FAILED", { applicantId: applicantId, debugId: debugId, stage: "student", error: studentRes.message || "" });
    return { ok: false, status: "failed", code: studentRes.code || "EMAIL_SEND_FAILED", message: studentRes.message || "Student email failed", warnings: warnings };
  }
  if (studentRes.warning) warnings.push(studentRes.warning);

  var adminRes = sendPaymentVerifiedAdminReleaseEmail_(row, debugId);
  if (!adminRes.ok) {
    logAdminEvent_("PAYMENT_EMAIL_SEND_FAILED", { applicantId: applicantId, debugId: debugId, stage: "admin", error: adminRes.message || "" });
    return { ok: false, status: "failed", code: adminRes.code || "EMAIL_SEND_FAILED", message: adminRes.message || "Admin email failed", warnings: warnings };
  }

  var ts = nowIso_();
  props.setProperty(key, ts);
  logAdminEvent_("PAYMENT_VERIFIED_EMAIL_SENT", {
    applicantId: applicantId,
    studentEmail: getRowEmailForStudent_(row),
    adminTo: parseCsvEmails_(CONFIG.EMAIL_RELEASE_ADMIN_TO || ""),
    sentKey: key,
    dbg: debugId
  });
  return { ok: true, status: "sent", warnings: warnings, sentKey: key, sentAt: ts };
}

function handleInvoiceTrigger_(sh, rowNumber, idx, rowObj, debugId) {
  var row = rowObj || {};
  var applicantId = safeStr_(row.ApplicantID);
  if (hasValue_(row.CRM_Invoice_Triggered)) {
    logAdminEvent_("INVOICE_TRIGGER_SKIPPED_ALREADY", { applicantId: applicantId, debugId: debugId });
    return { status: "skipped", reason: "already_triggered" };
  }
  if (CONFIG.INVOICE_TRIGGER_ENABLED !== true) {
    logAdminEvent_("INVOICE_TRIGGER_DISABLED", { applicantId: applicantId, debugId: debugId });
    return { status: "disabled", reason: "config_disabled" };
  }

  var trig = triggerInvoiceWebhook_(row, debugId);
  if (!trig.ok) {
    logAdminEvent_("INVOICE_TRIGGER_FAILED", { applicantId: applicantId, debugId: debugId, code: trig.code || "", message: trig.message || "" });
    return { status: "failed", code: trig.code || "INVOICE_TRIGGER_FAILED", message: trig.message || "Invoice trigger failed" };
  }

  var ts = nowIso_();
  patchIfHeadersPresent_(sh, rowNumber, idx, {
    CRM_Invoice_Triggered: "Yes",
    Invoice_Sent_At: ts
  });
  row.CRM_Invoice_Triggered = "Yes";
  row.Invoice_Sent_At = ts;
  var paymentEmail = sendPaymentEmail_(row, debugId);
  if (!paymentEmail.ok) {
    logAdminEvent_("INVOICE_TRIGGERED_EMAIL_FAILED", { applicantId: applicantId, debugId: debugId, code: paymentEmail.code || "", message: paymentEmail.message || "" });
    return { status: "failed", code: paymentEmail.code || "PAYMENT_EMAIL_FAILED", message: paymentEmail.message || "Payment email failed" };
  }
  logAdminEvent_("INVOICE_TRIGGERED", { applicantId: applicantId, debugId: debugId, mode: trig.mode || "", httpStatus: trig.httpStatus || 0 });
  return { status: "triggered", mode: trig.mode || "LOG_ONLY" };
}

function runVerificationAutomations_(sh, rowNumber, idx, beforeRowObj, afterRowObj, debugId) {
  var beforeRow = beforeRowObj || {};
  var afterRow = afterRowObj || {};
  var actions = {
    quoteEmail: "skipped",
    invoice: "skipped",
    paymentVerifiedEmails: "skipped",
    warnings: []
  };
  var applicantId = safeStr_(afterRow.ApplicantID || beforeRow.ApplicantID || "");
  try {
    var docsBefore = isYes_(beforeRow.Docs_Verified);
    var docsAfter = isYes_(afterRow.Docs_Verified) || computeDocVerificationStatus_(afterRow) === "Verified";
    if (!docsBefore && docsAfter) {
      actions.quoteEmail = "manual_only";
      logAdminEvent_("QUOTE_EMAIL_SKIPPED", { applicantId: applicantId, debugId: debugId, reason: "manual_only_cis91" });
    }
  } catch (quoteErr) {
    actions.quoteEmail = "failed";
    logAdminEvent_("QUOTE_EMAIL_FAILED", { applicantId: applicantId, debugId: debugId, message: String(quoteErr && quoteErr.message ? quoteErr.message : quoteErr) });
  }

  try {
    var payBefore = isYes_(beforeRow.Payment_Verified) || isPaymentVerifiedDerived_(beforeRow) === true;
    var payAfter = isYes_(afterRow.Payment_Verified) || isPaymentVerifiedDerived_(afterRow) === true;
    if (!payBefore && payAfter) {
      // Payment verified email workflow is triggered explicitly in save handlers.
      actions.paymentVerifiedEmails = "handled_in_save_handler";
      var invRes = handleInvoiceTrigger_(sh, rowNumber, idx, afterRow, debugId);
      actions.invoice = safeStr_(invRes && invRes.status) || "failed";
      if (invRes && invRes.code) actions.invoiceCode = safeStr_(invRes.code);
      if (invRes && invRes.message) actions.invoiceMessage = safeStr_(invRes.message);
    }
  } catch (invErr) {
    actions.invoice = "failed";
    actions.invoiceCode = "INVOICE_TRIGGER_FAILED";
    actions.invoiceMessage = String(invErr && invErr.message ? invErr.message : invErr);
    logAdminEvent_("INVOICE_TRIGGER_FAILED", { applicantId: applicantId, debugId: debugId, message: actions.invoiceMessage });
  }
  actions.emailTriggered = (safeStr_(actions.paymentVerifiedEmails) === "sent");
  return actions;
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

function buildQueueRow_(rowNumber, applicantId, name, extra) {
  var out = {
    rowNumber: Number(rowNumber || 0),
    applicantId: clean_(applicantId || ""),
    name: clean_(name || "")
  };
  var more = (extra && typeof extra === "object") ? extra : {};
  for (var k in more) {
    if (Object.prototype.hasOwnProperty.call(more, k)) out[k] = more[k];
  }
  return out;
}

function nonEmpty_(v) {
  var s = String(v === null || v === undefined ? "" : v).trim();
  if (!s) return false;
  var n = s.toLowerCase();
  if (n === "0" || n === "false" || n === "n/a") return false;
  return true;
}

function hasAnyRequiredDoc_(rowObj) {
  var row = rowObj || {};
  var required = [
    "Birth_ID_Passport_File",
    "Latest_School_Report_File",
    "Transfer_Certificate_File",
    "Passport_Photo_File"
  ];
  for (var i = 0; i < required.length; i++) {
    if (nonEmpty_(row[required[i]])) return true;
  }
  return false;
}

function parseTime_(v) {
  if (v instanceof Date) return v.getTime();
  var s = String(v === null || v === undefined ? "" : v).trim();
  if (!s) return 0;
  var t = Date.parse(s);
  return Number.isFinite(t) ? t : 0;
}

function hasStudentActivity_(row) {
  var r = row || {};
  var lastUpdateRaw = r.PortalLastUpdateAt;
  var lastUpdateTs = parseTime_(lastUpdateRaw);
  if (!(lastUpdateTs > 0)) return false;
  var submitted = clean_(r.Portal_Submitted || "");
  if (!submitted) return true;
  return submitted === "Yes" || submitted.indexOf("T") > 0;
}

function applicantSuffix_(id) {
  var m = String(id || "").match(/(\d+)\s*$/);
  return m ? (Number(m[1]) || 0) : 0;
}

function compareQueueItems_(a, b) {
  var aTime = parseTime_(a && a.portalLastUpdateAt);
  var bTime = parseTime_(b && b.portalLastUpdateAt);
  if (bTime !== aTime) return bTime - aTime;
  var aToken = parseTime_(a && a.portalTokenIssuedAt);
  var bToken = parseTime_(b && b.portalTokenIssuedAt);
  if (bToken !== aToken) return bToken - aToken;
  var aSuffix = applicantSuffix_(a && a.applicantId);
  var bSuffix = applicantSuffix_(b && b.applicantId);
  if (bSuffix !== aSuffix) return bSuffix - aSuffix;
  return Number(a && a.rowNumber || 0) - Number(b && b.rowNumber || 0);
}

function hasMandatoryDocIssue_(rowObj, idx) {
  var row = rowObj || {};
  var mappings = [
    { file: "Birth_ID_Passport_File", status: "Birth_ID_Status" },
    { file: "Latest_School_Report_File", status: "Report_Status" }
  ];
  for (var i = 0; i < mappings.length; i++) {
    var m = mappings[i];
    if (idx && (!idx[m.file] || !idx[m.status])) continue;
    var hasFile = clean_(row[m.file] || "");
    var status = clean_(row[m.status] || "");
    if (hasFile && normalizeDocStatus_(status || "Pending") !== "Verified") return true;
  }
  return false;
}

function getDashboardCacheKey_(adminEmail) {
  return "ADMIN_DASHBOARD::" + clean_(adminEmail || "").toLowerCase();
}

function sliceQueueByOffset_(rows, offset, limit) {
  var list = Array.isArray(rows) ? rows : [];
  var from = Math.max(0, Number(offset || 0));
  var size = Math.max(1, Number(limit || 20));
  return list.slice(from, from + size);
}

function mergeQueuePageMeta_(queues, offset, limit) {
  var names = ["docs", "awaitingPayment", "payments", "anomalies", "paidApproved", "postPaymentIssues"];
  var hasMore = false;
  var nextOffset = "";
  for (var i = 0; i < names.length; i++) {
    var name = names[i];
    var total = Number((queues && queues.counts && queues.counts[name]) || 0);
    if (offset + limit < total) {
      hasMore = true;
      nextOffset = offset + limit;
      break;
    }
  }
  return {
    hasMore: hasMore,
    nextOffset: hasMore ? nextOffset : ""
  };
}

function admin_getReviewQueues(payload) {
  var adminEmail = getActiveUserEmail_();
  if (!isAdmin_(adminEmail)) throw new Error("Access denied");
  payload = payload || {};
  var offset = Math.max(0, Number(payload.offset || 0));
  var limit = Math.max(1, Number(payload.limit || 20));
  var force = payload && (payload.force === 1 || payload.force === true);
  var cache = CacheService.getUserCache();
  var cacheKey = getDashboardCacheKey_(adminEmail);
  var fullData = null;
  if (!force) {
    try {
      var cached = cache.get(cacheKey);
      if (cached) fullData = JSON.parse(cached);
    } catch (_cacheReadErr) {}
  }

  if (!fullData || typeof fullData !== "object") {
    var sheet = openDataSheet_();
    var data = sheet.getDataRange().getValues();
    if (!data || data.length < 2) {
      fullData = {
        docs: [],
        awaitingPayment: [],
        payments: [],
        anomalies: [],
        paidApproved: [],
        postPaymentIssues: [],
        counts: { payments: 0, docs: 0, awaitingPayment: 0, anomalies: 0, paidApproved: 0, postPaymentIssues: 0 }
      };
    } else {
      var headers = data[0];
      var idx = headerIndex_(headers);
      var payments = [];
      var docs = [];
      var awaitingPayment = [];
      var anomalies = [];
      var paidApproved = [];
      var postPaymentIssues = [];
      var debugRows = [];
      function pushQueueItem_(target, item) {
        var rowNum = Number(item && item.rowNumber || 0);
        if (!Number.isFinite(rowNum) || rowNum < 2) return;
        target.push(item);
      }

      Logger.log("QUEUE_SCAN_START " + JSON.stringify({
        user: Session.getEffectiveUser().getEmail(),
        force: force
      }));

      for (var r = 1; r < data.length; r++) {
        var row = data[r] || [];
        var rowObj = {};
        for (var c = 0; c < headers.length; c++) {
          var h = clean_(headers[c]);
          if (!h) continue;
          rowObj[h] = row[c];
        }

        var applicantId = clean_(rowObj.ApplicantID || "");
        if (!applicantId) continue;
        var firstName = clean_(rowObj.First_Name || "");
        var lastName = clean_(rowObj.Last_Name || "");
        var name = (firstName + " " + lastName).trim();
        var parentEmail = clean_(rowObj.Parent_Email || "");
        var correctedEmail = clean_(rowObj.Parent_Email_Corrected || "");
        var effectiveEmail = correctedEmail || parentEmail;

        var paymentVerifiedRaw = clean_(rowObj.Payment_Verified || "") === "Yes";
        var receiptUrl = clean_(rowObj.Fee_Receipt_File || "");
        var docsVerifiedRaw = clean_(rowObj.Docs_Verified || "");
        var mandatoryDocIssue = hasMandatoryDocIssue_(rowObj, idx);

        var docsVerifiedForFollowup = isYes_(rowObj.Docs_Verified) || computeDocVerificationStatus_(rowObj) === "Verified";
        var hasValidEmailForFollowup = !!getRowEmailForStudent_(rowObj);
        var docsFollowupEligibleBase = CONFIG.DOCS_FOLLOWUP_ENABLE === true && docsVerifiedForFollowup && hasValidEmailForFollowup;
        var docsFollowupSentAt = getDocsFollowupSentAt_(rowObj);
        var eligibleDocsFollowUp = docsFollowupEligibleBase && !safeStr_(docsFollowupSentAt || "");
        var qItem = {
          rowNumber: r + 1,
          applicantId: applicantId,
          name: name,
          ApplicantID: applicantId,
          parentEmail: parentEmail,
          correctedEmail: correctedEmail,
          effectiveEmail: effectiveEmail,
          portalLastUpdateAt: rowObj.PortalLastUpdateAt || "",
          portalTokenIssuedAt: rowObj.PortalTokenIssuedAt || "",
          docsFollowupEligibleBase: !!docsFollowupEligibleBase,
          eligibleDocsFollowUp: !!eligibleDocsFollowUp,
          docsFollowupSentAt: safeStr_(docsFollowupSentAt || "")
        };

        var hasActivity = hasStudentActivity_(rowObj);
        var portalSubmittedRaw = clean_(rowObj.Portal_Submitted || "");
        var portalSubmitted = nonEmpty_(portalSubmittedRaw) && portalSubmittedRaw !== "No";
        var docsVerified = docsVerifiedRaw === "Yes";
        var paymentEvidencePresent = nonEmpty_(receiptUrl) || nonEmpty_(clean_(rowObj.Receipt_Status || ""));
        var paymentReceived = paymentEvidencePresent;
        var paymentVerified = paymentVerifiedRaw;
        var enrolledConfirmed = paymentVerified;
        var docsQueueMatch = portalSubmitted && !docsVerified;
        var awaitingPaymentQueueMatch = docsVerified && !paymentVerified && !paymentEvidencePresent;
        var paymentsQueueMatch = docsVerified && !paymentVerified && paymentEvidencePresent;
        var anomaliesQueueMatch = paymentVerified && !docsVerified;
        var paidApprovedQueueMatch = paymentVerified;

        qItem.Portal_Submitted = portalSubmitted ? "Yes" : "No";
        qItem.Docs_Verified = docsVerified ? "Yes" : "No";
        qItem.Payment_Received = paymentReceived ? "Yes" : "No";
        qItem.Payment_Verified = paymentVerified ? "Yes" : "No";
        qItem.Enrolled_Confirmed = enrolledConfirmed ? "Yes" : "No";
        qItem.Fee_Receipt_File = receiptUrl;
        qItem.Registration_Complete = clean_(rowObj.Registration_Complete || "") === "Yes" ? "Yes" : "No";

        debugRows.push({
          id: clean_(rowObj.ApplicantID || rowObj.ID || rowObj["Applicant ID"] || "unknown"),
          activity: hasActivity,
          portalSubmitted: portalSubmitted,
          docsVerified: docsVerified,
          paymentVerified: paymentVerified,
          paymentEvidencePresent: paymentEvidencePresent,
          receipt: paymentEvidencePresent,
          portalTs: clean_(rowObj.PortalLastUpdateAt || ""),
          docsQueue: docsQueueMatch,
          awaitingPaymentQueue: awaitingPaymentQueueMatch,
          paymentsQueue: paymentsQueueMatch,
          anomaliesQueue: anomaliesQueueMatch,
          paidApprovedQueue: paidApprovedQueueMatch
        });
        Logger.log("QUEUE_CLASSIFY " + JSON.stringify({
          applicantId: rowObj.ApplicantID,
          portalSubmitted: portalSubmitted,
          docsVerifiedRaw: rowObj.Docs_Verified,
          docsVerified: docsVerified,
          paymentVerifiedRaw: rowObj.Payment_Verified,
          paymentVerified: paymentVerified,
          paymentEvidencePresent: paymentEvidencePresent,
          awaitingPaymentQueue: awaitingPaymentQueueMatch,
          hasActivity: hasActivity
        }));

        if (paidApprovedQueueMatch) {
          pushQueueItem_(paidApproved, qItem);
        } else if (paymentsQueueMatch) {
          pushQueueItem_(payments, qItem);
        } else if (awaitingPaymentQueueMatch) {
          pushQueueItem_(awaitingPayment, qItem);
        } else if (docsQueueMatch) {
          pushQueueItem_(docs, qItem);
        }

        if (anomaliesQueueMatch) {
          pushQueueItem_(anomalies, qItem);
        }
        if (paymentVerified && mandatoryDocIssue) {
          pushQueueItem_(postPaymentIssues, qItem);
        }
      }
      docs.sort(compareQueueItems_);
      awaitingPayment.sort(compareQueueItems_);
      payments.sort(compareQueueItems_);
      anomalies.sort(compareQueueItems_);
      paidApproved.sort(compareQueueItems_);
      postPaymentIssues.sort(compareQueueItems_);

      function stripQueue_(items) {
        return (items || []).map(function (it) {
          return buildQueueRow_(it.rowNumber, it.applicantId, it.name, {
            ApplicantID: clean_(it.ApplicantID || it.applicantId || ""),
            parentEmail: clean_(it.parentEmail || ""),
            correctedEmail: clean_(it.correctedEmail || ""),
            effectiveEmail: clean_(it.effectiveEmail || ""),
            docsFollowupEligibleBase: !!it.docsFollowupEligibleBase,
            eligibleDocsFollowUp: !!it.eligibleDocsFollowUp,
            docsFollowupSentAt: safeStr_(it.docsFollowupSentAt || ""),
            Portal_Submitted: clean_(it.Portal_Submitted || ""),
            Docs_Verified: clean_(it.Docs_Verified || ""),
            Payment_Received: clean_(it.Payment_Received || ""),
            Payment_Verified: clean_(it.Payment_Verified || ""),
            Enrolled_Confirmed: clean_(it.Enrolled_Confirmed || ""),
            Fee_Receipt_File: clean_(it.Fee_Receipt_File || ""),
            Registration_Complete: clean_(it.Registration_Complete || "")
          });
        });
      }
      fullData = {
        docs: stripQueue_(docs),
        awaitingPayment: stripQueue_(awaitingPayment),
        payments: stripQueue_(payments),
        anomalies: stripQueue_(anomalies),
        paidApproved: stripQueue_(paidApproved),
        postPaymentIssues: stripQueue_(postPaymentIssues),
        counts: {
          payments: payments.length,
          docs: docs.length,
          awaitingPayment: awaitingPayment.length,
          anomalies: anomalies.length,
          paidApproved: paidApproved.length,
          postPaymentIssues: postPaymentIssues.length
        }
      };
      debugRows.forEach(function (d) {
        if (d.id === "FODE-26-000084" || d.id === "FODE-26-000007") {
          Logger.log("CIS-r228 QUEUE DEBUG for %s: %s", d.id, JSON.stringify(d));
        }
      });
      Logger.log("QUEUE_SUMMARY " + JSON.stringify({
        docs: docs.length,
        awaitingPayment: awaitingPayment.length,
        payments: payments.length,
        anomalies: anomalies.length,
        paidApproved: paidApproved.length
      }));
    }
    try {
      cache.put(cacheKey, JSON.stringify(fullData), 60);
    } catch (_cacheWriteErr) {}
  }

  var pageMeta = mergeQueuePageMeta_(fullData, offset, limit);
  function refreshDocsFollowupRuntime_(rows) {
    return (rows || []).map(function (row) {
      var out = {};
      var src = row && typeof row === "object" ? row : {};
      for (var k in src) {
        if (Object.prototype.hasOwnProperty.call(src, k)) out[k] = src[k];
      }
      var applicantId = clean_(out.ApplicantID || out.applicantId || "");
      var key = buildDocsFollowupKey_(applicantId);
      var sentAt = "";
      try { sentAt = safeStr_(PropertiesService.getScriptProperties().getProperty(key) || ""); } catch (_propErr) {}
      out.docsFollowupSentAt = sentAt;
      var eligibleBase = !!out.docsFollowupEligibleBase;
      out.eligibleDocsFollowUp = eligibleBase && !sentAt;
      return out;
    });
  }

  return {
    ok: true,
    docs: refreshDocsFollowupRuntime_(sliceQueueByOffset_(fullData.docs, offset, limit)),
    awaitingPayment: refreshDocsFollowupRuntime_(sliceQueueByOffset_(fullData.awaitingPayment, offset, limit)),
    payments: refreshDocsFollowupRuntime_(sliceQueueByOffset_(fullData.payments, offset, limit)),
    anomalies: refreshDocsFollowupRuntime_(sliceQueueByOffset_(fullData.anomalies, offset, limit)),
    paidApproved: refreshDocsFollowupRuntime_(sliceQueueByOffset_(fullData.paidApproved, offset, limit)),
    postPaymentIssues: refreshDocsFollowupRuntime_(sliceQueueByOffset_(fullData.postPaymentIssues, offset, limit)),
    counts: fullData.counts || { payments: 0, docs: 0, awaitingPayment: 0, anomalies: 0, paidApproved: 0, postPaymentIssues: 0 },
    offset: offset,
    limit: limit,
    hasMore: pageMeta.hasMore,
    nextOffset: pageMeta.nextOffset
  };
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

  if (scope === "range") {
    var rangeOut = [];
    var endRowRange = Math.min(lastRow, startRow + batchSize - 1);
    for (var rr = startRow; rr <= endRowRange; rr++) rangeOut.push(rr);
    return rangeOut;
  }

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
  Logger.log("EXPORT_PORTAL_LINKS " + JSON.stringify({
    scope: clean_(payload.scope || ""),
    startRow: Number(payload.startRow || 0),
    batchSize: Number(payload.batchSize || 0)
  }));
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
