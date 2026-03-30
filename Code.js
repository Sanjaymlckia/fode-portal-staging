/************************************************************
FODE ADMISSIONS — STAGING SCRIPT (PORTAL + DRIVE UPLOADS + ALLOWLIST)

Fixes included:
1) Student portal shows ONLY allowlisted fields (no CRM/IDs/logs/tokens)
2) Portal POST blank-screen fix: hardcode WEBAPP_URL for <form action=...>
3) doGet accepts email param aliases: email= OR Parent_Email=
4) Portal updates DO NOT overwrite Parent_Email (prevents lookup breaking)
   - saves to Parent_Email_Corrected instead
5) Drive uploads from portal replace the *_File URL and append File_Log
6) Portal locks when Payment_Verified == "Yes"
7) Subjects always prefill correctly (JSON / map style / CSV) + display nicely

Sheet:
- Spreadsheet: CONFIG.SHEET_ID
- Data sheet: CONFIG.DATA_SHEET
- Log sheet:  CONFIG.LOG_SHEET
************************************************************/


var PORTAL_SECRETS_SPREADSHEET_ID = "1HEJPtSov-iE5YTpSWWZ89YLIQAw4Eju9DDMG46HkTRc";
var PORTAL_SECRETS_TAB = "PortalSecrets";
var STUDENT_EXEC_BASE = "https://script.google.com/macros/s/AKfycbx2ve4bfCEofF_pJnra-UR02BaoumJaUeDS19Amftm2con2e7ggblMfHRzcn6fYAC4g";

/******************** ENTRYPOINT: POST ********************/
function doPost(e) {
  var reqId = makeReqId_();
  var params = (e && e.parameter && typeof e.parameter === "object") ? e.parameter : {};
  var postView = clean_(typeof getParam_ === "function" ? getParam_(e, "view") : (params.view || "")).toLowerCase();
  if (postView === "portalupload") {
    return doPost_portalUpload_(e);
  }
  var paramKeys = Object.keys(params);
  var rawPostData = (e && e.postData) ? e.postData : null;
  var postType = clean_(rawPostData && (rawPostData.type || rawPostData.contentType) || "");
  var postLen = rawPostData && rawPostData.contents ? String(rawPostData.contents).length : 0;
  var payload;
  try {
    payload = parseRequestPayload_(e);
  } catch (err) {
    var dbgFromParams = clean_(params.dbg || "") === "1";
    logPortalPostEvent_("PORTAL_POST_START", {
      reqId: reqId,
      keys: paramKeys,
      view: clean_(params.view || ""),
      action: clean_(params.action || params._action || params.route || ""),
      id: clean_(params.id || params.ApplicantID || ""),
      s: redactToken_(params.s || params.secret || ""),
      postDataType: postType,
      postDataLength: postLen,
      parseError: String(err && err.message ? err.message : err)
    });
    var dbgBadPayload = newDebugId_();
    var badPayloadResult = {
      ok: false,
      debugId: dbgBadPayload,
      applicantId: "",
      error: { message: "Invalid request payload.", code: "BAD_PAYLOAD" }
    };
    var badPayloadRedirect = buildPortalRedirectUrl_("", "", { error: true, dbg: dbgBadPayload });
    logPortalPostEvent_("PORTAL_POST_REDIRECT", {
      reqId: reqId,
      redirectUrl: badPayloadRedirect,
      id: "",
      s: "",
      saved: false
    });
    return returnPortalRedirectOutput_(badPayloadRedirect, {
      debug: CONFIG.DEBUG_PORTAL_SHOW_ON_PAGE === true && dbgFromParams,
      reqId: reqId,
      applicantId: "",
      secret: "",
      redirectUrl: badPayloadRedirect,
      tokenValidationPassed: false,
      result: badPayloadResult,
      debugId: dbgBadPayload
    });
  }
  var action = clean_(payload.action || payload._action || payload.route || "");
  payload.__reqId = reqId;
  payload.__paramKeys = paramKeys.slice();
  var debugPost = CONFIG.DEBUG_PORTAL_SHOW_ON_PAGE === true && clean_(payload.dbg || params.dbg || "") === "1";
  logPortalPostEvent_("PORTAL_POST_START", {
    reqId: reqId,
    keys: paramKeys,
    view: clean_(params.view || ""),
    action: action,
    id: clean_(payload.id || payload.ApplicantID || params.id || ""),
    s: redactToken_(payload.s || payload.secret || params.s || ""),
    postDataType: postType,
    postDataLength: postLen
  });

  if (action === "portal_update") {
    var debugId = newDebugId_();
    var applicantId = clean_(payload.id || payload.ApplicantID || "");
    var secret = clean_(payload.s || payload.secret || "");
    try {
      if (hasOwn_(payload, "payload")) {
        payload = mergePortalPayload_(payload, parsePortalPayloadField_(payload.payload));
        applicantId = clean_(payload.id || payload.ApplicantID || applicantId);
        secret = clean_(payload.s || payload.secret || secret);
      }
      if (!applicantId || !secret) {
        try {
          var ssMissing = getWorkingSpreadsheet_();
          var logSheetMissing = mustGetSheet_(ssMissing, CONFIG.LOG_SHEET);
          log_(logSheetMissing, "PORTAL_UPDATE_FATAL", debugId + " Missing portal token (id/s)");
        } catch (logErr0) {}
        var missRedirect = buildPortalRedirectUrl_(applicantId, secret, {
          error: true,
          dbg: debugId
        });
        var missResult = {
          ok: false,
          debugId: debugId,
          applicantId: applicantId,
          error: { message: "Missing portal token (id/s).", code: "MISSING_TOKEN" }
        };
        logPortalPostEvent_("PORTAL_POST_REDIRECT", {
          reqId: reqId,
          redirectUrl: missRedirect,
          id: applicantId,
          s: redactToken_(secret),
          saved: false
        });
        return returnPortalRedirectOutput_(missRedirect, {
          debug: debugPost,
          reqId: reqId,
          applicantId: applicantId,
          secret: secret,
          redirectUrl: missRedirect,
          tokenValidationPassed: false,
          result: missResult,
          debugId: debugId
        });
      }
      var ssPortal = getWorkingSpreadsheet_();
      var dataSheetPortal = mustGetDataSheet_(ssPortal);
      var logSheetPortal = mustGetSheet_(ssPortal, CONFIG.LOG_SHEET);
      log_(logSheetPortal, "PORTAL_UPDATE_DEBUG", debugId + " " + applicantId);
      var resultObj = handlePortalUpdate_(ssPortal, dataSheetPortal, logSheetPortal, payload, params, debugId);
      resultObj = outputToJsonObject_(resultObj) || {};
      if (resultObj.ok === true) {
        var okRedirect = buildPortalRedirectUrl_(clean_(resultObj.applicantId || applicantId), secret, { saved: true });
        logPortalPostEvent_("PORTAL_POST_REDIRECT", {
          reqId: reqId,
          redirectUrl: okRedirect,
          id: applicantId,
          s: redactToken_(secret),
          saved: okRedirect.indexOf("saved=1") >= 0
        });
        return returnPortalRedirectOutput_(okRedirect, {
          debug: debugPost,
          reqId: reqId,
          applicantId: applicantId,
          secret: secret,
          redirectUrl: okRedirect,
          tokenValidationPassed: true,
          result: resultObj,
          debugId: clean_(resultObj.debugId || debugId)
        });
      }
      var failDbgId = clean_(resultObj.debugId || debugId) || debugId;
      var failCode = clean_(resultObj && resultObj.error && resultObj.error.code || "");
      var failValidationErrors = Array.isArray(resultObj.validationErrors) ? resultObj.validationErrors : [];
      var failFields = failValidationErrors.map(function (item) { return clean_(item && item.field || ""); }).filter(function (item) { return !!item; });
      var failCodes = failValidationErrors.map(function (item) { return clean_(item && item.code || ""); }).filter(function (item) { return !!item; });
      var isPaymentVerifiedLock = failCode === "PAYMENT_VERIFIED_LOCK";
      var failRedirect = isPaymentVerifiedLock
        ? buildPortalRedirectUrl_(applicantId, secret, { locked: true, msg: "enrolled" })
        : buildPortalRedirectUrl_(applicantId, secret, { error: true, dbg: failDbgId, val: failValidationErrors.length > 0, fields: failFields.join(","), errCode: failCodes.join(",") });
      logPortalPostEvent_("PORTAL_POST_REDIRECT", {
        reqId: reqId,
        redirectUrl: failRedirect,
        id: applicantId,
        s: redactToken_(secret),
        saved: false
      });
      return returnPortalRedirectOutput_(failRedirect, {
        debug: debugPost,
        reqId: reqId,
        applicantId: applicantId,
        secret: secret,
        redirectUrl: failRedirect,
        tokenValidationPassed: isPaymentVerifiedLock ? true : false,
        result: resultObj,
        debugId: failDbgId
      });
    } catch (errPortal) {
      try {
        var ssLog = getWorkingSpreadsheet_();
        var logSheetFatal = mustGetSheet_(ssLog, CONFIG.LOG_SHEET);
        log_(logSheetFatal, "PORTAL_UPDATE_FATAL", debugId + " " + String(errPortal && errPortal.message ? errPortal.message : errPortal));
      } catch (logErr) {}
      var excResult = {
        ok: false,
        debugId: debugId,
        applicantId: applicantId,
        error: {
          code: "PORTAL_UPDATE_EXCEPTION",
          message: String(errPortal && errPortal.message ? errPortal.message : errPortal)
        }
      };
      var excRedirect = buildPortalRedirectUrl_(applicantId, secret, { error: true, dbg: debugId });
      logPortalPostEvent_("PORTAL_POST_REDIRECT", {
        reqId: reqId,
        redirectUrl: excRedirect,
        id: applicantId,
        s: redactToken_(secret),
        saved: false
      });
      return returnPortalRedirectOutput_(excRedirect, {
        debug: debugPost,
        reqId: reqId,
        applicantId: applicantId,
        secret: secret,
        redirectUrl: excRedirect,
        tokenValidationPassed: false,
        result: excResult,
        debugId: debugId
      });
    }
  }

  var ss = getWorkingSpreadsheet_();
  var dataSheet = mustGetDataSheet_(ss);
  var logSheet = mustGetSheet_(ss, CONFIG.LOG_SHEET);
  appendPortalLog_({ route: "doPost", status: "HIT", message: "doPost called", email: payload.email || payload.Parent_Email || "", applicantId: payload.id || payload.ApplicantID || "" });


  log_(logSheet, "doPost HIT", payloadSummary_(payload));
  log_(logSheet, "ACTION", action || "(blank)");

  // Intake webhook (FormDesigner)
  log_(logSheet, "POST HIT", payloadSummary_(payload));
  log_(logSheet, "PAYLOAD KEYS", Object.keys(payload).join(", "));

  ensureHeaders_(dataSheet, payload);

  var correlationId = clean_(payload.correlation_id || "");
  var activationStage = "START";
  var activationCode = "ACTIVATION_FAILED";
  var targetRow = 0;
  var applicantId = "";
  var folder = null;
  var folderUrl = "";
  var tokenState = null;
  var rowCommitted = false;
  var verification = null;
  var intakeLock = LockService.getScriptLock();

  try {
    intakeLock.waitLock(30000);
    logActivation_(logSheet, "ACTIVATION_START", {
      correlation_id: correlationId,
      payloadKeyCount: Object.keys(payload || {}).length,
      spreadsheetId: clean_(getWorkingSpreadsheetId_() || ""),
      dataSheet: clean_(dataSheet.getName() || "")
    });

    activationStage = "APPLICANTID_PREPARE";
    var applicantIdState = scanApplicantIdState_(dataSheet);
    applicantId = clean_(applicantIdState.applicantId || "");
    if (!applicantId) {
      activationCode = "APPLICANTID_PREPARE_FAILED";
      throw new Error("APPLICANTID_PREPARE_FAILED");
    }
    logActivation_(logSheet, "ACTIVATION_ID_PREPARED", {
      correlation_id: correlationId,
      applicantId: applicantId,
      validCount: applicantIdState.validCount,
      maxSuffix: applicantIdState.maxSuffix,
      skippedBlankCount: applicantIdState.skippedBlankCount,
      skippedMalformedCount: applicantIdState.skippedMalformedCount
    });

    activationStage = "FOLDER_PREPARE";
    folder = createApplicantFolder_(payload);
    folderUrl = clean_(folder && folder.getUrl ? folder.getUrl() : "");
    if (!folderUrl) {
      activationCode = "FOLDER_PREPARE_FAILED";
      throw new Error("FOLDER_PREPARE_FAILED");
    }
    logActivation_(logSheet, "ACTIVATION_FOLDER_PREPARED", {
      correlation_id: correlationId,
      folderUrl: folderUrl,
      folderId: clean_(folder && folder.getId ? folder.getId() : "")
    });

    activationStage = "FILE_CANONICALIZE";
    payload = canonicalizeFdIntakeFiles_(payload, folder, logSheet, {
      correlationId: correlationId,
      applicantId: applicantId
    });
    payload = maybeStampActivationSubmitState_(payload, logSheet, {
      applicantId: applicantId
    });

    activationStage = "TOKEN_PREPARE";
    tokenState = preparePortalActivationState_(dataSheet, applicantId);
    logActivation_(logSheet, "ACTIVATION_TOKEN_PREPARED", {
      correlation_id: correlationId,
      hasPortalTokenHashHeader: tokenState.hasTokenHashHeader === true,
      hasPortalTokenIssuedAtHeader: tokenState.hasTokenIssuedAtHeader === true,
      portalSecretsPrepared: tokenState.portalSecretsRequired === true
    });

    activationStage = "ROW_COMMIT";
    targetRow = dataSheet.getLastRow() + 1;
    var activatedRow = buildActivatedIntakeRow_(dataSheet, payload, folderUrl, applicantId, tokenState);
    insertActivatedRowAt_(dataSheet, targetRow, activatedRow);
    rowCommitted = true;
    logActivation_(logSheet, "ACTIVATION_ROW_COMMIT", {
      correlation_id: correlationId,
      targetRow: targetRow,
      applicantId: applicantId
    });

    activationStage = "PORTALSECRETS_COMMIT";
    if (tokenState.portalSecretsRequired === true) {
      commitPortalActivationState_(payload, applicantId, tokenState);
    }

    activationStage = "VERIFY";
    verification = verifyActivatedState_(dataSheet, targetRow, applicantId, folderUrl, tokenState);
    logActivation_(logSheet, "ACTIVATION_VERIFY", {
      correlation_id: correlationId,
      targetRow: targetRow,
      applicantIdActual: clean_(verification.applicantIdActual || ""),
      folderUrlPresent: verification.folderUrlPresent === true,
      portalTokenHashPresent: verification.portalTokenHashPresent === true,
      portalTokenIssuedAtPresent: verification.portalTokenIssuedAtPresent === true,
      portalSecretsResolvable: verification.portalSecretsResolvable === true
    });
    if (verification.ok !== true) {
      activationCode = clean_(verification.code || "ACTIVATION_VERIFY_FAILED") || "ACTIVATION_VERIFY_FAILED";
      throw new Error(clean_(verification.message || activationCode) || activationCode);
    }

    activationStage = "OK";
    logActivation_(logSheet, "ACTIVATION_OK", {
      correlation_id: correlationId,
      targetRow: targetRow,
      applicantId: applicantId
    });
    return jsonOut_({ status: "ok", ApplicantID: applicantId });
  } catch (errActivation) {
    if (rowCommitted && targetRow >= 2) {
      try { dataSheet.deleteRow(targetRow); } catch (_deleteErr) {}
    }
    logActivation_(logSheet, "ACTIVATION_FAIL", {
      correlation_id: correlationId,
      stage: activationStage,
      targetRow: targetRow || 0,
      applicantId: applicantId,
      error: String(errActivation && errActivation.message ? errActivation.message : errActivation)
    });
    return jsonOut_({
      status: "error",
      code: clean_(activationCode || "ACTIVATION_FAILED") || "ACTIVATION_FAILED",
      message: String(errActivation && errActivation.message ? errActivation.message : errActivation),
      correlation_id: correlationId
    });
  } finally {
    try { intakeLock.releaseLock(); } catch (_lockErr) {}
  }}

/******************** ENTRYPOINT: GET ********************/
function maybeRedirectToCanonical_(e) {
  var params = (e && e.parameter && typeof e.parameter === "object") ? e.parameter : {};
  var view = clean_(params.view || "").toLowerCase();
  var currentUrl = clean_(ScriptApp.getService().getUrl() || "");
  var canonicalBase = pickCanonicalExecBase_(e);

  if (!currentUrl || !canonicalBase) return null;

  var force = String(params.force || "") !== "";

  if (view === "admin" && !force) {
    var canonical = "https://script.google.com/macros/s/AKfycbwLz4rLrVzk-NriJAovTbifpg8YpQguFJiY-l02qkRrahH1ayX_2qBh3bk_rc8dVnPp/exec";
    var target = canonical + "?view=admin&force=1";
    Object.keys(params).forEach(function(key) {
      if (key && key !== "view" && key !== "force") {
        target += "&" + encodeURIComponent(key) + "=" + encodeURIComponent(String(params[key] || ""));
      }
    });
    return HtmlService.createHtmlOutput(
  '<!DOCTYPE html><html><body>' +
  '<script>try{top.location.replace(' + JSON.stringify(target) + ');}catch(e){location.replace(' + JSON.stringify(target) + ');}</script>' +
  '</body></html>'
);
  }

  if (currentUrl.indexOf("/a/macros/") !== -1) {
    var redirectHtml = HtmlService.createHtmlOutput(
      '<script>location.replace("' + canonicalBase + '" + location.search);</script>'
    );

    redirectHtml
      .addMetaTag("Cache-Control", "no-store, no-cache, must-revalidate, max-age=0")
      .addMetaTag("Pragma", "no-cache")
      .addMetaTag("Expires", "0");

    return redirectHtml;
  }

  return null;
}

function doGet(e) {
  var dbg = (typeof newDebugId_ === "function") ? newDebugId_() : ("DBG-" + Utilities.getUuid().slice(0, 8));
  var params = (e && e.parameter && typeof e.parameter === "object") ? e.parameter : {};
  var view = clean_(params.view || "").toLowerCase();
  var serviceUrl = "";
  var currentUrl = "";
  var isAdminDeployment = false;

  try {
    var redirect = maybeRedirectToCanonical_(e);
    if (redirect) return redirect;

    try {
      serviceUrl = clean_(ScriptApp.getService().getUrl() || "");
      currentUrl = serviceUrl;
    } catch (_serviceErr) {}

    isAdminDeployment = isAdminDeploymentRequest_();
    Logger.log("ROUTE doGet START dbg=%s view=%s isAdmin=%s url=%s", dbg, view || "(blank)", isAdminDeployment ? "true" : "false", currentUrl || "");

    var handler = resolveDoGetHandler_(view, isAdminDeployment);
    var result = handler(e);
    if (!result) {
      throw new Error("Route handler returned empty response. view=" + (view || "(blank)"));
    }

    Logger.log("ROUTE doGet OK dbg=%s view=%s", dbg, view || "(blank)");
    return result;
  } catch (err) {
    var errMsg = stringifyGsError_(err);
    try {
      Logger.log("ROUTE doGet FAIL dbg=%s view=%s url=%s err=%s", dbg, view || "(blank)", currentUrl || "", errMsg);
    } catch (_logErr) {}
    return renderDoGetFatalHtml_(dbg, view, currentUrl, errMsg);
  }
}

function resolveDoGetHandler_(view, isAdminDeployment) {
  var route = clean_(view || "").toLowerCase();
  if (route === "diag") return respondDiag_;
  if (route === "whoami") return doGet_whoami_;
  if (route === "file") return doGet_file_;
  if (route === "driveapiprobe") return doGet_driveApiProbe_;
  if (route === "drivedeepprobe") return doGet_driveDeepProbe_;
  if (route === "driveprobe") return doGet_driveProbe_;
  if (route === "portalsmoke") return doGet_portalSmoke_;
  if (route === "uploadsmoke") return doGet_uploadSmoke_;
  if (route === "admin") return renderAdminApp_;
  if (!route) return isAdminDeployment ? renderAdminApp_ : renderPortalAppFromDoGet_;
  return isAdminDeployment ? renderAdminApp_ : renderPortalAppFromDoGet_;
}

function renderPortalAppFromDoGet_(e) {
  return renderPortalPageResponse_(e, { uploadResult: null, viewName: "portal" });
}

function renderDoGetFatalHtml_(dbg, view, url, errMsg) {
  var html = ''
    + '<!doctype html><html><head><meta charset="utf-8"><base target="_top">'
    + '<meta name="viewport" content="width=device-width, initial-scale=1">'
    + '<meta http-equiv="Cache-Control" content="no-store, no-cache, must-revalidate, max-age=0">'
    + '<meta http-equiv="Pragma" content="no-cache">'
    + '<meta http-equiv="Expires" content="0">'
    + '<style>body{font-family:Arial,Helvetica,sans-serif;background:#0f172a;color:#e2e8f0;margin:0;padding:24px}.card{max-width:900px;margin:0 auto;background:#111827;border:1px solid #7f1d1d;border-radius:12px;padding:16px}.k{color:#93c5fd;font-weight:700}.v{word-break:break-word}.err{margin-top:12px;padding:10px;border:1px solid #7f1d1d;background:#3f1d1d;border-radius:8px;color:#fee2e2}</style>'
    + '</head><body><div class="card">'
    + '<h2 style="margin:0 0 12px 0;color:#fecaca">ROUTE FAILURE</h2>'
    + '<div><span class="k">Debug ID:</span> <span class="v">' + esc_(clean_(dbg || "")) + '</span></div>'
    + '<div style="margin-top:8px"><span class="k">View:</span> <span class="v">' + esc_(clean_(view || "(blank)")) + '</span></div>'
    + '<div style="margin-top:8px"><span class="k">URL:</span> <span class="v">' + esc_(clean_(url || "")) + '</span></div>'
    + '<div class="err"><span class="k">Error:</span> ' + esc_(clean_(errMsg || "Unknown error")) + '</div>'
    + '</div></body></html>';
  return HtmlService.createHtmlOutput(html)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function renderPortalPageResponse_(e, opts) {
  var params = (e && e.parameter && typeof e.parameter === "object") ? e.parameter : {};
  var o = (opts && typeof opts === "object") ? opts : {};
  var uploadRes = (o.uploadResult && typeof o.uploadResult === "object" && !Array.isArray(o.uploadResult)) ? o.uploadResult : null;
  var id = clean_(o.applicantId || params.id || "");
  var secret = clean_(o.secret || params.s || "");
  var saved = clean_(params.saved || "") === "1";
  var errorFlag = clean_(params.error || "") === "1";
  var lockedFlag = clean_(params.locked || "") === "1";
  var msgToken = clean_(params.msg || "");
  var dbg = clean_(params.dbg || "");
  var uploadFail = clean_(params.uploadFail || "") === "1";
  var uploadField = clean_(params.field || "");
  var uploadResult = clean_(params.u || "") === "1";
  var uploadOk = clean_(params.ok || "") === "1";
  var uploadDocKey = clean_(params.docKey || "");
  var uploadErrCode = clean_(params.errCode || "");
  var validationFlag = clean_(params.val || "") === "1";
  var validationFields = parsePortalCsvParam_(params.fields || "");
  var validationCodes = parsePortalCsvParam_(params.errCode || "");
  if (uploadRes) {
    uploadResult = true;
    uploadOk = uploadRes.ok === true;
    uploadDocKey = clean_(uploadRes.docKey || uploadDocKey || "");
    uploadErrCode = clean_(uploadRes.errCode || uploadErrCode || "");
    if (clean_(uploadRes.dbg || "")) dbg = clean_(uploadRes.dbg || "");
    uploadFail = uploadOk !== true;
  }
  var reqId = makeReqId_();
  var debugPage = CONFIG.DEBUG_PORTAL_SHOW_ON_PAGE === true && dbg === "1";
  var reqMeta = getPortalRequestMeta_(e);
  var isAdminDeployment = isAdminDeploymentRequest_();
  function invalidPortalLinkMsg_(reasonCode, dbgId, extra) {
    var reason = clean_(reasonCode || "") || "invalid";
    var debugId = clean_(dbgId || "");
    var extraText = clean_(extra || "");
    var msg = "Invalid portal link (" + reason + ")";
    if (extraText) msg += " - " + extraText;
    if (debugId) msg += " Debug: " + debugId;
    return msg;
  }
  if (!id || !secret) {
    var missingTokenDbg = newDebugId_();
    safePortalLog_({
      route: "doGet:portal",
      applicantId: id || "",
      email: reqMeta.ip || "",
      status: "invalid_token",
      message: "missing_params dbg=" + missingTokenDbg + " | ua=" + (reqMeta.ua || "")
    }, false);
    var msg = invalidPortalLinkMsg_("missingToken", errorFlag && dbg ? dbg : missingTokenDbg, "Please reopen your portal link.");
    return htmlOutput_(renderErrorHtml_(msg));
  }

  var secretRes = getPortalSecretForApplicant_(id);
  if (!secretRes || secretRes.ok !== true || clean_(secretRes.secret || "") !== secret) {
    var secretMismatchDbg = newDebugId_();
    var badCount = incrementInvalidPortalAttempt_(id);
    if (badCount > 10) return htmlOutput_(renderErrorHtml_("Too many invalid attempts. Please try again later."));
    safePortalLog_({
      route: "doGet:portal",
      applicantId: id,
      email: reqMeta.ip || "",
      status: "invalid_token",
      message: "hash_mismatch attempts=" + badCount + " dbg=" + secretMismatchDbg + " | ua=" + (reqMeta.ua || "")
    }, false);
    return htmlOutput_(renderErrorHtml_("Invalid or expired link. Please request a new link."));
  }

  var ss = getWorkingSpreadsheet_();
  var sheet = mustGetDataSheet_(ss);
  var rowNum = findRowByApplicantId_(sheet, id);
  if (!rowNum) {
    var applicantNotFoundDbg = newDebugId_();
    var missingCount = incrementInvalidPortalAttempt_(id);
    if (missingCount > 10) return htmlOutput_(renderErrorHtml_("Too many invalid attempts. Please try again later."));
    safePortalLog_({
      route: "doGet:portal",
      applicantId: id,
      email: reqMeta.ip || "",
      status: "invalid_token",
      message: "row_not_found attempts=" + missingCount + " dbg=" + applicantNotFoundDbg + " | ua=" + (reqMeta.ua || "")
    }, false);
    return htmlOutput_(renderErrorHtml_(invalidPortalLinkMsg_("applicantNotFound", applicantNotFoundDbg)));
  }
  var rowObj = getRowObject_(sheet, rowNum);
  var record = rowObj;
  if (String(record[SCHEMA.PORTAL_ACCESS_STATUS] || "").trim() === "Locked") {
    var lockedDbg = newDebugId_();
    safePortalLog_({
      route: "doGet:portal",
      applicantId: id,
      email: reqMeta.ip || "",
      status: "locked",
      message: "portal_locked dbg=" + lockedDbg + " | ua=" + (reqMeta.ua || "")
    }, false);
    return htmlOutput_(renderErrorHtml_(invalidPortalLinkMsg_("locked", lockedDbg, "Access suspended")));
  }

  record._PortalLockReason = getPortalLockReason_(record);
  record._PortalLocked = isPortalLocked_(record);
  safePortalLog_({
    route: "doGet:portal",
    applicantId: id,
    email: reqMeta.ip || "",
    status: "success",
    message: "open_ok | ua=" + (reqMeta.ua || "")
  }, false);

  var canonical = clean_(record.Subjects_Selected_Canonical || "");
  var fallbackCsv = subjectsToCsv_(record.Subjects_Selected || "");
  record._SubjectsCsv = canonical || fallbackCsv;

  var examSites = getExamSites_(ss);
  var portalHtml = renderPortalHtml_({
    id: id,
    secret: secret,
    reqId: reqId,
    debugPage: debugPage,
    saved: saved,
    errorFlag: errorFlag,
    lockedFlag: lockedFlag,
    msgToken: msgToken,
    dbg: dbg,
    uploadFail: uploadFail,
    uploadField: uploadField,
    uploadResult: uploadResult,
    uploadOk: uploadOk,
    uploadDocKey: uploadDocKey,
    uploadErrCode: uploadErrCode,
    validationFlag: validationFlag,
    validationFields: validationFields,
    validationCodes: validationCodes,
    record: record,
    subjects: CONFIG.PORTAL_SUBJECTS,
    examSites: examSites,
    editFields: getPortalEditableFields_(),
    docs: getDocUiFields_(),
    visibleFields: CONFIG.PORTAL_VISIBLE_FIELDS,
    subjectsLocked: isDocsVerified_(record),
    version: CONFIG.VERSION,
    versionShort: portalVersionShort_(CONFIG.VERSION),
    buildRenderedAt: new Date().toISOString(),
    buildScriptId: ScriptApp.getScriptId()
  });

  Logger.log(
    "PORTAL_RENDER_DIAG version=%s view=%s applicantId=%s isAdmin=%s cache=%s",
    clean_(CONFIG.VERSION || ""),
    clean_(clean_(o.viewName || "portal")),
    clean_(id || "-") || "-",
    isAdminDeployment ? "true" : "false",
    "DISABLED"
  );

  return htmlOutput_(portalHtml);
}

function diagStatus_(e) {
  var activeUser = "";
  var effectiveUser = "";
  var serviceUrl = "";
  var isAdmin = false;
  if (CONFIG.DIAG_RUNTIME !== true) {
    return { ok: false, err: "diag disabled" };
  }
  try { activeUser = clean_(Session.getActiveUser().getEmail() || ""); } catch (_au) {}
  try { effectiveUser = clean_(Session.getEffectiveUser().getEmail() || ""); } catch (_eu) {}
  try { serviceUrl = clean_(ScriptApp.getService().getUrl() || ""); } catch (_su) {}
  try {
    var allowlist = (CONFIG.ADMIN_EMAILS || []).map(function (x) { return clean_(x).toLowerCase(); });
    isAdmin = allowlist.indexOf(clean_(activeUser).toLowerCase()) >= 0;
  } catch (_adm) {}
  var out = {
    ok: true,
    version: CONFIG.VERSION,
    changelog: CONFIG.CHANGELOG_LAST || "",
    nowIso: new Date().toISOString(),
    scriptId: ScriptApp.getScriptId(),
    serviceUrl: serviceUrl,
    studentBaseUrl: getStudentBaseUrl_(),
    user: activeUser,
    effectiveUser: effectiveUser
  };
  if (isAdmin) {
    var propKey = clean_(CONFIG.SCRIPT_PROP_UPLOAD_ROOT_ID || "FODE_UPLOAD_ROOT_ID") || "FODE_UPLOAD_ROOT_ID";
    out.rootPrimaryId = clean_(CONFIG.APPLICANT_ROOT_FOLDER_ID_PRIMARY || "");
    out.rootFallbackId = clean_(CONFIG.APPLICANT_ROOT_FOLDER_ID_FALLBACK || "");
    out.yearFolderName = clean_(CONFIG.APPLICANT_ROOT_YEAR_FOLDER_NAME || "");
    out.autoUploadRootEnabled = CONFIG.AUTO_UPLOAD_ROOT_ENABLED === true;
    out.scriptPropUploadRootKey = propKey;
    out.scriptPropUploadRootId = (typeof getScriptProp_ === "function") ? clean_(getScriptProp_(propKey) || "") : "";
    out.driveAuthHint = "If Drive operations fail with server error, run authDrive() in the Apps Script editor as the owner.";
  }
  return out;
}

function driveProbeFolder_(folderId) {
  var out = {
    ok: false,
    folderId: clean_(folderId || ""),
    canIterate: false
  };
  try {
    var id = clean_(folderId || "");
    if (!id) throw new Error("Missing folderId");
    var folder = DriveApp.getFolderById(id);
    out.name = clean_(folder.getName() || "");
    out.url = clean_(folder.getUrl() || "");
    try {
      var it = folder.getFolders();
      out.canIterate = !!it;
      if (it && typeof it.hasNext === "function") it.hasNext();
    } catch (_iterErr) {
      out.canIterate = false;
    }
    out.ok = true;
    return out;
  } catch (e) {
    var se = safeErr_(e);
    out.errCode = "drive_probe_failed";
    out.errName = clean_(se.name || "Error") || "Error";
    out.errMessage = clean_(se.message || "Drive probe failed") || "Drive probe failed";
    return out;
  }
}

function authDrive() {
  var rootId = clean_(CONFIG.APPLICANT_ROOT_FOLDER_ID_PRIMARY || "");
  var out = {
    ok: false,
    rootId: rootId,
    version: CONFIG.VERSION
  };
  try {
    var myDriveRoot = DriveApp.getRootFolder();
    var _myDriveRootName = clean_(myDriveRoot.getName() || "");
    var root = DriveApp.getFolderById(rootId);
    out.rootName = clean_(root.getName() || "");
    out.rootUrl = clean_(root.getUrl() || "");
    try {
      var it = root.getFolders();
      if (it && typeof it.hasNext === "function") it.hasNext();
    } catch (_iterErr) {}
    out.ok = true;
  } catch (e) {
    var se = safeErr_(e);
    out.errName = clean_(se.name || "Error") || "Error";
    out.errMessage = clean_(se.message || "Drive auth failed") || "Drive auth failed";
  }
  var text = JSON.stringify(out);
  Logger.log(text);
  return text;
}

function authDriveYearFolder() {
  var out = {
    ok: false,
    version: CONFIG.VERSION
  };
  try {
    var rootId = clean_(CONFIG.APPLICANT_ROOT_FOLDER_ID_PRIMARY || "");
    var yearFolderName = clean_(CONFIG.APPLICANT_ROOT_YEAR_FOLDER_NAME || CONFIG.YEAR_FOLDER || "");
    if (!rootId) throw new Error("Missing CONFIG.APPLICANT_ROOT_FOLDER_ID_PRIMARY");
    if (!yearFolderName) throw new Error("Missing year folder config");
    var root = DriveApp.getFolderById(rootId);
    var yearFolder = (typeof getOrCreateFolderByName_ === "function")
      ? getOrCreateFolderByName_(root, yearFolderName, "AUTH")
      : getOrCreateFolder_(root, yearFolderName);
    out.ok = true;
    out.rootId = rootId;
    out.yearFolderId = clean_(yearFolder.getId() || "");
    out.yearFolderName = clean_(yearFolder.getName() || yearFolderName);
  } catch (e) {
    var se = safeErr_(e);
    out.errName = clean_(se.name || "Error") || "Error";
    out.errMessage = clean_(se.message || "Drive year-folder auth failed") || "Drive year-folder auth failed";
  }
  var text = JSON.stringify(out);
  Logger.log(text);
  return text;
}


/******************** PORTAL UPDATE HANDLER ********************/
function parsePortalCsvParam_(raw) {
  return clean_(raw).split(",").map(function (part) {
    return clean_(part);
  }).filter(function (part) {
    return !!part;
  });
}

function portalValidationMessageForCode_(code) {
  var key = clean_(code || "");
  if (key === "DOB_REQUIRED") return "Date of Birth is required.";
  if (key === "DOB_INVALID") return "Enter a valid Date of Birth.";
  if (key === "SUBJECTS_REQUIRED") return "Select at least one subject.";
  if (key === "SUBJECTS_INVALID_FOR_GRADE") return "Selected subjects are not valid for the chosen grade.";
  if (key === "SUBJECT_LOCK_DOCS_VERIFIED") return "Subjects are locked because documents have been verified by Admin.";
  return "Please correct the highlighted fields before submitting.";
}

function sanitizePortalUpdateValue_(field, value) {
  var raw = value;
  var cleaned = clean_(value);
  if (field === "Parent_Phone") cleaned = cleaned.replace(/\s+/g, " ");
  if (field === "Date_Of_Birth") {
    if (!cleaned) return { raw: raw, sanitized: "", omit: false, blank: true, typed: true };
    var iso = toIsoDateInput_(cleaned);
    if (!iso) return { raw: raw, sanitized: cleaned, omit: false, invalid: true, typed: true };
    return { raw: raw, sanitized: iso, omit: false, typed: true };
  }
  if (field === "Physical_Exam_Site") {
    if (!cleaned) return { raw: raw, sanitized: "", omit: true, typed: true };
    return { raw: raw, sanitized: cleaned, omit: false, typed: true };
  }
  if (!cleaned) return { raw: raw, sanitized: "", omit: true };
  return { raw: raw, sanitized: cleaned, omit: false };
}

function normalizePortalSubjectsCsv_(raw) {
  var csv = subjectsToCsv_(raw);
  if (!csv) return "";
  var known = {};
  var ordered = [];
  var configured = CONFIG.PORTAL_SUBJECTS || [];
  for (var i = 0; i < configured.length; i++) {
    var configuredName = clean_(configured[i]);
    if (!configuredName) continue;
    known[configuredName.toLowerCase()] = configuredName;
  }
  var seen = {};
  var parts = csv.split(",");
  for (var j = 0; j < parts.length; j++) {
    var part = clean_(parts[j]);
    if (!part) continue;
    var key = part.toLowerCase();
    if (seen[key]) continue;
    seen[key] = true;
    ordered.push(known[key] || part);
  }
  ordered.sort(function (a, b) {
    var ai = configured.indexOf(a);
    var bi = configured.indexOf(b);
    if (ai >= 0 && bi >= 0) return ai - bi;
    if (ai >= 0) return -1;
    if (bi >= 0) return 1;
    return a.toLowerCase() < b.toLowerCase() ? -1 : (a.toLowerCase() > b.toLowerCase() ? 1 : 0);
  });
  return ordered.join(", ");
}

function validatePortalSubjectsForGrade_(gradeRaw, subjectsCsv) {
  var csv = normalizePortalSubjectsCsv_(subjectsCsv);
  if (!csv) return { ok: false, invalidSubjects: [], reason: "SUBJECTS_REQUIRED" };
  var gradeMatch = clean_(gradeRaw).match(/(\d{1,2})/);
  var gradeNum = gradeMatch ? Number(gradeMatch[1]) : 0;
  var disallow = {};
  if (gradeNum === 7 || gradeNum === 8) {
    ["Biology", "Chemistry", "Physics", "History", "Geography", "Economics", "Business Studies", "Accounting"].forEach(function (name) {
      disallow[name.toLowerCase()] = true;
    });
  } else if (gradeNum === 11 || gradeNum === 12) {
    ["Science", "Social Science"].forEach(function (name) {
      disallow[name.toLowerCase()] = true;
    });
  }
  var invalid = csv.split(",").map(function (part) { return clean_(part); }).filter(function (part) {
    return !!disallow[part.toLowerCase()];
  });
  return {
    ok: invalid.length === 0,
    invalidSubjects: invalid,
    reason: invalid.length ? "SUBJECTS_INVALID_FOR_GRADE" : ""
  };
}

function handlePortalUpdate_(ss, dataSheet, logSheet, payload, postParams, debugId) {
  if (!dataSheet || dataSheet.getName() !== CONFIG.DATA_SHEET) {
    throw new Error("DATA_SHEET mismatch");
  }
  log_(logSheet, "PORTAL_UPDATE payload", payloadSummary_(payload));
  var effectiveEditFields = getPortalEditableFields_().slice();
  ["Date_Of_Birth", "Physical_Exam_Site", "Subjects_Selected_Canonical"].forEach(function (f) {
    if (effectiveEditFields.indexOf(f) === -1) effectiveEditFields.push(f);
  });
  log_(logSheet, "PORTAL_UPDATE editFields", JSON.stringify(effectiveEditFields));

  var id = clean_(payload.id || "");
  var secret = clean_(payload.s || "");
  var safeDebugId = clean_(debugId || newDebugId_()) || newDebugId_();
  var rowIndex = 0;
  var failResult = function (message, code, extras) {
    var extra = (extras && typeof extras === "object") ? extras : {};
    return {
      ok: false,
      debugId: safeDebugId,
      applicantId: id,
      error: {
        message: clean_(message || "Portal update failed."),
        code: clean_(code || "PORTAL_UPDATE_FAILED")
      },
      validationErrors: Array.isArray(extra.validationErrors) ? extra.validationErrors : []
    };
  };
  var logValidationError = function (field, rawValue, sanitizedValue, code, reason) {
    var entry = {
      dbgId: safeDebugId,
      applicantId: id,
      field: clean_(field || ""),
      rawValue: rawValue == null ? "" : String(rawValue),
      sanitizedValue: sanitizedValue == null ? "" : String(sanitizedValue),
      reason: clean_(reason || code || "")
    };
    try { log_(logSheet, "PORTAL_UPDATE_VALIDATION_ERRORS", JSON.stringify(entry)); } catch (_logValidationErr) {}
    return {
      field: entry.field,
      rawValue: entry.rawValue,
      sanitizedValue: entry.sanitizedValue,
      reason: entry.reason,
      code: clean_(code || "VALIDATION_FAILED"),
      message: portalValidationMessageForCode_(code)
    };
  };
  if (!id || !secret) {
    return failResult("Missing portal link parameters. Please reopen your portal link.", "MISSING_TOKEN");
  }

  var found = findPortalRowByIdSecret_(dataSheet, id, secret);
  if (!found) {
    return failResult("No matching record found. Please reopen your portal link.", "ROW_NOT_FOUND");
  }
  rowIndex = found.rowNum;
  log_(logSheet, "PORTAL_UPDATE_TARGET", "row=" + rowIndex + " applicantId=" + id);
  log_(logSheet, "PORTAL_UPDATE rowIndex", String(rowIndex));
  portalDebugLog_("PORTAL_UPDATE_TARGET", {
    applicantId: id,
    rowNumber: rowIndex,
    email: clean_(found.record.Parent_Email_Corrected || found.record.Parent_Email || ""),
    sheet: CONFIG.DATA_SHEET
  });

  if (String(found.record[SCHEMA.PORTAL_ACCESS_STATUS] || "").trim() === "Locked") {
    return failResult("Access suspended. Please contact admissions.", "ACCESS_LOCKED");
  }

  if (isPortalLocked_(found.record)) {
    return failResult("Your record is locked because payment has been verified. No further changes are allowed.", "PAYMENT_VERIFIED_LOCK");
  }

  var posted = (postParams && typeof postParams === "object") ? postParams : {};
  var postKeys = Object.keys(posted || {}).sort();
  if (!postKeys.length && payload.__paramKeys && Array.isArray(payload.__paramKeys)) {
    postKeys = payload.__paramKeys.slice().sort();
  }
  logPortalPostEvent_("PORTAL_POST_KEYS", {
    reqId: clean_(payload.__reqId || ""),
    applicantId: id,
    keys: postKeys
  });
  var postedSample = function (key) {
    if (hasOwn_(posted, key)) return clean_(posted[key]);
    return clean_(payload[key] || "");
  };
  logPortalPostEvent_("PORTAL_POST_SAMPLE", {
    reqId: clean_(payload.__reqId || ""),
    applicantId: id,
    Gender: postedSample("Gender"),
    Date_Of_Birth: postedSample("Date_Of_Birth"),
    Grade_Applying_For: postedSample("Grade_Applying_For"),
    Parent_Phone: postedSample("Parent_Phone"),
    Subjects_Selected_Canonical: postedSample("Subjects_Selected_Canonical"),
    dbg: postedSample("dbg"),
    id: postedSample("id") || clean_(payload.ApplicantID || ""),
    s: redactToken_(hasOwn_(posted, "s") ? posted.s : (payload.s || payload.secret || ""))
  });

  var sourceFields = postKeys.length ? posted : payload;
  var validatedUpdates = {};
  var rawByField = {};
  var validationErrors = [];
  var includeUpdate = function (field, rawValue, sanitizedValue) {
    validatedUpdates[field] = sanitizedValue;
    rawByField[field] = rawValue == null ? "" : String(rawValue);
  };
  var addValidationError = function (field, rawValue, sanitizedValue, code, reason) {
    validationErrors.push(logValidationError(field, rawValue, sanitizedValue, code, reason));
  };

  if (effectiveEditFields.indexOf("Subjects_Selected_Canonical") >= 0 && isDocsVerified_(found.record)) {
    var attemptedSubjectsCanonical = hasOwn_(sourceFields, "Subjects_Selected_Canonical")
      ? clean_(sourceFields.Subjects_Selected_Canonical)
      : "";
    var attemptedSubjectsLegacy = hasOwn_(sourceFields, "Subjects_Selected")
      ? sourceFields.Subjects_Selected
      : (payload.Subjects_Selected || payload.field_Subjects_Selected || "");
    var attemptedSubjectsCsv = normalizePortalSubjectsCsv_(attemptedSubjectsCanonical || subjectsToCsv_(attemptedSubjectsLegacy));
    var existingSubjectsCsv = normalizePortalSubjectsCsv_(clean_(found.record.Subjects_Selected_Canonical || "") || subjectsToCsv_(found.record.Subjects_Selected || ""));
    if (attemptedSubjectsCsv && attemptedSubjectsCsv !== existingSubjectsCsv) {
      var attemptedFields = Object.keys(sourceFields || {}).filter(function (k) {
        return k === "Subjects_Selected_Canonical" || k === "Subjects_Selected" || k === "subj";
      });
      try {
        log_(logSheet, "SUBJECT_LOCK_BLOCK", JSON.stringify({
          applicantId: id,
          debugId: safeDebugId,
          actor: "portal_student",
          attemptedFields: attemptedFields
        }));
      } catch (_subjectLockBlockLogErr) {}
      addValidationError("Subjects_Selected_Canonical", attemptedSubjectsCanonical || attemptedSubjectsLegacy, attemptedSubjectsCsv, "SUBJECT_LOCK_DOCS_VERIFIED", "docs_verified_locked");
      return failResult("Subjects are locked because documents have been verified by Admin.", "SUBJECT_LOCK_DOCS_VERIFIED", {
        validationErrors: validationErrors
      });
    }
  }

  var genericSkip = {
    Parent_Email: true,
    Date_Of_Birth: true,
    Physical_Exam_Site: true,
    Subjects_Selected_Canonical: true
  };
  for (var i = 0; i < effectiveEditFields.length; i++) {
    var h = effectiveEditFields[i];
    if (genericSkip[h]) continue;
    if (!hasOwn_(sourceFields, h)) continue;
    var sanitizedGeneric = sanitizePortalUpdateValue_(h, sourceFields[h]);
    if (sanitizedGeneric.omit) continue;
    includeUpdate(h, sourceFields[h], sanitizedGeneric.sanitized);
  }

  var dobSubmitted = hasOwn_(sourceFields, "Date_Of_Birth");
  var storedDobIso = toIsoDateInput_(found.record.Date_Of_Birth);
  if (dobSubmitted) {
    var dobSanitized = sanitizePortalUpdateValue_("Date_Of_Birth", sourceFields.Date_Of_Birth);
    if (dobSanitized.blank) {
      addValidationError("Date_Of_Birth", sourceFields.Date_Of_Birth, "", "DOB_REQUIRED", "submitted_blank");
    } else if (dobSanitized.invalid) {
      addValidationError("Date_Of_Birth", sourceFields.Date_Of_Birth, dobSanitized.sanitized, "DOB_INVALID", "submitted_invalid");
    } else {
      includeUpdate("Date_Of_Birth", sourceFields.Date_Of_Birth, dobSanitized.sanitized);
    }
  } else if (!storedDobIso) {
    addValidationError("Date_Of_Birth", found.record.Date_Of_Birth, "", "DOB_REQUIRED", "effective_blank_on_submit");
  }

  if (hasOwn_(sourceFields, "Physical_Exam_Site")) {
    var examSanitized = sanitizePortalUpdateValue_("Physical_Exam_Site", sourceFields.Physical_Exam_Site);
    if (!examSanitized.omit) includeUpdate("Physical_Exam_Site", sourceFields.Physical_Exam_Site, examSanitized.sanitized);
  }

  var storedSubjectsCsv = normalizePortalSubjectsCsv_(clean_(found.record.Subjects_Selected_Canonical || "") || subjectsToCsv_(found.record.Subjects_Selected || ""));
  var hasSubmittedSubjects = hasOwn_(sourceFields, "Subjects_Selected_Canonical") || hasOwn_(sourceFields, "Subjects_Selected") || hasOwn_(sourceFields, "subj");
  var submittedSubjectsRaw = hasOwn_(sourceFields, "Subjects_Selected_Canonical")
    ? sourceFields.Subjects_Selected_Canonical
    : (hasOwn_(sourceFields, "Subjects_Selected") ? sourceFields.Subjects_Selected : (payload.Subjects_Selected || payload.field_Subjects_Selected || ""));
  var submittedSubjectsCsv = normalizePortalSubjectsCsv_(submittedSubjectsRaw);
  var effectiveGrade = clean_(validatedUpdates.Grade_Applying_For || found.record.Grade_Applying_For || "");
  if (hasSubmittedSubjects) {
    if (!submittedSubjectsCsv) {
      addValidationError("Subjects_Selected_Canonical", submittedSubjectsRaw, submittedSubjectsCsv, "SUBJECTS_REQUIRED", "submitted_blank");
    } else {
      var subjectsValidation = validatePortalSubjectsForGrade_(effectiveGrade, submittedSubjectsCsv);
      if (!subjectsValidation.ok) {
        if (submittedSubjectsCsv !== storedSubjectsCsv) {
          addValidationError("Subjects_Selected_Canonical", submittedSubjectsRaw, submittedSubjectsCsv, "SUBJECTS_INVALID_FOR_GRADE", subjectsValidation.invalidSubjects.join(", "));
        }
      } else {
        includeUpdate("Subjects_Selected_Canonical", submittedSubjectsRaw, submittedSubjectsCsv);
      }
    }
  } else if (!storedSubjectsCsv) {
    addValidationError("Subjects_Selected_Canonical", "", "", "SUBJECTS_REQUIRED", "effective_blank_on_submit");
  }

  if (validationErrors.length) {
    return failResult("Please correct the highlighted fields before submitting.", validationErrors[0].code || "VALIDATION_FAILED", {
      validationErrors: validationErrors
    });
  }
  validatedUpdates.PortalLastUpdateAt = new Date().toISOString();
  rawByField.PortalLastUpdateAt = validatedUpdates.PortalLastUpdateAt;
  if (!clean_(found.record.Portal_Submitted)) {
    validatedUpdates.Portal_Submitted = new Date().toISOString();
    rawByField.Portal_Submitted = validatedUpdates.Portal_Submitted;
  }

  var updateKeys = Object.keys(validatedUpdates);
  var emailBefore = clean_(found.record.Parent_Email_Corrected || "");
  var emailAfter = hasOwn_(validatedUpdates, "Parent_Email_Corrected") ? clean_(validatedUpdates.Parent_Email_Corrected) : emailBefore;
  var emailChanged = hasOwn_(validatedUpdates, "Parent_Email_Corrected")
    && emailAfter.toLowerCase() !== emailBefore.toLowerCase();
  var patchSample = {};
  for (var ps = 0; ps < updateKeys.length && ps < 5; ps++) {
    var patchKey = updateKeys[ps];
    var patchVal = clean_(validatedUpdates[patchKey]);
    patchSample[patchKey] = patchVal.length > 120 ? patchVal.slice(0, 120) : patchVal;
  }
  portalDebugLog_("PORTAL_UPDATE_PATCH", {
    applicantId: id,
    rowNumber: rowIndex,
    keys: updateKeys,
    patchSample: patchSample
  });

  log_(logSheet, "PORTAL_UPDATE_PATCH", "keys=" + updateKeys.join(","));
  log_(logSheet, "PORTAL_UPDATE updates", JSON.stringify(validatedUpdates));
  var beforeReceiptRow = {
    ApplicantID: clean_(found.record.ApplicantID || id || ""),
    First_Name: clean_(found.record.First_Name || ""),
    Last_Name: clean_(found.record.Last_Name || ""),
    Fee_Receipt_File: clean_(found.record.Fee_Receipt_File || "")
  };
  var headers = dataSheet.getRange(1, 1, 1, dataSheet.getLastColumn()).getValues()[0];
  try {
    for (var uk = 0; uk < updateKeys.length; uk++) {
      var key = updateKeys[uk];
      var colIndex = headers.indexOf(key);
      if (colIndex < 0) continue;
      try {
        dataSheet.getRange(rowIndex, colIndex + 1).setValue(validatedUpdates[key]);
      } catch (writeFieldErr) {
        var writeEntry = {
          dbgId: safeDebugId,
          applicantId: id,
          rowIndex: rowIndex,
          field: key,
          rawValue: rawByField[key] == null ? "" : String(rawByField[key]),
          sanitizedValue: validatedUpdates[key] == null ? "" : String(validatedUpdates[key]),
          error: String(writeFieldErr && writeFieldErr.message ? writeFieldErr.message : writeFieldErr),
          stack: clean_(writeFieldErr && writeFieldErr.stack ? writeFieldErr.stack : "")
        };
        try { log_(logSheet, "PORTAL_UPDATE_WRITE_ERROR", JSON.stringify(writeEntry)); } catch (_writeLogErr) {}
        portalDebugLog_("PORTAL_UPDATE_WRITE_ERROR", writeEntry);
        throw writeFieldErr;
      }
    }
    SpreadsheetApp.flush();
  } catch (e) {
    try {
      log_(logSheet, "PORTAL_UPDATE_WRITE_ERROR", JSON.stringify({
        dbgId: safeDebugId,
        applicantId: id,
        rowIndex: rowIndex,
        error: String(e && e.message ? e.message : e),
        stack: clean_(e && e.stack ? e.stack : "")
      }));
    } catch (_writeSummaryErr) {}
    return failResult("We could not save your update. Please try again or contact admissions.", "WRITE_FAILED");
  }
  if (emailChanged) {
    var docsKey = buildDocsFollowupKey_(id);
    try {
      PropertiesService.getScriptProperties().deleteProperty(docsKey);
    } catch (_propDelErr) {}
    try {
      log_(logSheet, "DOCS_FOLLOWUP_RESET_EMAIL_CHANGE", JSON.stringify({
        applicantId: id,
        rowNumber: rowIndex,
        key: docsKey,
        oldEmail: emailBefore,
        newEmail: emailAfter,
        debugId: safeDebugId
      }));
    } catch (_docsResetLogErr) {}
  }
  try {
    if (Object.prototype.hasOwnProperty.call(validatedUpdates, "Fee_Receipt_File")) {
      var afterReceiptRow = getRowObject_(dataSheet, rowIndex);
      maybeNotifyPaymentReceiptUploadTransition_(beforeReceiptRow, afterReceiptRow, rowIndex, { source: "portal_update" });
    }
  } catch (receiptAlertErr) {
    portalDebugLog_("PAYMENT_RECEIPT_ALERT_ERROR", {
      applicantId: id,
      rowNumber: rowIndex,
      error: String(receiptAlertErr && receiptAlertErr.message ? receiptAlertErr.message : receiptAlertErr)
    });
  }
  portalDebugLog_("PORTAL_UPDATE_RESULT", {
    applicantId: id,
    rowNumber: rowIndex,
    ok: true,
    saved: 1
  });
  log_(logSheet, "PORTAL_UPDATE_RESULT", "updated=" + updateKeys.length);
  log_(logSheet, "PORTAL UPDATE", "ApplicantID=" + id + " viaSecret=yes");
  return {
    ok: true,
    debugId: safeDebugId,
    applicantId: id,
    rowNumber: rowIndex,
    saved: 1
  };
}

function isAllowedPortalUploadField_(fieldKey) {
  var key = clean_(fieldKey || "");
  var allow = CONFIG.PORTAL_UPLOAD_KEYS || [];
  return allow.indexOf(key) >= 0;
}

function savePortalUpload_(applicantId, fieldKey, fileName, mimeType, bytes, ctx) {
  var id = clean_(applicantId || "");
  var key = clean_(fieldKey || "");
  var context = ctx || {};
  if (!id) throw new Error("Missing ApplicantID");
  if (!isAllowedPortalUploadField_(key)) throw new Error("Invalid upload field");

  var ss = context.ss || getWorkingSpreadsheet_();
  var sheet = context.sheet || mustGetSheet_(ss, CONFIG.DATA_SHEET);
  var dbg = clean_(context.dbg || "");
  var preferRestOnly = context.preferRest === true && CONFIG.DRIVE_REST_FALLBACK_ENABLED === true && CONFIG.PORTAL_UPLOAD_PREFER_REST === true;
  var onStage = (typeof context.onStage === "function") ? context.onStage : null;
  function emitUploadStage_(name, extra) {
    if (!onStage) return;
    try { onStage(clean_(name), extra && typeof extra === "object" ? extra : {}); } catch (_stageErr) {}
  }
  var rowNumber = Number(context.rowNumber || 0);
  var rowObj = context.rowObj || (rowNumber >= 2 ? getRowObject_(sheet, rowNumber) : null) || {};
  if (!rowNumber && id) rowNumber = findRowByApplicantId_(sheet, id);
  if (!rowNumber || rowNumber < 2) throw new Error("Applicant row not found");

  var tFolder = nowMs_();
  var folderUrl = clean_(rowObj.Folder_Url || "");
  var folderId = folderIdFromUrl_(folderUrl);
  var folder = null;
  var folderHandle = null;
  var folderIdKnown = !!folderId;
  emitUploadStage_("folder", {
    applicantId: id,
    field: key,
    folderIdKnown: folderIdKnown
  });
  if (preferRestOnly) {
    if (folderId) {
      folderHandle = driveApiBuildFolderHandleById_(folderId, dbg, folderUrl);
    }
  } else if (folderId) {
    try {
      folder = withRetries_(function () { return DriveApp.getFolderById(folderId); }, { dbg: dbg, label: "savePortalUpload:getFolderById" });
      withRetries_(function () { return folder.getName(); }, { dbg: dbg, label: "savePortalUpload:folderGetName" });
      folderHandle = {
        kind: "driveapp",
        folder: folder,
        id: clean_(folder.getId() || folderId),
        url: clean_(folder.getUrl() || folderUrl)
      };
    } catch (e) {
      folder = null;
      if (typeof isDriveServerError_ === "function" && isDriveServerError_(e) && CONFIG.DRIVE_REST_FALLBACK_ENABLED === true) {
        folderHandle = driveApiBuildFolderHandleById_(folderId, dbg, folderUrl);
      }
    }
  }
  if (!folderHandle) {
    if (preferRestOnly) {
      folderHandle = createApplicantFolderHandleWithRestFallback_(rowObj, dbg, id);
    } else {
      try {
        folder = createApplicantFolder_(rowObj, { dbg: dbg });
        folderHandle = {
          kind: "driveapp",
          folder: folder,
          id: clean_(folder.getId() || ""),
          url: clean_(folder.getUrl() || "")
        };
      } catch (createErr) {
        if (typeof isDriveServerError_ === "function" && isDriveServerError_(createErr) && CONFIG.DRIVE_REST_FALLBACK_ENABLED === true) {
          folderHandle = createApplicantFolderHandleWithRestFallback_(rowObj, dbg, id);
        } else {
          throw createErr;
        }
      }
    }
    if (!folderHandle) throw new Error("folder_root_unusable: folder handle missing");
    writeBack_(sheet, rowNumber, { Folder_Url: clean_(folderHandle.url || "") });
    rowObj.Folder_Url = clean_(folderHandle.url || "");
  }
  logExecTrace_("UPLOAD_T_FOLDER", dbg, {
    applicantId: id,
    field: key,
    ms: elapsedMs_(tFolder),
    folderIdKnown: folderIdKnown
  });

  var meta = docMetaByField_(key) || {};
  var prefix = safeFileName_(id) + "__" + safeFileName_(key) + "__";
  var finalName = prefix + safeFileName_(fileName || ((meta.label || key) + ".bin"));
  var blob = Utilities.newBlob(bytes || [], clean_(mimeType || "application/octet-stream") || "application/octet-stream", finalName);
  emitUploadStage_("drive", {
    applicantId: id,
    field: key
  });
  var tDrive = nowMs_();
  var file = null;
  var fileInfo = null;
  if (folderHandle.kind === "driveapp") {
    file = withRetries_(function () {
      return folderHandle.folder.createFile(blob);
    }, { dbg: dbg, label: "savePortalUpload:createFile" });
    fileInfo = {
      fileId: clean_(file.getId() || ""),
      fileUrl: clean_(file.getUrl() || ""),
      fileName: clean_(file.getName() || finalName)
    };
  } else if (folderHandle.kind === "rest") {
    fileInfo = driveApiUploadBlobToFolder_(clean_(folderHandle.id || ""), finalName, blob, dbg);
  } else {
    throw new Error("folder_root_unusable: unknown folder handle kind");
  }
  logExecTrace_("UPLOAD_T_DRIVE", dbg, {
    applicantId: id,
    field: key,
    ms: elapsedMs_(tDrive)
  });
  return {
    fileId: clean_(fileInfo && fileInfo.fileId || ""),
    fileUrl: clean_(fileInfo && fileInfo.fileUrl || ""),
    fileName: clean_(fileInfo && fileInfo.fileName || finalName),
    rowNumber: rowNumber,
    folderUrl: clean_(folderHandle && folderHandle.url || ""),
    driveMode: folderHandle && folderHandle.kind === "rest" ? "REST" : "DRIVEAPP"
  };
}

function applyPortalUploadSheetUpdate_(sheet, rowNumber, rowObj, fieldKey, uploadResult, opts) {
  var sh = sheet;
  var rowNum = Number(rowNumber || 0);
  var record = rowObj || getRowObject_(sh, rowNum);
  var key = clean_(fieldKey || "");
  var res = uploadResult || {};
  var logSheet = (opts && opts.logSheet) || mustGetSheet_(getWorkingSpreadsheet_(), CONFIG.LOG_SHEET);
  var dbg = clean_((opts && opts.dbg) || "");
  var docMeta = docMetaByField_(key);
  if (!docMeta) throw new Error("Invalid document field");

  var oldCell = clean_(record[key] || "");
  var isMultiple = docMeta.multiple === true;
  var fileUrl = clean_(res.fileUrl || "");
  if (!fileUrl) throw new Error("Missing uploaded file URL");
  var updates = {};
  updates[key] = isMultiple ? appendUrlToCell_(oldCell, fileUrl) : fileUrl;
  updates.PortalLastUpdateAt = new Date().toISOString();
  if (!clean_(record.Portal_Submitted)) updates.Portal_Submitted = new Date().toISOString();
  if (docMeta.status && hasHeader_(sh, docMeta.status)) updates[docMeta.status] = "PENDING_REVIEW";
  var line = new Date().toISOString()
    + " | " + key
    + " | " + (isMultiple ? "uploaded" : "replaced")
    + " | old=" + (oldCell || "-")
    + " | new=" + fileUrl;
  updates.File_Log = appendLog_(clean_(record.File_Log || ""), line);

  var receiptBeforeRow = null;
  if (key === "Fee_Receipt_File") {
    receiptBeforeRow = {
      ApplicantID: clean_(record.ApplicantID || ""),
      First_Name: clean_(record.First_Name || ""),
      Last_Name: clean_(record.Last_Name || ""),
      Fee_Receipt_File: clean_(record.Fee_Receipt_File || "")
    };
  }

  var tSheet = nowMs_();
  writeBack_(sh, rowNum, updates);
  SpreadsheetApp.flush();
  var verifyRow = getRowObject_(sh, rowNum);
  logExecTrace_("UPLOAD_T_SHEET", dbg, {
    applicantId: clean_(record.ApplicantID || ""),
    field: key,
    ms: elapsedMs_(tSheet)
  });
  var verifyCell = clean_(verifyRow[key] || "");
  if (verifyCell.indexOf(fileUrl) < 0) throw new Error("Upload URL was not saved. Please try again.");

  log_(logSheet, "PORTAL_UPLOAD_OK", JSON.stringify({
    applicantId: clean_(record.ApplicantID || ""),
    fieldKey: key,
    fileId: clean_(res.fileId || ""),
    rowNumber: rowNum
  }));

  if (key === "Fee_Receipt_File") {
    try {
      maybeNotifyPaymentReceiptUploadTransition_(receiptBeforeRow || {}, verifyRow, rowNum, { source: "portalUploadFile_" });
    } catch (receiptAlertErr) {
      log_(logSheet, "PAYMENT_RECEIPT_ALERT_ERROR", String(receiptAlertErr && receiptAlertErr.message ? receiptAlertErr.message : receiptAlertErr));
    }
  }

  return {
    fileUrl: fileUrl,
    fileId: clean_(res.fileId || ""),
    rowNumber: rowNum,
    fieldKey: key,
    multiple: isMultiple,
    currentUrls: normalizeToUrlList_(verifyRow[key] || "")
  };
}

/******************** PORTAL B64 UPLOAD (student portal) ********************/
function portalUploadExt_(fileName) {
  var n = clean_(fileName || "").toLowerCase();
  var idx = n.lastIndexOf(".");
  if (idx < 0 || idx === n.length - 1) return "";
  return n.slice(idx + 1);
}

function sanitizePortalUploadFileName_(fileName, fieldName) {
  var raw = clean_(fileName || "");
  raw = raw.replace(/[\\\/:*?"<>|]+/g, "_");
  raw = raw.replace(/\s+/g, " ").trim();
  if (!raw) raw = clean_(fieldName || "upload") + ".bin";
  var ext = portalUploadExt_(raw);
  var base = raw;
  if (ext) base = raw.slice(0, -(ext.length + 1));
  base = safeFileName_(base || "upload");
  if (!base) base = "upload";
  if (base.length > 80) base = base.slice(0, 80);
  if (ext) {
    ext = ext.replace(/[^a-z0-9]/g, "");
    if (ext.length > 10) ext = ext.slice(0, 10);
    return ext ? (base + "." + ext) : base;
  }
  return base;
}

function isPortalUploadTypeAllowed_(mimeType, fileName) {
  var mime = clean_(mimeType || "").toLowerCase();
  var ext = portalUploadExt_(fileName);
  var mimeAllow = (CONFIG.PORTAL_ALLOWED_MIME || []).map(function (x) { return clean_(x).toLowerCase(); });
  var extAllow = (CONFIG.PORTAL_ALLOWED_EXT || []).map(function (x) { return clean_(x).toLowerCase(); });
  return (mime && mimeAllow.indexOf(mime) >= 0) || (ext && extAllow.indexOf(ext) >= 0);
}

function portalUploadBase64(data) {
  Logger.log("UPLOAD_B64_RPC_ENTER " + JSON.stringify({
    id: data && data.id,
    field: data && data.field,
    name: data && data.name,
    mime: data && data.mime,
    b64Len: data && data.base64 ? data.base64.length : 0
  }));
  return portalUpload_handleBase64_(data);
}

function portalUpload_handleBase64_(data) {
  var payload = (data && typeof data === "object") ? data : {};
  var dbg = makeDebugId_();
  var applicantId = clean_(payload.id || payload.applicantId || "");
  var secret = clean_(payload.s || payload.secret || "");
  var fieldName = clean_(payload.field || payload.docKey || "");
  var fileName = clean_(payload.name || "");
  var mimeType = clean_(payload.mime || payload.mimeType || "");
  var b64 = String(payload.base64 || payload.b64 || "");

  function fail_(code, message, extra) {
    var out = {
      ok: false,
      dbg: dbg,
      code: clean_(code || "UPLOAD_FAILED"),
      message: clean_(message || "Upload failed.")
    };
    if (extra && typeof extra === "object") {
      if (extra.field) out.field = clean_(extra.field);
      if (extra.fileUrl) out.fileUrl = clean_(extra.fileUrl);
    }
    logExecTrace_("UPLOAD_B64_FAIL", dbg, {
      dbg: dbg,
      code: out.code,
      field: fieldName,
      id: applicantId
    });
    return out;
  }

  logExecTrace_("UPLOAD_B64_ENTER", dbg, {
    dbg: dbg,
    id: applicantId,
    field: fieldName,
    name: fileName,
    mime: mimeType,
    b64Len: Number(b64.length || 0)
  });

  if (!applicantId || !secret || !fieldName || !fileName || !b64) {
    return fail_("BAD_REQUEST", "Missing upload fields.", { field: fieldName });
  }
  if (!isAllowedPortalUploadField_(fieldName) || !docMetaByField_(fieldName)) {
    return fail_("INVALID_FIELD", "Invalid document field.", { field: fieldName });
  }
  if (/^data:/i.test(b64) && b64.indexOf(",") >= 0) b64 = b64.split(",").slice(1).join(",");

  var bytes = [];
  try {
    bytes = Utilities.base64Decode(b64);
  } catch (_b64Err) {
    return fail_("BAD_BASE64", "Upload payload is invalid. Please retry.", { field: fieldName });
  }
  if (!bytes || !bytes.length) {
    return fail_("NO_FILE", "Please select a file.", { field: fieldName });
  }
  if (Number(bytes.length || 0) > Number(CONFIG.PORTAL_MAX_UPLOAD_BYTES || (5 * 1024 * 1024))) {
    return fail_("FILE_TOO_LARGE", "File too large (max 5 MB). Please compress and retry.", { field: fieldName });
  }
  if (!isPortalUploadTypeAllowed_(mimeType, fileName)) {
    return fail_("UNSUPPORTED_TYPE", "Unsupported file type. Use PDF/JPG/PNG (or DOC/DOCX if enabled).", { field: fieldName });
  }

  var ss = getWorkingSpreadsheet_();
  var sheet = mustGetDataSheet_(ss);
  var logSheet = mustGetSheet_(ss, CONFIG.LOG_SHEET);
  var found = findPortalRowByIdSecret_(sheet, applicantId, secret);
  if (!found) return fail_("TOKEN_INVALID", "Invalid or expired portal link token.", { field: fieldName });
  if (String(found.record[SCHEMA.PORTAL_ACCESS_STATUS] || "").trim() === "Locked") {
    return fail_("ACCESS_LOCKED", "Portal access is locked. Please contact admissions.", { field: fieldName });
  }
  if (isPortalLocked_(found.record)) {
    return fail_("ACCESS_LOCKED", "Portal access is locked. Please contact admissions.", { field: fieldName });
  }
  if (isPaymentFreezeActive_(found.record)) {
    return fail_("PAYMENT_FREEZE", "Uploads are disabled after payment verification.", { field: fieldName });
  }

  try {
    var safeName = sanitizePortalUploadFileName_(fileName, fieldName);
    var uploadRes = savePortalUpload_(applicantId, fieldName, safeName, mimeType || "application/octet-stream", bytes, {
      ss: ss,
      sheet: sheet,
      rowNumber: found.rowNum,
      rowObj: found.record,
      dbg: dbg
    });
    var applyRes = applyPortalUploadSheetUpdate_(sheet, found.rowNum, found.record, fieldName, uploadRes, {
      logSheet: logSheet,
      dbg: dbg
    });
    logExecTrace_("UPLOAD_B64_OK", dbg, {
      dbg: dbg,
      fileId: clean_(applyRes.fileId || uploadRes.fileId || "")
    });
    return {
      ok: true,
      dbg: dbg,
      field: fieldName,
      fileId: clean_(applyRes.fileId || uploadRes.fileId || ""),
      fileUrl: clean_(applyRes.fileUrl || uploadRes.fileUrl || ""),
      name: clean_(uploadRes.fileName || safeName),
      currentUrls: Array.isArray(applyRes.currentUrls) ? applyRes.currentUrls : []
    };
  } catch (e) {
    return fail_("UPLOAD_FAILED", clean_(stringifyGsError_(e) || "Upload failed."), { field: fieldName });
  }
}

/******************** DRIVE UPLOAD (called via google.script.run) ********************/
function uploadPortalFile(applicantId, secret, fieldName, fileName, mimeType, base64Data) {
  applicantId = clean_(applicantId);
  secret = clean_(secret);
  fieldName = clean_(fieldName);
  var fileNames = Array.isArray(fileName) ? fileName : [fileName];
  var mimeTypes = Array.isArray(mimeType) ? mimeType : [mimeType];
  var base64List = Array.isArray(base64Data) ? base64Data : [base64Data];
  var dbgId = newDebugId_();

  try {
    var ss = getWorkingSpreadsheet_();
    var sheet = mustGetDataSheet_(ss);
    var logSheet = mustGetSheet_(ss, CONFIG.LOG_SHEET);

    var found = findPortalRowByIdSecret_(sheet, applicantId, secret);
    if (!found) throw new Error("Record not found.");
    if (String(found.record[SCHEMA.PORTAL_ACCESS_STATUS] || "").trim() === "Locked") throw new Error("Access suspended.");
    if (isPortalLocked_(found.record)) throw new Error("Record locked.");

    var docMeta = docMetaByField_(fieldName);
    var isMultiple = !!(docMeta && docMeta.multiple === true);
    if (fileNames.length !== mimeTypes.length || fileNames.length !== base64List.length) {
      throw new Error("Upload payload mismatch.");
    }
    if (!fileNames.length) throw new Error("No files selected.");

    var folderUrl = clean_(found.record.Folder_Url || "");
    var folderId = folderIdFromUrl_(folderUrl);
    if (!folderId) throw new Error("Applicant folder missing. Please contact admissions.");
    var folder;
    try {
      folder = DriveApp.getFolderById(folderId);
      folder.getName();
    } catch (folderErr) {
      throw new Error("Applicant folder is unavailable or inaccessible.");
    }

    var createdUrls = [];
    try {
      for (var i = 0; i < fileNames.length; i++) {
        var fName = clean_(fileNames[i]) || ("upload_" + Date.now() + "_" + i);
        var fType = clean_(mimeTypes[i]) || "application/octet-stream";
        var b64 = String(base64List[i] || "");
        if (!b64) continue;
        var bytes = Utilities.base64Decode(b64);
        var blob = Utilities.newBlob(bytes, fType, fName);
        var file = folder.createFile(blob);
        createdUrls.push(file.getUrl());
      }
    } catch (driveErr) {
      throw new Error("Drive upload failed: " + String(driveErr && driveErr.message ? driveErr.message : driveErr));
    }
    if (!createdUrls.length) throw new Error("No valid files to upload.");

    var updates = {};
    var oldCell = clean_(found.record[fieldName] || "");
    if (isMultiple) {
      var merged = oldCell;
      for (var j = 0; j < createdUrls.length; j++) {
        merged = appendUrlToCell_(merged, createdUrls[j]);
      }
      updates[fieldName] = merged;
    } else {
      updates[fieldName] = createdUrls[createdUrls.length - 1];
    }
    updates.PortalLastUpdateAt = new Date().toISOString();
    if (!clean_(found.record.Portal_Submitted)) updates.Portal_Submitted = new Date().toISOString();

    // If status column exists, set PENDING_REVIEW
    if (docMeta && hasHeader_(sheet, docMeta.status)) updates[docMeta.status] = "PENDING_REVIEW";

    var fileLog = clean_(found.record.File_Log || "");
    for (var k = 0; k < createdUrls.length; k++) {
      var line = new Date().toISOString()
        + " | " + fieldName
        + " | " + (isMultiple ? "uploaded" : "replaced")
        + " | old=" + (oldCell || "-")
        + " | new=" + createdUrls[k];
      fileLog = appendLog_(fileLog, line);
    }
    updates.File_Log = fileLog;

    var receiptBeforeRow = null;
    if (fieldName === "Fee_Receipt_File") {
      receiptBeforeRow = {
        ApplicantID: clean_(found.record.ApplicantID || applicantId || ""),
        First_Name: clean_(found.record.First_Name || ""),
        Last_Name: clean_(found.record.Last_Name || ""),
        Fee_Receipt_File: clean_(found.record.Fee_Receipt_File || "")
      };
    }
    writeBack_(sheet, found.rowNum, updates);
    SpreadsheetApp.flush();
    var verifyRow = getRowObject_(sheet, found.rowNum);
    var verifyCell = clean_(verifyRow[fieldName] || "");
    var latestUrl = clean_(createdUrls[createdUrls.length - 1] || "");
    if (!latestUrl || verifyCell.indexOf(latestUrl) < 0) {
      throw new Error("Upload URL was not saved. Please try again.");
    }
    log_(logSheet, "PORTAL UPLOAD", "ApplicantID=" + applicantId + " field=" + fieldName + " files=" + createdUrls.length);
    if (fieldName === "Fee_Receipt_File") {
      try {
        maybeNotifyPaymentReceiptUploadTransition_(receiptBeforeRow || {}, verifyRow, found.rowNum, { source: "uploadPortalFile" });
      } catch (receiptUploadAlertErr) {
        log_(logSheet, "PAYMENT_RECEIPT_ALERT_ERROR", String(receiptUploadAlertErr && receiptUploadAlertErr.message ? receiptUploadAlertErr.message : receiptUploadAlertErr));
      }
    }

    return {
      ok: true,
      field: fieldName,
      url: createdUrls[createdUrls.length - 1],
      urls: createdUrls,
      multiple: isMultiple
    };
  } catch (e) {
    var msg = String(e && e.message ? e.message : e);
    var stack = String(e && e.stack ? e.stack : "");
    logPortalUploadError_(dbgId, applicantId, fieldName, msg, {
      code: "UPLOAD_FAILED",
      stack: stack,
      fileCount: fileNames.length,
      hasToken: !!(applicantId && secret)
    });
    return {
      ok: false,
      debugId: dbgId,
      field: fieldName,
      error: "Upload failed. Please try again. Debug: " + dbgId,
      redirectUrl: buildPortalRedirectUrl_(applicantId, secret, {
        error: true,
        dbg: dbgId,
        uploadFail: true,
        field: fieldName
      })
    };
  }
}

function portal_deleteUploadedFile(payload) {
  payload = payload || {};
  var applicantId = clean_(payload.applicantId || "");
  var secret = clean_(payload.secret || payload.s || "");
  var fieldName = clean_(payload.field || "");
  var targetUrl = clean_(payload.url || "");
  var rowNumber = Number(payload.rowNumber || 0);
  var dbgId = newDebugId_();

  if (!applicantId || !secret || !fieldName || !targetUrl) {
    return {
      ok: false,
      debugId: dbgId,
      error: "Missing delete payload fields",
      redirectUrl: buildPortalRedirectUrl_(applicantId, secret, { error: true, dbg: dbgId })
    };
  }

  var ss = getWorkingSpreadsheet_();
  var sheet = mustGetDataSheet_(ss);
  var logSheet = mustGetSheet_(ss, CONFIG.LOG_SHEET);
  var found = findPortalRowByIdSecret_(sheet, applicantId, secret);
  if (!found) return { ok: false, debugId: dbgId, error: "Record not found.", redirectUrl: buildPortalRedirectUrl_(applicantId, secret, { error: true, dbg: dbgId }) };
  if (rowNumber >= 2 && rowNumber !== found.rowNum) return { ok: false, debugId: dbgId, error: "Row mismatch.", redirectUrl: buildPortalRedirectUrl_(applicantId, secret, { error: true, dbg: dbgId }) };
  portalDebugLog_("PORTAL_DELETE_TARGET", {
    applicantId: applicantId,
    rowNumber: found.rowNum,
    fileField: fieldName,
    url: targetUrl
  });
  if (String(found.record[SCHEMA.PORTAL_ACCESS_STATUS] || "").trim() === "Locked") return { ok: false, debugId: dbgId, error: "Locked", redirectUrl: buildPortalRedirectUrl_(applicantId, secret, { error: true, dbg: dbgId }) };
  if (isPortalLocked_(found.record)) return { ok: false, debugId: dbgId, error: "Locked", redirectUrl: buildPortalRedirectUrl_(applicantId, secret, { error: true, dbg: dbgId }) };

  try {
    var docMeta = docMetaByField_(fieldName);
    if (!docMeta) return { ok: false, debugId: dbgId, error: "Invalid field.", redirectUrl: buildPortalRedirectUrl_(applicantId, secret, { error: true, dbg: dbgId }) };

    var existing = clean_(found.record[fieldName] || "");
    var updatedCell = removeUrlFromCell_(existing, targetUrl);
    var remainingUrls = normalizeToUrlList_(updatedCell);
    var removed = existing !== updatedCell;
    if (!removed) return { ok: false, debugId: dbgId, error: "File URL not found.", redirectUrl: buildPortalRedirectUrl_(applicantId, secret, { error: true, dbg: dbgId }) };

    var updates = {};
    updates[fieldName] = updatedCell;
    updates.PortalLastUpdateAt = new Date().toISOString();
    if (docMeta.status && hasHeader_(sheet, docMeta.status)) updates[docMeta.status] = "PENDING_REVIEW";
    var line = new Date().toISOString() + " | " + fieldName + ": DELETE " + targetUrl;
    updates.File_Log = appendLog_(clean_(found.record.File_Log || ""), line);
    writeBack_(sheet, found.rowNum, updates);

    var trashed = false;
    var warning = "";
    var fileId = extractDriveFileId_(targetUrl);
    if (fileId) {
      try {
        var file = DriveApp.getFileById(fileId);
        var folderId = folderIdFromUrl_(clean_(found.record.Folder_Url || ""));
        if (folderId && isFileInFolderChain_(file, folderId)) {
          file.setTrashed(true);
          trashed = true;
        } else {
          warning = "File not in applicant folder; URL removed only.";
        }
      } catch (e) {
        warning = "Could not trash file: " + (e && e.message ? e.message : String(e));
      }
    } else {
      warning = "Could not parse file id; URL removed only.";
    }

    portalDebugLog_("PORTAL_DELETE_RESULT", {
      applicantId: applicantId,
      rowNumber: found.rowNum,
      fileField: fieldName,
      removed: removed,
      trashed: trashed,
      warning: warning
    });
    log_(logSheet, "PORTAL_DOC_DELETE", "ApplicantID=" + applicantId + " field=" + fieldName + " trashed=" + trashed);
    return { ok: true, remainingUrls: remainingUrls, trashed: trashed, warning: warning };
  } catch (e2) {
    portalDebugLog_("PORTAL_DELETE_ERROR", {
      applicantId: applicantId,
      rowNumber: found.rowNum,
      fileField: fieldName,
      error: String(e2 && e2.message ? e2.message : e2)
    });
    return {
      ok: false,
      debugId: dbgId,
      error: String(e2 && e2.message ? e2.message : e2),
      redirectUrl: buildPortalRedirectUrl_(applicantId, secret, { error: true, dbg: dbgId })
    };
  }
}

/******************** PORTAL HTML ********************/
function portalVersionShort_(rawVersion) {
  var v = String(rawVersion || "");
  var tail = v.split("-").pop();
  if (/^r\d+$/i.test(tail)) return tail;
  var m = v.match(/-r(\d+)\b/i);
  return m ? ("r" + m[1]) : "r?";
}

function renderPortalHtml_(opts) {
  var id = opts.id, secret = opts.secret, record = opts.record;
  var reqId = clean_(opts.reqId || "");
  var debugPage = opts.debugPage === true;
  var saved = opts.saved === true;
  var hasErr = opts.errorFlag === true;
  var lockedFlag = opts.lockedFlag === true;
  var msgToken = clean_(opts.msgToken || "");
  var dbg = clean_(opts.dbg || "");
  var uploadFail = opts.uploadFail === true;
  var uploadField = clean_(opts.uploadField || "");
  var uploadResult = opts.uploadResult === true;
  var uploadOk = opts.uploadOk === true;
  var uploadDocKey = clean_(opts.uploadDocKey || "");
  var uploadErrCode = clean_(opts.uploadErrCode || "");
  var validationFlag = opts.validationFlag === true;
  var validationFields = Array.isArray(opts.validationFields) ? opts.validationFields : [];
  var validationCodes = Array.isArray(opts.validationCodes) ? opts.validationCodes : [];
  var subjects = opts.subjects || [];
  var examSites = opts.examSites || [];
  var editFields = opts.editFields || [];
  var docs = opts.docs || [];
  var visibleFields = opts.visibleFields || [];
  var subjectsLocked = opts.subjectsLocked === true;
  var error = opts.error || "";
  var milestoneStatus = computeOverallStatus_(record);
  var actionUrl = canonicalStudentExecBase_() + "?view=portal";
  var actionWarn = actionUrl ? "" : "Student URL not configured. Saving may not work for external users.";
  var scriptId = clean_(CONFIG.SCRIPT_ID || ScriptApp.getScriptId());
  var deploymentId = clean_(CONFIG.DEPLOYMENT_ID_STUDENT || "");
  var actionUrlShort = actionUrl.length > 140 ? (actionUrl.slice(0, 137) + "...") : actionUrl;
  var buildVersion = clean_(opts.version || CONFIG.VERSION || "");
  var shortVersion = clean_(opts.versionShort || portalVersionShort_(buildVersion));
  var buildRenderedAt = clean_(opts.buildRenderedAt || new Date().toISOString());
  var buildScriptId = clean_(opts.buildScriptId || scriptId || "");
  var buildLabel = clean_(CONFIG.BUILD_LABEL || "");
  var redactedSecret = redactToken_(secret);

  var locked = record._PortalLocked === true;
  var lockReason = clean_(record._PortalLockReason || "");
  var isPaymentVerifiedLock = locked && lockReason === "payment_verified";
  var dis = locked ? "disabled" : "";
  var subjectsDis = (locked || subjectsLocked) ? "disabled" : "";
  var ro = locked ? "readonly" : "";

  // subject selections: canonical preferred, else fallback
  var csv = clean_(record.Subjects_Selected_Canonical || record._SubjectsCsv || "");
  var selected = parseSubjects_(csv);

  var dobVal = esc_(toIsoDateInput_(record.Date_Of_Birth));

  var examVal = clean_(record.Physical_Exam_Site || "");
  var validationFieldSet = {};
  for (var vf = 0; vf < validationFields.length; vf++) validationFieldSet[validationFields[vf]] = true;
  var validationCodeSet = {};
  for (var vc = 0; vc < validationCodes.length; vc++) validationCodeSet[validationCodes[vc]] = true;
  var dobAttention = !!(validationFieldSet.Date_Of_Birth || validationCodeSet.DOB_REQUIRED || validationCodeSet.DOB_INVALID || (!dobVal && !locked));
  var dobMessage = validationCodeSet.DOB_REQUIRED
    ? "Date of Birth is required."
    : (validationCodeSet.DOB_INVALID
        ? "Enter a valid Date of Birth."
        : ((!dobVal && !locked) ? "Date of Birth is required to complete your application." : ""));
  var subjectsAttention = !!(validationFieldSet.Subjects_Selected_Canonical || validationCodeSet.SUBJECTS_REQUIRED || validationCodeSet.SUBJECTS_INVALID_FOR_GRADE || validationCodeSet.SUBJECT_LOCK_DOCS_VERIFIED);
  var subjectsMessage = validationCodeSet.SUBJECTS_REQUIRED
    ? "Select at least one subject."
    : ((validationCodeSet.SUBJECTS_INVALID_FOR_GRADE || validationCodeSet.SUBJECT_LOCK_DOCS_VERIFIED)
        ? portalValidationMessageForCode_(validationCodes[0] || "SUBJECTS_INVALID_FOR_GRADE")
        : "");
  var dobInputStyle = 'padding:8px;width:260px;' + (dobAttention ? 'border:2px solid #b30000;background:#fff7f7;' : '');
  var examInputStyle = 'padding:8px;width:520px;';
  var subjectsBoxStyle = 'margin-top:8px;padding:12px;border:' + (subjectsAttention ? '2px solid #b30000' : '1px solid #eee') + ';border-radius:10px;' + (subjectsAttention ? 'background:#fff7f7;' : '');

  // exam site options
  var examList = (examSites.length ? examSites : ["Port Moresby - HQ"]);
  var examOptions = "";
  for (var i = 0; i < examList.length; i++) {
    var s = examList[i];
    var sel = (s === examVal) ? "selected" : "";
    examOptions += '<option value="' + esc_(s) + '" ' + sel + ">" + esc_(s) + "</option>";
  }

  // subject checkboxes
  var subjectChecks = "";
  for (var j = 0; j < subjects.length; j++) {
    var subj = subjects[j];
    var checked = selected.has(subj.toLowerCase()) ? "checked" : "";
    subjectChecks += ''
      + '<label style="display:block;margin:6px 0;">'
      + '<input type="checkbox" name="subj" value="' + esc_(subj) + '" ' + checked + " " + subjectsDis + " /> "
      + esc_(subj)
      + "</label>";
  }
  var subjectsLockedNotice = subjectsLocked
    ? '<div style="margin-top:8px;padding:10px;border:1px solid #dbeafe;border-radius:8px;background:#eff6ff;color:#1e3a8a;font-size:13px;">Subjects are locked because your documents have been verified. Please contact administration for any changes.</div>'
    : "";

  var errorBlock = error
    ? '<div style="background:#ffecec;border:1px solid #ffb3b3;padding:12px;border-radius:10px;margin-bottom:16px;"><b>Action required:</b> ' + esc_(error) + "</div>"
    : "";

  var lockedMsg = (lockReason === "portal_access_locked")
    ? "Portal access is locked. Please contact admissions."
    : "Payment is verified. No further changes are needed.";
  var lockedBlock = locked
    ? '<div style="background:#e8f0ff;border:1px solid #b6ccff;padding:12px;border-radius:10px;margin-bottom:16px;"><b>' + (lockReason === "portal_access_locked" ? "Locked:" : "Locked for processing:") + '</b> ' + esc_(lockedMsg) + "</div>"
    : "";
  var enrollmentConfirmedBlock = (isPaymentVerifiedLock || (lockedFlag && msgToken === "enrolled" && locked))
    ? '<div style="background:#eaf7ea;border:1px solid #2e7d32;padding:12px;border-radius:10px;margin-bottom:16px;color:#000;"><b>Enrollment confirmed:</b> Your payment has been verified. Your application is now locked for processing. We will shortly provide you access to online studies.</div>'
    : "";
  var actionWarnBlock = actionWarn
    ? '<div style="background:#fff6e5;border:1px solid #f5c26b;padding:12px;border-radius:10px;margin-bottom:16px;"><b>Warning:</b> ' + esc_(actionWarn) + "</div>"
    : "";
  var savedBlock = saved
    ? '<div id="savedBanner" style="background:#eaf7ea;border:1px solid #2e7d32;padding:8px;margin-bottom:12px;color:#000;">Saved. Your updates are now shown below.</div>'
    : "";
  var errText = uploadFail
    ? ("Upload failed. Please try again." + (uploadField ? (" Field: " + uploadField + ".") : ""))
    : "Portal update failed.";
  var showErrorBanner = hasErr && !isPaymentVerifiedLock && !validationFlag;
  var errBlock = showErrorBanner
    ? '<div style="background:#ffecec;border:1px solid #b30000;padding:8px;margin-bottom:12px;color:#000;">' + esc_(errText) + (dbg ? (" Debug: " + esc_(dbg)) : "") + "</div>"
    : "";
  var validationSummaryBlock = validationFlag
    ? '<div style="background:#ffecec;border:1px solid #b30000;padding:10px;border-radius:8px;margin-bottom:12px;color:#000;"><b>Please correct the highlighted fields before submitting.</b></div>'
    : "";
  var uploadBannerId = "portalUploadResultBanner";
  var uploadResultBlock = "";
  if (uploadResult) {
    var uploadText = uploadOk
      ? ("Upload successful." + (dbg ? (" Debug: " + dbg) : ""))
      : ("Upload failed" + (uploadErrCode ? (" (" + uploadErrCode + ")") : "") + ". Please retry." + (dbg ? (" Debug: " + dbg) : ""));
    if (uploadDocKey) uploadText += " Field: " + uploadDocKey + ".";
    uploadResultBlock = ''
      + '<div id="' + uploadBannerId + '" style="'
      + (uploadOk
        ? 'background:#eaf7ea;border:1px solid #2e7d32;'
        : 'background:#ffecec;border:1px solid #b30000;')
      + 'padding:10px;border-radius:8px;margin:0 0 12px 0;color:#000;">'
      + '<span>' + esc_(uploadText) + '</span>'
      + '</div>';
  }
  var debugComment = debugPage
    ? '<!-- DEBUG_PORTAL_GET reqId=' + esc_(reqId) + ' id=' + esc_(id) + ' s(redacted)=' + esc_(redactedSecret) + ' -->'
    : "";
  var debugFooter = debugPage
    ? '<div id="debugPortalFooter" style="margin-top:10px;padding:8px;border:1px dashed #999;color:#000;font-size:12px;">'
      + "<div><b>DEBUG_PORTAL_RENDER</b></div>"
      + "<div>reqId: " + esc_(reqId) + "</div>"
      + "<div>applicantId: " + esc_(id) + "</div>"
      + "<div>form.action: " + esc_(actionUrl) + "</div>"
      + "<div>hidden id: " + esc_(id) + "</div>"
      + "<div>hidden s(redacted): " + esc_(redactedSecret) + "</div>"
      + "<div>form named elements count: <span id='dbgFormNameCount'>(loading...)</span></div>"
      + "<div>form field names (first 10): <span id='dbgFormNameList'>(loading...)</span></div>"
      + "<div>cleanup ran: <span id='dbgCleanupRan'>(loading...)</span></div>"
      + "<div>current URL (after cleanup): <span id='dbgHrefAfterCleanup'>(loading...)</span></div>"
      + "<div>window.location.href: <span id='dbgCurrentHref'>(loading...)</span></div>"
      + "</div>"
    : "";

  // allowlist-only summary
  var summaryHtml = renderAllowlistSummary_(record, visibleFields);

  // editable fields UI
  var extraInputs = renderEditableFields_(record, editFields, dis);

  // docs upload UI
  var docsHtml = renderDocsSection_(id, secret, record, docs, locked);

  var saveButton = locked
    ? '<div style="margin-top:12px;color:#1a4fb3;"><b>No action available.</b></div>'
    : '<button type="submit" style="padding:10px 16px;border:0;border-radius:10px;background:#1a73e8;color:#fff;">Save Updates</button>';

  return ''
    + '<!doctype html><html><head><meta charset="utf-8" />'
    + debugComment
    + '<!-- BUILD: ' + esc_(buildVersion) + " | " + esc_(buildLabel) + " -->"
    + "<title>FODE Student Portal</title>"
    + '<base target="_top" />'
    + '<meta name="viewport" content="width=device-width, initial-scale=1" />'
    + '<meta http-equiv="Cache-Control" content="no-store, no-cache, must-revalidate, max-age=0" />'
    + '<meta http-equiv="Pragma" content="no-cache" />'
    + '<meta http-equiv="Expires" content="0" />'
    + "<style>"
    + ".milestone{padding:12px;margin-bottom:16px;border-radius:6px;font-size:14px;}"
    + ".docs-stage{background:#1f2937;color:#fbbf24;}"
    + ".payment-stage{background:#064e3b;color:#34d399;}"
    + ".pending-stage{background:#1e3a8a;color:#93c5fd;}"
    + "</style>"
    + "</head>"
    + '<body style="font-family:Arial,Helvetica,sans-serif;max-width:980px;margin:24px auto;padding:0 16px;">'
    + '<div style="position:fixed;top:10px;right:12px;z-index:9;padding:2px 8px;border-radius:999px;font-size:11px;font-weight:700;color:#0f172a;background:#dbeafe;border:1px solid #93c5fd;" title="' + esc_(buildVersion) + '">' + esc_(shortVersion) + "</div>"
    + "<h2>FODE Student Portal</h2>"
    + '<div id="portalFlashMount"></div>'
    + '<div id="portalErrorBanner" style="display:none;background:#ffecec;border:1px solid #b30000;padding:10px;border-radius:8px;margin-bottom:12px;color:#000;"></div>'
    + savedBlock
    + errBlock
    + validationSummaryBlock
    + enrollmentConfirmedBlock
    + '<div style="padding:12px;border:1px solid #ddd;border-radius:10px;margin-bottom:16px;">'
    + "<div><b>Applicant ID:</b> " + esc_(id) + "</div>"
    + "<div><b>Secure Link:</b> verified</div>"
    + '<div style="margin-top:6px;color:#555;font-size:12px;" title="Script ID: ' + esc_(buildScriptId) + '">Build: ' + esc_(buildVersion) + ' | Rendered: ' + esc_(buildRenderedAt) + '</div>'
    + "</div>"
    + lockedBlock
    + errorBlock
    + actionWarnBlock

    + '<div style="padding:12px;border:1px solid #ddd;border-radius:10px;margin-bottom:16px;">'
    + '<h3 style="margin-top:0;">Submitted Details (read-only)</h3>'
    + summaryHtml
    + "</div>"

    + '<div id="milestoneBanner" class="milestone pending-stage">Your application is under review.</div>'

    + '<div style="padding:12px;border:1px solid #ddd;border-radius:10px;margin-bottom:16px;">'
    + '<h3 style="margin-top:0;">Documents & Payment Proof</h3>'
    + uploadResultBlock
    + docsHtml
    + "</div>"

    // ✅ hardcoded action URL to prevent blank screen / doPost not firing
    + '<form id="portalForm" method="post" target="_top" action="' + esc_(actionUrl) + '" onsubmit="return beforePortalSubmit(event,this);"'
    + ' style="padding:12px;border:1px solid #ddd;border-radius:10px;">'
    + '<input type="hidden" name="action" value="portal_update" />'
    + '<input type="hidden" name="route" value="portal_update" />'
    + '<input type="hidden" name="id" value="' + esc_(id) + '" />'
    + '<input type="hidden" name="s" value="' + esc_(secret) + '" />'
    + (CONFIG.DEBUG_PORTAL_POST === true ? '<input type="hidden" name="dbg" value="1" />' : "")
    + '<input type="hidden" id="portal_payload" name="payload" value="" />'
    + '<input type="hidden" id="Subjects_Selected_Canonical" name="Subjects_Selected_Canonical" value="" />'

    + '<h3 style="margin-top:0;">Update / Confirm Information</h3>'

    + '<div style="margin:12px 0;">'
    + "<label for=\"portalDobInput\"><b>Date of Birth <span style=\"color:#b30000;\">*</span></b></label><br/>"
    + '<input id="portalDobInput" type="date" name="Date_Of_Birth" value="' + dobVal + '" style="' + dobInputStyle + '" ' + ro + " />"
    + '<div id="portalDobError" style="margin-top:6px;color:#b30000;display:' + (dobMessage ? 'block' : 'none') + ';">' + esc_(dobMessage) + '</div>'
    + "</div>"

    + '<div style="margin:12px 0;">'
    + "<label><b>Physical Exam Site (optional):</b></label><br/>"
    + '<select name="Physical_Exam_Site" style="' + examInputStyle + '" ' + dis + ">"
    + '<option value="">-- Select Exam Site --</option>'
    + examOptions
    + "</select>"
    + (locked ? ('<input type="hidden" name="Physical_Exam_Site" value="' + esc_(examVal) + '" />') : "")
    + "</div>"

    + '<div style="margin:12px 0;">'
    + "<label><b>Select Subjects <span style=\"color:#b30000;\">*</span></b></label>"
    + '<div id="portalSubjectsBox" style="' + subjectsBoxStyle + '">'
    + subjectChecks
    + subjectsLockedNotice
    + '</div>'
    + '<div id="portalSubjectsError" style="margin-top:6px;color:#b30000;display:' + (subjectsMessage ? 'block' : 'none') + ';">' + esc_(subjectsMessage) + '</div>'
    + "</div>"

    + (editFields.length ? ('<div style="margin:12px 0;">'
      + "<h4 style='margin:8px 0;'>Additional Editable Fields</h4>"
      + extraInputs
      + "</div>") : "")

    + saveButton
    + "</form>"
    + debugFooter
    + '<div class="footer-version" id="studentVersion" style="margin-top:16px;color:#666;font-size:12px;" title="Script ID: ' + esc_(buildScriptId) + '"></div>'
    + '<div style="margin-top:8px;color:#000;">Script ID: ' + esc_(scriptId) + " | Deployment: " + esc_(deploymentId || "-") + " | View: portal</div>"
    + '<div style="margin-top:8px;color:#666;font-size:12px;">URL: ' + esc_(actionUrlShort) + "</div>"

    + "<script>"
    + "console.log('PORTAL BUILD: ' + " + JSON.stringify(buildVersion) + ");"
    + "function packSubjects(){"
    + "var boxes=[].slice.call(document.querySelectorAll('input[name=\"subj\"]:checked'));"
    + "var vals=boxes.map(function(b){return b.value;}).filter(Boolean);"
    + "document.getElementById('Subjects_Selected_Canonical').value=vals.join(', ');"
    + "return vals.length>0;}"
    + "function setPortalFieldMessage_(field,msg){"
    + "var el=null;"
    + "if(field==='Date_Of_Birth') el=document.getElementById('portalDobError');"
    + "else if(field==='Subjects_Selected_Canonical') el=document.getElementById('portalSubjectsError');"
    + "if(!el) return;"
    + "var text=String(msg||'').trim();"
    + "el.textContent=text;"
    + "el.style.display=text?'block':'none';"
    + "}"
    + "function markPortalFieldInvalid_(field){"
    + "if(field==='Date_Of_Birth'){ var dob=document.getElementById('portalDobInput'); if(dob){ dob.style.border='2px solid #b30000'; dob.style.background='#fff7f7'; } return; }"
    + "if(field==='Subjects_Selected_Canonical'){ var box=document.getElementById('portalSubjectsBox'); if(box){ box.style.border='2px solid #b30000'; box.style.background='#fff7f7'; } return; }"
    + "var form=document.getElementById('portalForm');"
    + "if(!form) return;"
    + "var nodes=form.querySelectorAll('[name=\"'+field+'\"]');"
    + "[].slice.call(nodes).forEach(function(node){ if(node){ node.style.border='2px solid #b30000'; node.style.background='#fff7f7'; } });"
    + "}"
    + "function applyPortalValidationState_(){"
    + "var fields=" + JSON.stringify(validationFields) + ";"
    + "var codes=" + JSON.stringify(validationCodes) + ";"
    + "fields.forEach(function(field){ markPortalFieldInvalid_(field); });"
    + "if(codes.indexOf('DOB_REQUIRED')>=0) setPortalFieldMessage_('Date_Of_Birth','Date of Birth is required.');"
    + "else if(codes.indexOf('DOB_INVALID')>=0) setPortalFieldMessage_('Date_Of_Birth','Enter a valid Date of Birth.');"
    + "if(codes.indexOf('SUBJECTS_REQUIRED')>=0) setPortalFieldMessage_('Subjects_Selected_Canonical','Select at least one subject.');"
    + "else if(codes.indexOf('SUBJECTS_INVALID_FOR_GRADE')>=0) setPortalFieldMessage_('Subjects_Selected_Canonical','Selected subjects are not valid for the chosen grade.');"
    + "else if(codes.indexOf('SUBJECT_LOCK_DOCS_VERIFIED')>=0) setPortalFieldMessage_('Subjects_Selected_Canonical','Subjects are locked because documents have been verified by Admin.');"
    + "}"
    + "function ensurePortalFormSerialization(form){"
    + "if(!form) return;"
    + "var oldClones=[].slice.call(form.querySelectorAll('input[data-portal-clone=\"1\"]'));"
    + "oldClones.forEach(function(n){ if(n && n.parentNode) n.parentNode.removeChild(n); });"
    + "var fd=new FormData(form);"
    + "var els=[].slice.call(form.querySelectorAll('input[name],select[name],textarea[name]'));"
    + "els.forEach(function(el){"
    + "  if(!el || !el.name || el.disabled) return;"
    + "  var type=(el.type||'').toLowerCase();"
    + "  if((type==='checkbox' || type==='radio') && !el.checked) return;"
    + "  if(fd.has(el.name)) return;"
    + "  var hidden=document.createElement('input');"
    + "  hidden.type='hidden';"
    + "  hidden.name=el.name;"
    + "  hidden.value=(el.value===undefined || el.value===null)?'':String(el.value);"
    + "  hidden.setAttribute('data-portal-clone','1');"
    + "  form.appendChild(hidden);"
    + "});"
    + "}"
    + "function beforePortalSubmit(evt,form){"
    + "clearPortalError_();"
    + "setPortalFieldMessage_('Date_Of_Birth','');"
    + "setPortalFieldMessage_('Subjects_Selected_Canonical','');"
    + "if(!packSubjects()){"
    + "setPortalError_('Please correct the highlighted fields before submitting.');"
    + "setPortalFieldMessage_('Subjects_Selected_Canonical','Select at least one subject.');"
    + "markPortalFieldInvalid_('Subjects_Selected_Canonical');"
    + "return false;}"
    + "var dobInput=document.getElementById('portalDobInput');"
    + "var dobValue=dobInput?String(dobInput.value||'').trim():'';"
    + "if(!dobValue){"
    + "setPortalError_('Please correct the highlighted fields before submitting.');"
    + "setPortalFieldMessage_('Date_Of_Birth','Date of Birth is required.');"
    + "markPortalFieldInvalid_('Date_Of_Birth');"
    + "if(dobInput && typeof dobInput.focus==='function') dobInput.focus();"
    + "return false;}"
    + "ensurePortalFormSerialization(form);"
    + "var fd=new FormData(form);"
    + "var obj={};"
    + "fd.forEach(function(v,k){"
    + "  if(k==='payload' || k==='subj') return;"
    + "  if(Object.prototype.hasOwnProperty.call(obj,k)){"
    + "    if(Array.isArray(obj[k])) obj[k].push(v);"
    + "    else obj[k]=[obj[k],v];"
    + "  } else obj[k]=v;"
    + "});"
    + "var p=document.getElementById('portal_payload');"
    + "if(p) p.value=JSON.stringify(obj);"
    + "var submitBtn=form && form.querySelector ? form.querySelector('button[type=\"submit\"]') : null;"
    + "if(submitBtn) submitBtn.disabled=true;"
    + "return true;"
    + "}"
    + "function escText(v){"
    + "return String(v===undefined||v===null?'':v).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/\\\"/g,'&quot;').replace(/'/g,'&#039;');"
    + "}"
    + "function renderPortalFlash(type,dbg,mode,field){"
    + "var mount=document.getElementById('portalFlashMount');"
    + "if(!mount) return;"
    + "if(type==='success'){"
    + "mount.innerHTML='<div id=\"portalFlashBanner\" style=\"background:#eaf7ea;border:1px solid #2e7d32;padding:8px;margin-bottom:12px;color:#000;\">Saved. Your updates are now shown below.</div>';"
    + "return;"
    + "}"
    + "if(type==='error'){"
    + "var text=(mode==='upload')?'Upload failed. Please try again.'+((field&&String(field).trim())?(' Field: '+escText(field)+'.'):''):'Portal update failed.';"
    + "var d=dbg?(' Debug: '+escText(dbg)):'';"
    + "mount.innerHTML='<div id=\"portalFlashBanner\" style=\"background:#ffecec;border:1px solid #b30000;padding:8px;margin-bottom:12px;color:#000;\">'+text+d+'</div>';"
    + "}"
    + "}"
    + "function initPortalFlashAndCleanup(){"
    + "var cleanupRan=false;"
    + "var hasSaved=false;"
    + "var hasError=false;"
    + "try{"
    + "var url=new URL(window.location.href);"
    + "var params=url.searchParams;"
    + "hasSaved=(params.get('saved')==='1');"
    + "hasError=(params.get('error')==='1');"
    + "var dbgQ=params.get('dbg')||'';"
    + "var uploadFailQ=(params.get('uploadFail')==='1');"
    + "var fieldQ=params.get('field')||'';"
    + "if(hasSaved){ sessionStorage.setItem('portalFlash', JSON.stringify({type:'success',ts:Date.now()})); }"
    + "var validationQ=(params.get('val')==='1');"
    + "if(hasError && !validationQ){ sessionStorage.setItem('portalFlash', JSON.stringify({type:'error',dbg:dbgQ,mode:(uploadFailQ?'upload':'update'),field:fieldQ,ts:Date.now()})); }"
    + "if(params.get('u')==='1'){"
    + "params.delete('u');"
    + "params.delete('ok');"
    + "params.delete('docKey');"
    + "params.delete('errCode');"
    + "}"
    + "if(params.has('id') || params.has('s')){"
    + "params.delete('id');"
    + "params.delete('s');"
    + "var q=params.toString();"
    + "var newUrl=url.pathname + (q?('?'+q):'') + url.hash;"
    + "history.replaceState(null,'',newUrl);"
    + "cleanupRan=true;"
    + "}"
    + "}catch(e){}"
    + "if(!hasSaved && !hasError){"
    + "try{"
    + "var raw=sessionStorage.getItem('portalFlash');"
    + "if(raw){"
    + "var flash=JSON.parse(raw);"
    + "var age=Date.now()-Number((flash&&flash.ts)||0);"
    + "if(age>=0 && age<=120000){"
    + "renderPortalFlash(String(flash.type||''), String((flash&&flash.dbg)||''), String((flash&&flash.mode)||''), String((flash&&flash.field)||''));"
    + "}"
    + "sessionStorage.removeItem('portalFlash');"
    + "}"
    + "}catch(e2){"
    + "try{sessionStorage.removeItem('portalFlash');}catch(e3){}"
    + "}"
    + "}"
    + "var dbgCleanupEl=document.getElementById('dbgCleanupRan'); if(dbgCleanupEl){ dbgCleanupEl.textContent=cleanupRan ? 'true' : 'false'; }"
    + "var dbgAfterEl=document.getElementById('dbgHrefAfterCleanup'); if(dbgAfterEl){ dbgAfterEl.textContent=window.location.href; }"
    + "var dbgHrefEl=document.getElementById('dbgCurrentHref'); if(dbgHrefEl){ dbgHrefEl.textContent=window.location.href; }"
    + "var uploadBanner=document.getElementById('portalUploadResultBanner');"
    + "if(uploadBanner){ setTimeout(function(){ try{ uploadBanner.style.display='none'; }catch(_e){} },10000); }"
    + "}"
    + "function initStudentVersionFooter(){"
    + "var full=" + JSON.stringify(buildVersion) + ";"
    + "var shortV=" + JSON.stringify(shortVersion) + ";"
    + "var el=document.getElementById('studentVersion');"
    + "if(el){ el.textContent='Version ' + shortV + ' | Rendered ' + " + JSON.stringify(buildRenderedAt) + "; el.title=full; }"
    + "}"
    + "function renderMilestone(status){"
    + "var el=document.getElementById('milestoneBanner');"
    + "if(!el) return;"
    + "if(status==='Docs_Verified'){"
    + "el.innerHTML='<strong>Documents Verified.</strong><br>Your documents have been verified. Please allow 10-15 working days after payment confirmation for login access.';"
    + "el.className='milestone docs-stage';"
    + "} else if(status==='Verified'){"
    + "el.innerHTML='<strong>Payment Confirmed.</strong><br>Your admission is confirmed. Login credentials will be issued within 10-15 working days.';"
    + "el.className='milestone payment-stage';"
    + "} else {"
    + "el.innerHTML='Your application is under review.';"
    + "el.className='milestone pending-stage';"
    + "}"
    + "}"
    + "function initPortalPage(){"
    + "initPortalFlashAndCleanup();"
    + "applyPortalValidationState_();"
    + "initStudentVersionFooter();"
    + "renderMilestone(" + JSON.stringify(milestoneStatus) + ");"
    + "}"
    + "if(document.readyState==='loading'){ document.addEventListener('DOMContentLoaded', initPortalPage); } else { initPortalPage(); }"
    + "setTimeout(function(){var b=document.getElementById('savedBanner');if(b){b.style.display='none';}},5000);"
    + "var dbgForm=document.getElementById('portalForm');"
    + "if(dbgForm){"
    + "  var named=[].slice.call(dbgForm.querySelectorAll('[name]'));"
    + "  var c=document.getElementById('dbgFormNameCount');"
    + "  var l=document.getElementById('dbgFormNameList');"
    + "  if(c) c.textContent=String(named.length);"
    + "  if(l) l.textContent=named.map(function(n){return n.name;}).slice(0,10).join(', ');"
    + "}"
    + "</script>"

    + "</body></html>";
}

function renderDocsSection_(id, secret, record, docs, locked) {
  var out = "";
  var portalReloadUrl = buildPortalRedirectUrl_(id, secret, {});
  var execUrl = clean_(typeof getExecUrl_ === "function" ? getExecUrl_() : "");
  if (!execUrl) execUrl = clean_(canonicalStudentExecBase_() || "");
  var uploadAction = execUrl ? (execUrl + (execUrl.indexOf("?") >= 0 ? "&" : "?") + "view=portalUpload") : "?view=portalUpload";
  for (var i = 0; i < docs.length; i++) {
    var d = docs[i];
    var cur = clean_(record[d.field] || "");
    var st = mapDocStatusForDisplay_(record[d.status]);
    var cm = clean_(record[d.comment] || "");

    var stBadge = "<b>Status:</b> " + esc_(st || "Pending Review");
    var cmBlock = cm ? ("<div style='margin-top:6px;'><b>Admin comment:</b> " + esc_(cm) + "</div>") : "";

    var urlList = normalizeToUrlList_(cur);
    var secureOpenUrl = buildTokenGatedFileUrl_(execUrl, id, secret, d.field, "open");
    var secureDownloadUrl = buildTokenGatedFileUrl_(execUrl, id, secret, d.field, "download");
    Logger.log("PORTAL_DOC_LINK " + JSON.stringify({
      applicantId: clean_(id || ""),
      field: clean_(d.field || ""),
      hasValidUrl: urlList.length > 0,
      urlCount: urlList.length
    }));
    var curLinks = "";
    if (urlList.length) {
      var deleteControls = [];
      for (var u = 0; u < urlList.length; u++) {
        var delBtn = (!locked)
          ? " <button type='button' onclick=\"deleteDocUrl('" + esc_(d.field) + "','" + esc_(encodeURIComponent(urlList[u])) + "')\">Delete " + String(u + 1) + "</button>"
          : "";
        if (delBtn) deleteControls.push(delBtn);
      }
      curLinks = "<div style='margin-top:6px;'>"
        + "<b>Current files:</b> " + String(urlList.length) + " uploaded."
        + (secureOpenUrl ? (" <a href='" + esc_(secureOpenUrl) + "' target='_blank' rel='noopener noreferrer'>Open</a>") : "")
        + (secureDownloadUrl ? (" | <a href='" + esc_(secureDownloadUrl) + "' target='_blank' rel='noopener noreferrer'>Download</a>") : "")
        + (deleteControls.length ? ("<div style='margin-top:6px;'>" + deleteControls.join(" ") + "</div>") : "")
        + "</div>";
    } else {
      curLinks = "<div style='margin-top:6px;'><b>Current files:</b> Not uploaded</div>";
    }
    var multipleAttr = d.multiple ? " multiple" : "";
    var multiBadge = d.multiple ? "<div class='muted' style='font-size:12px;margin-top:4px;'>Upload one file at a time (multi-file upload temporarily disabled).</div>" : "";
    var noteHtml = "";

    var uploadUi = locked
      ? "<div style='margin-top:10px;color:#666;'><i>Uploads disabled (locked).</i></div>"
      : "<div style='margin-top:10px;'>"
        + "<div id='uf_" + esc_(d.field) + "' style='margin:0;'>"
        + "<input type='hidden' name='dbg' id='dbg_" + esc_(d.field) + "' value='' />"
        + "<input type='file' name='file' required id='f_" + esc_(d.field) + "' data-upload-input='1' data-field='" + esc_(d.field) + "' data-multi='" + (d.multiple ? "1" : "0") + "' /> "
        + "<button type='button' id='btn_" + esc_(d.field) + "' data-upload-btn='1' data-field='" + esc_(d.field) + "'>Upload / Replace</button>"
        + "</div>"
        + noteHtml
        + "<div id='cur_" + esc_(d.field) + "' style='margin-top:6px;font-size:12px;'>" + curLinks + "</div>"
        + "<div id='msg_" + esc_(d.field) + "' style='margin-top:6px;font-size:12px;'></div>"
        + "<div id='ust_" + esc_(d.field) + "' style='margin-top:6px;padding:8px;border:1px solid #e5e7eb;border-radius:8px;background:#f8fafc;font-size:11px;'>"
        + "<div><b>Upload Status:</b> <span id='ust_text_" + esc_(d.field) + "'>Idle</span></div>"
        + "<div>dbg: <span id='ust_dbg_" + esc_(d.field) + "'>-</span></div>"
        + "<div>stage: <span id='ust_stage_" + esc_(d.field) + "'>-</span></div>"
        + "<div>errCode: <span id='ust_err_" + esc_(d.field) + "'>-</span></div>"
        + "<div>driveMode: <span id='ust_mode_" + esc_(d.field) + "'>-</span></div>"
        + "<div>serverMs: <span id='ust_ms_" + esc_(d.field) + "'>-</span></div>"
        + "<div>fileUrl: <span id='ust_url_" + esc_(d.field) + "'>-</span></div>"
        + "</div>"
        + multiBadge
        + "</div>";

    out += ""
      + "<div style='padding:10px;border:1px solid #eee;border-radius:10px;margin:10px 0;'>"
      + "<div><b>" + esc_(d.label) + "</b></div>"
      + "<div id='st_" + esc_(d.field) + "' style='margin-top:6px;'>" + stBadge + "</div>"
      + cmBlock
      + curLinks
      + uploadUi
      + "</div>";
  }

  // uploader script
  out += ""
    + "<script>"
    + "var PORTAL_AUTO_UPLOAD=false;"
    + "var PORTAL_LOCKED=" + (locked ? "true" : "false") + ";"
    + "var PORTAL_UPLOAD_MAX_MB=" + String(Number(CONFIG.PORTAL_UPLOAD_MAX_MB || 5)) + ";"
    + "var PORTAL_MAX_UPLOAD_BYTES=" + String(Number(CONFIG.PORTAL_MAX_UPLOAD_BYTES || (5*1024*1024))) + ";"
    + "var PORTAL_UPLOAD_MAX_SERVER_MS=" + String(Number(CONFIG.PORTAL_UPLOAD_MAX_SERVER_MS || 20000)) + ";"
    + "var PORTAL_UPLOAD_TIMEOUT_MS=" + String(Number(CONFIG.PORTAL_UPLOAD_TIMEOUT_MS || 25000)) + ";"
    + "var PORTAL_UPLOAD_ID=" + JSON.stringify(clean_(id || "")) + ";"
    + "var PORTAL_UPLOAD_S=" + JSON.stringify(clean_(secret || "")) + ";"
    + "var DOC_MULTI_MAP={"
    + (docs || []).map(function (x) { return "'" + esc_(x.field) + "':" + (x.multiple ? "true" : "false"); }).join(",")
    + "};"
    + "var PORTAL_UPLOAD_PENDING={};"
    + "function escHtml(s){return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/\\\"/g,'&quot;').replace(/'/g,'&#039;');}"
    + "function makeClientDebugId_(){"
    + "  var ts='';"
    + "  try{ ts=new Date().toISOString().replace(/[-:.TZ]/g,'').slice(0,14); }catch(e){ ts=String(Date.now()); }"
    + "  return 'CDBG-'+ts+'-'+Math.random().toString(16).slice(2,10);"
    + "}"
    + "function stringifyGsError_(err){"
    + "  if(err===null || err===undefined) return 'Unknown error';"
    + "  if(typeof err==='string') return err;"
    + "  if(err && typeof err.message==='string' && err.message) return err.message;"
    + "  try{return JSON.stringify(err);}catch(e){}"
    + "  try{return String(err);}catch(e2){}"
    + "  return 'Unknown error';"
    + "}"
    + "function setPortalError_(msg){"
    + "  var el=document.getElementById('portalErrorBanner');"
    + "  if(!el) return;"
    + "  var t=String(msg||'').trim();"
    + "  if(!t){ el.style.display='none'; el.textContent=''; return; }"
    + "  el.textContent=t;"
    + "  el.style.display='block';"
    + "}"
    + "function clearPortalError_(){ setPortalError_(''); }"
    + "function getPortalTokens_(){"
    + "  var q=new URLSearchParams(window.location.search||'');"
    + "  var id=(q.get('id')||'').trim();"
    + "  var s=(q.get('s')||'').trim();"
    + "  if((!id || !s)){"
    + "    var form=document.getElementById('portalForm');"
    + "    if(form){"
    + "      var idEl=form.querySelector('input[name=\"id\"]');"
    + "      var sEl=form.querySelector('input[name=\"s\"]');"
    + "      if(!id && idEl) id=String(idEl.value||'').trim();"
    + "      if(!s && sEl) s=String(sEl.value||'').trim();"
    + "    }"
    + "  }"
    + "  if(!id || !s){"
    + "    setPortalError_('Missing portal link token. Please reopen your portal link.');"
    + "    return null;"
    + "  }"
    + "  return {id:id,s:s};"
    + "}"
    + "function setRowMsg(fieldName, txt, isErr){"
    + "  var msg=document.getElementById('msg_'+fieldName);"
    + "  if(!msg) return;"
    + "  msg.style.color=isErr ? '#b30000' : '';"
    + "  msg.innerHTML=txt||'';"
    + "}"
    + "function setUploadStatusPanel_(fieldName, info){"
    + "  var x=info||{};"
    + "  function set_(id,val,isLink){"
    + "    var el=document.getElementById(id+'_'+fieldName);"
    + "    if(!el) return;"
    + "    var v=(val===undefined||val===null||val==='')?'-':String(val);"
    + "    if(isLink && v && v!=='-'){ el.innerHTML='<a target=\"_blank\" href=\"'+escHtml(v)+'\">'+escHtml(v)+'</a>'; }"
    + "    else { el.textContent=v; }"
    + "  }"
    + "  set_('ust_text', x.text||'');"
    + "  set_('ust_dbg', x.dbg||'');"
    + "  set_('ust_stage', x.stage||'');"
    + "  set_('ust_err', x.errCode||'');"
    + "  set_('ust_mode', x.driveMode||'');"
    + "  set_('ust_ms', x.serverMs!==undefined&&x.serverMs!==null ? String(x.serverMs) : '');"
    + "  set_('ust_url', x.fileUrl||'', true);"
    + "}"
    + "function setUploadUiError_(fieldName, txt, meta){"
    + "  var m=meta||{};"
    + "  setPortalError_(txt);"
    + "  setRowMsg(fieldName, txt, true);"
    + "  setUploadStatusPanel_(fieldName, { text:'Error', dbg:m.dbg||'', stage:m.stage||'', errCode:m.errCode||'', driveMode:m.driveMode||'', serverMs:m.serverMs, fileUrl:m.fileUrl||'' });"
    + "}"
    + "function setUploadUiSuccess_(fieldName, txt, meta){"
    + "  var m=meta||{};"
    + "  clearPortalError_();"
    + "  setRowMsg(fieldName, txt, false);"
    + "  setUploadStatusPanel_(fieldName, { text:'Uploaded', dbg:m.dbg||'', stage:m.stage||'done', errCode:m.errCode||'', driveMode:m.driveMode||'', serverMs:m.serverMs, fileUrl:m.fileUrl||'' });"
    + "}"
    + "function markDocStatusPendingReviewUi_(fieldName){"
    + "  var el=document.getElementById('st_'+fieldName);"
    + "  if(el) el.innerHTML='<b>Status:</b> Pending Review';"
    + "}"
    + "function setUploadBusy(fieldName,busy){"
    + "  var input=document.getElementById('f_'+fieldName);"
    + "  var btn=document.getElementById('btn_'+fieldName);"
    + "  if(input) input.disabled=!!busy;"
    + "  if(btn) btn.disabled=!!busy;"
    + "}"
    + "function getCurrentUrlsFromDom(fieldName){"
    + "  var links=[].slice.call(document.querySelectorAll('#cur_'+fieldName+' a'));"
    + "  return links.map(function(a){return a.getAttribute('href');}).filter(Boolean);"
    + "}"
    + "function renderCurrentUrls(fieldName, urls){"
    + "  var box=document.getElementById('cur_'+fieldName);"
    + "  if(!box) return;"
    + "  if(!urls || !urls.length){ box.innerHTML='<b>Current files:</b> Not uploaded'; return; }"
    + "  var parts=urls.map(function(u,i){"
    + "    var line='<a href=\"'+escHtml(u)+'\" target=\"_blank\" rel=\"noopener noreferrer\">Open '+(i+1)+'</a>';"
    + "    if(!PORTAL_LOCKED){ line += ' <button type=\"button\" onclick=\"deleteDocUrl(\\''+fieldName+'\\',\\''+encodeURIComponent(u)+'\\')\">Delete</button>'; }"
    + "    return line;"
    + "  });"
    + "  box.innerHTML='<b>Current files:</b><br/>'+parts.join('<br/>');"
    + "}"
    + "function onDocFileChange(fieldName, isMultiple){"
    + "  var input=document.getElementById('f_'+fieldName);"
    + "  var btn=document.getElementById('btn_'+fieldName);"
    + "  if(!input){ return; }"
    + "  var files=[].slice.call(input.files||[]);"
    + "  if(!files.length){ if(btn) btn.disabled=false; setRowMsg(fieldName,'Select a file first.',true); return; }"
    + "  if(!isMultiple) files=[files[0]];"
    + "  var tooBig=files.some(function(f){ return Number(f && f.size || 0) > Math.max(1,Number(PORTAL_MAX_UPLOAD_BYTES||0)); });"
    + "  if(tooBig){ if(btn) btn.disabled=false; setUploadUiError_(fieldName,'File too large. Maximum '+String(Math.round((Number(PORTAL_MAX_UPLOAD_BYTES||0)/(1024*1024))||5))+' MB allowed.',{errCode:'FILE_TOO_LARGE',stage:'validate'}); return; }"
    + "  var names=files.map(function(f){ return f.name; }).join(', ');"
    + "  clearPortalError_();"
    + "  setRowMsg(fieldName,'Selected: '+escHtml(names),false);"
    + "  setUploadStatusPanel_(fieldName,{text:'Ready',stage:'validate',driveMode:'B64'});"
    + "  if(btn) btn.disabled=false;"
    + "  setRowMsg(fieldName,'Selected: '+escHtml(names)+'<br/>Click Upload / Replace to submit.',false);"
    + "}"
    + "function uploadFile(fieldName){"
    + "  if(PORTAL_LOCKED){ setUploadUiError_(fieldName,'Uploads are disabled (locked).',{errCode:'LOCKED',stage:'validate'}); return; }"
    + "  var input=document.getElementById('f_'+fieldName);"
    + "  var btn=document.getElementById('btn_'+fieldName);"
    + "  if(!input || !btn){ setUploadUiError_(fieldName,'Upload controls not found.',{errCode:'UI_MISSING',stage:'validate'}); return; }"
    + "  if(PORTAL_UPLOAD_PENDING[fieldName]){ return; }"
    + "  var files=[].slice.call(input.files||[]);"
    + "  if(!files.length){ setUploadUiError_(fieldName,'Please select a file.',{errCode:'NO_FILE',stage:'validate'}); return; }"
    + "  var file=files[0];"
    + "  if(Number(file && file.size || 0) > Math.max(1,Number(PORTAL_MAX_UPLOAD_BYTES||0))){ setUploadUiError_(fieldName,'File too large. Maximum '+String(Math.round((Number(PORTAL_MAX_UPLOAD_BYTES||0)/(1024*1024))||5))+' MB allowed.',{errCode:'FILE_TOO_LARGE',stage:'validate'}); return; }"
    + "  var dbg=makeClientDebugId_();"
    + "  var dbgEl=document.getElementById('dbg_'+fieldName);"
    + "  if(dbgEl) dbgEl.value=dbg;"
    + "  PORTAL_UPLOAD_PENDING[fieldName]={dbg:dbg,startedAt:Date.now()};"
    + "  setUploadBusy(fieldName,true);"
    + "  clearPortalError_();"
    + "  setRowMsg(fieldName,'Reading file...',false);"
    + "  setUploadStatusPanel_(fieldName,{text:'Reading...',dbg:dbg,stage:'read',driveMode:'B64'});"
    + "  var reader=new FileReader();"
    + "  reader.onerror=function(){"
    + "    delete PORTAL_UPLOAD_PENDING[fieldName];"
    + "    setUploadBusy(fieldName,false);"
    + "    setUploadUiError_(fieldName,'Could not read file.',{dbg:dbg,errCode:'READ_FAIL',stage:'read',driveMode:'B64'});"
    + "  };"
    + "  reader.onload=function(){"
    + "    var result=String(reader.result||'');"
    + "    var comma=result.indexOf(',');"
    + "    var base64=(comma>=0)?result.slice(comma+1):result;"
    + "    if(!base64){"
    + "      delete PORTAL_UPLOAD_PENDING[fieldName];"
    + "      setUploadBusy(fieldName,false);"
    + "      setUploadUiError_(fieldName,'Could not read file.',{dbg:dbg,errCode:'READ_EMPTY',stage:'read',driveMode:'B64'});"
    + "      return;"
    + "    }"
    + "    setRowMsg(fieldName,'Uploading...',false);"
    + "    setUploadStatusPanel_(fieldName,{text:'Uploading...',dbg:dbg,stage:'call',driveMode:'B64'});"
    + "    google.script.run"
    + "      .withSuccessHandler(function(res){"
    + "        delete PORTAL_UPLOAD_PENDING[fieldName];"
    + "        setUploadBusy(fieldName,false);"
    + "        if(!res || res.ok!==true){"
    + "          var code=String((res&&res.code)||'UPLOAD_FAILED');"
    + "          var dbgOut=String((res&&res.dbg)||dbg);"
    + "          var msg=String((res&&res.message)||('Upload failed ('+code+').'));"
    + "          setUploadUiError_(fieldName,msg+(dbgOut?(' Debug: '+dbgOut):''),{dbg:dbgOut,errCode:code,stage:'call',driveMode:'B64'});"
    + "          return;"
    + "        }"
    + "        var urls=(res.currentUrls && res.currentUrls.length)?res.currentUrls:getCurrentUrlsFromDom(fieldName);"
    + "        if(res.fileUrl && urls.indexOf(String(res.fileUrl))<0){ urls.push(String(res.fileUrl)); }"
    + "        renderCurrentUrls(fieldName, urls);"
    + "        markDocStatusPendingReviewUi_(fieldName);"
    + "        setUploadUiSuccess_(fieldName,'Uploaded (Pending Review).',{dbg:String(res.dbg||dbg),stage:'done',driveMode:'B64',fileUrl:String(res.fileUrl||'')});"
    + "        setRowMsg(fieldName,'Uploaded: '+escHtml(String(res.name||file.name||'')),false);"
    + "        try{ input.value=''; }catch(_e){}"
    + "      })"
    + "      .withFailureHandler(function(err){"
    + "        delete PORTAL_UPLOAD_PENDING[fieldName];"
    + "        setUploadBusy(fieldName,false);"
    + "        setUploadUiError_(fieldName,'Upload failed: '+(err&&err.message?err.message:err),{dbg:dbg,errCode:'RPC_FAIL',stage:'call',driveMode:'B64'});"
    + "      })"
    + "      .portalUploadBase64({id:PORTAL_UPLOAD_ID,s:PORTAL_UPLOAD_S,field:fieldName,name:String(file&&file.name||''),mime:String(file&&file.type||''),base64:base64});"
    + "  };"
    + "  reader.readAsDataURL(file);"
    + "}"
    + "function bindPortalUploadUi_(){"
    + "  var inputs=[].slice.call(document.querySelectorAll('[data-upload-input]'));"
    + "  inputs.forEach(function(input){"
    + "    if(!input || input.__fodeBoundUploadChange) return;"
    + "    input.__fodeBoundUploadChange=true;"
    + "    input.addEventListener('change', function(){"
    + "      var field=String(input.getAttribute('data-field')||'');"
    + "      var isMultiple=String(input.getAttribute('data-multi')||'0')==='1';"
    + "      onDocFileChange(field, isMultiple);"
    + "    });"
    + "  });"
    + "  var btns=[].slice.call(document.querySelectorAll('[data-upload-btn]'));"
    + "  btns.forEach(function(btn){"
    + "    if(!btn || btn.__fodeBoundUploadClick) return;"
    + "    btn.__fodeBoundUploadClick=true;"
    + "    btn.addEventListener('click', function(){"
    + "      var field=String(btn.getAttribute('data-field')||'');"
    + "      if(!field) return;"
    + "      uploadFile(field);"
    + "    });"
    + "  });"
    + "}"
    + "if(document.readyState==='loading'){ document.addEventListener('DOMContentLoaded', bindPortalUploadUi_); } else { bindPortalUploadUi_(); }"
    + "function deleteDocUrl(fieldName, encodedUrl){"
    + "  if(PORTAL_LOCKED){ return; }"
    + "  var msg=document.getElementById('msg_'+fieldName);"
    + "  var url=decodeURIComponent(encodedUrl||'');"
    + "  if(!url){ if(msg) msg.innerHTML='Invalid file URL.'; return; }"
    + "  setUploadBusy(fieldName,true);"
    + "  if(msg) msg.innerHTML='Deleting...';"
    + "  google.script.run"
    + "    .withSuccessHandler(function(res){"
    + "      if(!res || res.ok!==true){"
    + "        var redirect=(res&&res.redirectUrl)?String(res.redirectUrl):'';"
    + "        if(msg) msg.innerHTML='Delete failed: '+((res&&res.error)?res.error:'Unknown error');"
    + "        if(redirect){ window.location.replace(redirect); return; }"
        + "        setUploadBusy(fieldName,false);"
        + "        return;"
        + "      }"
    + "      if(msg) msg.innerHTML='Deleted.'+(res.warning?(' '+res.warning):'');"
    + "      window.location.replace(" + JSON.stringify(portalReloadUrl) + ");"
    + "    })"
    + "    .withFailureHandler(function(err){"
    + "      if(msg) msg.innerHTML='Delete failed: '+(err&&err.message?err.message:err);"
    + "      setUploadBusy(fieldName,false);"
    + "    })"
    + "    .portal_deleteUploadedFile({ applicantId:'" + esc_(id) + "', secret:'" + esc_(secret) + "', field:fieldName, url:url });"
    + "}"
    + "</script>";

  return out || "<i>No document fields configured.</i>";
}

function renderAllowlistSummary_(record, visibleFields) {
  var keys = visibleFields || [];
  var deny = { Physical_Exam_Site: true };
  var docFields = CONFIG.DOC_FIELDS || [];
  for (var df = 0; df < docFields.length; df++) {
    var fileField = clean_(docFields[df] && docFields[df].file);
    if (fileField) deny[fileField] = true;
  }
  var rows = "";

  for (var i = 0; i < keys.length; i++) {
    var k = keys[i];
    if (deny[k]) continue;
    if (/_File$/i.test(k)) continue;
    var rawVal = record[k];
    var v = clean_(rawVal);
    var display = v ? v : "-";

    // Pretty subjects display
    if (k === "Subjects_Selected_Canonical") {
      var csv = clean_(record.Subjects_Selected_Canonical || "") || subjectsToCsv_(record.Subjects_Selected || "");
      display = csv ? csv : "-";
    }

    // Keep DOB summary human-friendly while date input remains yyyy-mm-dd.
    if (k === "Date_Of_Birth") {
      if (rawVal instanceof Date && !isNaN(rawVal.getTime())) {
        var tz = Session.getScriptTimeZone() || "Pacific/Port_Moresby";
        display = Utilities.formatDate(rawVal, tz, "dd/MM/yyyy");
      } else {
        display = v || "-";
      }
    }

    // make URLs clickable
    if (/^https?:\/\//i.test(display)) {
      display = "<a target='_blank' href='" + esc_(display) + "'>Open</a>";
    } else {
      display = esc_(display);
    }

    rows += "<tr>"
      + "<td style='padding:6px 10px;border-bottom:1px solid #f0f0f0;vertical-align:top;width:34%;'><b>" + esc_(k) + "</b></td>"
      + "<td style='padding:6px 10px;border-bottom:1px solid #f0f0f0;vertical-align:top;'>" + display + "</td>"
      + "</tr>";
  }

  return "<table style='width:100%;border-collapse:collapse;font-size:13px;'>" + rows + "</table>";
}

function mapDocStatusForDisplay_(raw) {
  var s = clean_(raw);
  if (!s) return "Pending Review";
  var upper = s.toUpperCase();
  if (upper === "VERIFIED") return "Verified";
  if (upper === "REJECTED") return "Rejected";
  if (upper === "FRAUDULENT") return "Fraudulent";
  if (upper === "PENDING_REVIEW") return "Pending Review";
  if (upper === "PENDING REVIEW") return "Pending Review";
  if (upper === "PENDING") return "Pending Review";
  if (upper === "VERIFIED" || upper === "REJECTED" || upper === "FRAUDULENT") return s;
  if (s === "Verified" || s === "Rejected" || s === "Fraudulent" || s === "Pending Review") return s;
  return "Pending Review";
}

function renderEditableFields_(record, editFields, dis) {
  var out = "";
  var ro = (dis === "disabled") ? "readonly" : "";
  for (var i = 0; i < editFields.length; i++) {
    var h = editFields[i];
    var val = clean_(record[h] || "");
    var isUrl = isHttpUrl_(val);
    var linkHtml = isUrl ? (" <a target='_blank' rel='noopener' href='" + esc_(val) + "'>Open</a>") : "";
    out += "<div style='margin:10px 0;'>"
      + "<label><b>" + esc_(h) + ":</b></label><br/>"
      + "<input type='text' name='" + esc_(h) + "' value='" + esc_(val) + "' style='padding:8px;width:520px;' " + ro + " />"
      + linkHtml
      + "</div>";
  }
  return out || "<div><i>No additional editable fields configured.</i></div>";
}

function renderErrorHtml_(msg) {
  return '<!doctype html><html><body style="font-family:Arial;max-width:780px;margin:24px auto;padding:0 16px;">'
    + "<h2>FODE Student Portal</h2>"
    + '<div style="background:#ffecec;border:1px solid #ffb3b3;padding:12px;border-radius:10px;">'
    + esc_(msg) + "</div></body></html>";
}

function renderExpiredHtml_() {
  return '<!doctype html><html><body style="font-family:Arial;max-width:780px;margin:24px auto;padding:0 16px;">'
    + "<h2>FODE Student Portal</h2>"
    + '<div style="background:#fff6e5;border:1px solid #f5c26b;padding:12px;border-radius:10px;">'
    + "This portal link has expired. Please contact admissions for a new link."
    + "</div></body></html>";
}

function renderSuccessHtml_(applicantId) {
  return '<!doctype html><html><body style="font-family:Arial;max-width:780px;margin:24px auto;padding:0 16px;">'
    + "<h2>FODE Student Portal</h2>"
    + '<div style="background:#e8fff0;border:1px solid #9be7b2;padding:12px;border-radius:10px;">'
    + "Updates saved successfully for <b>" + esc_(applicantId) + "</b>."
    + "</div></body></html>";
}

function getStudentActionUrl_() {
  var studentExec = clean_(CONFIG.WEBAPP_URL_STUDENT_EXEC || "");
  var hasStudentExec = /^https:\/\/script\.google\.com\//i.test(studentExec);
  if (hasStudentExec) return { url: getStudentBaseUrl_() || studentExec, isStudentReady: true, warning: "" };
  var studentUrl = clean_(CONFIG.WEBAPP_URL_STUDENT || "");
  var isStudentReady = /^https:\/\/script\.google\.com\//i.test(studentUrl);
  var url = isStudentReady ? (getStudentBaseUrl_() || studentUrl) : "";
  var warning = isStudentReady ? "" : "Student URL not configured. Saving may not work for external users.";
  return {
    url: url,
    isStudentReady: isStudentReady,
    warning: warning
  };
}

function boundaryFromContentType_(ct) {
  var s = String(ct || "");
  var m = s.match(/boundary=(?:"([^"]+)"|([^;]+))/i);
  return clean_((m && (m[1] || m[2])) || "");
}

function boundaryFromContents_(contents) {
  var line = (typeof firstLine_ === "function")
    ? firstLine_(contents, 400)
    : String(contents || "").split(/\r?\n/, 1)[0];
  if (line.indexOf("--") !== 0) return "";
  return clean_(line.substring(2));
}

function parseMultipartForm_(e) {
  try {
    var pd = (e && e.postData) ? e.postData : null;
    var contentType = String((pd && (pd.type || pd.contentType)) || "");
    var raw = (pd && typeof pd.contents === "string") ? pd.contents : "";
    var contentLen = raw ? raw.length : 0;
    var firstLinePrefix = (typeof firstLine_ === "function") ? firstLine_(raw, 80) : String(raw || "").slice(0, 80);
    var boundary = boundaryFromContentType_(contentType);
    var boundaryFoundFrom = boundary ? "type" : "none";
    var boundaryDerived = false;
    if (!boundary) {
      boundary = boundaryFromContents_(raw);
      boundaryDerived = !!boundary;
      boundaryFoundFrom = boundary ? "contents" : "none";
    }
    if (!boundary) {
      return {
        ok: false,
        code: "MULTIPART_PARSE_FAIL",
        message: "Multipart parse failed. Missing boundary.",
        contentTypeSeen: (typeof truncate_ === "function") ? truncate_(contentType, 120) : String(contentType || "").slice(0, 120),
        contentLen: contentLen,
        boundaryDerived: false,
        boundaryFoundFrom: "none",
        firstLinePrefix: clean_(firstLinePrefix || "")
      };
    }
    if (!raw) {
      return {
        ok: false,
        code: "MULTIPART_PARSE_FAIL",
        message: "Empty multipart body.",
        contentTypeSeen: (typeof truncate_ === "function") ? truncate_(contentType, 120) : String(contentType || "").slice(0, 120),
        contentLen: contentLen,
        boundaryDerived: boundaryDerived,
        boundaryFoundFrom: boundaryFoundFrom,
        firstLinePrefix: clean_(firstLinePrefix || "")
      };
    }

    var out = {
      ok: true,
      fields: {},
      files: {},
      contentTypeSeen: (typeof truncate_ === "function") ? truncate_(contentType, 120) : String(contentType || "").slice(0, 120),
      contentLen: contentLen,
      boundaryDerived: boundaryDerived,
      boundaryFoundFrom: boundaryFoundFrom,
      firstLinePrefix: clean_(firstLinePrefix || "")
    };
    var delimiter = "--" + boundary;
    var parts = raw.split(delimiter);
    for (var i = 0; i < parts.length; i++) {
      var part = parts[i];
      if (!part) continue;
      if (part === "--" || part === "--\r\n" || part === "--\n") continue;
      if (/^\r?\n/.test(part)) part = part.replace(/^\r?\n/, "");
      if (!part || part === "--") continue;
      if (/--\r?\n?$/.test(part)) part = part.replace(/--\r?\n?$/, "");
      var sepIdx = part.indexOf("\r\n\r\n");
      var sepLen = 4;
      if (sepIdx < 0) {
        sepIdx = part.indexOf("\n\n");
        sepLen = 2;
      }
      if (sepIdx < 0) continue;
      var headerText = part.slice(0, sepIdx);
      var body = part.slice(sepIdx + sepLen);
      body = body.replace(/\r?\n$/, "");

      var headers = {};
      var lines = headerText.split(/\r?\n/);
      for (var h = 0; h < lines.length; h++) {
        var line = String(lines[h] || "");
        var colon = line.indexOf(":");
        if (colon < 0) continue;
        var hk = line.slice(0, colon).trim().toLowerCase();
        var hv = line.slice(colon + 1).trim();
        headers[hk] = hv;
      }
      var disp = String(headers["content-disposition"] || "");
      if (!disp) continue;
      var nameMatch = disp.match(/name="([^"]*)"/i);
      var fileMatch = disp.match(/filename="([^"]*)"/i);
      var fieldName = clean_((nameMatch && nameMatch[1]) || "");
      if (!fieldName) continue;
      if (fileMatch && fileMatch[1] !== undefined) {
        var fileName = String(fileMatch[1] || "");
        var ct = clean_(headers["content-type"] || "application/octet-stream");
        var blob = Utilities.newBlob(body || "", ct || "application/octet-stream", fileName || "upload.bin");
        var bytes = blob.getBytes();
        var fileObj = {
          fileName: fileName,
          contentType: ct,
          bytes: bytes,
          blob: blob
        };
        if (hasOwn_(out.files, fieldName)) {
          if (!Array.isArray(out.files[fieldName])) out.files[fieldName] = [out.files[fieldName]];
          out.files[fieldName].push(fileObj);
        } else {
          out.files[fieldName] = fileObj;
        }
      } else {
        out.fields[fieldName] = body;
      }
    }
    return out;
  } catch (err) {
    return {
      ok: false,
      code: "MULTIPART_PARSE_FAIL",
      message: String(err && err.message ? err.message : err),
      contentTypeSeen: "",
      contentLen: 0,
      boundaryDerived: false,
      boundaryFoundFrom: "none",
      firstLinePrefix: ""
    };
  }
}

function portal_uploadMultipart_(e) {
  var t0 = nowMs_();
  var dbg = makeDebugId_();
  var reqParams = (e && e.parameter && typeof e.parameter === "object") ? e.parameter : {};
  var redirectApplicantId = clean_(reqParams.id || reqParams.applicantId || "");
  var redirectSecret = clean_(reqParams.s || reqParams.secret || "");
  var diagBase = {
    runtimeVersion: clean_(CONFIG.VERSION || ""),
    portalBuildParam: clean_(typeof getParam_ === "function" ? getParam_(e, "portalBuild") : (reqParams.portalBuild || "")),
    viewParam: clean_(typeof getParam_ === "function" ? getParam_(e, "view") : (reqParams.view || "")),
    method: "POST"
  };
  function diag_(extra) {
    var out = {};
    var k;
    for (k in diagBase) if (Object.prototype.hasOwnProperty.call(diagBase, k)) out[k] = diagBase[k];
    if (extra && typeof extra === "object") {
      for (k in extra) if (Object.prototype.hasOwnProperty.call(extra, k)) out[k] = extra[k];
    }
    return out;
  }
  function out_(ok, code, message, extra) {
    var x = (extra && typeof extra === "object") ? extra : {};
    var dbgOut = clean_(x.debugId || dbg);
    var errCodeOut = clean_(x.errCode || code || "");
    var idOut = clean_(x.applicantId || redirectApplicantId || "");
    var secretOut = clean_(x.secret || redirectSecret || "");
    var docKeyOut = clean_(x.docKey || "");
    logExecTrace_("PORTAL_UPLOAD_RESULT_RENDER", dbgOut || dbg, {
      dbg: dbgOut || dbg,
      ok: ok ? 1 : 0,
      docKey: docKeyOut,
      runtimeVersion: clean_(CONFIG.VERSION || "")
    });
    try {
      return renderPortalPageResponse_(e, {
        applicantId: idOut,
        secret: secretOut,
        viewName: "portal",
        uploadResult: {
          ok: !!ok,
          dbg: dbgOut || dbg,
          errCode: ok ? "" : errCodeOut,
          docKey: docKeyOut,
          message: clean_(message || "")
        }
      });
    } catch (renderErr) {
      logExecTrace_("PORTAL_UPLOAD_FAIL", dbgOut || dbg, diag_({
        code: "PORTAL_UPLOAD_RENDER_FAIL",
        stage: "render",
        err: clean_(stringifyGsError_(renderErr) || "")
      }));
      return htmlOutput_(renderErrorHtml_("Upload processed, but portal page could not be rendered. Debug: " + (dbgOut || dbg)));
    }
  }

  try {
    var pd = (e && e.postData) ? e.postData : null;
    var pdKeys = (typeof safeKeys_ === "function") ? safeKeys_(pd) : (pd && typeof pd === "object" ? Object.keys(pd) : []);
    var parameterKeys = (typeof safeKeys_ === "function") ? safeKeys_(reqParams) : (reqParams && typeof reqParams === "object" ? Object.keys(reqParams) : []);
    var nativeParamFile = (reqParams && typeof reqParams === "object") ? reqParams.file : null;
    var nativeFileName = "";
    var nativeFileBytes = [];
    var nativeFileBytesLen = 0;
    var nativeFileContentType = "";
    var hasNativeParameterFile = !!nativeParamFile;
    try { nativeFileName = (nativeParamFile && nativeParamFile.getName) ? clean_(nativeParamFile.getName()) : ""; } catch (_eName) {}
    try { nativeFileContentType = (nativeParamFile && nativeParamFile.getContentType) ? clean_(nativeParamFile.getContentType()) : ""; } catch (_eCt) {}
    try {
      if (nativeParamFile && nativeParamFile.getBytes) {
        nativeFileBytes = nativeParamFile.getBytes() || [];
        nativeFileBytesLen = Number(nativeFileBytes.length || 0);
      }
    } catch (_eBytes) {
      nativeFileBytes = [];
      nativeFileBytesLen = 0;
    }
    var pdContentLen = (pd && typeof pd.contents === "string") ? pd.contents.length : 0;
    var parsed = null;
    if (nativeParamFile && nativeFileBytesLen > 0) {
      parsed = {
        ok: true,
        fields: reqParams,
        files: {
          file: {
            fileName: nativeFileName || "portal-upload.bin",
            contentType: nativeFileContentType || "application/octet-stream",
            bytes: nativeFileBytes,
            blob: nativeParamFile
          }
        },
        contentTypeSeen: clean_((pd && (pd.type || pd.contentType)) || ""),
        contentLen: pdContentLen,
        boundaryDerived: false,
        boundaryFoundFrom: "native_parameter_file",
        firstLinePrefix: ""
      };
    } else if (!pd || pdContentLen <= 0) {
      logExecTrace_("PORTAL_UPLOAD_ENTER", dbg, {
        dbg: dbg,
        runtimeVersion: clean_(CONFIG.VERSION || ""),
        view: clean_(typeof getParam_ === "function" ? getParam_(e, "view") : (reqParams.view || "")),
        hasPostData: !!pd,
        postDataKeys: pdKeys,
        hasParameterFile: hasNativeParameterFile,
        parameterKeys: parameterKeys,
        fileName: nativeFileName,
        fileBytesLen: nativeFileBytesLen,
        contentTypeSeen: clean_((pd && (pd.type || pd.contentType)) || ""),
        contentLen: Number(pdContentLen || 0),
        applicantId: clean_(reqParams.id || reqParams.applicantId || ""),
        docKey: clean_(reqParams.docKey || reqParams.field || "")
      });
      logExecTrace_("PORTAL_UPLOAD_FAIL", dbg, diag_({
        code: "NO_POSTDATA",
        stage: "parse",
        contentTypeSeen: clean_((pd && (pd.type || pd.contentType)) || ""),
        contentLen: Number(pdContentLen || 0)
      }));
      return out_(false, "NO_POSTDATA", "Upload payload missing (empty POST). This usually means the browser did not submit the form containing the file input.", {
        stage: "parse",
        contentTypeSeen: clean_((pd && (pd.type || pd.contentType)) || ""),
        contentLen: Number(pdContentLen || 0),
        postDataKeys: pdKeys,
        applicantId: clean_(reqParams.id || reqParams.applicantId || ""),
        secret: clean_(reqParams.s || reqParams.secret || ""),
        docKey: clean_(reqParams.docKey || reqParams.field || "")
      });
    } else {
      parsed = parseMultipartForm_(e);
    }
    var enterFields = (parsed && parsed.fields && typeof parsed.fields === "object") ? parsed.fields : {};
    logExecTrace_("PORTAL_UPLOAD_ENTER", dbg, {
      dbg: dbg,
      runtimeVersion: clean_(CONFIG.VERSION || ""),
      view: clean_(typeof getParam_ === "function" ? getParam_(e, "view") : (reqParams.view || "")),
      hasPostData: !!pd,
      postDataKeys: pdKeys,
      hasParameterFile: hasNativeParameterFile,
      parameterKeys: parameterKeys,
      fileName: nativeFileName,
      fileBytesLen: nativeFileBytesLen,
      contentTypeSeen: clean_(parsed && parsed.contentTypeSeen || ((pd && (pd.type || pd.contentType)) || "")),
      contentLen: Number(parsed && parsed.contentLen || ((pd && typeof pd.contents === "string") ? pd.contents.length : 0)),
      applicantId: clean_(enterFields.id || enterFields.applicantId || reqParams.id || reqParams.applicantId || ""),
      docKey: clean_(enterFields.docKey || enterFields.field || reqParams.docKey || reqParams.field || "")
    });
    logExecTrace_("PORTAL_UPLOAD_DIAG", dbg, diag_({
      hasPostData: !!pd,
      postDataKeys: pdKeys,
      contentTypeSeen: clean_(parsed && parsed.contentTypeSeen || ((pd && (pd.type || pd.contentType)) || "")),
      contentLen: Number(parsed && parsed.contentLen || ((pd && typeof pd.contents === "string") ? pd.contents.length : 0)),
      firstLinePrefix: clean_(parsed && parsed.firstLinePrefix || ((typeof firstLine_ === "function") ? firstLine_((pd && pd.contents) || "", 80) : "")),
      boundaryFoundFrom: clean_(parsed && parsed.boundaryFoundFrom || "none"),
      boundaryDerived: !!(parsed && parsed.boundaryDerived)
    }));
    if (!parsed || parsed.ok !== true) {
      logExecTrace_("PORTAL_UPLOAD_FAIL", dbg, diag_({
        code: "MULTIPART_PARSE_FAIL",
        stage: "parse",
        err: clean_(parsed && parsed.message || ""),
        contentTypeSeen: clean_(parsed && parsed.contentTypeSeen || ""),
        contentLen: Number(parsed && parsed.contentLen || 0),
        boundaryDerived: !!(parsed && parsed.boundaryDerived),
        boundaryFoundFrom: clean_(parsed && parsed.boundaryFoundFrom || "none"),
        firstLinePrefix: clean_(parsed && parsed.firstLinePrefix || "")
      }));
      return out_(false, "MULTIPART_PARSE_FAIL", clean_(parsed && parsed.message || "Multipart parse failed. Missing boundary."), {
        stage: "parse",
        contentTypeSeen: clean_(parsed && parsed.contentTypeSeen || ""),
        boundaryDerived: !!(parsed && parsed.boundaryDerived),
        boundaryFoundFrom: clean_(parsed && parsed.boundaryFoundFrom || "none"),
        contentLen: Number(parsed && parsed.contentLen || 0),
        postDataKeys: pdKeys,
        firstLinePrefix: clean_(parsed && parsed.firstLinePrefix || "")
      });
    }

    var fields = parsed.fields || {};
    var queryParams = (e && e.parameter && typeof e.parameter === "object") ? e.parameter : {};
    var queryId = clean_(queryParams.id || queryParams.applicantId || "");
    var queryS = clean_(queryParams.s || queryParams.secret || "");
    var multipartId = clean_(fields.id || fields.applicantId || "");
    var multipartS = clean_(fields.s || fields.secret || "");
    var recvIdSource = "missing";
    var applicantId = "";
    var secret = "";
    if (queryId && queryS) {
      applicantId = queryId;
      secret = queryS;
      recvIdSource = "query";
    } else if (multipartId && multipartS) {
      applicantId = multipartId;
      secret = multipartS;
      recvIdSource = "multipart";
    } else {
      applicantId = clean_(queryId || multipartId || "");
      secret = clean_(queryS || multipartS || "");
      recvIdSource = "missing";
    }
    redirectApplicantId = applicantId || redirectApplicantId;
    redirectSecret = secret || redirectSecret;
    var recvSHashed = (typeof secretHashPrefix_ === "function")
      ? secretHashPrefix_(secret)
      : clean_(hashPortalSecret_(secret || "")).slice(0, 8);
    var docKey = clean_(fields.docKey || fields.field || "");
    var postedDbg = clean_(fields.dbg || "");
    if (postedDbg) dbg = postedDbg;
    var fileEntry = parsed.files && parsed.files.file ? parsed.files.file : null;
    if (Array.isArray(fileEntry)) fileEntry = fileEntry[0] || null;
    if (!applicantId || !secret) {
      return out_(false, "TOKEN_MISSING", "Missing portal token.", {
        stage: "auth",
        docKey: docKey,
        recvIdSource: recvIdSource,
        recvSHashed: recvSHashed,
        contentTypeSeen: clean_(parsed.contentTypeSeen || ""),
        boundaryDerived: !!parsed.boundaryDerived,
        boundaryFoundFrom: clean_(parsed.boundaryFoundFrom || "none"),
        contentLen: Number(parsed.contentLen || 0)
      });
    }
    if (!docKey) {
      return out_(false, "INVALID_FIELD", "Missing document field.", {
        stage: "validate",
        recvIdSource: recvIdSource,
        recvSHashed: recvSHashed,
        contentTypeSeen: clean_(parsed.contentTypeSeen || ""),
        boundaryDerived: !!parsed.boundaryDerived,
        boundaryFoundFrom: clean_(parsed.boundaryFoundFrom || "none"),
        contentLen: Number(parsed.contentLen || 0)
      });
    }

    if (!fileEntry || !Array.isArray(fileEntry.bytes) || !fileEntry.bytes.length) {
      logExecTrace_("PORTAL_UPLOAD_FAIL", dbg, diag_({ code: "NO_FILE", stage: "parse", applicantId: applicantId, docKey: docKey }));
      return out_(false, "NO_FILE", "Please select a file.", {
        stage: "parse",
        docKey: docKey,
        recvIdSource: recvIdSource,
        recvSHashed: recvSHashed,
        contentTypeSeen: clean_(parsed.contentTypeSeen || ""),
        boundaryDerived: !!parsed.boundaryDerived,
        boundaryFoundFrom: clean_(parsed.boundaryFoundFrom || "none"),
        contentLen: Number(parsed.contentLen || 0)
      });
    }

    var ss = getWorkingSpreadsheet_();
    var sheet = mustGetDataSheet_(ss);
    var found = findPortalRowByIdSecret_(sheet, applicantId, secret);
    if (!found) {
      logExecTrace_("PORTAL_UPLOAD_FAIL", dbg, diag_({ code: "TOKEN_INVALID", stage: "auth", applicantId: applicantId, docKey: docKey }));
      return out_(false, "TOKEN_INVALID", "Invalid or expired portal link token.", {
        stage: "auth",
        docKey: docKey,
        recvIdSource: recvIdSource,
        recvSHashed: recvSHashed,
        contentTypeSeen: clean_(parsed.contentTypeSeen || ""),
        boundaryDerived: !!parsed.boundaryDerived,
        boundaryFoundFrom: clean_(parsed.boundaryFoundFrom || "none"),
        contentLen: Number(parsed.contentLen || 0)
      });
    }
    if (String(found.record[SCHEMA.PORTAL_ACCESS_STATUS] || "").trim() === "Locked") {
      return out_(false, "ACCESS_LOCKED", "Portal access is locked. Please contact admissions.", {
        stage: "auth",
        docKey: docKey,
        recvIdSource: recvIdSource,
        recvSHashed: recvSHashed,
        contentTypeSeen: clean_(parsed.contentTypeSeen || ""),
        boundaryDerived: !!parsed.boundaryDerived,
        boundaryFoundFrom: clean_(parsed.boundaryFoundFrom || "none"),
        contentLen: Number(parsed.contentLen || 0)
      });
    }
    if (isPaymentFreezeActive_(found.record)) {
      logExecTrace_("PORTAL_UPLOAD_FAIL", dbg, diag_({ code: "PAYMENT_FREEZE", stage: "auth", applicantId: applicantId, docKey: docKey }));
      return out_(false, "PAYMENT_FREEZE", "Uploads are disabled after payment verification.", {
        stage: "auth",
        docKey: docKey,
        recvIdSource: recvIdSource,
        recvSHashed: recvSHashed
      });
    }

    var fileName = clean_(fileEntry.fileName || fields.name || "portal-upload.bin");
    var mimeType = clean_(fileEntry.contentType || fields.mimeType || "application/octet-stream");
    var b64 = Utilities.base64Encode(fileEntry.bytes);
    var res = portalUpload_fromUi_({
      id: applicantId,
      s: secret,
      field: docKey,
      name: fileName,
      mimeType: mimeType,
      b64: b64,
      dbg: dbg
    }) || {};

    var code = clean_(res.errCode || (res.ok ? "OK" : "UPLOAD_FAILED")) || (res.ok ? "OK" : "UPLOAD_FAILED");
    if (res.ok === true) {
      logExecTrace_("PORTAL_UPLOAD", dbg, {
        applicantId: applicantId,
        docKey: docKey,
        fileName: fileName,
        byteSize: Number(fileEntry.bytes.length || 0),
        serverMs: Number(res.serverMs || elapsedMs_(t0)),
        driveMode: clean_(res.driveMode || "")
      });
      logExecTrace_("PORTAL_UPLOAD_OK", dbg, {
        dbg: clean_(res.dbg || dbg),
        docKey: docKey,
        runtimeVersion: clean_(CONFIG.VERSION || ""),
        applicantId: applicantId,
        driveMode: clean_(res.driveMode || ""),
        serverMs: Number(res.serverMs || elapsedMs_(t0))
      });
      return out_(true, "OK", "Uploaded", {
        debugId: clean_(res.dbg || dbg),
        applicantId: applicantId,
        secret: secret,
        docKey: docKey,
        fileUrl: res.fileUrl,
        fileId: res.fileId,
        currentUrls: res.currentUrls,
        driveMode: res.driveMode,
        stage: res.stage || "done",
        serverMs: res.serverMs,
        recvIdSource: recvIdSource,
        recvSHashed: recvSHashed,
        contentTypeSeen: clean_(parsed.contentTypeSeen || ""),
        boundaryDerived: !!parsed.boundaryDerived,
        boundaryFoundFrom: clean_(parsed.boundaryFoundFrom || "none"),
        contentLen: Number(parsed.contentLen || 0)
      });
    }

    logExecTrace_("PORTAL_UPLOAD_FAIL", dbg, diag_({
      applicantId: applicantId,
      docKey: docKey,
      fileName: fileName,
      byteSize: Number(fileEntry.bytes.length || 0),
      stage: clean_(res.stage || "call"),
      errCode: clean_(res.errCode || code),
      serverMs: Number(res.serverMs || elapsedMs_(t0))
    }));
    return out_(false, code.toUpperCase(), clean_(res.err || "Upload failed."), {
      debugId: clean_(res.dbg || dbg),
      applicantId: applicantId,
      secret: secret,
      docKey: docKey,
      fileUrl: res.fileUrl,
      fileId: res.fileId,
      currentUrls: res.currentUrls,
      driveMode: res.driveMode,
      stage: res.stage || "call",
      serverMs: res.serverMs,
      errCode: res.errCode || code,
      recvIdSource: recvIdSource,
      recvSHashed: recvSHashed,
      contentTypeSeen: clean_(parsed.contentTypeSeen || ""),
      boundaryDerived: !!parsed.boundaryDerived,
      boundaryFoundFrom: clean_(parsed.boundaryFoundFrom || "none"),
      contentLen: Number(parsed.contentLen || 0)
    });
  } catch (err) {
    var msg = String(err && err.message ? err.message : err);
    logExecTrace_("PORTAL_UPLOAD_FAIL", dbg, diag_({ code: "PORTAL_UPLOAD_EXCEPTION", stage: "call", e: safeErr_(err) }));
    return out_(false, "PORTAL_UPLOAD_EXCEPTION", msg || "Upload failed.", { stage: "call" });
  }
}

function buildPortalRedirectUrl_(applicantId, secret) {
  var baseUrl = canonicalStudentExecBase_();
  var opts = (arguments.length > 2 && arguments[2]) ? arguments[2] : {};
  var sep = baseUrl.indexOf("?") === -1 ? "?" : "&";
  var url = baseUrl
    + sep + "view=portal";
  var idNorm = clean_(applicantId);
  var secretNorm = clean_(secret);
  if (idNorm) url += "&id=" + encodeURIComponent(idNorm);
  if (secretNorm) url += "&s=" + encodeURIComponent(secretNorm);
  if (opts.saved === true) url += "&saved=1";
  var hasError = opts.error === true || opts.err === true;
  if (hasError) url += "&error=1";
  if (opts.locked === true) url += "&locked=1";
  if (opts.msg) url += "&msg=" + encodeURIComponent(clean_(opts.msg));
  if (opts.dbg) url += "&dbg=" + encodeURIComponent(clean_(opts.dbg));
  if (opts.uploadFail === true) url += "&uploadFail=1";
  if (opts.field) url += "&field=" + encodeURIComponent(clean_(opts.field));
  if (opts.val === true) url += "&val=1";
  if (opts.fields) url += "&fields=" + encodeURIComponent(clean_(opts.fields));
  if (opts.errCode) url += "&errCode=" + encodeURIComponent(clean_(opts.errCode));
  return url;
}

function returnPortalRedirectOutput_(url) {
  var opts = (arguments.length > 1 && arguments[1]) ? arguments[1] : {};
  var u = canonicalizePortalRedirectUrl_(url);
  if (!u) return htmlOutput_(renderErrorHtml_("Missing redirect URL"));
  var showDebugBlock = CONFIG.DEBUG_PORTAL_SHOW_ON_PAGE === true && opts.debug === true;
  var debugBlock = showDebugBlock
    ? '<div style="font-family:Arial,Helvetica,sans-serif;max-width:720px;margin:12px auto;padding:8px;border:1px dashed #999;color:#000;font-size:12px;">'
      + "<div><b>DEBUG_PORTAL_POST</b></div>"
      + "<div>reqId: " + esc_(clean_(opts.reqId || "")) + "</div>"
      + "<div>received id: " + esc_(clean_(opts.applicantId || "")) + "</div>"
      + "<div>received s(redacted): " + esc_(redactToken_(opts.secret || "")) + "</div>"
      + "<div>result.ok: " + (opts.result && opts.result.ok === true ? "true" : "false") + "</div>"
      + "<div>debugId: " + esc_(clean_(opts.debugId || (opts.result && opts.result.debugId) || "")) + "</div>"
      + "<div>redirectUrl: " + esc_(clean_(opts.redirectUrl || u)) + "</div>"
      + "<div>token validation passed: " + (opts.tokenValidationPassed === true ? "true" : "false") + "</div>"
      + "</div>"
    : "";
  var html = '<!doctype html><html><head>'
    + '<meta charset="utf-8" />'
    + '<base target="_top" />'
    + '<meta name="viewport" content="width=device-width, initial-scale=1" />'
    + '<meta http-equiv="refresh" content="0;url=' + esc_(u) + '" />'
    + '<title>Redirecting</title>'
    + '</head><body>'
    + debugBlock
    + '<div style="font-family:Arial,Helvetica,sans-serif;max-width:720px;margin:24px auto;padding:0 16px;">'
    + '<h3>Redirecting...</h3>'
    + '<p>If you are not redirected, click <a href="' + esc_(u) + '" target="_top">Continue</a>.</p>'
    + "</div>"
    + "<script>"
    + "(function(){"
    + "var t=" + JSON.stringify(u) + ";"
    + "try{ if(window.top && window.top.location){ window.top.location.replace(t); return; } }catch(e){}"
    + "window.location.replace(t);"
    + "})();"
    + "</script>"
    + '<noscript><p><a href="' + esc_(u) + '" target="_top">Continue</a></p></noscript>'
    + "</body></html>";
  return HtmlService.createHtmlOutput(html)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function canonicalizePortalRedirectUrl_(url) {
  var raw = clean_(url);
  if (!raw) return "";
  var canonicalBase = getStudentBaseUrl_() || clean_(CONFIG.WEBAPP_URL_STUDENT_EXEC || "");
  var qIndex = raw.indexOf("?");
  var query = qIndex >= 0 ? raw.slice(qIndex) : "";
  if (/^https:\/\/script\.google\.com\/a\//i.test(raw)) {
    if (canonicalBase) return canonicalBase + query;
    return raw.replace(/^https:\/\/script\.google\.com\/a\/[^/]+\//i, "https://script.google.com/");
  }
  if (canonicalBase && /^https:\/\/script\.google\.com\//i.test(raw)) {
    return canonicalBase + query;
  }
  return raw;
}

function canonicalStudentExecBase_() {
  return getStudentBaseUrl_() || clean_(CONFIG.WEBAPP_URL_STUDENT_EXEC || CONFIG.WEBAPP_URL_STUDENT || "");
}

function normalizeWebAppUrl_(url) {
  var out = clean_(url || "");
  if (!out) return "";
  out = out.replace(/\?.*$/, "");
  out = out.replace(/\/+$/, "");
  return out;
}

function isAdminDeploymentRequest_() {
  try {
    var current = normalizeWebAppUrl_(ScriptApp.getService().getUrl() || "");
    var adminBase = normalizeWebAppUrl_(CONFIG.WEBAPP_URL_ADMIN || CONFIG.WEBAPP_URL || "");

    // normalize domain-specific paths
    current = current.replace("/a/macros/", "/macros/");
    adminBase = adminBase.replace("/a/macros/", "/macros/");

    return !!(current && adminBase && current === adminBase);
  } catch (e) {
    Logger.log("ADMIN_URL_MATCH_FAIL " + (e.message || e));
    return false;
  }
}

function parsePortalPayloadField_(raw) {
  var txt = clean_(raw);
  if (!txt) return {};
  try {
    var parsed = JSON.parse(txt);
    if (!parsed || typeof parsed !== "object" || Array.isArray(parsed)) {
      throw new Error("payload must be object");
    }
    return parsed;
  } catch (e) {
    throw new Error("Invalid payload JSON");
  }
}

function mergePortalPayload_(basePayload, payloadObj) {
  var out = {};
  var src = basePayload || {};
  Object.keys(src).forEach(function (k) { out[k] = src[k]; });
  Object.keys(payloadObj || {}).forEach(function (k2) { out[k2] = payloadObj[k2]; });
  return out;
}

function outputToJsonObject_(res) {
  if (!res) return null;
  if (typeof res.getContent === "function") {
    var txt = clean_(res.getContent());
    if (!txt) return null;
    try {
      return JSON.parse(txt);
    } catch (e) {
      return null;
    }
  }
  if (typeof res === "object") return res;
  return null;
}

function logPortalPostEvent_(label, payload) {
  try {
    var ss = getWorkingSpreadsheet_();
    var sh = mustGetSheet_(ss, CONFIG.LOG_SHEET);
    log_(sh, label, JSON.stringify(payload || {}));
  } catch (e) {
    // Diagnostic logging must never break request flow.
  }
}

function mustGetDataSheet_(ss) {
  var expectedName = clean_(CONFIG.SHEET_TAB_WORKING || CONFIG.DATA_SHEET || "FODE_Data");
  var sheet = mustGetSheet_(ss, expectedName);
  if (sheet.getName() !== expectedName) {
    throw new Error("DATA_SHEET mismatch");
  }
  return sheet;
}
function htmlOutput_(html) {
  return HtmlService.createHtmlOutput(html)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/******************** LOOKUP ********************/
function resolvePortalSecretColumnIndex_(idx) {
  var map = idx || {};
  if (map.Secret) return map.Secret;
  if (map.Secret_Plain) return map.Secret_Plain;
  if (map.Secret_Hash) return map.Secret_Hash;
  return 0;
}

function getPortalSecretForApplicant_(applicantId) {
  var debugId = newDebugId_();
  var idNorm = clean_(applicantId || "");
  if (!idNorm) return { ok: false, code: "NO_SECRET", debugId: debugId };
  try {
    var ss = SpreadsheetApp.openById(PORTAL_SECRETS_SPREADSHEET_ID);
    var sh = ss.getSheetByName(PORTAL_SECRETS_TAB);
    if (!sh) return { ok: false, code: "NO_SECRET", debugId: debugId };
    var lastCol = sh.getLastColumn();
    var lastRow = sh.getLastRow();
    if (lastCol < 1 || lastRow < 2) return { ok: false, code: "NO_SECRET", debugId: debugId };
    var headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
    var idx = {};
    for (var i = 0; i < headers.length; i++) {
      var h = clean_(headers[i]);
      if (h) idx[h] = i + 1;
    }
    var secretCol = resolvePortalSecretColumnIndex_(idx);
    if (!idx.ApplicantID) return { ok: false, code: "NO_SECRET", debugId: debugId };
    if (!secretCol) return { ok: false, code: "SECRET_COLUMN_MISSING", debugId: debugId };
    var data = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
    var idLower = idNorm.toLowerCase();
    for (var r = 0; r < data.length; r++) {
      var rowId = clean_(data[r][idx.ApplicantID - 1]);
      if (rowId && rowId.toLowerCase() === idLower) {
        var secret = clean_(data[r][secretCol - 1]);
        if (!secret) return { ok: false, code: "NO_SECRET", debugId: debugId };
        return { ok: true, secret: secret };
      }
    }
    return { ok: false, code: "NO_SECRET", debugId: debugId };
  } catch (_e) {
    return { ok: false, code: "NO_SECRET", debugId: debugId };
  }
}

function setPortalSecretForApplicant_(applicantId, newSecret) {
  var debugId = newDebugId_();
  var idNorm = clean_(applicantId || "");
  var secretNorm = clean_(newSecret || "");
  if (!idNorm || !secretNorm) return { ok: false, code: "NO_SECRET", debugId: debugId };
  try {
    var ss = SpreadsheetApp.openById(PORTAL_SECRETS_SPREADSHEET_ID);
    var sh = ss.getSheetByName(PORTAL_SECRETS_TAB);
    if (!sh) return { ok: false, code: "NO_SECRET", debugId: debugId };
    var lastCol = sh.getLastColumn();
    var lastRow = sh.getLastRow();
    if (lastCol < 1 || lastRow < 2) return { ok: false, code: "NO_SECRET", debugId: debugId };
    var headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
    var idx = {};
    for (var i = 0; i < headers.length; i++) {
      var h = clean_(headers[i]);
      if (h) idx[h] = i + 1;
    }
    var secretCol = resolvePortalSecretColumnIndex_(idx);
    if (!idx.ApplicantID) return { ok: false, code: "NO_SECRET", debugId: debugId };
    if (!secretCol) return { ok: false, code: "SECRET_COLUMN_MISSING", debugId: debugId };
    var data = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
    var idLower = idNorm.toLowerCase();
    for (var r = 0; r < data.length; r++) {
      var rowId = clean_(data[r][idx.ApplicantID - 1]);
      if (rowId && rowId.toLowerCase() === idLower) {
        sh.getRange(r + 2, secretCol).setValue(secretNorm);
        return { ok: true };
      }
    }
    return { ok: false, code: "NO_SECRET", debugId: debugId };
  } catch (_e) {
    return { ok: false, code: "NO_SECRET", debugId: debugId };
  }
}

function buildStudentPortalUrl_(applicantId, secret) {
  var base = canonicalExecBase_(CONFIG.DEPLOYMENT_ID_STUDENT || CONFIG.WEBAPP_URL_STUDENT || "");
  if (!base) throw new Error("Missing canonical student exec base");
  return base
    + "?view=portal&id="
    + encodeURIComponent(clean_(applicantId || ""))
    + "&s="
    + encodeURIComponent(clean_(secret || ""));
}

function buildRuntimeTruth_(e, surfaceHint) {
  var params = (e && e.parameter && typeof e.parameter === "object") ? e.parameter : {};
  var requestedView = clean_(params.view || "").toLowerCase();
  var rawServiceUrl = "";
  var activeUser = "";
  var effectiveUser = "";
  try { rawServiceUrl = clean_(ScriptApp.getService().getUrl() || ""); } catch (_rawErr) {}
  try { activeUser = clean_(Session.getActiveUser().getEmail() || ""); } catch (_activeErr) {}
  try { effectiveUser = clean_(Session.getEffectiveUser().getEmail() || ""); } catch (_effectiveErr) {}

  var deployVersion = Number(CONFIG.DEPLOY_VERSION_NUMBER || 0);
  var adminBase = canonicalExecBase_(CONFIG.DEPLOYMENT_ID_ADMIN || CONFIG.WEBAPP_URL_ADMIN || "");
  var studentBase = canonicalExecBase_(CONFIG.DEPLOYMENT_ID_STUDENT || CONFIG.WEBAPP_URL_STUDENT || "");
  var serviceUrl = canonicalExecBase_(rawServiceUrl || adminBase || studentBase || "");
  var configAdminUrl = canonicalExecBase_(CONFIG.WEBAPP_URL_ADMIN || "");
  var configStudentUrl = canonicalExecBase_(CONFIG.WEBAPP_URL_STUDENT || CONFIG.WEBAPP_URL_STUDENT_EXEC || "");
  var warnings = [];

  if (rawServiceUrl && rawServiceUrl.indexOf('/a/') >= 0) warnings.push('ScriptApp service URL resolved as domain-scoped and was canonicalized for reporting.');

  var requestedSurface = clean_(surfaceHint || '');
  if (!requestedSurface) {
    if (requestedView === 'admin') requestedSurface = 'admin';
    else if (requestedView === 'portal') requestedSurface = 'student';
    else if (serviceUrl && adminBase && serviceUrl === adminBase) requestedSurface = 'admin';
    else if (serviceUrl && studentBase && serviceUrl === studentBase) requestedSurface = 'student';
    else requestedSurface = requestedView || 'unknown';
  }

  var runtime = {
    ok: true,
    endpoint: 'whoami',
    version: clean_(CONFIG.VERSION || ''),
    deployVersion: deployVersion,
    deploymentIdAdmin: clean_(CONFIG.DEPLOYMENT_ID_ADMIN || ''),
    deploymentIdStudent: clean_(CONFIG.DEPLOYMENT_ID_STUDENT || ''),
    serviceUrl: serviceUrl,
    canonicalAdminUrl: adminBase,
    canonicalStudentUrl: studentBase,
    activeUser: activeUser,
    effectiveUser: effectiveUser,
    requestedView: requestedView,
    requestedSurface: requestedSurface,
    scriptId: clean_(CONFIG.SCRIPT_ID || ScriptApp.getScriptId() || ''),
    timestamp: new Date().toISOString(),
    mismatch: false,
    warning: '',
    warnings: warnings,
    mismatches: []
  };

  if (runtime.deployVersion !== Number(CONFIG.DEPLOY_VERSION_NUMBER || 0)) {
    runtime.mismatch = true;
    runtime.mismatches.push('deployVersion mismatch');
  }
  if (runtime.deploymentIdAdmin !== clean_(CONFIG.DEPLOYMENT_ID_ADMIN || '')) {
    runtime.mismatch = true;
    runtime.mismatches.push('Admin deployment mismatch');
  }
  if (runtime.deploymentIdStudent !== clean_(CONFIG.DEPLOYMENT_ID_STUDENT || '')) {
    runtime.mismatch = true;
    runtime.mismatches.push('Student deployment mismatch');
  }
  if (runtime.canonicalAdminUrl !== configAdminUrl) {
    runtime.mismatch = true;
    runtime.mismatches.push('Admin URL mismatch');
  }
  if (runtime.canonicalStudentUrl !== configStudentUrl) {
    runtime.mismatch = true;
    runtime.mismatches.push('Student URL mismatch');
  }
  if (!runtime.deployVersion || !runtime.deploymentIdAdmin || !runtime.deploymentIdStudent) {
    runtime.mismatch = true;
    runtime.mismatches.push('Missing runtime fields');
  }

  runtime.warning = runtime.mismatches.concat(runtime.warnings).join(' | ');
  return runtime;
}

function admin_getRuntimeInfo() {
  return buildRuntimeTruth_({ parameter: { view: 'admin' } }, 'admin');
}

// SV GROK Mar 7 function overwrite
function admin_getStudentPortalLink(payload) {
  var debugId = newDebugId_();
  try {
    var caller = "";
    try { caller = clean_(Session.getEffectiveUser().getEmail() || ""); } catch (_callerErr) {}

    if (!isAdminCaller_()) {
      return {
        ok: false,
        code: "ACCESS_DENIED",
        debugId: debugId,
        error: "Access denied"
      };
    }

    var p = (payload && typeof payload === "object")
      ? payload
      : { applicantId: payload };

    var applicantId = clean_(p.applicantId || p.id || "");
    var rowNumber = Number(p.rowNumber || 0);

    Logger.log("admin_getStudentPortalLink START " + JSON.stringify({
      debugId: debugId,
      applicantId: applicantId,
      rowNumber: rowNumber,
      caller: caller
    }));

    var ss = getWorkingSpreadsheet_();
    var sheet = mustGetDataSheet_(ss);

    if (!rowNumber || rowNumber < 2) {
      if (!applicantId) {
        throw new Error("Portal link error. Debug: missing applicant id");
      }
      rowNumber = findRowByApplicantId_(sheet, applicantId);
    }

    if (!rowNumber || rowNumber < 2) {
      throw new Error("Portal link error. Debug: applicant not found");
    }

    var rowObj = getRowObject_(sheet, rowNumber);
    applicantId = clean_(rowObj.ApplicantID || applicantId || "");
    if (!applicantId) {
      throw new Error("Portal link error. Debug: missing applicant id");
    }

    var secretRes = getPortalSecretForApplicant_(applicantId);
    if (!secretRes || secretRes.ok !== true || !clean_(secretRes.secret || "")) {
      throw new Error("Portal link error. Debug: token missing");
    }

    var portalUrl = buildStudentPortalUrl_(applicantId, clean_(secretRes.secret || ""));

    Logger.log("admin_getStudentPortalLink OK " + JSON.stringify({
      debugId: debugId,
      applicantId: applicantId,
      rowNumber: rowNumber,
      url: portalUrl
    }));

    return {
      ok: true,
      url: portalUrl,
      applicantId: applicantId,
      rowNumber: rowNumber,
      debugId: debugId
    };

  } catch (e) {
    var rawMsg = String(e && e.message ? e.message : e);
    var lowerMsg = String(rawMsg || "").toLowerCase();
    var isPermissionDoc = lowerMsg.indexOf("permission to access the requested document") >= 0;
    var isPortalSecretsHint = lowerMsg.indexOf("portalsecrets") >= 0 || lowerMsg.indexOf(String(CONFIG.PORTAL_SECRETS_SHEET_ID || "").toLowerCase()) >= 0;
    var isSecretsPermission = isPermissionDoc || isPortalSecretsHint;
    var userMsg = isSecretsPermission
      ? "Portal link cannot be generated because this admin account does not have access to the PortalSecrets store. Share the PortalSecrets spreadsheet with this user and retry."
      : rawMsg;

    Logger.log("admin_getStudentPortalLink FAIL " + JSON.stringify({
      debugId: debugId,
      message: rawMsg,
      classifiedCode: isSecretsPermission ? "PORTAL_SECRETS_ACCESS_DENIED" : "PORTAL_LINK_ERROR",
      stack: String((e && e.stack) || "")
    }));

    return {
      ok: false,
      code: isSecretsPermission ? "PORTAL_SECRETS_ACCESS_DENIED" : "PORTAL_LINK_ERROR",
      debugId: debugId,
      error: userMsg
    };
  }
}
function findPortalRowByIdSecret_(sheet, applicantId, secret) {
  var rowNum = findRowByApplicantId_(sheet, applicantId);
  if (!rowNum) return null;
  var secretRes = getPortalSecretForApplicant_(applicantId);
  if (!secretRes || secretRes.ok !== true) return null;
  if (clean_(secretRes.secret || "") !== clean_(secret || "")) return null;
  var record = getRowObject_(sheet, rowNum);
  return { rowNum: rowNum, record: record };
}

function isPortalTokenExpired_(issuedAtValue, maxAgeDays) {
  var maxDays = Number(maxAgeDays || 0);
  if (!maxDays || maxDays <= 0) return false;
  if (!issuedAtValue) return true;
  var issuedAt = new Date(issuedAtValue);
  if (isNaN(issuedAt.getTime())) return true;
  var ageMs = new Date().getTime() - issuedAt.getTime();
  return ageMs > (maxDays * 24 * 60 * 60 * 1000);
}

function getPortalRequestMeta_(e) {
  var p = (e && e.parameter) ? e.parameter : {};
  return {
    ip: clean_(p.ip || p.clientIp || p.client_ip || p.x_forwarded_for || ""),
    ua: clean_(p.ua || p.userAgent || p.user_agent || "")
  };
}

function incrementInvalidPortalAttempt_(applicantId) {
  var cache = CacheService.getScriptCache();
  var key = "portal_invalid_" + clean_(applicantId || "unknown");
  var current = Number(cache.get(key) || 0);
  var next = current + 1;
  cache.put(key, String(next), 3600);
  return next;
}

function findRowByApplicantId_(sheet, applicantId) {
  var headerMap = getHeaderIndexMap_(sheet);
  var idCol = headerMap[SCHEMA.APPLICANT_ID];
  if (!idCol) throw new Error("Missing header: " + SCHEMA.APPLICANT_ID);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  var idNorm = clean_(applicantId);
  var idVals = sheet.getRange(2, idCol, lastRow - 1, 1).getValues();
  for (var r = 0; r < idVals.length; r++) {
    if (clean_(idVals[r][0]) === idNorm) return r + 2;
  }
  return null;
}

function findRowByIdEmail_(sheet, applicantId, parentEmailLower) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var idCol = headers.indexOf(CONFIG.APPLICANT_ID_HEADER);
  var emailCol = headers.indexOf("Parent_Email");
  if (idCol === -1 || emailCol === -1) return null;

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;

  var data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  for (var i = 0; i < data.length; i++) {
    var rowId = clean_(data[i][idCol]);
    var rowEmail = clean_(data[i][emailCol]).toLowerCase();
    if (rowId === applicantId && rowEmail === parentEmailLower) {
      var record = {};
      for (var c = 0; c < headers.length; c++) record[headers[c]] = data[i][c];
      return { rowNum: i + 2, record: record };
    }
  }
  return null;
}

/******************** EXAM SITES ********************/
function getExamSites_(ss) {
  var sh = ss.getSheetByName(CONFIG.EXAM_SITES_SHEET);
  if (!sh) return [];
  var values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  var headers = values[0].map(function(h){ return String(h || "").trim(); });
  var nameIdx = headers.indexOf("Site_Name");
  var activeIdx = headers.indexOf("Active");
  if (nameIdx === -1) return [];

  var out = [];
  for (var i = 1; i < values.length; i++) {
    var name = String(values[i][nameIdx] || "").trim();
    var active = (activeIdx >= 0) ? String(values[i][activeIdx] || "").trim().toLowerCase() : "true";
    var isActive = (active === "true" || active === "yes" || active === "1");
    if (name && isActive) out.push(name);
  }
  return out;
}

/******************** SUBJECT NORMALIZATION ********************/
function subjectsToCsv_(raw) {
  var s = clean_(raw);
  if (!s) return "";

  // JSON map like {"102":"Math","103":"Biology"}
  if (s.charAt(0) === "{" && (s.indexOf('":"') !== -1 || s.indexOf('": "') !== -1)) {
    try {
      var obj = JSON.parse(s);
      var vals = [];
      for (var k in obj) vals.push(clean_(obj[k]));
      return uniqCsv_(vals);
    } catch (e) {}
  }

  // "{102=English, 103=Math}" style
  if (s.indexOf("{") === 0 && s.indexOf("=") !== -1) {
    var inner = s.substring(1, s.length - 1);
    var parts = inner.split(",");
    var vals2 = [];
    for (var i = 0; i < parts.length; i++) {
      var p = parts[i].trim();
      if (p.indexOf("=") !== -1) vals2.push(clean_(p.split("=").slice(1).join("=")));
    }
    return uniqCsv_(vals2);
  }

  // CSV already
  return uniqCsv_(s.split(",").map(function(x){ return x.trim(); }));
}

// parse csv to Set(lowercased)
function parseSubjects_(csv) {
  var set = new Set();
  var s = clean_(csv);
  if (!s) return set;
  var parts = s.split(",");
  for (var i = 0; i < parts.length; i++) {
    var p = clean_(parts[i]).toLowerCase();
    if (p) set.add(p);
  }
  return set;
}

function uniqCsv_(arr) {
  var seen = {};
  var out = [];
  for (var i = 0; i < (arr || []).length; i++) {
    var v = clean_(arr[i]);
    if (!v) continue;
    var key = v.toLowerCase();
    if (!seen[key]) { seen[key] = true; out.push(v); }
  }
  return out.join(", ");
}

/******************** LOCK RULE ********************/
function isDocsVerified_(row) {
  var r = row || {};
  return r["Docs_Verified"] === "Yes";
}

function getPortalLockReason_(record) {
  var row = record || {};
  if (isPaymentVerifiedDerived_(row)) return "payment_verified";
  if (String(row[SCHEMA.PORTAL_ACCESS_STATUS] || "").trim() === "Locked") return "portal_access_locked";
  if (row._PortalHardLocked === true) return "hard_locked";
  return "";
}

function isPortalLocked_(record) {
  return !!getPortalLockReason_(record);
}

function isPaymentVerified_(record) {
  return isPaymentVerifiedDerived_(record) === true;
}

function isPaymentVerifiedDerived_(row) {
  row = row || {};
  if (derivePaymentBadge_(row) === "Verified") return true;
  return clean_(row.Payment_Verified).toLowerCase() === "yes";
}

function isPaymentFreezeActive_(row) {
  return isPaymentVerifiedDerived_(row) === true;
}

function resolveDocStatusKeys_(row) {
  row = row || {};
  function pick_(keys) {
    for (var i = 0; i < keys.length; i++) {
      var k = keys[i];
      if (hasOwn_(row, k)) return k;
    }
    return keys[0];
  }
  return {
    birth: pick_(["Birth_ID_Status", "Birth_Status"]),
    report: pick_(["Report_Status"]),
    photo: pick_(["Photo_Status"]),
    transfer: pick_(["Transfer_Status"]),
    receipt: pick_(["Receipt_Status"])
  };
}

function derivePaymentVerified_(row) {
  row = row || {};
  var paymentVerified = isPaymentVerifiedDerived_(row);
  // Keep legacy compatibility column aligned when present in row object.
  if (hasOwn_(row, "Payment_Verified")) row.Payment_Verified = paymentVerified ? "Yes" : "";
  return paymentVerified;
}

function normalizeOverallDocValue_(v) {
  var s = clean_(v).toLowerCase();
  if (s === "verified" || s === "yes" || s === "true" || s === "1") return "Verified";
  if (s === "rejected" || s === "reject") return "Rejected";
  if (s === "fraudulent" || s === "fraud") return "Rejected";
  return "Pending";
}

function computeDocVerificationStatus_(row) {
  row = row || {};
  var keys = resolveDocStatusKeys_(row);
  var requiredDocs = [keys.birth, keys.report, keys.photo];
  var hasRejected = false;
  var allVerified = true;
  for (var i = 0; i < requiredDocs.length; i++) {
    var st = normalizeOverallDocValue_(row[requiredDocs[i]]);
    if (st === "Rejected") hasRejected = true;
    if (st !== "Verified") allVerified = false;
  }
  if (hasRejected) return "Rejected";
  if (allVerified) return "Verified";
  return "Pending";
}

function derivePaymentBadge_(row) {
  row = row || {};
  var keys = resolveDocStatusKeys_(row);
  var st = normalizeOverallDocValue_(row[keys.receipt]);
  if (st === "Verified") return "Verified";
  if (st === "Rejected") return "Rejected";
  return "Pending";
}

function computeOverallStatus_(row) {
  row = row || {};
  var docStage = computeDocVerificationStatus_(row);
  var paymentBadge = derivePaymentBadge_(row);
  var paymentVerified = isPaymentVerifiedDerived_(row);
  // keep compatibility alignment
  if (hasOwn_(row, "Payment_Verified")) row.Payment_Verified = paymentVerified ? "Yes" : "";
  // Payment verification is the final milestone and must not be downgraded by doc edits.
  if (paymentVerified) return "Verified";
  if (docStage === "Verified" && paymentBadge !== "Verified") return "Docs_Verified";
  return "Pending";
}

function canOverrideOverall_(email) {
  var e = clean_(email).toLowerCase();
  if (!e) return false;
  var superList = (CONFIG.SUPER_ADMIN_EMAILS || []).map(function (x) { return clean_(x).toLowerCase(); });
  var elevatedList = (CONFIG.ELEVATED_OVERRIDE_EMAILS || []).map(function (x) { return clean_(x).toLowerCase(); });
  return superList.indexOf(e) >= 0 || elevatedList.indexOf(e) >= 0;
}

function canBypassPaymentFreeze_(email) {
  var e = clean_(email).toLowerCase();
  if (!e) return false;
  var superList = (CONFIG.SUPER_ADMIN_EMAILS || []).map(function (x) { return clean_(x).toLowerCase(); });
  return superList.indexOf(e) >= 0;
}

function computeFodeFeeQuote_(rowObj) {
  var row = rowObj || {};
  var registrationK = Number(CONFIG.FEE_REGISTRATION_KINA || 600);
  var perSubjectK = Number(CONFIG.FEE_PER_SUBJECT_KINA || 450);
  var csv = clean_(row.Subjects_Selected_Canonical || row[SCHEMA.SUBJECTS_CANONICAL] || "");
  if (!csv) csv = subjectsToCsv_(row.Subjects_Selected || "");
  var parts = csv ? csv.split(",") : [];
  var subjects = [];
  for (var i = 0; i < parts.length; i++) {
    var s = clean_(parts[i]);
    if (s) subjects.push(s);
  }
  var subjectCount = subjects.length;
  var subjectFeeK = perSubjectK * subjectCount;
  return {
    registrationK: registrationK,
    subjectCount: subjectCount,
    subjectFeeK: subjectFeeK,
    totalK: registrationK + subjectFeeK,
    subjectsList: subjects.join(", ")
  };
}

function buildAdminApplicantDeepLink_(applicantId) {
  var base = clean_(CONFIG.WEBAPP_URL_ADMIN || CONFIG.WEBAPP_URL || "");
  var id = clean_(applicantId || "");
  if (!base || !id) return "";
  return base + "?view=admin&open=" + encodeURIComponent(id);
}

function formatKina_(n) {
  var num = Number(n || 0);
  if (!isFinite(num)) num = 0;
  return "K" + String(Math.round(num));
}

function sendEmailBestEffort_(toEmail, subject, body, logLabel, meta) {
  var dbgId = newDebugId_();
  var to = clean_(toEmail || "");
  var lbl = clean_(logLabel || "EMAIL_SEND");
  var payload = meta && typeof meta === "object" ? meta : {};
  function safeEmailLogPayload_(obj) {
    try {
      log_(mustGetSheet_(getWorkingSpreadsheet_(), CONFIG.LOG_SHEET), lbl, JSON.stringify(obj || {}));
    } catch (_logErr) {
      try { Logger.log("%s %s", lbl, JSON.stringify(obj || {})); } catch (_e) {}
    }
  }
  try {
    if (!to) throw new Error("Missing recipient email");
    MailApp.sendEmail({
      to: to,
      subject: String(subject || ""),
      body: String(body || ""),
      name: clean_(CONFIG.EMAIL_FROM_NAME || "FODE")
    });
    safeEmailLogPayload_({
      ok: true,
      debugId: dbgId,
      to: to,
      subject: String(subject || ""),
      meta: payload
    });
    return { ok: true, debugId: dbgId, to: to };
  } catch (e) {
    safeEmailLogPayload_({
      ok: false,
      debugId: dbgId,
      to: to,
      subject: String(subject || ""),
      error: String(e && e.message ? e.message : e),
      meta: payload
    });
    return { ok: false, debugId: dbgId, to: to, error: String(e && e.message ? e.message : e) };
  }
}

function sendDocsVerifiedPaymentRequiredEmail_(rowObj, rowNumber, actorEmail) {
  var row = rowObj || {};
  var applicantId = clean_(row.ApplicantID || row[SCHEMA.APPLICANT_ID] || "");
  var recipient = clean_(row.Parent_Email_Corrected || row[SCHEMA.PARENT_EMAIL_CORRECTED] || row.Parent_Email || row[SCHEMA.PARENT_EMAIL] || "");
  var quote = computeFodeFeeQuote_(row);
  var sh = mustGetDataSheet_(getWorkingSpreadsheet_());
  var rowNum = Number(rowNumber || 0) || findRowByApplicantId_(sh, applicantId);
  var portalUrl = "";
  if (rowNum >= 2 && applicantId) {
    var emailForSecret = clean_(row.Parent_Email_Corrected || row.Parent_Email || "");
    var fullName = (clean_(row.First_Name || "") + " " + clean_(row.Last_Name || "")).trim();
    var secretInfo = getOrCreateActivePortalSecret_(applicantId, emailForSecret, fullName, sh, rowNum, {});
    portalUrl = buildPortalLinkFromBase_(clean_(getStudentBaseUrl_() || CONFIG.WEBAPP_URL_STUDENT || ""), applicantId, secretInfo.secretPlain);
  }
  var subject = "FODE Documents Verified - Payment Required - " + applicantId;
  var body = [
    "Dear Parent/Guardian,",
    "",
    "Your FODE application documents have been verified. Payment is now required to proceed.",
    "",
    "ApplicantID: " + applicantId,
    "",
    "Fee breakdown:",
    "- Registration Fee: " + formatKina_(quote.registrationK),
    "- Subjects Selected: " + quote.subjectCount + (quote.subjectsList ? (" (" + quote.subjectsList + ")") : ""),
    "- Subject Fee: " + formatKina_(quote.subjectFeeK) + " (" + formatKina_(CONFIG.FEE_PER_SUBJECT_KINA || 450) + " x " + quote.subjectCount + ")",
    "- Total Payable: " + formatKina_(quote.totalK),
    "",
    String(CONFIG.PAYMENT_INSTRUCTIONS_TEXT || "").trim(),
    "",
    portalUrl ? ("Upload payment receipt here: " + portalUrl) : "Portal link unavailable. Please contact admissions.",
    "",
    "ApplicantID: " + applicantId
  ].join("\n");
  var sendRes = sendEmailBestEffort_(recipient, subject, body, "DOCS_VERIFIED_EMAIL_SENT", {
    applicantId: applicantId,
    rowNumber: rowNum,
    by: clean_(actorEmail || ""),
    recipient: recipient,
    feeQuote: quote
  });
  sendRes.portalUrl = portalUrl;
  sendRes.feeQuote = quote;
  return sendRes;
}

function notifyAdminPaymentReceiptUploaded_(rowObj, rowNumber, opts) {
  var row = rowObj || {};
  opts = opts || {};
  var applicantId = clean_(row.ApplicantID || row[SCHEMA.APPLICANT_ID] || "");
  var fullName = (clean_(row.First_Name || "") + " " + clean_(row.Last_Name || "")).trim();
  var toEmail = clean_(CONFIG.EMAIL_ADMIN_ALERTS_TO || "");
  var subject = "PAYMENT RECEIPT UPLOADED - " + applicantId + " - " + (fullName || "Unknown");
  var adminUrl = buildAdminApplicantDeepLink_(applicantId);
  var body = [
    "Payment receipt uploaded.",
    "",
    "ApplicantID: " + applicantId,
    "Name: " + (fullName || "-"),
    "Timestamp: " + (new Date()).toISOString(),
    "RowNumber: " + String(Number(rowNumber || 0) || ""),
    "",
    "Admin review URL:",
    adminUrl || "(admin URL not configured)"
  ].join("\n");
  try {
    log_(mustGetSheet_(getWorkingSpreadsheet_(), CONFIG.LOG_SHEET), "PAYMENT_RECEIPT_UPLOADED", JSON.stringify({
      applicantId: applicantId,
      rowNumber: Number(rowNumber || 0) || "",
      source: clean_(opts.source || ""),
      oldValue: clean_(opts.oldValue || ""),
      newValue: clean_(opts.newValue || "")
    }));
  } catch (_alertLogErr) {}
  return sendEmailBestEffort_(toEmail, subject, body, "PAYMENT_RECEIPT_ALERT_EMAIL", {
    applicantId: applicantId,
    rowNumber: Number(rowNumber || 0) || "",
    source: clean_(opts.source || ""),
    adminUrl: adminUrl
  });
}

function maybeNotifyPaymentReceiptUploadTransition_(beforeRow, afterRow, rowNumber, opts) {
  var prevUrl = clean_((beforeRow || {}).Fee_Receipt_File || "");
  var nextUrl = clean_((afterRow || {}).Fee_Receipt_File || "");
  if (!nextUrl) return { notified: false, reason: "empty" };
  if (prevUrl === nextUrl) return { notified: false, reason: "unchanged" };
  var alertRes = notifyAdminPaymentReceiptUploaded_(afterRow || beforeRow || {}, rowNumber, Object.assign({
    oldValue: prevUrl,
    newValue: nextUrl
  }, opts || {}));
  return { notified: true, alert: alertRes };
}

function buildCrmPayloadFromRow_(rowObj) {
  var row = rowObj || {};
  var applicantId = clean_(row.ApplicantID || row[SCHEMA.APPLICANT_ID] || "");
  var firstName = clean_(row.First_Name || "");
  var lastName = clean_(row.Last_Name || "");
  var emailCorrected = clean_(row.Parent_Email_Corrected || row[SCHEMA.PARENT_EMAIL_CORRECTED] || "");
  var emailRaw = clean_(row.Parent_Email || row[SCHEMA.PARENT_EMAIL] || "");
  var effectiveEmail = emailCorrected || emailRaw;
  var intakeYear = clean_(row.Intake_Year || "");
  if (!intakeYear) intakeYear = String((new Date()).getFullYear() + 1);
  return {
    applicantId: applicantId,
    firstName: firstName,
    lastName: lastName,
    fullName: (firstName + " " + lastName).trim(),
    parentEmail: emailRaw,
    parentEmailCorrected: emailCorrected,
    effectiveEmail: effectiveEmail,
    parentPhone: clean_(row.Parent_Phone || ""),
    gradeApplyingFor: clean_(row.Grade_Applying_For || ""),
    intakeYear: intakeYear,
    crmPipeline: clean_(CONFIG.CRM_PIPELINE_FODE || ""),
    crmStage: clean_(CONFIG.CRM_STAGE_PAYMENT_VERIFIED || CONFIG.DEAL_STAGE || ""),
    subjects: clean_(row.Subjects_Selected_Canonical || subjectsToCsv_(row.Subjects_Selected || "")),
    folderUrl: clean_(row.Folder_Url || row[SCHEMA.FOLDER_URL] || ""),
    formId: clean_(row.FormID || row.FD_FormID || "")
  };
}

function ensureStableFormId_(rowObj, sh, rowNumber, idx) {
  var row = rowObj || {};
  var stable = clean_(row.FormID || row.FD_FormID || "");
  if (stable) return stable;
  var applicantId = clean_(row.ApplicantID || row[SCHEMA.APPLICANT_ID] || "");
  stable = applicantId ? ("FODE_" + applicantId) : ("FODE_" + Utilities.formatDate(new Date(), "UTC", "yyyyMMddHHmmss"));
  var formCol = idx && idx.FormID ? "FormID" : (idx && idx.FD_FormID ? "FD_FormID" : "");
  if (formCol) {
    sh.getRange(rowNumber, idx[formCol]).setValue(stable);
    row[formCol] = stable;
  }
  return stable;
}

function readRowSnapshot_(applicantId) {
  var id = clean_(applicantId || "");
  if (!id) return null;
  var ss = getWorkingSpreadsheet_();
  var sheet = mustGetDataSheet_(ss);
  var rowNum = findRowByApplicantId_(sheet, id);
  if (!rowNum) return null;
  var rowObj = getRowObject_(sheet, rowNum) || {};
  var out = {
    ApplicantID: clean_(rowObj.ApplicantID || ""),
    Birth_Status: clean_(rowObj.Birth_Status || ""),
    Report_Status: clean_(rowObj.Report_Status || ""),
    Photo_Status: clean_(rowObj.Photo_Status || ""),
    Transfer_Status: clean_(rowObj.Transfer_Status || ""),
    Receipt_Status: clean_(rowObj.Receipt_Status || "")
  };
  if (Object.prototype.hasOwnProperty.call(rowObj, "Doc_Verification_Status")) {
    out.Doc_Verification_Status = clean_(rowObj.Doc_Verification_Status || "");
  }
  if (Object.prototype.hasOwnProperty.call(rowObj, "Overall_Status")) {
    out.Overall_Status = clean_(rowObj.Overall_Status || "");
  }
  if (Object.prototype.hasOwnProperty.call(rowObj, "Payment_Verified")) {
    out.Payment_Verified = clean_(rowObj.Payment_Verified || "");
  }
  return out;
}

/******************** SHEET HELPERS ********************/
function ensureHeaders_(sheet, payload) {
  var meta = [
  CONFIG.APPLICANT_ID_HEADER,
  "Folder_Url",
  "PortalLastUpdateAt",
  "Portal_Submitted",
  SCHEMA.PORTAL_TOKEN_HASH,
  SCHEMA.PORTAL_TOKEN_ISSUED_AT,
  SCHEMA.PORTAL_ACCESS_STATUS,
  "Physical_Exam_Site",
  "Subjects_Selected_Canonical",
  CONFIG.PARENT_EMAIL_CORRECTED_HEADER,
  "File_Log",

  // ✅ NEW — verification tracking
  "Doc_Last_Verified_At",
  "Doc_Last_Verified_By"
];


  for (var i = 0; i < CONFIG.DOC_FIELDS.length; i++) {
    meta.push(CONFIG.DOC_FIELDS[i].status);
    meta.push(CONFIG.DOC_FIELDS[i].comment);
  }

  var headersWanted = Object.keys(payload).concat(meta);

  if (sheet.getLastRow() === 0) { sheet.appendRow(headersWanted); return; }

  var existing = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var changed = false;

  for (var j = 0; j < headersWanted.length; j++) {
    if (existing.indexOf(headersWanted[j]) === -1) { existing.push(headersWanted[j]); changed = true; }
  }

  if (changed) sheet.getRange(1, 1, 1, existing.length).setValues([existing]);
}

function logActivation_(logSheet, label, payload) {
  try {
    log_(logSheet, label, JSON.stringify(payload || {}));
  } catch (logErr) {
    try {
      Logger.log(label + " " + JSON.stringify(payload || {}));
    } catch (_loggerErr) {}
  }
}

function scanApplicantIdState_(sheet) {
  var idCol = findCol_(sheet, CONFIG.APPLICANT_ID_HEADER);
  if (!idCol) throw new Error(CONFIG.APPLICANT_ID_HEADER + " column not found.");
  var lastRow = sheet.getLastRow();
  var rowCount = Math.max(lastRow - 1, 0);
  var values = rowCount > 0 ? sheet.getRange(2, idCol, rowCount, 1).getValues() : [];
  var prefix = CONFIG.APPLICANT_PREFIX;
  var digits = CONFIG.APPLICANT_DIGITS;
  var re = new RegExp("^" + escapeRegExp_(prefix) + "(\\d{" + digits + "})$");
  var maxSuffix = 0;
  var validCount = 0;
  var skippedBlankCount = 0;
  var skippedMalformedCount = 0;
  for (var i = 0; i < values.length; i++) {
    var s = String(values[i][0] || "").trim();
    if (!s) {
      skippedBlankCount++;
      continue;
    }
    var m = s.match(re);
    if (!m) {
      skippedMalformedCount++;
      continue;
    }
    var n = parseInt(m[1], 10);
    if (isNaN(n)) {
      skippedMalformedCount++;
      continue;
    }
    validCount++;
    if (n > maxSuffix) maxSuffix = n;
  }
  return {
    applicantId: prefix + String(maxSuffix + 1).padStart(digits, "0"),
    validCount: validCount,
    maxSuffix: maxSuffix,
    skippedBlankCount: skippedBlankCount,
    skippedMalformedCount: skippedMalformedCount
  };
}

function nextApplicantId_(sheet) {
  return scanApplicantIdState_(sheet).applicantId;
}

function preparePortalActivationState_(sheet, applicantId) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var hasTokenHashHeader = headers.indexOf("PortalTokenHash") >= 0;
  var hasTokenIssuedAtHeader = headers.indexOf("PortalTokenIssuedAt") >= 0;
  var plainSecret = newPortalSecret_();
  var tokenHash = hashPortalSecret_(plainSecret);
  var tokenIssuedAt = new Date();
  return {
    hasTokenHashHeader: hasTokenHashHeader,
    hasTokenIssuedAtHeader: hasTokenIssuedAtHeader,
    tokenHash: hasTokenHashHeader ? tokenHash : "",
    tokenIssuedAt: hasTokenIssuedAtHeader ? tokenIssuedAt : "",
    plainSecret: plainSecret,
    secretHash: tokenHash,
    applicantId: clean_(applicantId || ""),
    portalSecretsRequired: true
  };
}

function ensurePortalActivationStoreHeaders_(sheet) {
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
  var current = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), expected.length)).getValues()[0].map(function(v) {
    return clean_(v);
  });
  var changed = false;
  for (var i = 0; i < expected.length; i++) {
    if (current.indexOf(expected[i]) === -1) {
      current.push(expected[i]);
      changed = true;
    }
  }
  if (changed) sheet.getRange(1, 1, 1, current.length).setValues([current]);
}

function commitPortalActivationState_(payload, applicantId, tokenState) {
  if (!tokenState || tokenState.portalSecretsRequired !== true) return { ok: true, skipped: true };
  var ss = SpreadsheetApp.openById(PORTAL_SECRETS_SPREADSHEET_ID);
  var sh = ss.getSheetByName(PORTAL_SECRETS_TAB);
  if (!sh) sh = ss.insertSheet(PORTAL_SECRETS_TAB);
  ensurePortalActivationStoreHeaders_(sh);
  var idx = {};
  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  for (var i = 0; i < headers.length; i++) {
    var h = clean_(headers[i]);
    if (h) idx[h] = i + 1;
  }
  var lastRow = sh.getLastRow();
  var rowIndex = 0;
  if (idx.ApplicantID && lastRow >= 2) {
    var ids = sh.getRange(2, idx.ApplicantID, lastRow - 1, 1).getValues();
    for (var r = 0; r < ids.length; r++) {
      if (clean_(ids[r][0]) === clean_(applicantId || "")) {
        rowIndex = r + 2;
        break;
      }
    }
  }
  var nowIso = new Date().toISOString();
  var email = clean_(payload.Parent_Email_Corrected || payload.Parent_Email || payload.email || "");
  var fullName = (clean_(payload.First_Name || "") + " " + clean_(payload.Last_Name || "")).trim();
  var patch = {
    ApplicantID: clean_(applicantId || ""),
    Email: email,
    Full_Name: fullName,
    Secret_Plain: clean_(tokenState.plainSecret || ""),
    Secret_Hash: clean_(tokenState.secretHash || ""),
    Last_Rotated_At: nowIso,
    Status: "Active"
  };
  if (rowIndex) {
    var createdAt = idx.Created_At ? sh.getRange(rowIndex, idx.Created_At).getValue() : "";
    if (!clean_(createdAt)) patch.Created_At = nowIso;
    for (var key in patch) {
      if (!Object.prototype.hasOwnProperty.call(patch, key)) continue;
      if (!idx[key]) continue;
      sh.getRange(rowIndex, idx[key]).setValue(patch[key]);
    }
    return { ok: true, rowIndex: rowIndex, created: false };
  }
  sh.appendRow([
    clean_(applicantId || ""),
    email,
    fullName,
    clean_(tokenState.plainSecret || ""),
    clean_(tokenState.secretHash || ""),
    nowIso,
    nowIso,
    "Active"
  ]);
  return { ok: true, rowIndex: sh.getLastRow(), created: true };
}

function fileExtensionFromName_(name) {
  var raw = clean_(name || "");
  var idx = raw.lastIndexOf(".");
  if (idx < 0 || idx === raw.length - 1) return "";
  return raw.slice(idx + 1).replace(/[^a-zA-Z0-9]+/g, "").toLowerCase();
}

function fileExtensionFromUrl_(url) {
  var raw = clean_(url || "");
  var match = raw.match(/\.([a-zA-Z0-9]{1,10})(?:[?#].*)?$/);
  return match ? clean_(match[1] || "").toLowerCase() : "";
}

function fileExtensionFromContentType_(contentType) {
  var type = clean_(contentType || "").toLowerCase();
  if (!type) return "";
  var map = {
    "application/pdf": "pdf",
    "image/jpeg": "jpg",
    "image/jpg": "jpg",
    "image/png": "png",
    "image/gif": "gif",
    "image/webp": "webp",
    "image/heic": "heic",
    "image/heif": "heif",
    "application/msword": "doc",
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document": "docx",
    "application/vnd.ms-excel": "xls",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": "xlsx",
    "text/plain": "txt"
  };
  return map[type] || "";
}

function canonicalizeFdIntakeFiles_(payload, applicantFolder, logSheet, context) {
  var sourcePayload = payload || {};
  var out = {};
  for (var key in sourcePayload) {
    if (!Object.prototype.hasOwnProperty.call(sourcePayload, key)) continue;
    out[key] = sourcePayload[key];
  }
  if (!applicantFolder) return out;

  var ctx = (context && typeof context === "object") ? context : {};
  var correlationId = clean_(ctx.correlationId || "");
  var applicantId = clean_(ctx.applicantId || "");
  var folderId = clean_(applicantFolder.getId ? applicantFolder.getId() : "");
  var folderUrl = clean_(applicantFolder.getUrl ? applicantFolder.getUrl() : "");
  var fileLog = clean_(out.File_Log || "");
  var fields = (CONFIG.DOC_FIELDS || []).map(function (doc) {
    return clean_(doc && doc.file || "");
  }).filter(function (field) {
    return !!field;
  });

  for (var i = 0; i < fields.length; i++) {
    var field = fields[i];
    var rawUrls = normalizeToUrlList_(out[field], field);
    if (!rawUrls.length) continue;
    var canonicalUrls = [];
    for (var u = 0; u < rawUrls.length; u++) {
      var rawUrl = clean_(rawUrls[u]);
      if (!rawUrl) continue;
      try {
        var response = UrlFetchApp.fetch(rawUrl, { muteHttpExceptions: true });
        var responseCode = Number(response && response.getResponseCode ? response.getResponseCode() : 0);
        if (responseCode != 200) {
          logActivation_(logSheet, "ACTIVATION_FILE_CANONICALIZE_SKIP", {
            correlation_id: correlationId,
            applicantId: applicantId,
            field: field,
            reason: "fetch_failed",
            rawUrl: rawUrl,
            responseCode: responseCode
          });
          continue;
        }
        var blob = response.getBlob();
        if (!blob) {
          logActivation_(logSheet, "ACTIVATION_FILE_CANONICALIZE_SKIP", {
            correlation_id: correlationId,
            applicantId: applicantId,
            field: field,
            reason: "blob_missing",
            rawUrl: rawUrl
          });
          continue;
        }
        var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "Pacific/Port_Moresby", "yyyyMMdd_HHmmss_SSS");
        var ext = fileExtensionFromContentType_(blob.getContentType()) || fileExtensionFromUrl_(rawUrl) || fileExtensionFromName_(blob.getName()) || "bin";
        var newName = field + "_" + timestamp + (ext ? ("." + ext) : "");
        blob.setName(newName);
        var newFile = applicantFolder.createFile(blob);
        newFile.setName(newName);
        var newUrl = clean_(newFile.getUrl() || "");
        if (!newUrl) {
          logActivation_(logSheet, "ACTIVATION_FILE_CANONICALIZE_SKIP", {
            correlation_id: correlationId,
            applicantId: applicantId,
            field: field,
            reason: "canonical_url_missing",
            rawUrl: rawUrl,
            newFileId: clean_(newFile.getId() || "")
          });
          continue;
        }
        canonicalUrls.push(newUrl);
        fileLog = appendLog_(fileLog, new Date().toISOString()
          + " | " + field
          + " | fetched_and_copied"
          + " | rawUrl=" + rawUrl
          + " | newFileId=" + clean_(newFile.getId() || "")
          + " | folder=" + folderId);
        logActivation_(logSheet, "ACTIVATION_FILE_CANONICALIZED", {
          correlation_id: correlationId,
          applicantId: applicantId,
          field: field,
          rawUrl: rawUrl,
          newFileId: clean_(newFile.getId() || ""),
          folderId: folderId,
          folderUrl: folderUrl
        });
      } catch (fileErr) {
        logActivation_(logSheet, "ACTIVATION_FILE_CANONICALIZE_SKIP", {
          correlation_id: correlationId,
          applicantId: applicantId,
          field: field,
          reason: "fetch_or_create_failed",
          rawUrl: rawUrl,
          error: String(fileErr && fileErr.message ? fileErr.message : fileErr)
        });
      }
    }
    out[field] = canonicalUrls.join("\n");
  }
  out.File_Log = fileLog;
  return out;
}

function maybeStampActivationSubmitState_(payload, logSheet, context) {
  var out = payload || {};
  var ctx = (context && typeof context === "object") ? context : {};
  var applicantId = clean_(ctx.applicantId || "");
  var qualifyingFieldsDetected = (CONFIG.DOC_FIELDS || []).filter(function (doc) {
    return clean_(doc && doc.file || "") !== "Fee_Receipt_File";
  }).map(function (doc) {
    return clean_(doc && doc.file || "");
  }).filter(function (field) {
    return !!field && normalizeToUrlList_(out[field], field).length > 0;
  });
  var shouldStampSubmitState = qualifyingFieldsDetected.length > 0;
  logActivation_(logSheet, "ACTIVATION_SUBMIT_STATE_DECISION", {
    applicantId: applicantId,
    shouldStampSubmitState: shouldStampSubmitState,
    qualifyingFieldsDetected: qualifyingFieldsDetected
  });
  if (!shouldStampSubmitState) return out;

  var nowIso = new Date().toISOString();
  if (!clean_(out.PortalLastUpdateAt || "")) out.PortalLastUpdateAt = nowIso;
  if (!clean_(out.Portal_Submitted || "")) out.Portal_Submitted = nowIso;
  logActivation_(logSheet, "ACTIVATION_SUBMIT_STATE_STAMPED", {
    applicantId: applicantId,
    PortalLastUpdateAt: clean_(out.PortalLastUpdateAt || ""),
    Portal_Submitted: clean_(out.Portal_Submitted || "")
  });
  return out;
}

function buildActivatedIntakeRow_(sheet, payload, folderUrl, applicantId, tokenState) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var row = [];
  for (var i = 0; i < headers.length; i++) {
    var h = headers[i];
    if (h === CONFIG.APPLICANT_ID_HEADER) row.push(clean_(applicantId || ""));
    else if (h === "Folder_Url") row.push(clean_(folderUrl || ""));
    else if (h === "PortalTokenHash" && tokenState && tokenState.hasTokenHashHeader) row.push(clean_(tokenState.tokenHash || ""));
    else if (h === "PortalTokenIssuedAt" && tokenState && tokenState.hasTokenIssuedAtHeader) row.push(tokenState.tokenIssuedAt || "");
    else row.push(normalize_(payload[h]));
  }
  return row;
}

function insertActivatedRowAt_(sheet, targetRow, rowArray) {
  var rowNum = Number(targetRow || 0);
  if (!rowNum || rowNum < 2) throw new Error("Invalid targetRow for activation commit");
  sheet.getRange(rowNum, 1, 1, rowArray.length).setValues([rowArray]);
  return rowNum;
}

function verifyActivatedState_(sheet, rowNum, applicantId, folderUrl, tokenState) {
  SpreadsheetApp.flush();
  var rowObj = getRowObject_(sheet, rowNum) || {};
  var applicantIdActual = clean_(rowObj.ApplicantID || "");
  var folderUrlActual = clean_(rowObj.Folder_Url || "");
  var portalTokenHashPresent = !!clean_(rowObj.PortalTokenHash || "");
  var portalTokenIssuedAtRaw = rowObj.PortalTokenIssuedAt;
  var portalTokenIssuedAtPresent = false;
  if (portalTokenIssuedAtRaw instanceof Date) {
    portalTokenIssuedAtPresent = !isNaN(portalTokenIssuedAtRaw.getTime());
  } else {
    portalTokenIssuedAtPresent = !!clean_(portalTokenIssuedAtRaw || "");
  }
  var secretRes = getPortalSecretForApplicant_(applicantId);
  var portalSecretsResolvable = !!(secretRes && secretRes.ok === true && clean_(secretRes.secret || ""));
  if (!applicantIdActual || applicantIdActual !== clean_(applicantId || "")) {
    return {
      ok: false,
      code: "FINALIZE_MISSING_APPLICANTID",
      message: "ApplicantID verification failed.",
      applicantIdActual: applicantIdActual,
      folderUrlPresent: !!folderUrlActual,
      portalTokenHashPresent: portalTokenHashPresent,
      portalTokenIssuedAtPresent: portalTokenIssuedAtPresent,
      portalSecretsResolvable: portalSecretsResolvable
    };
  }
  if (!folderUrlActual || folderUrlActual !== clean_(folderUrl || "")) {
    return {
      ok: false,
      code: "FINALIZE_MISSING_FOLDER_URL",
      message: "Folder_Url verification failed.",
      applicantIdActual: applicantIdActual,
      folderUrlPresent: !!folderUrlActual,
      portalTokenHashPresent: portalTokenHashPresent,
      portalTokenIssuedAtPresent: portalTokenIssuedAtPresent,
      portalSecretsResolvable: portalSecretsResolvable
    };
  }
  if (tokenState && tokenState.hasTokenHashHeader && !portalTokenHashPresent) {
    return {
      ok: false,
      code: "FINALIZE_MISSING_PORTAL_TOKEN",
      message: "PortalTokenHash verification failed.",
      applicantIdActual: applicantIdActual,
      folderUrlPresent: !!folderUrlActual,
      portalTokenHashPresent: portalTokenHashPresent,
      portalTokenIssuedAtPresent: portalTokenIssuedAtPresent,
      portalSecretsResolvable: portalSecretsResolvable
    };
  }
  if (tokenState && tokenState.hasTokenIssuedAtHeader && !portalTokenIssuedAtPresent) {
    return {
      ok: false,
      code: "FINALIZE_MISSING_PORTAL_TOKEN",
      message: "PortalTokenIssuedAt verification failed.",
      applicantIdActual: applicantIdActual,
      folderUrlPresent: !!folderUrlActual,
      portalTokenHashPresent: portalTokenHashPresent,
      portalTokenIssuedAtPresent: portalTokenIssuedAtPresent,
      portalSecretsResolvable: portalSecretsResolvable
    };
  }
  if (!portalSecretsResolvable) {
    return {
      ok: false,
      code: "FINALIZE_MISSING_PORTALSECRET",
      message: "PortalSecrets resolvability verification failed.",
      applicantIdActual: applicantIdActual,
      folderUrlPresent: !!folderUrlActual,
      portalTokenHashPresent: portalTokenHashPresent,
      portalTokenIssuedAtPresent: portalTokenIssuedAtPresent,
      portalSecretsResolvable: portalSecretsResolvable
    };
  }
  return {
    ok: true,
    applicantIdActual: applicantIdActual,
    folderUrlPresent: !!folderUrlActual,
    portalTokenHashPresent: portalTokenHashPresent,
    portalTokenIssuedAtPresent: portalTokenIssuedAtPresent,
    portalSecretsResolvable: portalSecretsResolvable
  };
}

function appendRow_(sheet, payload, folder) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var row = [];
  for (var i = 0; i < headers.length; i++) {
    var h = headers[i];
    if (h === CONFIG.APPLICANT_ID_HEADER) row.push("");
    else if (h === "Folder_Url") row.push(folder.getUrl());
    else row.push(normalize_(payload[h]));
  }
  sheet.appendRow(row);
  return sheet.getLastRow();
}

function ensurePortalTokenAtRow_(sheet, rowNum) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var hashCol = headers.indexOf("PortalTokenHash");
  var issuedCol = headers.indexOf("PortalTokenIssuedAt");
  if (hashCol < 0 || issuedCol < 0) return;

  var currentHash = clean_(sheet.getRange(rowNum, hashCol + 1).getValue());
  var currentIssued = sheet.getRange(rowNum, issuedCol + 1).getValue();
  if (currentHash && currentIssued) return;

  var secret = newPortalSecret_();
  sheet.getRange(rowNum, hashCol + 1).setValue(hashPortalSecret_(secret));
  sheet.getRange(rowNum, issuedCol + 1).setValue(new Date());
}

function setPortalTokenHashForRow_(sheet, rowNumber, tokenHash) {
  var targetRow = Number(rowNumber || 0);
  if (!targetRow || targetRow < 2) throw new Error("Invalid rowNumber for hash write");
  var hash = clean_(tokenHash);
  if (!hash) throw new Error("Missing token hash");

  var idx = getHeaderIndexMap_(sheet);
  if (!idx[SCHEMA.PORTAL_TOKEN_HASH]) throw new Error("Missing header: " + SCHEMA.PORTAL_TOKEN_HASH);
  sheet.getRange(targetRow, idx[SCHEMA.PORTAL_TOKEN_HASH]).setValue(hash);
}

function writeBack_(sheet, row, kv) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  for (var k in kv) {
    var idx = headers.indexOf(k);
    if (idx >= 0) sheet.getRange(row, idx + 1).setValue(kv[k]);
  }
}

function hasHeader_(sheet, headerName) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  return headers.indexOf(headerName) >= 0;
}

/******************** DRIVE ********************/
function resolveUploadRootFolderId_(dbg) {
  var primary = clean_(CONFIG.APPLICANT_ROOT_FOLDER_ID_PRIMARY || "");
  var fallback = clean_(CONFIG.APPLICANT_ROOT_FOLDER_ID_FALLBACK || "");
  var propKey = clean_(CONFIG.SCRIPT_PROP_UPLOAD_ROOT_ID || "FODE_UPLOAD_ROOT_ID") || "FODE_UPLOAD_ROOT_ID";
  var propRoot = getScriptProp_(propKey);
  var autoEnabled = CONFIG.AUTO_UPLOAD_ROOT_ENABLED === true;
  var autoName = clean_(CONFIG.AUTO_UPLOAD_ROOT_NAME || "FODE Upload Root (Auto)") || "FODE Upload Root (Auto)";
  var candidateInfo = [
    { source: "PRIMARY", id: primary },
    { source: "FALLBACK", id: fallback },
    { source: "SCRIPT_PROP", id: propRoot }
  ];
  var attemptedSources = [];
  var unusableDetails = [];
  var anyConfiguredCandidate = false;

  for (var i = 0; i < candidateInfo.length; i++) {
    var c = candidateInfo[i];
    var cid = clean_(c.id || "");
    if (!cid) continue;
    anyConfiguredCandidate = true;
    attemptedSources.push(c.source);
    var probe = safeFolderProbe_(cid);
    if (probe && probe.ok) {
      logExecTrace_("ROOT_PROBE_OK", dbg, {
        event: "ROOT_PROBE_OK",
        candidateSource: c.source,
        candidateId: cid
      });
      return {
        primary: primary,
        fallback: fallback,
        propRoot: propRoot,
        chosenRoot: cid,
        source: c.source,
        name: clean_(probe.name || ""),
        rootAttemptSummary: attemptedSources.slice()
      };
    }
    var errName = clean_(probe && probe.errName || "Error") || "Error";
    var errMessage = clean_(probe && probe.errMessage || "probe_failed") || "probe_failed";
    logExecTrace_("ROOT_PROBE_FAIL", dbg, {
      event: "ROOT_PROBE_FAIL",
      candidateSource: c.source,
      candidateId: cid,
      probeErr: {
        errName: errName,
        errMessage: errMessage
      }
    });
    unusableDetails.push(c.source + " msg=" + errName + ": " + errMessage);
  }

  if (autoEnabled) {
    attemptedSources.push("AUTO");
    try {
      var ts = Utilities.formatDate(new Date(), "UTC", "yyyyMMddHHmm");
      var scriptId = "";
      try { scriptId = clean_(ScriptApp.getScriptId() || ""); } catch (_sidErr) {}
      var autoFolderName = autoName + " - " + ts + (scriptId ? " - " + scriptId : "");
      var autoFolder = DriveApp.createFolder(autoFolderName);
      var autoId = clean_(autoFolder.getId() || "");
      var autoActualName = clean_(autoFolder.getName() || autoFolderName);
      if (!autoId) throw new Error("missing_folder_id");
      setScriptProp_(propKey, autoId);
      logExecTrace_("DRIVE_ROOT_PROVISIONED", dbg, {
        event: "DRIVE_ROOT_PROVISIONED",
        folderId: autoId,
        folderName: autoActualName
      });
      return {
        primary: primary,
        fallback: fallback,
        propRoot: propRoot,
        chosenRoot: autoId,
        source: "AUTO",
        name: autoActualName,
        rootAttemptSummary: attemptedSources.slice()
      };
    } catch (autoErr) {
      var autoMsg = clean_(stringifyGsError_(autoErr) || "auto_create_failed");
      unusableDetails.push("AUTO msg=" + autoMsg);
      var eAuto = new Error("folder_root_unusable: " + autoMsg);
      eAuto.errCode = "folder_root_unusable";
      eAuto.rootAttemptSummary = attemptedSources.slice();
      throw eAuto;
    }
  }

  if (!anyConfiguredCandidate) {
    var eUnset = new Error("folder_root_unset");
    eUnset.errCode = "folder_root_unset";
    eUnset.rootAttemptSummary = attemptedSources.slice();
    throw eUnset;
  }

  var eUnusable = new Error("folder_root_unusable: " + (unusableDetails.join(" | ") || "no_usable_root"));
  eUnusable.errCode = "folder_root_unusable";
  eUnusable.rootAttemptSummary = attemptedSources.slice();
  throw eUnusable;
}

function driveApiErrCodeFromStatus_(status) {
  var n = Number(status || 0);
  if (n === 401) return "drive_api_401";
  if (n === 403) return "drive_api_403";
  if (n === 429) return "drive_api_429";
  if (n >= 500 && n < 600) return "drive_api_5xx";
  return "drive_api_error";
}

function throwDriveApiHttpError_(label, resp) {
  var safe = safeHttpErr_(resp && resp.text || "", resp && resp.status || 0);
  var msg = clean_(label || "drive_api") + " failed status=" + String(Number(safe.status || 0));
  if (safe.bodySnippet) msg += " body=" + safe.bodySnippet;
  var err = new Error(msg);
  err.errCode = driveApiErrCodeFromStatus_(safe.status);
  err.status = Number(safe.status || 0);
  err.errorBodySnippet = clean_(safe.bodySnippet || "");
  throw err;
}

function escapeDriveQueryValue_(s) {
  return String(s || "").replace(/\\/g, "\\\\").replace(/'/g, "\\'");
}

function driveApiGetRoot_(dbg) {
  var resp = driveApiGet_("/files/root", { fields: clean_(CONFIG.DRIVE_FIELDS_FOLDER || "id,name,webViewLink") });
  if (!resp.ok) throwDriveApiHttpError_("driveApiGetRoot_", resp);
  return resp.json || {};
}

function driveApiGetFile_(fileId, dbg) {
  var id = encodeURIComponent(clean_(fileId || ""));
  var resp = driveApiGet_("/files/" + id, { fields: clean_(CONFIG.DRIVE_FIELDS_FILE || "id,name,webViewLink,parents") });
  if (!resp.ok) throwDriveApiHttpError_("driveApiGetFile_", resp);
  return resp.json || {};
}

function driveApiFindFolderByName_(parentId, name, dbg) {
  var q = [
    "mimeType='application/vnd.google-apps.folder'",
    "name='" + escapeDriveQueryValue_(name) + "'",
    "'" + escapeDriveQueryValue_(parentId) + "' in parents",
    "trashed=false"
  ].join(" and ");
  var resp = driveApiGet_("/files", {
    q: q,
    pageSize: 1,
    fields: "files(" + clean_(CONFIG.DRIVE_FIELDS_FOLDER || "id,name,webViewLink") + ")"
  });
  if (!resp.ok) throwDriveApiHttpError_("driveApiFindFolderByName_", resp);
  var files = (resp.json && resp.json.files && Array.isArray(resp.json.files)) ? resp.json.files : [];
  return files.length ? files[0] : null;
}

function driveApiCreateFolder_(parentId, name, dbg) {
  var body = {
    name: clean_(name || ""),
    mimeType: "application/vnd.google-apps.folder",
    parents: [clean_(parentId || "")]
  };
  var resp = driveApiPost_("/files", body, { fields: clean_(CONFIG.DRIVE_FIELDS_FOLDER || "id,name,webViewLink") });
  if (!resp.ok) throwDriveApiHttpError_("driveApiCreateFolder_", resp);
  return resp.json || {};
}

function driveApiGetOrCreateFolder_(parentId, name, dbg) {
  var found = driveApiFindFolderByName_(parentId, name, dbg);
  if (found && found.id) return found;
  return driveApiCreateFolder_(parentId, name, dbg);
}

function driveApiUploadBlobToFolder_(parentId, name, blob, dbg) {
  var metadata = {
    name: clean_(name || "upload.bin"),
    parents: [clean_(parentId || "")]
  };
  var resp = driveApiMultipartUpload_(metadata, blob);
  if (!resp.ok) throwDriveApiHttpError_("driveApiUploadBlobToFolder_", resp);
  var j = resp.json || {};
  return {
    fileId: clean_(j.id || ""),
    fileUrl: clean_(j.webViewLink || (j.id ? ("https://drive.google.com/file/d/" + j.id + "/view") : "")),
    fileName: clean_(j.name || metadata.name)
  };
}

function driveApiCreateTextFile_(parentId, name, contentString, dbg) {
  var blob = Utilities.newBlob(String(contentString || ""), "text/plain", clean_(name || "probe.txt"));
  return driveApiUploadBlobToFolder_(parentId, clean_(name || "probe.txt"), blob, dbg);
}

function resolveUploadRootIdForRest_() {
  var primary = clean_(CONFIG.APPLICANT_ROOT_FOLDER_ID_PRIMARY || "");
  var fallback = clean_(CONFIG.APPLICANT_ROOT_FOLDER_ID_FALLBACK || "");
  var propKey = clean_(CONFIG.SCRIPT_PROP_UPLOAD_ROOT_ID || "FODE_UPLOAD_ROOT_ID") || "FODE_UPLOAD_ROOT_ID";
  var propRoot = (typeof getScriptProp_ === "function") ? clean_(getScriptProp_(propKey) || "") : "";
  var candidates = [primary, fallback, propRoot];
  for (var i = 0; i < candidates.length; i++) {
    if (clean_(candidates[i])) return clean_(candidates[i]);
  }
  var e = new Error("folder_root_unset");
  e.errCode = "folder_root_unset";
  throw e;
}

function buildApplicantFolderName_(record, applicantIdHint) {
  var row = record || {};
  var applicantId = clean_(applicantIdHint || row.ApplicantID || row[CONFIG.APPLICANT_ID_HEADER] || "");
  if (applicantId) return applicantId;
  var first = slug_(row.First_Name);
  var last = slug_(row.Last_Name);
  var date = new Date().toISOString().slice(0, 10);
  return first + "_" + last + "_" + date;
}

function driveApiBuildFolderHandleById_(folderId, dbg, existingUrl) {
  var file = driveApiGetFile_(folderId, dbg);
  return {
    kind: "rest",
    id: clean_(file.id || folderId),
    url: clean_(file.webViewLink || existingUrl || (folderId ? ("https://drive.google.com/drive/folders/" + folderId) : ""))
  };
}

function createApplicantFolderHandleWithRestFallback_(payloadOrRecord, dbg, applicantIdHint) {
  var record = payloadOrRecord || {};
  var rootId = resolveUploadRootIdForRest_();
  var yearFolderName = clean_(CONFIG.APPLICANT_ROOT_YEAR_FOLDER_NAME || CONFIG.YEAR_FOLDER || "");
  if (!yearFolderName) {
    var eYear = new Error("folder_root_unusable: missing year folder config");
    eYear.errCode = "folder_root_unusable";
    throw eYear;
  }
  var year = driveApiGetOrCreateFolder_(rootId, yearFolderName, dbg);
  var applicantFolderName = buildApplicantFolderName_(record, applicantIdHint);
  var applicant = driveApiGetOrCreateFolder_(clean_(year.id || ""), applicantFolderName, dbg);
  return {
    kind: "rest",
    id: clean_(applicant.id || ""),
    url: clean_(applicant.webViewLink || (applicant.id ? ("https://drive.google.com/drive/folders/" + applicant.id) : ""))
  };
}

function createApplicantFolder_(payloadOrRecord, opts) {
  var record = payloadOrRecord || {};
  var dbg = clean_(opts && opts.dbg || "");
  var first = slug_(record.First_Name);
  var last = slug_(record.Last_Name);
  var date = new Date().toISOString().slice(0, 10);
  var rootInfo = resolveUploadRootFolderId_(dbg);
  var chosenRoot = clean_(rootInfo && rootInfo.chosenRoot || "");

  logExecTrace_("FOLDER_CREATE_ENTER", dbg, {
    primary: clean_(rootInfo && rootInfo.primary || ""),
    fallback: clean_(rootInfo && rootInfo.fallback || ""),
    propRoot: clean_(rootInfo && rootInfo.propRoot || ""),
    chosenRoot: chosenRoot
  });

  try {
    if (!chosenRoot) {
      var eMissing = new Error("folder_root_unset");
      eMissing.errCode = "folder_root_unset";
      eMissing.rootAttemptSummary = rootInfo && rootInfo.rootAttemptSummary ? rootInfo.rootAttemptSummary : [];
      throw eMissing;
    }
    var root = withRetries_(function () {
      return DriveApp.getFolderById(chosenRoot);
    }, { dbg: dbg, label: "createApplicantFolder:getRootById" });
    var yearFolderName = clean_(CONFIG.APPLICANT_ROOT_YEAR_FOLDER_NAME || CONFIG.YEAR_FOLDER || "");
    if (!yearFolderName) {
      var eYear = new Error("folder_root_unusable: missing year folder config");
      eYear.errCode = "folder_root_unusable";
      eYear.rootAttemptSummary = rootInfo && rootInfo.rootAttemptSummary ? rootInfo.rootAttemptSummary : [];
      throw eYear;
    }
    var year = withRetries_(function () {
      return (typeof getOrCreateFolderByName_ === "function")
        ? getOrCreateFolderByName_(root, yearFolderName, dbg)
        : getOrCreateFolder_(root, yearFolderName);
    }, { dbg: dbg, label: "createApplicantFolder:getOrCreateYear" });
    var applicantFolderName = first + "_" + last + "_" + date;
    return withRetries_(function () {
      return getOrCreateFolder_(year, applicantFolderName);
    }, { dbg: dbg, label: "createApplicantFolder:getOrCreateApplicant" });
  } catch (e) {
    var code = classifyUploadErr_(e);
    if (code === "folder_root_unset" || code === "folder_root_unusable") throw e;
    if (typeof isDriveServerError_ === "function" && isDriveServerError_(e)) {
      var eDriveApp = new Error("driveapp_unavailable: " + clean_(stringifyGsError_(e) || "DriveApp server error"));
      eDriveApp.errCode = "driveapp_unavailable";
      eDriveApp.rootAttemptSummary = rootInfo && rootInfo.rootAttemptSummary ? rootInfo.rootAttemptSummary : [];
      throw eDriveApp;
    }
    var msg = clean_(stringifyGsError_(e) || "Drive error").replace(/\s+/g, " ");
    if (msg.length > 180) msg = msg.slice(0, 180);
    var eDrive = new Error("folder_root_unusable: rootId=" + (chosenRoot || "none") + " msg=" + (msg || "drive_error"));
    eDrive.errCode = "folder_root_unusable";
    eDrive.rootAttemptSummary = rootInfo && rootInfo.rootAttemptSummary ? rootInfo.rootAttemptSummary : [];
    throw eDrive;
  }
}

function driveDeepProbe_(opts) {
  var cfg = opts && typeof opts === "object" ? opts : {};
  var folderId = clean_(cfg.folderId || "");
  var dbg = clean_(cfg.dbg || "");
  var results = {};
  var target = null;
  var tempFolder = null;

  function step_(key, fn) {
    try {
      var value = fn();
      results[key] = { ok: true };
      if (value && typeof value === "object" && !Array.isArray(value)) {
        for (var k in value) results[key][k] = value[k];
      } else if (value !== undefined) {
        results[key].value = value;
      }
      return results[key];
    } catch (e) {
      var se = safeErr_(e);
      results[key] = {
        ok: false,
        errName: clean_(se.name || "Error") || "Error",
        errMessage: clean_(se.message || "Probe step failed") || "Probe step failed"
      };
      return results[key];
    }
  }

  step_("step1_rootName", function () {
    var root = withRetries_(function () { return DriveApp.getRootFolder(); }, { dbg: dbg, label: "driveDeepProbe:getRootFolder#1" });
    return { rootName: withRetries_(function () { return clean_(root.getName() || ""); }, { dbg: dbg, label: "driveDeepProbe:rootGetName" }) };
  });
  step_("step2_rootId", function () {
    var root2 = withRetries_(function () { return DriveApp.getRootFolder(); }, { dbg: dbg, label: "driveDeepProbe:getRootFolder#2" });
    return { rootId: withRetries_(function () { return clean_(root2.getId() || ""); }, { dbg: dbg, label: "driveDeepProbe:rootGetId" }) };
  });
  step_("step3_createTempFolder", function () {
    var name = "FODE_DRIVE_PROBE_" + Utilities.formatDate(new Date(), "UTC", "yyyyMMddHHmmss");
    tempFolder = withRetries_(function () { return DriveApp.createFolder(name); }, { dbg: dbg, label: "driveDeepProbe:createTempFolder" });
    return {
      tempFolderId: clean_(tempFolder.getId() || ""),
      tempFolderUrl: clean_(tempFolder.getUrl() || ""),
      tempFolderName: clean_(tempFolder.getName() || name)
    };
  });
  step_("step4_deleteTempFolder", function () {
    if (!tempFolder) throw new Error("No temp folder to delete");
    withRetries_(function () { return tempFolder.setTrashed(true); }, { dbg: dbg, label: "driveDeepProbe:trashTempFolder" });
    return { trashed: true };
  });
  if (folderId) {
    step_("canOpenTargetFolder", function () {
      target = withRetries_(function () { return DriveApp.getFolderById(folderId); }, { dbg: dbg, label: "driveDeepProbe:getTargetFolder" });
      return {
        folderId: folderId,
        name: withRetries_(function () { return clean_(target.getName() || ""); }, { dbg: dbg, label: "driveDeepProbe:targetGetName" }),
        url: withRetries_(function () { return clean_(target.getUrl() || ""); }, { dbg: dbg, label: "driveDeepProbe:targetGetUrl" })
      };
    });
    step_("canIterateTargetChildren", function () {
      if (!target) target = withRetries_(function () { return DriveApp.getFolderById(folderId); }, { dbg: dbg, label: "driveDeepProbe:getTargetFolderIter" });
      var hasAny = withRetries_(function () {
        var it = target.getFolders();
        return !!(it && it.hasNext && it.hasNext());
      }, { dbg: dbg, label: "driveDeepProbe:targetIterChildren" });
      return { canIterate: true, hasChildFolders: !!hasAny };
    });
    step_("canCreateFileInTarget", function () {
      if (!target) target = withRetries_(function () { return DriveApp.getFolderById(folderId); }, { dbg: dbg, label: "driveDeepProbe:getTargetFolderCreateFile" });
      var f = withRetries_(function () {
        return target.createFile("probe.txt", "ok");
      }, { dbg: dbg, label: "driveDeepProbe:createFileInTarget" });
      try {
        withRetries_(function () { return f.setTrashed(true); }, { dbg: dbg, label: "driveDeepProbe:trashTargetProbeFile" });
      } catch (_trashFileErr) {}
      return { fileId: clean_(f.getId() || ""), fileUrl: clean_(f.getUrl() || "") };
    });
  }

  return {
    ok: true,
    folderId: folderId,
    results: results
  };
}

function getOrCreateFolder_(parent, name) {
  var it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

function folderIdFromUrl_(url) {
  url = String(url || "");
  var m = url.match(/\/folders\/([a-zA-Z0-9_-]+)/);
  return m ? m[1] : "";
}

/******************** ApplicantID ********************/
function assignApplicantIdIfBlank_(sheet, rowNum) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var idCol = findCol_(sheet, CONFIG.APPLICANT_ID_HEADER);
    if (!idCol) throw new Error(CONFIG.APPLICANT_ID_HEADER + " column not found.");

    var idCell = sheet.getRange(rowNum, idCol);
    var existing = String(idCell.getValue() || "").trim();
    if (existing) return existing;

    var lastRow = sheet.getLastRow();
    var values = sheet.getRange(2, idCol, Math.max(lastRow - 1, 1), 1).getValues();
    var flat = values.map(function(r){ return r[0]; });

    var prefix = CONFIG.APPLICANT_PREFIX;
    var digits = CONFIG.APPLICANT_DIGITS;
    var re = new RegExp("^" + escapeRegExp_(prefix) + "(\\d{" + digits + "})$");

    var maxN = 0;
    for (var i = 0; i < flat.length; i++) {
      var s = String(flat[i] || "").trim();
      var mm = s.match(re);
      if (mm) {
        var n = parseInt(mm[1], 10);
        if (!isNaN(n)) maxN = Math.max(maxN, n);
      }
    }

    var nextId = prefix + String(maxN + 1).padStart(digits, "0");
    idCell.setValue(nextId);
    return nextId;
  } finally {
    lock.releaseLock();
  }
}

function findCol_(sheet, headerName) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  for (var i = 0; i < headers.length; i++) {
    if (String(headers[i] || "").trim() === headerName) return i + 1;
  }
  return null;
}

function escapeRegExp_(s) {
  return String(s).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

/******************** DOC META ********************/
function docMetaByField_(fieldName) {
  for (var i = 0; i < CONFIG.DOC_FIELDS.length; i++) {
    if (CONFIG.DOC_FIELDS[i].file === fieldName) {
      return {
        label: CONFIG.DOC_FIELDS[i].label,
        field: CONFIG.DOC_FIELDS[i].file,
        status: CONFIG.DOC_FIELDS[i].status,
        comment: CONFIG.DOC_FIELDS[i].comment,
        multiple: CONFIG.DOC_FIELDS[i].multiple === true
      };
    }
  }
  return null;
}

function getDocUiFields_() {
  return (CONFIG.DOC_FIELDS || []).map(function(d) {
    return { label: d.label, field: d.file, status: d.status, comment: d.comment, multiple: d.multiple === true };
  });
}

/******************** LOG APPEND ********************/
function appendLog_(existing, line) {
  if (!existing) return line;
  return existing + "\n" + line;
}

/******************** OUTPUT ********************/
function jsonOutput_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}


// test code
function test_ShowConfig() {
  Logger.log("SHEET_ID=" + CONFIG.SHEET_ID);
  Logger.log("LOG_SHEET_ID=" + CONFIG.LOG_SHEET_ID);
  Logger.log("LOG_SHEET_NAME=" + CONFIG.LOG_SHEET_NAME);
}

function test_DumpConfigKeys() {
  Logger.log(JSON.stringify(Object.keys(CONFIG).sort(), null, 2));
}

function test_LogSheetWrite() {
  // Open Portal Log spreadsheet by ID
  var ss = SpreadsheetApp.openById(CONFIG.LOG_SHEET_ID);
  var sheet = ss.getSheetByName(CONFIG.LOG_SHEET_NAME);

  if (!sheet) {
    throw new Error(
      "Sheet not found: " + CONFIG.LOG_SHEET_NAME +
      " (Spreadsheet: " + ss.getName() + ")"
    );
  }

  // Minimal write first (don’t push a 50-column row until we confirm wiring)
  sheet.appendRow([new Date(), "TEST", "CONFIG_OK"]);

  Logger.log("Log sheet write successful -> " + ss.getName() + " / " + sheet.getName());
}

function _claspPing() { return "pong"; }
/************************************************************
ADMIN QUEUE RPC (restored using existing row helpers)
************************************************************/

function legacy_admin_getReviewQueues() {
  try {
    var ss = getWorkingSpreadsheet_();
    var sheet = mustGetDataSheet_(ss);
    var rows = admin_listQueueRowObjects_(sheet);

    var counts = {
      docs_pending: 0,
      payment_pending: 0,
      payment_first_anomalies: 0,
      enrolled_ready: 0
    };

    rows.forEach(function(row) {
      var q = classifyAdminQueue_(row);
      if (q && Object.prototype.hasOwnProperty.call(counts, q)) counts[q]++;
    });

    return {
      ok: true,
      queues: [
        {
          id: "docs_pending",
          title: "Documents to Review",
          description: "Applicants waiting for document verification",
          count: counts.docs_pending
        },
        {
          id: "payment_pending",
          title: "Payments to Verify",
          description: "Applicants with documents cleared and payment pending",
          count: counts.payment_pending
        },
        {
          id: "payment_first_anomalies",
          title: "Payment-First Anomalies",
          description: "Payment marked before document completion",
          count: counts.payment_first_anomalies
        },
        {
          id: "enrolled_ready",
          title: "Paid & Approved for Enrollment",
          description: "Applicants ready for next downstream action",
          count: counts.enrolled_ready
        }
      ]
    };

  } catch (e) {
    return {
      ok: false,
      error: e && e.message ? e.message : String(e)
    };
  }
}

function legacy_admin_getQueueItems(queueId, limit, offset) {
  try {
    var ss = getWorkingSpreadsheet_();
    var sheet = mustGetDataSheet_(ss);
    var rows = admin_listQueueRowObjects_(sheet);

    var safeLimit = Math.max(1, Math.min(Number(limit || 50), 200));
    var safeOffset = Math.max(0, Number(offset || 0));

    var filtered = rows.filter(function(row) {
      return classifyAdminQueue_(row) === queueId;
    });

    var page = filtered.slice(safeOffset, safeOffset + safeLimit).map(function(row) {
      return mapAdminQueueRow_(row);
    });

    return {
      ok: true,
      queueId: queueId,
      items: page,
      total: filtered.length,
      offset: safeOffset,
      limit: safeLimit,
      hasMore: (safeOffset + safeLimit) < filtered.length
    };

  } catch (e) {
    return {
      ok: false,
      error: e && e.message ? e.message : String(e),
      queueId: queueId
    };
  }
}

function admin_listQueueRowObjects_(sheet) {
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (!lastRow || lastRow < 2 || !lastCol) return [];

  var data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  if (!data || data.length < 2) return [];

  var headers = data[0].map(function(h) {
    return String(h == null ? "" : h).trim();
  });

  var out = [];
  for (var r = 1; r < data.length; r++) {
    var raw = data[r];
    var row = {};
    for (var c = 0; c < headers.length; c++) {
      row[headers[c]] = raw[c];
    }

    var applicantId = String(firstNonEmpty_(
      row.ApplicantID,
      row.Applicant_Id,
      row["Applicant ID"],
      row.ID
    ) || "").trim();

    if (!applicantId) continue;
    row.__rowNum = r + 1;
    out.push(row);
  }

  return out;
}

function classifyAdminQueue_(row) {
  var birthStatus = String(firstNonEmpty_(
    row.Birth_ID_Status,
    row.Birth_Status,
    row["Birth_ID_Status"],
    row["Birth Status"],
    row["Birth ID Status"]
  ) || "").trim();
  var reportStatus = String(firstNonEmpty_(row.Report_Status, row["Report Status"]) || "").trim();
  var photoStatus = String(firstNonEmpty_(row.Photo_Status, row["Photo Status"]) || "").trim();
  var transferStatus = String(firstNonEmpty_(row.Transfer_Status, row["Transfer Status"]) || "").trim();
  var receiptStatus = String(firstNonEmpty_(row.Receipt_Status, row["Receipt Status"]) || "").trim();

  var docVerificationStatus = String(firstNonEmpty_(
    row.Doc_Verification_Status,
    row["Doc Verification Status"],
    row.Overall_Document_Status
  ) || "").trim();

  var overallStatus = String(firstNonEmpty_(
    row.Overall_Status,
    row["Overall Status"],
    row.Status,
    row["Application Status"]
  ) || "").trim();

  var paymentRaw = String(firstNonEmpty_(
    row.Payment_Verified,
    row.Payment_Status,
    row["Payment Verified"]
  ) || "").trim();

  var docsComplete = [birthStatus, reportStatus, photoStatus, transferStatus, receiptStatus]
    .filter(function(v) { return v !== ""; })
    .every(function(v) { return /verified/i.test(v); });

  var docsVerified = /verified/i.test(docVerificationStatus) || docsComplete;
  var paymentVerified =
    /yes|true|verified|paid/i.test(paymentRaw);
  var hasAnyPayment = paymentVerified || /yes|true|received|paid/i.test(paymentRaw);

  if (!docsVerified && hasAnyPayment) return "payment_first_anomalies";
  if (docsVerified && paymentVerified) {
    return "enrolled_ready";
  }
  if (docsVerified && !paymentVerified) return "payment_pending";
  if (/approved|verified/i.test(overallStatus) && paymentVerified) return "enrolled_ready";
  return "docs_pending";
}

function mapAdminQueueRow_(row) {
  var applicantId = String(firstNonEmpty_(
    row.ApplicantID,
    row.Applicant_Id,
    row["Applicant ID"],
    row.ID
  ) || "").trim();

  var firstName = String(firstNonEmpty_(row.First_Name, row.FirstName, row["First Name"]) || "").trim();
  var lastName = String(firstNonEmpty_(row.Last_Name, row.LastName, row["Last Name"]) || "").trim();
  var fullName = String((firstName + " " + lastName).trim() || firstNonEmpty_(row.Student_Name, row.Name, row["Student Name"]) || "").trim();

  var docVerificationStatus = String(firstNonEmpty_(
    row.Doc_Verification_Status,
    row["Doc Verification Status"],
    row.Overall_Document_Status
  ) || "").trim();

  var paymentVerifiedRaw = String(firstNonEmpty_(
    row.Payment_Verified,
    row["Payment Verified"],
    row.PaymentStatus,
    row.Payment_Status,
    row.Payment
  ) || "").trim();

  var overallStatus = String(firstNonEmpty_(
    row.Overall_Status,
    row["Overall Status"],
    row.Status,
    row["Application Status"]
  ) || "").trim();

  var portalStatus = String(firstNonEmpty_(
    row.Portal_Status,
    row.PortalStatus,
    row.Portal,
    row["Portal Status"]
  ) || "").trim();

  var docsFollowUp = String(firstNonEmpty_(
    row.Docs_Follow_Up,
    row.DocsFollowUp,
    row["Docs Follow-Up"],
    row["Docs Follow Up"]
  ) || "").trim();

  return {
    applicantId: applicantId,
    name: fullName,
    docStatus: docVerificationStatus || overallStatus,
    paymentStatus: /yes|true|verified|paid/i.test(paymentVerifiedRaw) ? "Payment Verified" : "Pending",
    portalStatus: portalStatus,
    docsFollowUp: docsFollowUp,
    eligibleDocsFollowUp: /verified/i.test(docVerificationStatus || overallStatus) && !/yes|true|verified|paid/i.test(paymentVerifiedRaw)
  };
}

function firstNonEmpty_() {
  for (var i = 0; i < arguments.length; i++) {
    var v = arguments[i];
    if (v !== null && v !== undefined && String(v).trim() !== "") return v;
  }
  return "";
}




function campaignLog_(label, payload) {
  var tag = clean_(label || "CAMPAIGN_LOG");
  var data = payload && typeof payload === "object" ? payload : {};
  try {
    log_(mustGetSheet_(getWorkingSpreadsheet_(), CONFIG.LOG_SHEET), tag, JSON.stringify(data));
  } catch (_logErr) {
    try { Logger.log(tag + " " + JSON.stringify(data)); } catch (_e) {}
  }
}

function campaignGetContext_() {
  var ss = getWorkingSpreadsheet_();
  var sh = mustGetDataSheet_(ss);
  ensureCampaignColumns_(sh);
  var values = sh.getDataRange().getValues();
  var headers = values[0] || [];
  var idx = getHeaderIndexMap_(sh);
  return {
    spreadsheet: ss,
    sheet: sh,
    values: values,
    headers: headers,
    idx: idx,
    campaignCols: getCampaignColumnsMap_(headers)
  };
}

function campaignRowObjectFromValues_(headers, row) {
  var out = {};
  var head = Array.isArray(headers) ? headers : [];
  var vals = Array.isArray(row) ? row : [];
  for (var i = 0; i < head.length; i++) {
    var key = clean_(head[i]);
    if (key) out[key] = vals[i];
  }
  return out;
}

function campaignAttemptCount_(row) {
  var raw = Number(row && row.Email_Attempt_Count || 0);
  if (!isFinite(raw) || raw < 0) return 0;
  return Math.floor(raw);
}

function campaignBatchLabel_(baseDate) {
  var dt = baseDate instanceof Date ? baseDate : new Date();
  return "LEGACY-" + Utilities.formatDate(dt, Session.getScriptTimeZone(), "yyyyMMdd-HHmmss");
}

function campaignSubjectForAttempt_(attemptCount, rowNumber) {
  var subjects = Array.isArray(CONFIG.CAMPAIGN_EMAIL_SUBJECTS) ? CONFIG.CAMPAIGN_EMAIL_SUBJECTS : [];
  return subjects[0] || "Your FODE KIA Online Application - Status & Next Step";
}

function buildCampaignEmailBody_(row, portalUrl, applicantId) {
  return [
    "Dear Parent/Guardian,",
    "",
    "We are writing to you regarding your online application submitted to Kundu International Academy under the FODE program.",
    "",
    "Your application is currently on record, and you are now invited to proceed to the next stage through our fully online enrolment and learning system.",
    "",
    "Kundu International Academy has received formal approval from the FODE Head Office to deliver the FODE program through an approved online model. This approval was granted following a detailed review of our academic systems, delivery structure, and compliance measures, ensuring full alignment with national FODE curriculum standards and requirements.",
    "",
    "This means students across Papua New Guinea, including those in remote and rural areas, can now complete their enrolment, submit documents, and progress academically without the need for physical paperwork or travel.",
    "",
    "You may access your student record using the secure link below:",
    "",
    String(portalUrl || ""),
    "",
    "This link is unique to your application and should not be shared.",
    "",
    "Applicant ID: " + String(applicantId || ""),
    "",
    "What you need to do:",
    "",
    "- Review your personal and academic details",
    "- Upload all required documents clearly",
    "- Upload a recent passport-size photo",
    "- Provide accurate contact details",
    "- Submit your application for verification",
    "",
    "Fees and Payment:",
    "",
    "- Registration Fee: K600 (one-time)",
    "- Subject Fee: K450 per subject",
    "- Total cost depends on the number of subjects selected",
    "",
    "All fees are strictly non-refundable, so please ensure all details and documents are accurate before submission.",
    "",
    "Document Requirements:",
    "",
    "- All documents must be clear and readable",
    "- Photos must be recent and passport-style",
    "- Documents must belong to the correct applicant",
    "- Incomplete or incorrect uploads will not be accepted",
    "",
    "Important Information:",
    "",
    "- Submission of false or misleading information may result in cancellation of the application",
    "- Subject selections and placement are final once enrolment is completed",
    "- The application must be fully submitted before further processing can begin",
    "",
    "About the Program:",
    "",
    "Through Kundu FODE, students can:",
    "",
    "- Upgrade Grades 8, 10, or 12",
    "- Study through a structured and flexible system",
    "- Access core subjects including English, Mathematics, Science, ICT, and Business Studies",
    "- Progress towards national examinations and certification",
    "",
    "Kundu International Academy is a registered permitted school under the Papua New Guinea Department of Education (Registration No: PS557/1983) and is formally authorized by FODE Head Office to deliver FODE programs online.",
    "",
    "Your application is already in our system, and this is your opportunity to proceed using the newly approved online platform.",
    "",
    "We strongly recommend completing your submission as soon as possible to secure your place.",
    "",
    "If you require assistance, please contact us:",
    "",
    "FODE Admissions",
    "Kundu International Academy",
    "WhatsApp: +675 7860 4013",
    "Email: fode@kundu.ac"
  ].join("\n");
}

function campaignSendEmailGmail_(toEmail, subject, body) {
  var alias = clean_(CONFIG.CAMPAIGN_GMAIL_ALIAS || "");
  var replyTo = clean_(CONFIG.CAMPAIGN_REPLY_TO || "fode@kundu.ac");
  var to = clean_(toEmail || "");
  if (!to) return { ok: false, error: "Missing recipient email" };
  if (!alias) return { ok: false, error: "Missing campaign Gmail alias" };
  try {
    var aliases = GmailApp.getAliases();
    if (Array.isArray(aliases) && aliases.indexOf(alias) === -1) {
      return { ok: false, error: "Campaign alias not configured: " + alias };
    }
  } catch (_aliasErr) {}
  try {
    GmailApp.sendEmail(to, String(subject || ""), String(body || ""), {
      from: alias,
      replyTo: replyTo,
      name: clean_(CONFIG.EMAIL_FROM_NAME || "FODE Admissions") || "FODE Admissions"
    });
    return { ok: true, to: to, from: alias, replyTo: replyTo };
  } catch (e) {
    return { ok: false, error: String(e && e.message ? e.message : e), to: to, from: alias, replyTo: replyTo };
  }
}

function campaignBuildEmailPreview_(rowObj, rowNumber, attemptCount, batchLabel) {
  var applicantId = clean_(rowObj.ApplicantID || "");
  var secretRes = getActivePortalSecretForCampaign_(applicantId);
  if (!secretRes.ok) {
    return {
      ok: false,
      code: clean_(secretRes.code || "NO_SECRET"),
      applicantId: applicantId,
      rowNumber: rowNumber,
      effectiveEmail: getCampaignEffectiveEmail_(rowObj),
      error: clean_(secretRes.error || secretRes.code || "Missing active secret")
    };
  }
  var portalUrl = buildLegacyCampaignPortalUrl_(applicantId, secretRes.secretPlain);
  var subject = campaignSubjectForAttempt_(attemptCount, rowNumber);
  var body = buildCampaignEmailBody_(rowObj, portalUrl, applicantId);
  return {
    ok: true,
    applicantId: applicantId,
    rowNumber: rowNumber,
    effectiveEmail: getCampaignEffectiveEmail_(rowObj),
    attemptCount: attemptCount,
    batchLabel: batchLabel,
    portalUrl: portalUrl,
    subject: subject,
    body: body
  };
}

function normalizeApplicantMessageType_(messageType) {
  var raw = clean_(messageType || "").toLowerCase();
  var allowed = Array.isArray(CONFIG.COMMUNICATION_ALLOWED_MESSAGE_TYPES) ? CONFIG.COMMUNICATION_ALLOWED_MESSAGE_TYPES : [];
  return allowed.indexOf(raw) >= 0 ? raw : "";
}

function normalizeApplicantBatchFilterType_(filterType) {
  var raw = clean_(filterType || "").toLowerCase();
  var allowed = Array.isArray(CONFIG.COMMUNICATION_ALLOWED_BATCH_FILTER_TYPES) ? CONFIG.COMMUNICATION_ALLOWED_BATCH_FILTER_TYPES : [];
  return allowed.indexOf(raw) >= 0 ? raw : "";
}

function communicationCooldownMs_() {
  return Math.max(1, Number(CONFIG.COMMUNICATION_COOLDOWN_MINUTES || 60)) * 60 * 1000;
}

function communicationCooldownKey_(applicantId, messageType) {
  return "COMM_LAST::" + clean_(messageType || "") + "::" + clean_(applicantId || "");
}

function getLastCommunicationSentAt_(applicantId, messageType) {
  try {
    return clean_(PropertiesService.getScriptProperties().getProperty(communicationCooldownKey_(applicantId, messageType)) || "");
  } catch (_err) {
    return "";
  }
}

function setLastCommunicationSentAt_(applicantId, messageType, isoValue) {
  try {
    PropertiesService.getScriptProperties().setProperty(communicationCooldownKey_(applicantId, messageType), clean_(isoValue || ""));
  } catch (_err) {}
}

function communicationGetActorInfo_(opts) {
  var o = opts && typeof opts === "object" ? opts : {};
  var email = clean_(o.actorEmail || "");
  if (!email && typeof getCallerEmail_ === "function") email = clean_(getCallerEmail_() || "");
  var role = clean_(o.actorRole || "").toUpperCase();
  if (!role && email) {
    if (typeof getAdminRole_ === "function") role = clean_(getAdminRole_(email) || "").toUpperCase();
    if (!role) {
      var mapped = CONFIG.ADMIN_ROLES || {};
      role = clean_(mapped[String(email || "").toLowerCase()] || "VERIFIER").toUpperCase();
    }
  }
  var isAdmin = false;
  if (typeof isAdmin_ === "function") isAdmin = isAdmin_(email);
  else isAdmin = !!role;
  if (!role && isAdmin) role = "VERIFIER";
  return {
    email: email,
    role: role || "",
    isAdmin: !!isAdmin,
    isSuper: role === "SUPER"
  };
}

function communicationBlockReason_(code, messageType) {
  var map = {
    NO_EFFECTIVE_EMAIL: "No effective parent email is available.",
    BOUNCED: "This applicant email is marked as bounced.",
    DO_NOT_CONTACT: "This applicant is marked as do not contact.",
    PORTAL_ALREADY_SUBMITTED: "The portal has already been submitted for this applicant.",
    MISSING_PORTAL_SECRET: "No active portal link is available for this applicant.",
    COOLDOWN_ACTIVE: "A recent message of this type was already sent. Try again later.",
    ROLE_BLOCKED: "Your role is not allowed to perform this action.",
    UNKNOWN_MESSAGE_TYPE: "Unsupported message type.",
    APPLICANT_NOT_FOUND: "Applicant not found.",
    UNKNOWN_FILTER_TYPE: "Unsupported batch planning filter.",
    DOCS_ALREADY_COMPLETE: "Documents are already complete for this applicant.",
    PAYMENT_ALREADY_RESOLVED: "Payment is already resolved for this applicant."
  };
  return map[clean_(code || "")] || ("Action blocked for message type: " + clean_(messageType || "unknown"));
}

function communicationRequiresPortalUrl_(messageType) {
  return ["legacy_invite", "reminder", "docs_missing", "payment_followup"].indexOf(clean_(messageType || "")) >= 0;
}

function communicationDocsMissing_(rowObj) {
  var row = rowObj || {};
  return computeDocVerificationStatus_(row) !== "Verified";
}

function communicationPaymentOutstanding_(rowObj) {
  var row = rowObj || {};
  return !(derivePaymentBadge_(row) === "Verified" || clean_(row.Payment_Verified || "") === "Yes");
}

function communicationMessageTypeForFilter_(filterType) {
  var normalized = normalizeApplicantBatchFilterType_(filterType);
  if (normalized === "legacy_invite_eligible") return "legacy_invite";
  if (normalized === "docs_missing") return "docs_missing";
  if (normalized === "payment_pending") return "payment_followup";
  return "";
}

function communicationMatchesFilterPrecheck_(rowObj, filterType) {
  var row = rowObj || {};
  var applicantId = clean_(row.ApplicantID || "");
  if (!applicantId) return false;
  var normalized = normalizeApplicantBatchFilterType_(filterType);
  if (normalized === "legacy_invite_eligible") {
    var status = normalizeEmailStatus_(row.Email_Status || "");
    return !status || status === "NEW" || status === "READY";
  }
  if (normalized === "docs_missing") return communicationDocsMissing_(row);
  if (normalized === "payment_pending") return communicationPaymentOutstanding_(row);
  return false;
}

function buildReminderEmailBody_(context) {
  return [
    "Dear Parent/Guardian,",
    "",
    "This is a reminder that your FODE KIA online application is still pending completion.",
    "",
    "Please review and complete your application using the secure portal link below:",
    "",
    String(context.portalUrl || ""),
    "",
    "Applicant ID: " + String(context.applicantId || ""),
    "",
    "If you need assistance, contact FODE Admissions at fode@kundu.ac or WhatsApp +675 7860 4013.",
    "",
    "FODE Admissions",
    "Kundu International Academy"
  ].join("\n");
}

function buildDocsMissingEmailBody_(context) {
  return [
    "Dear Parent/Guardian,",
    "",
    "Your FODE KIA application is still missing required documents or has unresolved document checks.",
    "",
    "Please reopen the portal link below, review the required uploads, and submit the missing or corrected documents:",
    "",
    String(context.portalUrl || ""),
    "",
    "Applicant ID: " + String(context.applicantId || ""),
    "",
    "If you need help identifying the required documents, contact FODE Admissions at fode@kundu.ac.",
    "",
    "FODE Admissions",
    "Kundu International Academy"
  ].join("\n");
}

function buildPaymentFollowupEmailBody_(context) {
  return [
    "Dear Parent/Guardian,",
    "",
    "Your FODE KIA application is on record, but payment is still outstanding or pending verification.",
    "",
    "Please use the portal link below to review your application and complete the required payment follow-up steps:",
    "",
    String(context.portalUrl || ""),
    "",
    "Applicant ID: " + String(context.applicantId || ""),
    "",
    "If payment has already been made, please ensure the receipt is uploaded clearly in the portal.",
    "",
    "FODE Admissions",
    "Kundu International Academy"
  ].join("\n");
}

function resolveApplicantMessageContext_(applicantId, messageType, opts) {
  var options = opts && typeof opts === "object" ? opts : {};
  var debugId = clean_(options.debugId || newDebugId_());
  var normalizedType = normalizeApplicantMessageType_(messageType);
  var actor = communicationGetActorInfo_(options);
  var context = {
    ok: true,
    eligible: false,
    blockCode: "",
    blockReason: "",
    effectiveEmail: "",
    portalUrl: "",
    rowObj: null,
    applicantId: clean_(applicantId || ""),
    messageType: normalizedType || clean_(messageType || ""),
    emailStatus: "",
    portalSubmittedActive: false,
    docsVerified: false,
    paymentVerified: false,
    requiresPortalUrl: false,
    debugId: debugId,
    actorEmail: actor.email,
    actorRole: actor.role,
    rowNumber: 0,
    sheet: null,
    batchLabel: clean_(options.batchLabel || "")
  };

  function block(code) {
    context.eligible = false;
    context.blockCode = clean_(code || "");
    context.blockReason = communicationBlockReason_(context.blockCode, context.messageType);
    return context;
  }

  if (!normalizedType) return block("UNKNOWN_MESSAGE_TYPE");
  if (!actor.isAdmin) return block("ROLE_BLOCKED");
  if (clean_(options.action || "") === "planBatch" && !actor.isSuper) return block("ROLE_BLOCKED");

  var sheet = mustGetDataSheet_(getWorkingSpreadsheet_());
  var rowNumber = findRowByApplicantId_(sheet, applicantId);
  if (!rowNumber) return block("APPLICANT_NOT_FOUND");

  var rowObj = getRowObject_(sheet, rowNumber);
  context.sheet = sheet;
  context.rowObj = rowObj;
  context.rowNumber = rowNumber;
  context.applicantId = clean_(rowObj.ApplicantID || applicantId || "");
  context.effectiveEmail = getCampaignEffectiveEmail_(rowObj);
  context.emailStatus = normalizeEmailStatus_(rowObj.Email_Status || "");
  context.portalSubmittedActive = isCampaignPortalSubmittedActive_(rowObj);
  context.docsVerified = computeDocVerificationStatus_(rowObj) === "Verified" || clean_(rowObj.Docs_Verified || "") === "Yes";
  context.paymentVerified = derivePaymentBadge_(rowObj) === "Verified" || clean_(rowObj.Payment_Verified || "") === "Yes";
  context.requiresPortalUrl = communicationRequiresPortalUrl_(normalizedType);

  if (!isValidCampaignEmail_(context.effectiveEmail)) return block("NO_EFFECTIVE_EMAIL");
  if (isCampaignBounceFlagTrue_(rowObj.Email_Bounce_Flag)) return block("BOUNCED");
  if (context.emailStatus === "DO_NOT_CONTACT") return block("DO_NOT_CONTACT");

  var lastSentAt = getLastCommunicationSentAt_(context.applicantId, normalizedType);
  if (lastSentAt) {
    var cooldownRemaining = parseTime_(lastSentAt) + communicationCooldownMs_();
    if (cooldownRemaining > new Date().getTime()) return block("COOLDOWN_ACTIVE");
  }

  if (context.requiresPortalUrl) {
    var secretRes = getActivePortalSecretForCampaign_(context.applicantId);
    if (!secretRes.ok) return block("MISSING_PORTAL_SECRET");
    context.portalUrl = buildLegacyCampaignPortalUrl_(context.applicantId, secretRes.secretPlain);
  }

  if ((normalizedType === "legacy_invite" || normalizedType === "reminder") && context.portalSubmittedActive) {
    return block("PORTAL_ALREADY_SUBMITTED");
  }
  if (normalizedType === "docs_missing" && !communicationDocsMissing_(rowObj)) {
    return block("DOCS_ALREADY_COMPLETE");
  }
  if (normalizedType === "payment_followup" && !communicationPaymentOutstanding_(rowObj)) {
    return block("PAYMENT_ALREADY_RESOLVED");
  }

  context.eligible = true;
  return context;
}


function recordApplicantContactOutcome_(context, outcome, extra) {
  var ctx = context || {};
  var more = extra && typeof extra === "object" ? extra : {};
  if (!ctx.sheet || !ctx.rowNumber) return false;
  var actorEmail = clean_(more.actorEmail || ctx.actorEmail || "");
  var updates = {
    Last_Contact_Type: clean_(ctx.messageType || ""),
    Last_Contact_By: actorEmail,
    Last_Contact_Result: clean_(outcome || ""),
    Last_Contact_Batch: clean_(more.batchLabel || ctx.batchLabel || ""),
    Last_Contact_DebugId: clean_(ctx.debugId || more.debugId || "")
  };
  var subject = clean_(more.subject || "");
  if (subject) updates.Last_Contact_Subject = subject;
  if (clean_(outcome || "") === "SENT") {
    updates.Last_Contacted_At = clean_(more.sentAt || new Date().toISOString());
  }
  return writeApplicantContactTracking_(ctx.sheet, ctx.rowNumber, updates);
}

function buildApplicantMessage_(context) {
  var ctx = context || {};
  var type = normalizeApplicantMessageType_(ctx.messageType || "");
  if (!type) return { ok: false, code: "UNKNOWN_MESSAGE_TYPE", subject: "", body: "" };
  if (type === "legacy_invite") {
    return {
      ok: true,
      subject: campaignSubjectForAttempt_(0, ctx.rowNumber || 0),
      body: buildCampaignEmailBody_(ctx.rowObj || {}, ctx.portalUrl || "", ctx.applicantId || "")
    };
  }
  if (type === "reminder") {
    return {
      ok: true,
      subject: "Reminder: Complete Your FODE KIA Online Application",
      body: buildReminderEmailBody_(ctx)
    };
  }
  if (type === "docs_missing") {
    return {
      ok: true,
      subject: "FODE KIA Application - Missing Documents",
      body: buildDocsMissingEmailBody_(ctx)
    };
  }
  if (type === "payment_followup") {
    return {
      ok: true,
      subject: "FODE KIA Application - Payment Follow-Up",
      body: buildPaymentFollowupEmailBody_(ctx)
    };
  }
  return { ok: false, code: "UNKNOWN_MESSAGE_TYPE", subject: "", body: "" };
}

function dispatchApplicantMessage_(context, builtMessage, opts) {
  var ctx = context || {};
  var message = builtMessage || {};
  var options = opts && typeof opts === "object" ? opts : {};
  var actorEmail = clean_(options.actorEmail || ctx.actorEmail || (typeof getCallerEmail_ === "function" ? getCallerEmail_() : "") || "");
  if (!ctx.eligible) {
    return {
      ok: false,
      result: "BLOCKED",
      blockCode: clean_(ctx.blockCode || ""),
      blockReason: clean_(ctx.blockReason || ""),
      applicantId: clean_(ctx.applicantId || ""),
      messageType: clean_(ctx.messageType || ""),
      debugId: clean_(ctx.debugId || options.debugId || newDebugId_())
    };
  }
  if (!clean_(message.subject || "") || !clean_(message.body || "") || !clean_(ctx.effectiveEmail || "") || !ctx.sheet || !ctx.rowNumber) {
    recordApplicantContactOutcome_(ctx, "FAILED", {
      actorEmail: actorEmail,
      batchLabel: clean_(options.batchLabel || ctx.batchLabel || ""),
      subject: clean_(message.subject || "")
    });
    return {
      ok: false,
      result: "FAILED",
      code: "DISPATCH_INVALID",
      applicantId: clean_(ctx.applicantId || ""),
      messageType: clean_(ctx.messageType || ""),
      effectiveEmail: clean_(ctx.effectiveEmail || ""),
      debugId: clean_(ctx.debugId || options.debugId || newDebugId_())
    };
  }
  var sendRes = campaignSendEmailGmail_(ctx.effectiveEmail, message.subject, message.body);
  if (!sendRes.ok) {
    recordApplicantContactOutcome_(ctx, "FAILED", {
      actorEmail: actorEmail,
      batchLabel: clean_(options.batchLabel || ctx.batchLabel || ""),
      subject: clean_(message.subject || "")
    });
    return {
      ok: false,
      result: "FAILED",
      code: "SEND_FAILED",
      error: clean_(sendRes.error || "SEND_FAILED"),
      applicantId: clean_(ctx.applicantId || ""),
      messageType: clean_(ctx.messageType || ""),
      effectiveEmail: clean_(ctx.effectiveEmail || ""),
      subject: clean_(message.subject || ""),
      debugId: clean_(ctx.debugId || options.debugId || newDebugId_())
    };
  }
  var now = new Date();
  var nextAttempt = campaignAttemptCount_(ctx.rowObj) + 1;
  var patch = {
    Email_Status: "SENT",
    Email_Last_Sent_At: now.toISOString(),
    Email_Attempt_Count: nextAttempt,
    Email_Next_Action_Date: computeNextActionDate_(nextAttempt, now)
  };
  if (clean_(options.batchLabel || ctx.batchLabel || "")) patch.Email_Campaign_Batch = clean_(options.batchLabel || ctx.batchLabel || "");
  applyPatch_(ctx.sheet, ctx.rowNumber, patch);
  setLastCommunicationSentAt_(ctx.applicantId, ctx.messageType, now.toISOString());
  recordApplicantContactOutcome_(ctx, "SENT", {
    actorEmail: actorEmail,
    batchLabel: clean_(options.batchLabel || ctx.batchLabel || ""),
    subject: clean_(message.subject || ""),
    sentAt: now.toISOString()
  });
  return {
    ok: true,
    eligible: true,
    result: "SENT",
    applicantId: clean_(ctx.applicantId || ""),
    messageType: clean_(ctx.messageType || ""),
    effectiveEmail: clean_(ctx.effectiveEmail || ""),
    subject: clean_(message.subject || ""),
    sentAt: now.toISOString(),
    rowNumber: Number(ctx.rowNumber || 0),
    debugId: clean_(ctx.debugId || options.debugId || newDebugId_()),
    blockCode: "",
    blockReason: ""
  };
}

function previewApplicantMessage_(applicantId, messageType, opts) {
  var options = opts && typeof opts === "object" ? opts : {};
  var context = resolveApplicantMessageContext_(applicantId, messageType, Object.assign({}, options, { action: "preview" }));
  if (!context.eligible) {
    var blocked = {
      ok: true,
      action: "preview",
      eligible: false,
      result: "BLOCKED",
      blockCode: clean_(context.blockCode || ""),
      blockReason: clean_(context.blockReason || ""),
      applicantId: clean_(context.applicantId || applicantId || ""),
      messageType: clean_(context.messageType || messageType || ""),
      effectiveEmail: clean_(context.effectiveEmail || ""),
      debugId: clean_(context.debugId || newDebugId_())
    };
    campaignLog_("COMM_PREVIEW", {
      applicantId: blocked.applicantId,
      messageType: blocked.messageType,
      actorEmail: clean_(context.actorEmail || options.actorEmail || ""),
      actorRole: clean_(context.actorRole || options.actorRole || ""),
      blockCode: blocked.blockCode,
      result: "BLOCKED",
      debugId: blocked.debugId,
      batchLabel: clean_(options.batchLabel || "")
    });
    return blocked;
  }
  var built = buildApplicantMessage_(context);
  var preview = {
    ok: true,
    action: "preview",
    eligible: true,
    result: "PREVIEW",
    blockCode: "",
    blockReason: "",
    applicantId: clean_(context.applicantId || ""),
    messageType: clean_(context.messageType || ""),
    effectiveEmail: clean_(context.effectiveEmail || ""),
    portalUrl: clean_(context.portalUrl || ""),
    subject: clean_(built.subject || ""),
    body: String(built.body || ""),
    debugId: clean_(context.debugId || newDebugId_())
  };
  campaignLog_("COMM_PREVIEW", {
    applicantId: preview.applicantId,
    messageType: preview.messageType,
    actorEmail: clean_(context.actorEmail || options.actorEmail || ""),
    actorRole: clean_(context.actorRole || options.actorRole || ""),
    blockCode: "",
    result: "PREVIEW",
    debugId: preview.debugId,
    batchLabel: clean_(options.batchLabel || "")
  });
  return preview;
}

function sendApplicantMessage_(applicantId, messageType, opts) {
  var options = opts && typeof opts === "object" ? opts : {};
  var context = resolveApplicantMessageContext_(applicantId, messageType, Object.assign({}, options, { action: "send" }));
  if (!context.eligible) {
    recordApplicantContactOutcome_(context, "BLOCKED", {
      actorEmail: clean_(options.actorEmail || context.actorEmail || (typeof getCallerEmail_ === "function" ? getCallerEmail_() : "") || ""),
      batchLabel: clean_(options.batchLabel || "")
    });
    var blocked = {
      ok: true,
      action: "send",
      result: "BLOCKED",
      eligible: false,
      blockCode: clean_(context.blockCode || ""),
      blockReason: clean_(context.blockReason || ""),
      applicantId: clean_(context.applicantId || applicantId || ""),
      messageType: clean_(context.messageType || messageType || ""),
      effectiveEmail: clean_(context.effectiveEmail || ""),
      debugId: clean_(context.debugId || newDebugId_())
    };
    campaignLog_("COMM_BLOCKED", {
      applicantId: blocked.applicantId,
      messageType: blocked.messageType,
      actorEmail: clean_(context.actorEmail || options.actorEmail || ""),
      actorRole: clean_(context.actorRole || options.actorRole || ""),
      blockCode: blocked.blockCode,
      result: "BLOCKED",
      debugId: blocked.debugId,
      batchLabel: clean_(options.batchLabel || "")
    });
    return blocked;
  }
  var built = buildApplicantMessage_(context);
  var dispatched = dispatchApplicantMessage_(context, built, options);
  campaignLog_(dispatched.result === "SENT" ? "COMM_SENT" : "COMM_FAILED", {
    applicantId: clean_(context.applicantId || applicantId || ""),
    messageType: clean_(context.messageType || messageType || ""),
    actorEmail: clean_(context.actorEmail || options.actorEmail || ""),
    actorRole: clean_(context.actorRole || options.actorRole || ""),
    blockCode: clean_(dispatched.blockCode || dispatched.code || ""),
    result: clean_(dispatched.result || (dispatched.ok ? "SENT" : "FAILED")),
    debugId: clean_(dispatched.debugId || context.debugId || newDebugId_()),
    batchLabel: clean_(options.batchLabel || "")
  });
  return dispatched;
}

function planApplicantBatch_(filterType, limit, opts) {
  var options = opts && typeof opts === "object" ? opts : {};
  var debugId = clean_(options.debugId || newDebugId_());
  var normalizedFilter = normalizeApplicantBatchFilterType_(filterType);
  var actor = communicationGetActorInfo_(options);
  if (!normalizedFilter) {
    return {
      ok: true,
      eligible: 0,
      blocked: 0,
      selected: 0,
      sampleRecipients: [],
      blockCounts: { UNKNOWN_FILTER_TYPE: 1 },
      limit: Math.max(1, Math.floor(Number(limit || 20))),
      filterType: clean_(filterType || ""),
      debugId: debugId,
      blockCode: "UNKNOWN_FILTER_TYPE",
      blockReason: communicationBlockReason_("UNKNOWN_FILTER_TYPE", "")
    };
  }
  if (!actor.isSuper) {
    return {
      ok: true,
      eligible: 0,
      blocked: 0,
      selected: 0,
      sampleRecipients: [],
      blockCounts: { ROLE_BLOCKED: 1 },
      limit: Math.max(1, Math.floor(Number(limit || 20))),
      filterType: normalizedFilter,
      debugId: debugId,
      blockCode: "ROLE_BLOCKED",
      blockReason: communicationBlockReason_("ROLE_BLOCKED", "")
    };
  }
  var batchLimit = Math.max(1, Math.floor(Number(limit || 20)));
  var ctx = campaignGetContext_();
  var headers = ctx.headers;
  var messageType = communicationMessageTypeForFilter_(normalizedFilter);
  var selected = 0;
  var eligible = 0;
  var blocked = 0;
  var blockCounts = {};
  var sampleRecipients = [];
  var candidates = [];
  for (var r = 1; r < ctx.values.length; r++) {
    if (selected >= batchLimit) break;
    var rowObj = campaignRowObjectFromValues_(headers, ctx.values[r]);
    if (!communicationMatchesFilterPrecheck_(rowObj, normalizedFilter)) continue;
    var applicantId = clean_(rowObj.ApplicantID || "");
    if (!applicantId) continue;
    selected++;
    var resolved = resolveApplicantMessageContext_(applicantId, messageType, Object.assign({}, options, { action: "planBatch", actorEmail: actor.email, actorRole: actor.role, debugId: debugId }));
    if (resolved.eligible) eligible++;
    else {
      blocked++;
      var key = clean_(resolved.blockCode || "UNKNOWN");
      blockCounts[key] = Number(blockCounts[key] || 0) + 1;
    }
    var candidate = {
      applicantId: applicantId,
      eligible: !!resolved.eligible,
      blockCode: clean_(resolved.blockCode || ""),
      blockReason: clean_(resolved.blockReason || ""),
      effectiveEmail: clean_(resolved.effectiveEmail || ""),
      messageType: messageType,
      rowNumber: Number(resolved.rowNumber || 0)
    };
    candidates.push(candidate);
    if (sampleRecipients.length < 10) sampleRecipients.push(candidate);
  }
  var summary = {
    ok: true,
    selected: selected,
    eligible: eligible,
    blocked: blocked,
    sampleRecipients: sampleRecipients,
    blockCounts: blockCounts,
    limit: batchLimit,
    filterType: normalizedFilter,
    debugId: debugId,
    candidates: candidates
  };
  campaignLog_("COMM_BATCH_PLAN", {
    applicantId: "",
    messageType: messageType,
    actorEmail: actor.email,
    actorRole: actor.role,
    blockCode: "",
    result: "planned",
    debugId: debugId,
    batchLabel: clean_(options.batchLabel || ""),
    filterType: normalizedFilter,
    selected: selected,
    eligible: eligible,
    blocked: blocked,
    blockCounts: blockCounts
  });
  return summary;
}

function testCampaignPing() {
  return "OK";
}

function campaign_prepareLegacyRows_() {
  var ctx = campaignGetContext_();
  var sh = ctx.sheet;
  var headers = ctx.headers;
  var prepared = 0;
  var keptReady = 0;
  var skippedMissingSecret = 0;
  var skippedIneligible = 0;
  var scanned = Math.max(0, ctx.values.length - 1);
  for (var r = 1; r < ctx.values.length; r++) {
    var rowNumber = r + 1;
    var rowObj = campaignRowObjectFromValues_(headers, ctx.values[r]);
    var applicantId = clean_(rowObj.ApplicantID || "");
    var status = normalizeEmailStatus_(rowObj.Email_Status || "");
    if (status === "READY") {
      keptReady++;
      continue;
    }
    if (status && status !== "NEW") {
      skippedIneligible++;
      continue;
    }
    if (!applicantId) {
      skippedIneligible++;
      continue;
    }
    var resolved = resolveApplicantMessageContext_(applicantId, "legacy_invite", {
      actorEmail: clean_(getCallerEmail_ && getCallerEmail_()),
      actorRole: "SUPER",
      action: "prepare",
      debugId: newDebugId_()
    });
    if (!resolved.eligible) {
      if (resolved.blockCode === "MISSING_PORTAL_SECRET") skippedMissingSecret++;
      else skippedIneligible++;
      continue;
    }
    applyPatch_(sh, rowNumber, { Email_Status: "READY" });
    prepared++;
  }
  var summary = {
    ok: true,
    scanned: scanned,
    prepared: prepared,
    keptReady: keptReady,
    skippedMissingSecret: skippedMissingSecret,
    skippedIneligible: skippedIneligible
  };
  campaignLog_("CAMPAIGN_PREPARE_SUMMARY", summary);
  return summary;
}

function campaign_sendLegacyBatch_(limit, opts) {
  var options = opts && typeof opts === "object" ? opts : {};
  var dryRun = options.dryRun === true;
  var requestedId = clean_(options.applicantId || "");
  var batchLimit = Math.max(1, Math.floor(Number(limit || CONFIG.CAMPAIGN_BATCH_SIZE_DEFAULT || 50)));
  var batchLabel = clean_(options.batchLabel || "") || campaignBatchLabel_(new Date());
  var mergedOpts = Object.assign({}, options, { batchLabel: batchLabel });

  if (requestedId) {
    var single = dryRun
      ? previewApplicantMessage_(requestedId, "legacy_invite", mergedOpts)
      : sendApplicantMessage_(requestedId, "legacy_invite", mergedOpts);
    return {
      ok: true,
      dryRun: dryRun,
      requestedApplicantId: requestedId,
      requestedLimit: batchLimit,
      batchLabel: batchLabel,
      selected: single.eligible || single.result === "sent" ? 1 : 1,
      sent: single.result === "SENT" ? 1 : 0,
      dryRunCount: dryRun && single.eligible ? 1 : 0,
      skippedIneligible: (!single.eligible && single.blockCode) ? 1 : 0,
      skippedMissingSecret: single.blockCode === "MISSING_PORTAL_SECRET" ? 1 : 0,
      skippedNoStatus: 0,
      sendFailed: single.result === "FAILED" ? 1 : 0,
      preview: single.subject ? [{
        applicantId: clean_(single.applicantId || requestedId),
        effectiveEmail: clean_(single.effectiveEmail || ""),
        subject: clean_(single.subject || ""),
        portalUrl: clean_(single.portalUrl || ""),
        batchLabel: batchLabel,
        dryRun: dryRun
      }] : [],
      skipped: (!single.eligible && single.blockCode) || single.result === "FAILED"
        ? [{ applicantId: clean_(single.applicantId || requestedId), rowNumber: Number(single.rowNumber || 0), reason: clean_(single.blockCode || single.code || single.error || "BLOCKED") }]
        : []
    };
  }

  var plan = planApplicantBatch_("legacy_invite_eligible", batchLimit, mergedOpts);
  var sent = 0;
  var dryRunCount = 0;
  var sendFailed = 0;
  var previews = [];
  var skipped = [];
  var candidates = Array.isArray(plan.candidates) ? plan.candidates : [];
  for (var i = 0; i < candidates.length; i++) {
    var candidate = candidates[i];
    if (!candidate.eligible) {
      skipped.push({ applicantId: candidate.applicantId, rowNumber: candidate.rowNumber, reason: candidate.blockCode || "BLOCKED" });
      continue;
    }
    if (dryRun) {
      var preview = previewApplicantMessage_(candidate.applicantId, "legacy_invite", mergedOpts);
      if (preview.eligible) {
        dryRunCount++;
        previews.push({
          applicantId: preview.applicantId,
          effectiveEmail: preview.effectiveEmail,
          subject: preview.subject,
          portalUrl: preview.portalUrl,
          batchLabel: batchLabel,
          dryRun: true
        });
      } else {
        skipped.push({ applicantId: candidate.applicantId, rowNumber: candidate.rowNumber, reason: preview.blockCode || "BLOCKED" });
      }
      continue;
    }
    var sendResult = sendApplicantMessage_(candidate.applicantId, "legacy_invite", mergedOpts);
    if (sendResult.result === "SENT") sent++;
    else if (sendResult.result === "FAILED") {
      sendFailed++;
      skipped.push({ applicantId: candidate.applicantId, rowNumber: candidate.rowNumber, reason: sendResult.code || sendResult.error || "SEND_FAILED" });
    } else if (sendResult.blockCode) {
      skipped.push({ applicantId: candidate.applicantId, rowNumber: candidate.rowNumber, reason: sendResult.blockCode });
    }
  }
  return {
    ok: true,
    dryRun: dryRun,
    requestedApplicantId: requestedId,
    requestedLimit: batchLimit,
    batchLabel: batchLabel,
    selected: Number(plan.selected || 0),
    sent: sent,
    dryRunCount: dryRunCount,
    skippedIneligible: Number(plan.blocked || 0),
    skippedMissingSecret: Number((plan.blockCounts && plan.blockCounts.MISSING_PORTAL_SECRET) || 0),
    skippedNoStatus: 0,
    sendFailed: sendFailed,
    preview: previews,
    skipped: skipped,
    blockCounts: plan.blockCounts || {}
  };
}

function campaign_syncResponses_() {
  var ctx = campaignGetContext_();
  var sh = ctx.sheet;
  var headers = ctx.headers;
  var scanned = Math.max(0, ctx.values.length - 1);
  var updated = 0;
  for (var r = 1; r < ctx.values.length; r++) {
    var rowNumber = r + 1;
    var rowObj = campaignRowObjectFromValues_(headers, ctx.values[r]);
    if (!isCampaignPortalSubmittedActive_(rowObj)) continue;
    var status = normalizeEmailStatus_(rowObj.Email_Status || "");
    if (status === "RESPONDED") continue;
    if (status === "DO_NOT_CONTACT") continue;
    applyPatch_(sh, rowNumber, { Email_Status: "RESPONDED" });
    updated++;
  }
  var summary = { ok: true, scanned: scanned, updated: updated };
  campaignLog_("CAMPAIGN_SYNC_RESPONSES", summary);
  return summary;
}

function campaignExtractBounceEmails_(text) {
  var lower = String(text || "").toLowerCase();
  var matches = lower.match(/[a-z0-9._%+\-]+@[a-z0-9.\-]+\.[a-z]{2,}/g) || [];
  var seen = {};
  var out = [];
  for (var i = 0; i < matches.length; i++) {
    var email = clean_(matches[i]).toLowerCase();
    if (!email || seen[email]) continue;
    seen[email] = true;
    out.push(email);
  }
  return out;
}

function campaignExtractApplicantIds_(text) {
  var matches = String(text || "").match(/FODE-[A-Za-z0-9\-]+/g) || [];
  var seen = {};
  var out = [];
  for (var i = 0; i < matches.length; i++) {
    var id = clean_(matches[i]);
    if (!id || seen[id]) continue;
    seen[id] = true;
    out.push(id);
  }
  return out;
}

function campaignExtractBatchLabel_(text) {
  var match = String(text || "").match(/Campaign Batch:\s*([A-Za-z0-9\-]+)/i);
  return match ? clean_(match[1]) : "";
}

function campaignIsBounceMessage_(message) {
  var lower = String(message || "").toLowerCase();
  return lower.indexOf("mail delivery subsystem") >= 0
    || lower.indexOf("delivery status notification") >= 0
    || lower.indexOf("undeliverable") >= 0
    || lower.indexOf("failure notice") >= 0
    || lower.indexOf("delivery has failed") >= 0;
}

function campaign_processBounces_() {
  var ctx = campaignGetContext_();
  var sh = ctx.sheet;
  var headers = ctx.headers;
  var rowsByEmail = {};
  var rowsByApplicantId = {};
  for (var r = 1; r < ctx.values.length; r++) {
    var rowNumber = r + 1;
    var rowObj = campaignRowObjectFromValues_(headers, ctx.values[r]);
    var applicantId = clean_(rowObj.ApplicantID || "");
    var effectiveEmail = clean_(getCampaignEffectiveEmail_(rowObj)).toLowerCase();
    if (effectiveEmail) {
      rowsByEmail[effectiveEmail] = rowsByEmail[effectiveEmail] || [];
      rowsByEmail[effectiveEmail].push({ rowNumber: rowNumber, row: rowObj });
    }
    if (applicantId) rowsByApplicantId[applicantId] = { rowNumber: rowNumber, row: rowObj };
  }
  var lookbackDays = Math.max(1, Math.floor(Number(CONFIG.CAMPAIGN_BOUNCE_LOOKBACK_DAYS || 7)));
  var threads = GmailApp.search("newer_than:" + lookbackDays + "d", 0, 200);
  var bouncedRows = 0;
  var unmatchedBounceCount = 0;
  var processedMessages = 0;
  var skippedAlreadyBounced = 0;
  var seenRowNumbers = {};
  for (var t = 0; t < threads.length; t++) {
    var messages = threads[t].getMessages();
    for (var m = 0; m < messages.length; m++) {
      var msg = messages[m];
      var blob = [msg.getFrom(), msg.getSubject(), msg.getPlainBody()].join("\n");
      if (!campaignIsBounceMessage_(blob)) continue;
      processedMessages++;
      var emailMatches = campaignExtractBounceEmails_(blob).filter(function (email) {
        return !!rowsByEmail[email];
      });
      var matched = null;
      if (emailMatches.length === 1 && rowsByEmail[emailMatches[0]].length === 1) {
        matched = rowsByEmail[emailMatches[0]][0];
      }
      if (!matched) {
        var applicantIds = campaignExtractApplicantIds_(blob);
        var batchLabel = campaignExtractBatchLabel_(blob);
        var candidateMatches = [];
        for (var i = 0; i < applicantIds.length; i++) {
          var cand = rowsByApplicantId[applicantIds[i]];
          if (!cand) continue;
          if (batchLabel) {
            var rowBatch = clean_(cand.row.Email_Campaign_Batch || "");
            if (rowBatch && rowBatch !== batchLabel) continue;
          }
          candidateMatches.push(cand);
        }
        if (candidateMatches.length === 1) matched = candidateMatches[0];
      }
      if (!matched) {
        unmatchedBounceCount++;
        continue;
      }
      var rowNumber = Number(matched.rowNumber || 0);
      if (!rowNumber || seenRowNumbers[rowNumber]) continue;
      seenRowNumbers[rowNumber] = true;
      var currentStatus = normalizeEmailStatus_(matched.row.Email_Status || "");
      if (currentStatus === "BOUNCED" || isCampaignBounceFlagTrue_(matched.row.Email_Bounce_Flag)) {
        skippedAlreadyBounced++;
        continue;
      }
      var reason = clean_(msg.getSubject() || "") || clean_(msg.getPlainBody() || "").slice(0, 180);
      applyPatch_(sh, rowNumber, {
        Email_Bounce_Flag: true,
        Email_Bounce_Reason: reason,
        Email_Status: "BOUNCED"
      });
      bouncedRows++;
    }
  }
  var summary = {
    ok: true,
    lookbackDays: lookbackDays,
    processedMessages: processedMessages,
    bouncedRows: bouncedRows,
    skippedAlreadyBounced: skippedAlreadyBounced,
    unmatchedBounceCount: unmatchedBounceCount
  };
  campaignLog_("CAMPAIGN_BOUNCE_SUMMARY", summary);
  return summary;
}

function campaign_sendLegacyFollowups_(limit) {
  var batchLimit = Math.max(1, Math.floor(Number(limit || CONFIG.CAMPAIGN_BATCH_SIZE_DEFAULT || 50)));
  var ctx = campaignGetContext_();
  var headers = ctx.headers;
  var now = new Date();
  var todayTs = new Date(now.getFullYear(), now.getMonth(), now.getDate()).getTime();
  var batchLabel = campaignBatchLabel_(now);
  var selected = 0;
  var sent = 0;
  var skipped = [];
  for (var r = 1; r < ctx.values.length; r++) {
    if (selected >= batchLimit) break;
    var rowNumber = r + 1;
    var rowObj = campaignRowObjectFromValues_(headers, ctx.values[r]);
    if (normalizeEmailStatus_(rowObj.Email_Status || "") !== "SENT") continue;
    if (isCampaignPortalSubmittedActive_(rowObj)) {
      skipped.push({ applicantId: clean_(rowObj.ApplicantID || ""), rowNumber: rowNumber, reason: "PORTAL_ALREADY_SUBMITTED" });
      continue;
    }
    if (isCampaignBounceFlagTrue_(rowObj.Email_Bounce_Flag)) continue;
    var nextActionTs = parseTime_(rowObj.Email_Next_Action_Date || "");
    if (!(nextActionTs > 0) || nextActionTs > todayTs) continue;
    var attemptCount = campaignAttemptCount_(rowObj);
    if (attemptCount < 1 || attemptCount >= 3) continue;
    selected++;
    var sendRes = sendApplicantMessage_(clean_(rowObj.ApplicantID || ""), "reminder", { batchLabel: batchLabel });
    if (sendRes.result === "SENT") sent++;
    else skipped.push({ applicantId: clean_(rowObj.ApplicantID || ""), rowNumber: rowNumber, reason: clean_(sendRes.blockCode || sendRes.code || sendRes.error || "SEND_FAILED") });
  }
  var summary = {
    ok: true,
    selected: selected,
    sent: sent,
    batchLabel: batchLabel,
    skipped: skipped
  };
  campaignLog_("CAMPAIGN_FOLLOWUP_SUMMARY", summary);
  return summary;
}

function campaign_getLegacyEmailSummary_() {
  var ctx = campaignGetContext_();
  var headers = ctx.headers;
  var counts = {
    READY: 0,
    SENT: 0,
    BOUNCED: 0,
    RESPONDED: 0,
    DO_NOT_CONTACT: 0,
    NEW: 0,
    BLANK: 0
  };
  var eligibleForInitialSend = 0;
  var eligibleForFollowup = 0;
  var today = new Date();
  today.setHours(0, 0, 0, 0);
  var todayTs = today.getTime();
  for (var r = 1; r < ctx.values.length; r++) {
    var rowObj = campaignRowObjectFromValues_(headers, ctx.values[r]);
    var status = normalizeEmailStatus_(rowObj.Email_Status || "");
    if (!status) counts.BLANK++;
    else if (Object.prototype.hasOwnProperty.call(counts, status)) counts[status]++;
    var eligibility = computeCampaignEligibility_(rowObj);
    if (eligibility.eligible && (!status || status === "NEW" || status === "READY")) eligibleForInitialSend++;
    if (status === "SENT" && !isCampaignPortalSubmittedActive_(rowObj) && !isCampaignBounceFlagTrue_(rowObj.Email_Bounce_Flag)) {
      var attempts = campaignAttemptCount_(rowObj);
      var nextActionTs = parseTime_(rowObj.Email_Next_Action_Date || "");
      if (attempts < 3 && nextActionTs > 0 && nextActionTs <= todayTs) eligibleForFollowup++;
    }
  }
  return {
    ok: true,
    counts: counts,
    eligibleForInitialSend: eligibleForInitialSend,
    eligibleForFollowup: eligibleForFollowup
  };
}

function testCampaignGmailAuth() {
  return {
    ok: true,
    aliases: GmailApp.getAliases()
  };
}
