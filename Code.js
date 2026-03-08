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
      var isPaymentVerifiedLock = failCode === "PAYMENT_VERIFIED_LOCK";
      var failRedirect = isPaymentVerifiedLock
        ? buildPortalRedirectUrl_(applicantId, secret, { locked: true, msg: "enrolled" })
        : buildPortalRedirectUrl_(applicantId, secret, { error: true, dbg: failDbgId });
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

  var folder = createApplicantFolder_(payload);
  var rowNum = appendRow_(dataSheet, payload, folder);

  // Assign ApplicantID for this new row if blank
  var applicantId = "";
  try {
    applicantId = assignApplicantIdIfBlank_(dataSheet, rowNum);
  } catch (err) {
    log_(logSheet, "APPLICANTID ERROR", "ERROR: " + err.message);
  }

  writeBack_(dataSheet, rowNum, {
    Folder_Url: folder.getUrl(),
    ApplicantID: applicantId || ""
  });
  ensurePortalTokenAtRow_(dataSheet, rowNum);

  return jsonOut_({ status: "ok", ApplicantID: applicantId || "" });
}

/******************** ENTRYPOINT: GET ********************/
function maybeRedirectToCanonical_(e) {
  var currentUrl = ScriptApp.getService().getUrl();
  var canonicalBase = pickCanonicalExecBase_(e);

  if (!currentUrl || !canonicalBase) return null;

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

    if (view === "whoami") {
      return HtmlService.createHtmlOutput(
        JSON.stringify({
          version: CONFIG.VERSION,
          deployment: CONFIG.DEPLOY_VERSION_NUMBER,
          user: Session.getActiveUser().getEmail(),
          url: ScriptApp.getService().getUrl()
        }, null, 2)
      );
    }

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
  var failResult = function (message, code) {
    return {
      ok: false,
      debugId: safeDebugId,
      applicantId: id,
      error: {
        message: clean_(message || "Portal update failed."),
        code: clean_(code || "PORTAL_UPDATE_FAILED")
      }
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

  var updates = {};
  var editFields = effectiveEditFields;
  var sourceFields = postKeys.length ? posted : payload;

  if (editFields.indexOf("Subjects_Selected_Canonical") >= 0 && isDocsVerified_(found.record)) {
    var attemptedSubjectsCanonical = hasOwn_(sourceFields, "Subjects_Selected_Canonical")
      ? clean_(sourceFields.Subjects_Selected_Canonical)
      : "";
    var attemptedSubjectsLegacy = hasOwn_(sourceFields, "Subjects_Selected")
      ? sourceFields.Subjects_Selected
      : (payload.Subjects_Selected || payload.field_Subjects_Selected || "");
    var attemptedSubjectsCsv = attemptedSubjectsCanonical || subjectsToCsv_(attemptedSubjectsLegacy);
    var existingSubjectsCsv = clean_(found.record.Subjects_Selected_Canonical || "") || subjectsToCsv_(found.record.Subjects_Selected || "");
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
      return {
        ok: false,
        code: "SUBJECT_LOCK_DOCS_VERIFIED",
        message: "Subjects are locked because documents have been verified by Admin.",
        debugId: safeDebugId,
        applicantId: id,
        error: {
          message: "Subjects are locked because documents have been verified by Admin.",
          code: "SUBJECT_LOCK_DOCS_VERIFIED"
        }
      };
    }
  }

    // Core fields from portal
  if (hasOwn_(sourceFields, "Date_Of_Birth")) {
    updates.Date_Of_Birth = normalize_(sourceFields.Date_Of_Birth);
  }

  if (hasOwn_(sourceFields, "Physical_Exam_Site")) {
    updates.Physical_Exam_Site = normalize_(sourceFields.Physical_Exam_Site);
  }

    // Subjects canonical is authoritative; normalize legacy/raw payloads if canonical is absent.
  var subjectsCanonical = hasOwn_(sourceFields, "Subjects_Selected_Canonical")
    ? clean_(sourceFields.Subjects_Selected_Canonical)
    : "";
  var subjectsLegacyRaw = hasOwn_(sourceFields, "Subjects_Selected")
    ? sourceFields.Subjects_Selected
    : (payload.Subjects_Selected || payload.field_Subjects_Selected || "");
  var subjectsCsv = subjectsCanonical || subjectsToCsv_(subjectsLegacyRaw);
  if (subjectsCanonical || subjectsCsv) updates.Subjects_Selected_Canonical = subjectsCsv;
  if (hasOwn_(sourceFields, "Upgrade_Grade_Stream")) {
    updates.Upgrade_Grade_Stream = normalize_(sourceFields.Upgrade_Grade_Stream);
  }

    // Never overwrite Parent_Email directly; allow Parent_Email_Corrected only when explicitly provided.
    // Read editable values directly from posted keys that match sheet headers.
  for (var i = 0; i < editFields.length; i++) {
    var h = editFields[i];
    if (h === "Parent_Email") continue;
    if (!hasOwn_(sourceFields, h)) continue;
    updates[h] = normalize_(sourceFields[h]);
  }

    // Validation
  var missing = [];

  var effectiveDob = updates.Date_Of_Birth || clean_(found.record.Date_Of_Birth || "");
  if (!effectiveDob) missing.push("Date of Birth");

  var effectiveSubjects =
    updates.Subjects_Selected_Canonical ||
    clean_(found.record.Subjects_Selected_Canonical || "") ||
    subjectsToCsv_(found.record.Subjects_Selected || "");
  if (!effectiveSubjects) missing.push("Subjects");

  var effectiveSite = updates.Physical_Exam_Site || clean_(found.record.Physical_Exam_Site || "");
  if (!effectiveSite) missing.push("Physical Exam Site");

  if (missing.length) {
    return failResult("Please complete/fix: " + missing.join(", "), "VALIDATION_FAILED");
  }

  updates.PortalLastUpdateAt = new Date().toISOString();

    // mark first submit time if empty
  if (!clean_(found.record.Portal_Submitted)) updates.Portal_Submitted = new Date().toISOString();

  var updateKeys = Object.keys(updates);
  var emailBefore = clean_(found.record.Parent_Email_Corrected || "");
  var emailAfter = hasOwn_(updates, "Parent_Email_Corrected") ? clean_(updates.Parent_Email_Corrected) : emailBefore;
  var emailChanged = hasOwn_(updates, "Parent_Email_Corrected")
    && emailAfter.toLowerCase() !== emailBefore.toLowerCase();
  var patchSample = {};
  for (var ps = 0; ps < updateKeys.length && ps < 5; ps++) {
    var patchKey = updateKeys[ps];
    var patchVal = clean_(updates[patchKey]);
    patchSample[patchKey] = patchVal.length > 120 ? patchVal.slice(0, 120) : patchVal;
  }
  portalDebugLog_("PORTAL_UPDATE_PATCH", {
    applicantId: id,
    rowNumber: rowIndex,
    keys: updateKeys,
    patchSample: patchSample
  });

  log_(logSheet, "PORTAL_UPDATE_PATCH", "keys=" + updateKeys.join(","));
  log_(logSheet, "PORTAL_UPDATE updates", JSON.stringify(updates));
  var beforeReceiptRow = {
    ApplicantID: clean_(found.record.ApplicantID || id || ""),
    First_Name: clean_(found.record.First_Name || ""),
    Last_Name: clean_(found.record.Last_Name || ""),
    Fee_Receipt_File: clean_(found.record.Fee_Receipt_File || "")
  };
  try {
    writeBack_(dataSheet, rowIndex, updates);
    SpreadsheetApp.flush();
  } catch (e) {
    portalDebugLog_("PORTAL_UPDATE_ERROR", {
      applicantId: id,
      rowNumber: rowIndex,
      error: String(e && e.message ? e.message : e)
    });
    return failResult(String(e && e.message ? e.message : e), "WRITE_FAILED");
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
    if (Object.prototype.hasOwnProperty.call(updates, "Fee_Receipt_File")) {
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
  var showErrorBanner = hasErr && !isPaymentVerifiedLock;
  var errBlock = showErrorBanner
    ? '<div style="background:#ffecec;border:1px solid #b30000;padding:8px;margin-bottom:12px;color:#000;">' + esc_(errText) + (dbg ? (" Debug: " + esc_(dbg)) : "") + "</div>"
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
    + "<label><b>Date of Birth (mandatory):</b></label><br/>"
    + '<input type="date" name="Date_Of_Birth" value="' + dobVal + '" style="padding:8px;width:260px;" ' + ro + " />"
    + "</div>"

    + '<div style="margin:12px 0;">'
    + "<label><b>Physical Exam Site (mandatory):</b></label><br/>"
    + '<select name="Physical_Exam_Site" style="padding:8px;width:520px;" ' + dis + ">"
    + '<option value="">-- Select Exam Site --</option>'
    + examOptions
    + "</select>"
    + (locked ? ('<input type="hidden" name="Physical_Exam_Site" value="' + esc_(examVal) + '" />') : "")
    + "</div>"

    + '<div style="margin:12px 0;">'
    + "<label><b>Select Subjects (mandatory):</b></label>"
    + '<div style="margin-top:8px;padding:12px;border:1px solid #eee;border-radius:10px;">'
    + subjectChecks
    + subjectsLockedNotice
    + "</div>"
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
    + "if(!vals.length){alert('Please select at least one subject.');return false;}"
    + "return true;}"
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
    + "if(!packSubjects()) return false;"
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
    + "if(hasError){ sessionStorage.setItem('portalFlash', JSON.stringify({type:'error',dbg:dbgQ,mode:(uploadFailQ?'upload':'update'),field:fieldQ,ts:Date.now()})); }"
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

function admin_getRuntimeInfo() {
  var serviceUrl = "";
  try { serviceUrl = clean_(ScriptApp.getService().getUrl() || ""); } catch (_e) {}
  return {
    ok: true,
    version: clean_(CONFIG.VERSION || ""),
    deployVersion: Number(CONFIG.DEPLOY_VERSION_NUMBER || 0),
    scriptId: clean_(CONFIG.SCRIPT_ID || ScriptApp.getScriptId() || ""),
    serviceUrl: serviceUrl,
    adminBase: canonicalExecBase_(CONFIG.DEPLOYMENT_ID_ADMIN || CONFIG.WEBAPP_URL_ADMIN || ""),
    studentBase: canonicalExecBase_(CONFIG.DEPLOYMENT_ID_STUDENT || CONFIG.WEBAPP_URL_STUDENT || ""),
    nowIso: new Date().toISOString()
  };
}

// SV GROK Mar 7 function overwrite
function admin_getStudentPortalLink(applicantId) {
  Logger.log("admin_getStudentPortalLink CALLED - id=" + applicantId + " caller=" + Session.getEffectiveUser().getEmail());

  try {
    var ss = getWorkingSpreadsheet_();
    var sheet = mustGetDataSheet_(ss);
    var row = findApplicantRowById_(sheet, applicantId);
    if (!row) throw new Error("Applicant not found");

    var tokenHash = firstNonEmpty_(
      row.PORTAL_TOKEN_HASH,
      row.Portal_Token_Hash
    );

    var issuedAt = firstNonEmpty_(
      row.PORTAL_TOKEN_ISSUED_AT,
      row.Portal_Token_Issued_At
    );

    // access status is optional
    var access = firstNonEmpty_(
      row.PORTAL_ACCESS_STATUS,
      row.Portal_Access_Status
    );
    if (access === "Locked") throw new Error("Portal access locked");

    if (!tokenHash || !issuedAt) {
      throw new Error("Portal link error. Debug: token missing");
    }

    var secret = firstNonEmpty_(
      row.PortalSecret,
      row.Portal_Secret
    );
    if (!secret) {
      var secretRes = getPortalSecretForApplicant_(applicantId);
      secret = secretRes && secretRes.ok === true ? clean_(secretRes.secret || "") : "";
    }
    if (!secret) throw new Error("No portal token");

    var base = CONFIG.WEBAPP_URL_STUDENT || "https://script.google.com/macros/s/AKfycbx2ve4bfCEofF_pJnra-UR02BaoumJaUeDS19Amftm2con2e7ggblMfHRzcn6fYAC4g/exec";

    var url =
      base +
      "?view=portal&id=" +
      encodeURIComponent(applicantId) +
      "&s=" +
      encodeURIComponent(secret);

    Logger.log("LINK GENERATED: " + url);

    return {
      ok: true,
      url: url,
      secret: secret
    };

  } catch (e) {
    Logger.log("LINK FAIL: " + e.message);

    return {
      ok: false,
      error: e.message
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

function admin_getReviewQueues() {
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

function admin_getQueueItems(queueId, limit, offset) {
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

  var paymentVerifiedRaw = String(firstNonEmpty_(
    row.Payment_Verified,
    row["Payment Verified"],
    row.PaymentStatus,
    row.Payment_Status,
    row.Payment
  ) || "").trim();

  var docsComplete = [birthStatus, reportStatus, photoStatus, transferStatus, receiptStatus]
    .filter(function(v) { return v !== ""; })
    .every(function(v) { return /verified/i.test(v); });

  var docsVerified = /verified/i.test(docVerificationStatus) || docsComplete;
  var paymentVerified = /yes|true|verified|paid/i.test(paymentVerifiedRaw);
  var hasAnyPayment = paymentVerified || /yes|true|received|paid/i.test(paymentVerifiedRaw);

  if (!docsVerified && hasAnyPayment) return "payment_first_anomalies";
  if (docsVerified && paymentVerified) return "enrolled_ready";
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
