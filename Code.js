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


/******************** ENTRYPOINT: POST ********************/
function doPost(e) {
  var reqId = makeReqId_();
  var params = (e && e.parameter && typeof e.parameter === "object") ? e.parameter : {};
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
          var ssMissing = SpreadsheetApp.openById(CONFIG.SHEET_ID);
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
      var ssPortal = SpreadsheetApp.openById(CONFIG.SHEET_ID);
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
      var failRedirect = buildPortalRedirectUrl_(applicantId, secret, { error: true, dbg: failDbgId });
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
        tokenValidationPassed: false,
        result: resultObj,
        debugId: failDbgId
      });
    } catch (errPortal) {
      try {
        var ssLog = SpreadsheetApp.openById(CONFIG.SHEET_ID);
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

  var ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
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
function doGet(e) {
  var params = (e && e.parameter && typeof e.parameter === "object") ? e.parameter : {};
  var view = clean_(params.view || "").toLowerCase();
  var isAdminDeployment = isAdminDeploymentRequest_();
  var currentUrl = "";
  try {
    currentUrl = clean_(ScriptApp.getService().getUrl() || "");
  } catch (routeErr) {}
  Logger.log("ROUTE doGet view=%s isAdmin=%s url=%s", view || "(blank)", isAdminDeployment ? "true" : "false", currentUrl);

  if (view === "diag") return respondDiag_(e);
  if (!view) {
    if (isAdminDeployment) return renderAdminApp_(e);
    view = "portal";
  }
  if (view === "admin") return renderAdminApp_(e);
  if (view !== "portal") {
    if (isAdminDeployment) return renderAdminApp_(e);
    view = "portal";
  }

  var id = clean_(params.id || "");
  var secret = clean_(params.s || "");
  var saved = clean_(params.saved || "") === "1";
  var errorFlag = clean_(params.error || "") === "1";
  var dbg = clean_(params.dbg || "");
  var uploadFail = clean_(params.uploadFail || "") === "1";
  var uploadField = clean_(params.field || "");
  var reqId = makeReqId_();
  var debugPage = CONFIG.DEBUG_PORTAL_SHOW_ON_PAGE === true && dbg === "1";
  var reqMeta = getPortalRequestMeta_(e);
  if (!id || !secret) {
    safePortalLog_({
      route: "doGet:portal",
      applicantId: id || "",
      email: reqMeta.ip || "",
      status: "invalid_token",
      message: "missing_params | ua=" + (reqMeta.ua || "")
    }, false);
    var msg = "Missing portal link parameters";
    if (errorFlag) {
      msg = "Missing portal token (id/s). Please reopen your portal link.";
      if (dbg) msg += " Debug: " + dbg;
    }
    return htmlOutput_(renderErrorHtml_(msg));
  }

  var ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  var sheet = mustGetDataSheet_(ss);
  var rowNum = findRowByApplicantId_(sheet, id);
  if (!rowNum) {
    var missingCount = incrementInvalidPortalAttempt_(id);
    if (missingCount > 10) return htmlOutput_(renderErrorHtml_("Too many invalid attempts. Please try again later."));
    safePortalLog_({
      route: "doGet:portal",
      applicantId: id,
      email: reqMeta.ip || "",
      status: "invalid_token",
      message: "row_not_found attempts=" + missingCount + " | ua=" + (reqMeta.ua || "")
    }, false);
    return htmlOutput_(renderErrorHtml_("Invalid portal link"));
  }
  var rowObj = getRowObject_(sheet, rowNum);
  if (!clean_(rowObj[SCHEMA.PORTAL_TOKEN_HASH] || "")) {
    safePortalLog_({
      route: "doGet:portal",
      applicantId: id,
      email: reqMeta.ip || "",
      status: "invalid_token",
      message: "token_not_initialized | ua=" + (reqMeta.ua || "")
    }, false);
    return htmlOutput_(renderErrorHtml_("Link not initialized; contact admissions."));
  }
  var issuedAt = rowObj[SCHEMA.PORTAL_TOKEN_ISSUED_AT];
  if (isPortalTokenExpired_(issuedAt, CONFIG.PORTAL_TOKEN_MAX_AGE_DAYS)) {
    safePortalLog_({
      route: "doGet:portal",
      applicantId: id,
      email: reqMeta.ip || "",
      status: "invalid_token",
      message: "expired | ua=" + (reqMeta.ua || "")
    }, false);
    return htmlOutput_(renderExpiredHtml_());
  }

  var currentHash = clean_(rowObj[SCHEMA.PORTAL_TOKEN_HASH] || "");
  var inputHash = hashPortalSecret_(secret);
  if (!currentHash || !inputHash || currentHash !== inputHash) {
    var badCount = incrementInvalidPortalAttempt_(id);
    if (badCount > 10) return htmlOutput_(renderErrorHtml_("Too many invalid attempts. Please try again later."));
    safePortalLog_({
      route: "doGet:portal",
      applicantId: id,
      email: reqMeta.ip || "",
      status: "invalid_token",
      message: "hash_mismatch attempts=" + badCount + " | ua=" + (reqMeta.ua || "")
    }, false);
    return htmlOutput_(renderErrorHtml_("Invalid portal link"));
  }

  var record = rowObj;
  if (String(record[SCHEMA.PORTAL_ACCESS_STATUS] || "").trim() === "Locked") {
    safePortalLog_({
      route: "doGet:portal",
      applicantId: id,
      email: reqMeta.ip || "",
      status: "locked",
      message: "portal_locked | ua=" + (reqMeta.ua || "")
    }, false);
    return htmlOutput_(renderErrorHtml_("Access suspended"));
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

  // prefill subjects from canonical OR raw Subjects_Selected
  var canonical = clean_(record.Subjects_Selected_Canonical || "");
  var fallbackCsv = subjectsToCsv_(record.Subjects_Selected || "");
  record._SubjectsCsv = canonical || fallbackCsv;

  var examSites = getExamSites_(ss);

  return htmlOutput_(renderPortalHtml_({
    id: id,
    secret: secret,
    reqId: reqId,
    debugPage: debugPage,
    saved: saved,
    errorFlag: errorFlag,
    dbg: dbg,
    uploadFail: uploadFail,
    uploadField: uploadField,
    record: record,
    subjects: CONFIG.PORTAL_SUBJECTS,
    examSites: examSites,
    editFields: getPortalEditableFields_(),
    docs: getDocUiFields_(),
    visibleFields: CONFIG.PORTAL_VISIBLE_FIELDS,
    version: CONFIG.VERSION,
    versionShort: portalVersionShort_(CONFIG.VERSION),
    buildRenderedAt: new Date().toISOString(),
    buildScriptId: ScriptApp.getScriptId()
  }));
}

function diagStatus_() {
  return {
    ok: true,
    version: CONFIG.VERSION,
    now: new Date().toISOString(),
    scriptId: ScriptApp.getScriptId(),
    deployment: "DEV",
    note: "Use this endpoint to confirm deployment + runtime without clasp run"
  };
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
    var ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    var sheet = mustGetSheet_(ss, CONFIG.DATA_SHEET);
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

    writeBack_(sheet, found.rowNum, updates);
    SpreadsheetApp.flush();
    var verifyRow = getRowObject_(sheet, found.rowNum);
    var verifyCell = clean_(verifyRow[fieldName] || "");
    var latestUrl = clean_(createdUrls[createdUrls.length - 1] || "");
    if (!latestUrl || verifyCell.indexOf(latestUrl) < 0) {
      throw new Error("Upload URL was not saved. Please try again.");
    }
    log_(logSheet, "PORTAL UPLOAD", "ApplicantID=" + applicantId + " field=" + fieldName + " files=" + createdUrls.length);

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

  var ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  var sheet = mustGetSheet_(ss, CONFIG.DATA_SHEET);
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
  var dbg = clean_(opts.dbg || "");
  var uploadFail = opts.uploadFail === true;
  var uploadField = clean_(opts.uploadField || "");
  var subjects = opts.subjects || [];
  var examSites = opts.examSites || [];
  var editFields = opts.editFields || [];
  var docs = opts.docs || [];
  var visibleFields = opts.visibleFields || [];
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
  var dis = locked ? "disabled" : "";
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
      + '<input type="checkbox" name="subj" value="' + esc_(subj) + '" ' + checked + " " + dis + " /> "
      + esc_(subj)
      + "</label>";
  }

  var errorBlock = error
    ? '<div style="background:#ffecec;border:1px solid #ffb3b3;padding:12px;border-radius:10px;margin-bottom:16px;"><b>Action required:</b> ' + esc_(error) + "</div>"
    : "";

  var lockReason = clean_(record._PortalLockReason || "");
  var lockedMsg = (lockReason === "portal_access_locked")
    ? "Portal access is locked. Please contact admissions."
    : "Payment has been verified. No further changes are allowed.";
  var lockedBlock = locked
    ? '<div style="background:#e8f0ff;border:1px solid #b6ccff;padding:12px;border-radius:10px;margin-bottom:16px;"><b>Locked:</b> ' + esc_(lockedMsg) + "</div>"
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
  var errBlock = hasErr
    ? '<div style="background:#ffecec;border:1px solid #b30000;padding:8px;margin-bottom:12px;color:#000;">' + esc_(errText) + (dbg ? (" Debug: " + esc_(dbg)) : "") + "</div>"
    : "";
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
    + savedBlock
    + errBlock
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
    + "console.log('PORTAL BUILD:', " + JSON.stringify(buildVersion) + ", '|', " + JSON.stringify(buildLabel) + ");"
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
  for (var i = 0; i < docs.length; i++) {
    var d = docs[i];
    var cur = clean_(record[d.field] || "");
    var st = mapDocStatusForDisplay_(record[d.status]);
    var cm = clean_(record[d.comment] || "");

    var stBadge = "<b>Status:</b> " + esc_(st || "Pending Review");
    var cmBlock = cm ? ("<div style='margin-top:6px;'><b>Admin comment:</b> " + esc_(cm) + "</div>") : "";

    var urlList = normalizeToUrlList_(cur);
    var curLinks = "";
    if (urlList.length) {
      var linksHtml = [];
      for (var u = 0; u < urlList.length; u++) {
        var delBtn = (!locked)
          ? " <button type='button' onclick=\"deleteDocUrl('" + esc_(d.field) + "','" + esc_(encodeURIComponent(urlList[u])) + "')\">Delete</button>"
          : "";
        linksHtml.push("<span><a target='_blank' href='" + esc_(urlList[u]) + "'>Open " + String(u + 1) + "</a>" + delBtn + "</span>");
      }
      curLinks = "<div style='margin-top:6px;'><b>Current files:</b> " + linksHtml.join("<br/>") + "</div>";
    } else {
      curLinks = "<div style='margin-top:6px;'><b>Current files:</b> None uploaded</div>";
    }
    var multipleAttr = d.multiple ? " multiple" : "";
    var multiBadge = d.multiple ? "<div class='muted' style='font-size:12px;margin-top:4px;'>Multi-upload enabled</div>" : "";
    var noteHtml = "";

    var uploadUi = locked
      ? "<div style='margin-top:10px;color:#666;'><i>Uploads disabled (locked).</i></div>"
      : "<div style='margin-top:10px;'>"
        + "<input type='file' id='f_" + esc_(d.field) + "'" + multipleAttr + " onchange=\"onDocFileChange('" + esc_(d.field) + "', " + (d.multiple ? "true" : "false") + ")\" /> "
        + "<button type='button' id='btn_" + esc_(d.field) + "' onclick=\"uploadDoc('" + esc_(d.field) + "', " + (d.multiple ? "true" : "false") + ")\">Upload / Replace</button>"
        + noteHtml
        + "<div id='cur_" + esc_(d.field) + "' style='margin-top:6px;font-size:12px;'>" + curLinks + "</div>"
        + "<div id='msg_" + esc_(d.field) + "' style='margin-top:6px;font-size:12px;'></div>"
        + multiBadge
        + "</div>";

    out += ""
      + "<div style='padding:10px;border:1px solid #eee;border-radius:10px;margin:10px 0;'>"
      + "<div><b>" + esc_(d.label) + "</b></div>"
      + "<div style='margin-top:6px;'>" + stBadge + "</div>"
      + cmBlock
      + curLinks
      + uploadUi
      + "</div>";
  }

  // uploader script
  out += ""
    + "<script>"
    + "var PORTAL_AUTO_UPLOAD=true;"
    + "var PORTAL_LOCKED=" + (locked ? "true" : "false") + ";"
    + "var DOC_MULTI_MAP={"
    + (docs || []).map(function (x) { return "'" + esc_(x.field) + "':" + (x.multiple ? "true" : "false"); }).join(",")
    + "};"
    + "function escHtml(s){return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/\\\"/g,'&quot;').replace(/'/g,'&#039;');}"
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
    + "  if(!urls || !urls.length){ box.innerHTML='<b>Current files:</b> None uploaded'; return; }"
    + "  var parts=urls.map(function(u,i){"
    + "    var line='<a target=\"_blank\" href=\"'+escHtml(u)+'\">Open '+(i+1)+'</a>';"
    + "    if(!PORTAL_LOCKED){ line += ' <button type=\"button\" onclick=\"deleteDocUrl(\\''+fieldName+'\\',\\''+encodeURIComponent(u)+'\\')\">Delete</button>'; }"
    + "    return line;"
    + "  });"
    + "  box.innerHTML='<b>Current files:</b><br/>'+parts.join('<br/>');"
    + "}"
    + "function readFilesAsBase64(files, done, fail){"
    + "  var out=[];"
    + "  var idx=0;"
    + "  function step(){"
    + "    if(idx>=files.length){ done(out); return; }"
    + "    var f=files[idx++];"
    + "    var reader=new FileReader();"
    + "    reader.onload=function(e){"
    + "      var data=e.target && e.target.result ? e.target.result : '';"
    + "      out.push({name:f.name||('upload_'+Date.now()), type:f.type||'application/octet-stream', base64:String(data).split(',').pop()||''});"
    + "      step();"
    + "    };"
    + "    reader.onerror=function(){ fail('Failed to read file: '+(f && f.name ? f.name : 'unknown')); };"
    + "    reader.readAsDataURL(f);"
    + "  }"
    + "  step();"
    + "}"
    + "function onDocFileChange(fieldName, isMultiple){"
    + "  var input=document.getElementById('f_'+fieldName);"
    + "  var msg=document.getElementById('msg_'+fieldName);"
    + "  var btn=document.getElementById('btn_'+fieldName);"
    + "  if(!input || !msg){ return; }"
    + "  var files=[].slice.call(input.files||[]);"
    + "  if(!files.length){ if(btn) btn.disabled=false; msg.innerHTML='Select a file first.'; return; }"
    + "  if(!isMultiple) files=[files[0]];"
    + "  var names=files.map(function(f){ return f.name; }).join(', ');"
    + "  msg.innerHTML='Selected: '+escHtml(names);"
    + "  if(btn) btn.disabled=false;"
    + "  if(PORTAL_AUTO_UPLOAD){ uploadDoc(fieldName, isMultiple); }"
    + "  else { msg.innerHTML += '<br/>Selecting a file does NOT upload until Upload/Replace runs.'; }"
    + "}"
    + "function uploadDoc(fieldName, isMultiple){"
    + "  var input=document.getElementById('f_'+fieldName);"
    + "  var msg=document.getElementById('msg_'+fieldName);"
    + "  if(!input || !input.files || !input.files.length){ if(msg) msg.innerHTML='Select a file first.'; return; }"
    + "  var files=[].slice.call(input.files||[]);"
    + "  if(!isMultiple) files=[files[0]];"
    + "  setUploadBusy(fieldName,true);"
    + "  if(msg) msg.innerHTML='Uploading...';"
    + "  readFilesAsBase64(files,function(payloads){"
    + "    var names=payloads.map(function(p){return p.name;});"
    + "    var types=payloads.map(function(p){return p.type;});"
    + "    var base64s=payloads.map(function(p){return p.base64;});"
    + "    google.script.run"
    + "      .withSuccessHandler(function(res){"
    + "        if(!res || res.ok!==true){"
    + "          var dbg=(res&&res.debugId)?String(res.debugId):'';"
    + "          var redirect=(res&&res.redirectUrl)?String(res.redirectUrl):'';"
    + "          if(msg) msg.innerHTML=(res&&res.error)?String(res.error):('Upload failed.'+(dbg?(' Debug: '+dbg):''));"
    + "          if(redirect){ window.location.replace(redirect); return; }"
    + "          setUploadBusy(fieldName,false);"
    + "          return;"
    + "        }"
    + "        var uploaded=(res && res.urls && res.urls.length) ? res.urls : (res && res.url ? [res.url] : []);"
    + "        var current=getCurrentUrlsFromDom(fieldName);"
    + "        uploaded.forEach(function(u){ if(u && current.indexOf(u)===-1) current.push(u); });"
    + "        renderCurrentUrls(fieldName, current);"
    + "        if(msg) msg.innerHTML='Uploaded '+uploaded.length+' file(s).';"
    + "        input.value='';"
    + "        setUploadBusy(fieldName,false);"
    + "      })"
    + "      .withFailureHandler(function(err){"
    + "        if(msg) msg.innerHTML='Upload failed: '+(err && err.message ? err.message : err);"
    + "        setUploadBusy(fieldName,false);"
    + "      })"
    + "      .uploadPortalFile('" + esc_(id) + "','" + esc_(secret) + "', fieldName, names, types, base64s);"
    + "  }, function(readErr){"
    + "    if(msg) msg.innerHTML='Upload failed: '+readErr;"
    + "    setUploadBusy(fieldName,false);"
    + "  });"
    + "}"
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
  if (hasStudentExec) return { url: studentExec, isStudentReady: true, warning: "" };
  var studentUrl = clean_(CONFIG.WEBAPP_URL_STUDENT || "");
  var adminUrl = clean_(CONFIG.WEBAPP_URL_ADMIN || CONFIG.WEBAPP_URL || "");
  var isStudentReady = /^https:\/\/script\.google\.com\//i.test(studentUrl);
  var url = isStudentReady ? studentUrl : adminUrl;
  var warning = isStudentReady ? "" : "Student URL not configured. Saving may not work for external users.";
  return {
    url: url,
    isStudentReady: isStudentReady,
    warning: warning
  };
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
  var canonicalBase = clean_(CONFIG.WEBAPP_URL_STUDENT_EXEC || "");
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
  var raw = clean_(CONFIG.WEBAPP_URL_STUDENT_EXEC || "");
  if (!raw) raw = clean_(CONFIG.WEBAPP_URL_STUDENT || getStudentActionUrl_().url || "");
  if (!raw) return "";
  if (/^https:\/\/script\.google\.com\/a\//i.test(raw)) {
    return raw.replace(/^https:\/\/script\.google\.com\/a\/[^/]+\//i, "https://script.google.com/");
  }
  return raw;
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
    return !!(current && adminBase && current === adminBase);
  } catch (e) {
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
    var ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    var sh = mustGetSheet_(ss, CONFIG.LOG_SHEET);
    log_(sh, label, JSON.stringify(payload || {}));
  } catch (e) {
    // Diagnostic logging must never break request flow.
  }
}

function mustGetDataSheet_(ss) {
  var sheet = mustGetSheet_(ss, CONFIG.DATA_SHEET);
  if (sheet.getName() !== CONFIG.DATA_SHEET) {
    throw new Error("DATA_SHEET mismatch");
  }
  return sheet;
}
function htmlOutput_(html) {
  return HtmlService.createHtmlOutput(html)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/******************** LOOKUP ********************/
function findPortalRowByIdSecret_(sheet, applicantId, secret) {
  var rowNum = findRowByApplicantId_(sheet, applicantId);
  if (!rowNum) return null;
  var record = getRowObject_(sheet, rowNum);
  var currentHash = clean_(record[SCHEMA.PORTAL_TOKEN_HASH] || "");
  if (!currentHash) return null;
  var inputHash = hashPortalSecret_(clean_(secret || ""));
  if (!inputHash || currentHash !== inputHash) return null;
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
function getPortalLockReason_(record) {
  var row = record || {};
  if (derivePaymentBadge_(row) === "Verified") return "payment_verified";
  if (clean_(row.Payment_Verified).toLowerCase() === "yes") return "payment_verified_compat";
  if (String(row[SCHEMA.PORTAL_ACCESS_STATUS] || "").trim() === "Locked") return "portal_access_locked";
  if (row._PortalHardLocked === true) return "hard_locked";
  return "";
}

function isPortalLocked_(record) {
  return !!getPortalLockReason_(record);
}

function isPaymentVerified_(record) {
  return derivePaymentVerified_(record) === true;
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
  var paymentBadge = derivePaymentBadge_(row);
  var paymentVerified = paymentBadge === "Verified";
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
  // keep compatibility alignment
  if (hasOwn_(row, "Payment_Verified")) row.Payment_Verified = paymentBadge === "Verified" ? "Yes" : "";
  if (docStage === "Verified" && paymentBadge === "Verified") return "Verified";
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
  var ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  var sheet = mustGetSheet_(ss, CONFIG.DATA_SHEET);
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
function createApplicantFolder_(payloadOrRecord) {
  var first = slug_(payloadOrRecord.First_Name);
  var last = slug_(payloadOrRecord.Last_Name);
  var date = new Date().toISOString().slice(0, 10);

  var root = DriveApp.getFolderById(CONFIG.ROOT_FOLDER_ID);
  var year = getOrCreateFolder_(root, CONFIG.YEAR_FOLDER);
  return year.createFolder(first + "_" + last + "_" + date);
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

