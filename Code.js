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
  var ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  var dataSheet = mustGetSheet_(ss, CONFIG.DATA_SHEET);
  var logSheet = mustGetSheet_(ss, CONFIG.LOG_SHEET);

  var payload = getPayload_(e);
  appendPortalLog_({ route: "doPost", status: "HIT", message: "doPost called", email: payload.email || payload.Parent_Email || "", applicantId: payload.id || payload.ApplicantID || "" });

  var action = clean_(payload.action || payload._action || "");


  log_(logSheet, "doPost HIT", payloadSummary_(payload));
  log_(logSheet, "ACTION", action || "(blank)");

  // Portal update (normal fields)
  if (action === "portal_update") {
    return handlePortalUpdate_(ss, dataSheet, logSheet, payload);
  }

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

  return jsonOutput_({ status: "ok", ApplicantID: applicantId || "" });
}

/******************** ENTRYPOINT: GET ********************/
function doGet(e) {
  var view = String((e.parameter.view || "portal")).toLowerCase();

  // ROUTE FIRST
  if (view === "admin") return renderAdminApp_(e);
  if (view !== "portal") return htmlOutput_(renderErrorHtml_("Missing portal link"));

  var id = clean_(e.parameter.id || "");
  var secret = clean_(e.parameter.s || "");
  var reqMeta = getPortalRequestMeta_(e);
  if (!id || !secret) {
    safePortalLog_({
      route: "doGet:portal",
      applicantId: id || "",
      email: reqMeta.ip || "",
      status: "invalid_token",
      message: "missing_params | ua=" + (reqMeta.ua || "")
    }, false);
    return htmlOutput_(renderErrorHtml_("Missing portal link parameters"));
  }

  var ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  var sheet = mustGetSheet_(ss, CONFIG.DATA_SHEET);
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

  record._PortalLocked = isPaymentVerified_(record) || String(record[SCHEMA.PORTAL_ACCESS_STATUS] || "").trim() === "Locked";
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
    record: record,
    subjects: CONFIG.PORTAL_SUBJECTS,
    examSites: examSites,
    editFields: getPortalEditableFields_(),
    docs: getDocUiFields_(),
    visibleFields: CONFIG.PORTAL_VISIBLE_FIELDS
  }));
}


/******************** PORTAL UPDATE HANDLER ********************/
function handlePortalUpdate_(ss, dataSheet, logSheet, payload) {
  log_(logSheet, "PORTAL_UPDATE payload", payloadSummary_(payload));
  var effectiveEditFields = getPortalEditableFields_().slice();
  ["Date_Of_Birth", "Physical_Exam_Site", "Subjects_Selected_Canonical"].forEach(function (f) {
    if (effectiveEditFields.indexOf(f) === -1) effectiveEditFields.push(f);
  });
  log_(logSheet, "PORTAL_UPDATE editFields", JSON.stringify(effectiveEditFields));

  var id = clean_(payload.id || "");
  var secret = clean_(payload.s || "");
  var rowIndex = 0;
  try {
    if (!id || !secret) return htmlOutput_(renderErrorHtml_("Missing portal link parameters. Please reopen your portal link."));

    var found = findPortalRowByIdSecret_(dataSheet, id, secret);
    if (!found) return htmlOutput_(renderErrorHtml_("No matching record found. Please reopen your portal link."));
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
      return htmlOutput_(renderErrorHtml_("Access suspended. Please contact admissions."));
    }

    if (isPaymentVerified_(found.record)) {
      return htmlOutput_(renderErrorHtml_("Your record is locked because payment has been verified. No further changes are allowed."));
    }

    var updates = {};
    var editFields = effectiveEditFields;

    // Core fields from portal
    var dob = clean_(payload.Date_Of_Birth || payload.field_Date_Of_Birth || "");
    if (dob) updates.Date_Of_Birth = dob;

    var examSite = clean_(payload.Physical_Exam_Site || payload.field_Physical_Exam_Site || "");
    if (examSite) updates.Physical_Exam_Site = examSite;

    // Subjects canonical is authoritative; normalize legacy/raw payloads if canonical is absent.
    var subjectsCanonical = clean_(payload.Subjects_Selected_Canonical || payload.field_Subjects_Selected_Canonical || "");
    var subjectsLegacyRaw = payload.Subjects_Selected || payload.field_Subjects_Selected || "";
    var subjectsCsv = subjectsCanonical || subjectsToCsv_(subjectsLegacyRaw);
    if (subjectsCsv) updates.Subjects_Selected_Canonical = subjectsCsv;
    var stream = clean_(payload.Upgrade_Grade_Stream || payload.field_Upgrade_Grade_Stream || "");
    if (stream) updates.Upgrade_Grade_Stream = stream;

    // Do NOT overwrite Parent_Email_Corrected.

    // (Optional extra editable fields - currently none)
    for (var i = 0; i < editFields.length; i++) {
      var h = editFields[i];
      if (h === "Parent_Email") continue;
      var key = "field_" + h;
      if (Object.prototype.hasOwnProperty.call(payload, key)) {
        updates[h] = normalize_(payload[key]);
      }
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
      // re-render with typed values
      var rec = shallowCopy_(found.record);
      for (var k in updates) rec[k] = updates[k];

      var canonical = clean_(rec.Subjects_Selected_Canonical || "");
      rec._SubjectsCsv = canonical || subjectsToCsv_(rec.Subjects_Selected || "");
      rec._PortalLocked = false;

      var examSites = getExamSites_(ss);
      return htmlOutput_(renderPortalHtml_({
        id: id,
        secret: secret,
        record: rec,
        subjects: CONFIG.PORTAL_SUBJECTS,
        examSites: examSites,
        editFields: getPortalEditableFields_(),
        docs: getDocUiFields_(),
        visibleFields: CONFIG.PORTAL_VISIBLE_FIELDS,
        error: "Please complete/fix: " + missing.join(", ")
      }));
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
    writeBack_(dataSheet, rowIndex, updates);
    var verify = getRowObject_(dataSheet, rowIndex);
    var failed = [];
    updateKeys.forEach(function (k2) {
      var expected = clean_(updates[k2]);
      var actual = clean_(verify[k2]);
      if (expected && actual !== expected) failed.push(k2);
    });
    portalDebugLog_("PORTAL_UPDATE_RESULT", {
      applicantId: id,
      rowNumber: rowIndex,
      ok: failed.length === 0,
      mismatches: failed
    });
    if (failed.length) {
      return htmlOutput_(renderErrorHtml_("Update failed to persist for " + failed.join(", ")));
    }
    log_(logSheet, "PORTAL_UPDATE_RESULT", "updated=" + updateKeys.length);
    log_(logSheet, "PORTAL UPDATE", "ApplicantID=" + id + " viaSecret=yes");

    return htmlOutput_(renderSuccessHtml_(id));
  } catch (e) {
    portalDebugLog_("PORTAL_UPDATE_ERROR", {
      applicantId: id,
      rowNumber: rowIndex,
      error: String(e && e.message ? e.message : e)
    });
    throw e;
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

  var ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  var sheet = mustGetSheet_(ss, CONFIG.DATA_SHEET);
  var logSheet = mustGetSheet_(ss, CONFIG.LOG_SHEET);

  var found = findPortalRowByIdSecret_(sheet, applicantId, secret);
  if (!found) throw new Error("Record not found.");
  if (String(found.record[SCHEMA.PORTAL_ACCESS_STATUS] || "").trim() === "Locked") throw new Error("Access suspended.");

  if (isPaymentVerified_(found.record)) throw new Error("Record locked (payment verified).");

  // Determine applicant folder (prefer Folder_Url)
  var folderUrl = clean_(found.record.Folder_Url || "");
  var folderId = folderIdFromUrl_(folderUrl);
  var folder;

  if (folderId) {
    folder = DriveApp.getFolderById(folderId);
  } else {
    folder = createApplicantFolder_(found.record);
    writeBack_(sheet, found.rowNum, { Folder_Url: folder.getUrl() });
  }

  var docMeta = docMetaByField_(fieldName);
  var isMultiple = !!(docMeta && docMeta.multiple === true);
  if (fileNames.length !== mimeTypes.length || fileNames.length !== base64List.length) {
    throw new Error("Upload payload mismatch.");
  }
  if (!fileNames.length) throw new Error("No files selected.");

  var createdUrls = [];
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
  log_(logSheet, "PORTAL UPLOAD", "ApplicantID=" + applicantId + " field=" + fieldName + " files=" + createdUrls.length);

  return {
    ok: true,
    field: fieldName,
    url: createdUrls[createdUrls.length - 1],
    urls: createdUrls,
    multiple: isMultiple
  };
}

function portal_deleteUploadedFile(payload) {
  payload = payload || {};
  var applicantId = clean_(payload.applicantId || "");
  var secret = clean_(payload.secret || payload.s || "");
  var fieldName = clean_(payload.field || "");
  var targetUrl = clean_(payload.url || "");
  var rowNumber = Number(payload.rowNumber || 0);

  if (!applicantId || !secret || !fieldName || !targetUrl) {
    return { ok: false, error: "Missing delete payload fields" };
  }

  var ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  var sheet = mustGetSheet_(ss, CONFIG.DATA_SHEET);
  var logSheet = mustGetSheet_(ss, CONFIG.LOG_SHEET);
  var found = findPortalRowByIdSecret_(sheet, applicantId, secret);
  if (!found) return { ok: false, error: "Record not found." };
  if (rowNumber >= 2 && rowNumber !== found.rowNum) return { ok: false, error: "Row mismatch." };
  portalDebugLog_("PORTAL_DELETE_TARGET", {
    applicantId: applicantId,
    rowNumber: found.rowNum,
    fileField: fieldName,
    url: targetUrl
  });
  if (String(found.record[SCHEMA.PORTAL_ACCESS_STATUS] || "").trim() === "Locked") return { ok: false, error: "Locked" };
  if (isPaymentVerified_(found.record)) return { ok: false, error: "Locked" };

  try {
    var docMeta = docMetaByField_(fieldName);
    if (!docMeta) return { ok: false, error: "Invalid field." };
    if (!docMeta.multiple) return { ok: false, error: "Delete only allowed for multi-upload fields." };

    var existing = clean_(found.record[fieldName] || "");
    var updatedCell = removeUrlFromCell_(existing, targetUrl);
    var remainingUrls = normalizeToUrlList_(updatedCell);
    var removed = existing !== updatedCell;

    var updates = {};
    updates[fieldName] = updatedCell;
    updates.PortalLastUpdateAt = new Date().toISOString();
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
    log_(logSheet, "PORTAL DELETE", "ApplicantID=" + applicantId + " field=" + fieldName + " trashed=" + trashed);
    return { ok: true, remainingUrls: remainingUrls, trashed: trashed, warning: warning };
  } catch (e2) {
    portalDebugLog_("PORTAL_DELETE_ERROR", {
      applicantId: applicantId,
      rowNumber: found.rowNum,
      fileField: fieldName,
      error: String(e2 && e2.message ? e2.message : e2)
    });
    return { ok: false, error: String(e2 && e2.message ? e2.message : e2) };
  }
}

/******************** PORTAL HTML ********************/
function renderPortalHtml_(opts) {
  var id = opts.id, secret = opts.secret, record = opts.record;
  var subjects = opts.subjects || [];
  var examSites = opts.examSites || [];
  var editFields = opts.editFields || [];
  var docs = opts.docs || [];
  var visibleFields = opts.visibleFields || [];
  var error = opts.error || "";
  var actionMeta = getStudentActionUrl_();
  var actionUrl = actionMeta.url;

  var locked = record._PortalLocked === true;
  var dis = locked ? "disabled" : "";

  // subject selections: canonical preferred, else fallback
  var csv = clean_(record.Subjects_Selected_Canonical || record._SubjectsCsv || "");
  var selected = parseSubjects_(csv);

  // date input expects yyyy-mm-dd; if your sheet stores dd/mm/yyyy, keep it blank rather than breaking
  var dobVal = esc_(clean_(record.Date_Of_Birth || ""));
  if (dobVal && dobVal.indexOf("/") !== -1) dobVal = ""; // avoid invalid date showing wrong

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

  var lockedBlock = locked
    ? '<div style="background:#e8f0ff;border:1px solid #b6ccff;padding:12px;border-radius:10px;margin-bottom:16px;"><b>Locked:</b> Payment has been verified. No further changes are allowed.</div>'
    : "";
  var actionWarnBlock = actionMeta.warning
    ? '<div style="background:#fff6e5;border:1px solid #f5c26b;padding:12px;border-radius:10px;margin-bottom:16px;"><b>Warning:</b> ' + esc_(actionMeta.warning) + "</div>"
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
    + "<title>FODE Student Portal</title>"
    + '<meta name="viewport" content="width=device-width, initial-scale=1" />'
    + "</head>"
    + '<body style="font-family:Arial,Helvetica,sans-serif;max-width:980px;margin:24px auto;padding:0 16px;">'
    + "<h2>FODE Student Portal</h2>"
    + '<div style="padding:12px;border:1px solid #ddd;border-radius:10px;margin-bottom:16px;">'
    + "<div><b>Applicant ID:</b> " + esc_(id) + "</div>"
    + "<div><b>Secure Link:</b> verified</div>"
    + "</div>"
    + lockedBlock
    + errorBlock
    + actionWarnBlock

    + '<div style="padding:12px;border:1px solid #ddd;border-radius:10px;margin-bottom:16px;">'
    + '<h3 style="margin-top:0;">Submitted Details (read-only)</h3>'
    + summaryHtml
    + "</div>"

    + '<div style="padding:12px;border:1px solid #ddd;border-radius:10px;margin-bottom:16px;">'
    + '<h3 style="margin-top:0;">Documents & Payment Proof</h3>'
    + docsHtml
    + "</div>"

    // ✅ hardcoded action URL to prevent blank screen / doPost not firing
    + '<form method="post" action="' + esc_(actionUrl) + '" onsubmit="return packSubjects();"'
    + ' style="padding:12px;border:1px solid #ddd;border-radius:10px;">'
    + '<input type="hidden" name="action" value="portal_update" />'
    + '<input type="hidden" name="id" value="' + esc_(id) + '" />'
    + '<input type="hidden" name="s" value="' + esc_(secret) + '" />'
    + '<input type="hidden" id="Subjects_Selected_Canonical" name="Subjects_Selected_Canonical" value="" />'

    + '<h3 style="margin-top:0;">Update / Confirm Information</h3>'

    + '<div style="margin:12px 0;">'
    + "<label><b>Date of Birth (mandatory):</b></label><br/>"
    + '<input type="date" name="Date_Of_Birth" value="' + dobVal + '" style="padding:8px;width:260px;" ' + dis + " />"
    + "</div>"

    + '<div style="margin:12px 0;">'
    + "<label><b>Physical Exam Site (mandatory):</b></label><br/>"
    + '<select name="Physical_Exam_Site" style="padding:8px;width:520px;" ' + dis + ">"
    + '<option value="">-- Select Exam Site --</option>'
    + examOptions
    + "</select>"
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

    + "<script>"
    + "function packSubjects(){"
    + "var boxes=[].slice.call(document.querySelectorAll('input[name=\"subj\"]:checked'));"
    + "var vals=boxes.map(function(b){return b.value;}).filter(Boolean);"
    + "document.getElementById('Subjects_Selected_Canonical').value=vals.join(', ');"
    + "if(!vals.length){alert('Please select at least one subject.');return false;}"
    + "return true;}"
    + "</script>"

    + "</body></html>";
}

function renderDocsSection_(id, secret, record, docs, locked) {
  var out = "";
  for (var i = 0; i < docs.length; i++) {
    var d = docs[i];
    var cur = clean_(record[d.field] || "");
    var st = clean_(record[d.status] || "");
    var cm = clean_(record[d.comment] || "");

    var stBadge = st ? ("<b>Status:</b> " + esc_(st)) : "<b>Status:</b> (not set)";
    var cmBlock = cm ? ("<div style='margin-top:6px;'><b>Admin comment:</b> " + esc_(cm) + "</div>") : "";

    var urlList = normalizeToUrlList_(cur);
    var curLinks = "";
    if (urlList.length) {
      var linksHtml = [];
      for (var u = 0; u < urlList.length; u++) {
        var delBtn = (d.multiple && !locked)
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
    + "  var isMulti=!!DOC_MULTI_MAP[fieldName];"
    + "  var parts=urls.map(function(u,i){"
    + "    var line='<a target=\"_blank\" href=\"'+escHtml(u)+'\">Open '+(i+1)+'</a>';"
    + "    if(isMulti && !PORTAL_LOCKED){ line += ' <button type=\"button\" onclick=\"deleteDocUrl(\\''+fieldName+'\\',\\''+encodeURIComponent(u)+'\\')\">Delete</button>'; }"
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
    + "        if(msg) msg.innerHTML='Delete failed: '+((res&&res.error)?res.error:'Unknown error');"
    + "        setUploadBusy(fieldName,false);"
    + "        return;"
    + "      }"
    + "      renderCurrentUrls(fieldName, res.remainingUrls||[]);"
    + "      if(msg) msg.innerHTML='Deleted.'+(res.warning?(' '+res.warning):'');"
    + "      setUploadBusy(fieldName,false);"
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
  var rows = "";

  for (var i = 0; i < keys.length; i++) {
    var k = keys[i];
    var v = clean_(record[k]);
    var display = v ? v : "-";

    // Pretty subjects display
    if (k === "Subjects_Selected_Canonical") {
      var csv = clean_(record.Subjects_Selected_Canonical || "") || subjectsToCsv_(record.Subjects_Selected || "");
      display = csv ? esc_(csv) : "-";
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

function renderEditableFields_(record, editFields, dis) {
  var out = "";
  for (var i = 0; i < editFields.length; i++) {
    var h = editFields[i];
    var val = clean_(record[h] || "");
    var isUrl = isHttpUrl_(val);
    var linkHtml = isUrl ? (" <a target='_blank' rel='noopener' href='" + esc_(val) + "'>Open</a>") : "";
    out += "<div style='margin:10px 0;'>"
      + "<label><b>" + esc_(h) + ":</b></label><br/>"
      + "<input type='text' name='field_" + esc_(h) + "' value='" + esc_(val) + "' style='padding:8px;width:520px;' " + dis + " />"
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
function isPaymentVerified_(record) {
  return String(record.Payment_Verified || "").trim().toLowerCase() === "yes";
}

/******************** PAYLOAD ********************/
function getPayload_(e) {
  if (e && e.parameter && Object.keys(e.parameter).length) return e.parameter;
  if (e && e.postData && e.postData.contents) {
    try { return JSON.parse(e.postData.contents); } catch (err) {}
  }
  return {};
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



