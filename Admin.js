/******************** ADMIN APP (REWORKED FOR HASHED PORTAL TOKENS) ********************/

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
    "Doc_Verification_Status", "Payment_Verified", "Portal_Access_Status"
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

    out.push({
      rowNumber: r + 1,
      applicantId: rid,
      name: (clean_(row[idx.First_Name - 1]) + " " + clean_(row[idx.Last_Name - 1])).trim(),
      email: effectiveEmail,
      docStatus: clean_(row[idx.Doc_Verification_Status - 1]) || "Pending",
      paymentVerified: clean_(row[idx.Payment_Verified - 1]),
      portalAccess: clean_(row[idx.Portal_Access_Status - 1]) || "Open"
    });
  }

  return { ok: true, rows: out };
}

function admin_getApplicantDetail(payload) {
  try {
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
      "ApplicantID", "First_Name", "Last_Name", "Parent_Email_Corrected", "Payment_Verified",
      "Portal_Access_Status", "Doc_Verification_Status", "Doc_Last_Verified_At", "Doc_Last_Verified_By",
      "PortalTokenIssuedAt",
      "Birth_ID_Passport_File", "Latest_School_Report_File", "Transfer_Certificate_File", "Passport_Photo_File", "Fee_Receipt_File",
      "Birth_ID_Status", "Birth_ID_Comment", "Report_Status", "Report_Comment", "Transfer_Status", "Transfer_Comment",
      "Photo_Status", "Photo_Comment", "Receipt_Status", "Receipt_Comment"
    ]);

    var values = sheet.getRange(rowNumber, 1, 1, lastCol).getValues();
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
      Payment_Verified: clean_(row[idx.Payment_Verified - 1]),
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
      var rawUrl = row[idx[m.file] - 1];
      var url = toPlainString_(rawUrl);
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
    detailObj.Payment_Verified = String(detailObj.Payment_Verified || "");
    detailObj.Portal_Access_Status = String(detailObj.Portal_Access_Status || "");
    detailObj.Doc_Verification_Status = String(detailObj.Doc_Verification_Status || "");
    detailObj._docs = (detailObj._docs || []).map(function (d) {
      d.url = asStringUrl_(d.url);
      d.hasFile = /^https?:\/\//i.test(d.url);
      return d;
    });

    if (!detailObj) {
      return { ok: false, error: "Failed to build detail object" };
    }

    var resultObject = { ok: true, detail: detailObj };
    Logger.log("DETAIL RETURN SHAPE: %s", JSON.stringify(resultObject));
    return resultObject;
  } catch (e) {
    return { ok: false, error: "admin_getApplicantDetail failed: " + (e && e.message ? e.message : String(e)) };
  }
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
  var patch = {};
  patch["PortalTokenHash"] = secretHash;
  patch["PortalTokenIssuedAt"] = new Date();
  patch["Doc_Last_Verified_At"] = new Date();
  patch["Doc_Last_Verified_By"] = adminEmail || "admin";
  applyPatch_(sh, rowNumber, patch);
  syncPortalSecretsActive_(applicantId, emailForSecret, fullName, secret, secretHash);

  var link = buildPortalLink_(applicantId, secret);
  log_(openLogSheet_(), "ADMIN_PORTAL_LINK_RESET",
    "row=" + rowNumber + " applicantId=" + applicantId + " by=" + (adminEmail || "admin"));

  return {
    ok: true,
    link: link,
    applicantId: applicantId,
    issuedAt: new Date().toISOString(),
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
  var idx = headerIndex_(sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]);
  requireHeaders_(idx, ["ApplicantID", "Doc_Verification_Status", "Doc_Last_Verified_At", "Doc_Last_Verified_By", "Portal_Access_Status"]);
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
  if (["Pending", "Verified", "Rejected", "Fraudulent"].indexOf(action) === -1) throw new Error("Invalid action");
  if ((action === "Rejected" || action === "Fraudulent") && !reason) throw new Error("Reason required");

  var sh = openDataSheet_();
  var idx = headerIndex_(sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]);
  requireHeaders_(idx, ["Doc_Verification_Status", "Portal_Access_Status", "Doc_Last_Verified_At", "Doc_Last_Verified_By"]);
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
  requireSuperAdmin_(adminEmail);

  payload = payload || {};
  var rowNumber = Number(payload.rowNumber || 0);
  var status = clean_(payload.status || "");
  if (!rowNumber || rowNumber < 2) throw new Error("Invalid rowNumber");
  if (status !== "Open" && status !== "Locked") throw new Error("Invalid status");

  var sh = openDataSheet_();
  var idx = headerIndex_(sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]);
  requireHeaders_(idx, ["Portal_Access_Status", "Doc_Last_Verified_At", "Doc_Last_Verified_By"]);

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
  var requested = Array.isArray(payload.rowNumbers) ? payload.rowNumbers : [];
  var out = [];
  if (requested.length) {
    var seen = {};
    for (var i = 0; i < requested.length; i++) {
      var n = Number(requested[i] || 0);
      if (!n || n < 2 || n > lastRow || seen[n]) continue;
      seen[n] = true;
      out.push(n);
    }
    return out;
  }
  for (var row = 2; row <= lastRow; row++) out.push(row);
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
  var limit = Math.max(0, Number(payload.limit || 0));

  var sh = openDataSheet_();
  var secretsSheet = openPortalSecrets_();
  ensureHeadersExist_(sh, ["PortalTokenHash", "PortalTokenIssuedAt", "Portal_Access_Status"]);
  var idx = headerIndex_(sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]);
  requireHeaders_(idx, ["ApplicantID", "PortalTokenHash"]);

  var lastRow = sh.getLastRow();
  var checked = 0;
  var updated = 0;
  var updatedRows = [];
  var skipped = 0;
  var generatedCount = 0;
  var rekeyedCount = 0;
  var hashSyncedCount = 0;

  for (var rowNumber = 2; rowNumber <= lastRow; rowNumber++) {
    var rowObj = getRowObject_(sh, rowNumber);
    var applicantId = clean_(rowObj.ApplicantID || "");
    if (!applicantId) continue;
    checked++;
    var admissionsHash = clean_(rowObj.PortalTokenHash || "");
    var secretRowIndex = findPortalSecretsRowByApplicantId_(secretsSheet, applicantId);
    var hasActiveSecret = false;
    var activeSecretHash = "";
    if (secretRowIndex) {
      var rec = readPortalSecretsRecord_(secretsSheet, secretRowIndex);
      if (clean_(rec.Status) === "Active" && clean_(rec.Secret_Hash)) {
        hasActiveSecret = true;
        activeSecretHash = clean_(rec.Secret_Hash);
      }
    }

    if (admissionsHash && hasActiveSecret) {
      skipped++;
      continue;
    }

    if (limit > 0 && updated >= limit) break;
    updated++;
    updatedRows.push(rowNumber);

    if (!admissionsHash && hasActiveSecret) {
      hashSyncedCount++;
      if (!dryRun) setPortalTokenHashForRow_(sh, rowNumber, activeSecretHash);
      continue;
    }

    var emailForSecret = clean_(rowObj.Parent_Email_Corrected || "") || clean_(rowObj.Parent_Email || "");
    var fullName = (clean_(rowObj.First_Name || "") + " " + clean_(rowObj.Last_Name || "")).trim();
    var secretInfo = getOrCreateActivePortalSecret_(applicantId, emailForSecret, fullName, sh, rowNumber, {
      dryRun: dryRun,
      secretsSheet: secretsSheet
    });
    if (secretInfo.created) generatedCount++;
    if (admissionsHash && !hasActiveSecret) {
      rekeyedCount++;
    }
  }

  log_(openLogSheet_(), "ADMIN_TOKEN_BACKFILL",
    "dryRun=" + dryRun
    + " checked=" + checked
    + " updated=" + updated
    + " skipped=" + skipped
    + " generated=" + generatedCount
    + " rekeyed=" + rekeyedCount
    + " hashSynced=" + hashSyncedCount
    + " by=" + (adminEmail || "admin"));

  return {
    ok: true,
    dryRun: dryRun,
    checked: checked,
    updated: updated,
    skipped: skipped,
    generatedCount: generatedCount,
    rekeyedCount: rekeyedCount,
    hashSyncedCount: hashSyncedCount,
    updatedRows: updatedRows
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
  var out = [["ApplicantID", "Email", "Full_Name", "Student_Portal_Link"]];
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
    out.push([applicantId, email, fullName, link]);
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
