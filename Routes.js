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

function doPost_portalUpload_(e) {
  return portal_uploadMultipart_(e);
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

  var ss = getWorkingSpreadsheet_();
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

function doGet_file_(e) {
  var dbg = newDebugId_();
  var params = (e && e.parameter && typeof e.parameter === "object") ? e.parameter : {};
  var applicantId = clean_(params.id || "");
  var secret = clean_(params.s || "");
  var fieldKey = clean_(params.field || "");
  var mode = clean_(params.mode || "open").toLowerCase();
  if (mode !== "download") mode = "open";

  if (!applicantId || !secret || !fieldKey) {
    return htmlOutput_(renderFileProxyMessageHtml_("Missing file link parameters.", dbg));
  }

  var allowedFields = {};
  (CONFIG.DOC_FIELDS || []).forEach(function (d) {
    var k = clean_(d && d.file || "");
    if (k) allowedFields[k] = true;
  });
  if (!allowedFields[fieldKey]) {
    return htmlOutput_(renderFileProxyMessageHtml_("Invalid file field.", dbg));
  }

  try {
    var ss = getWorkingSpreadsheet_();
    var sheet = mustGetDataSheet_(ss);
    var found = findPortalRowByIdSecret_(sheet, applicantId, secret);
    if (!found) {
      Logger.log("FILE_PROXY_TOKEN_FAIL dbg=" + dbg + " id=" + applicantId + " field=" + fieldKey);
      return htmlOutput_(renderFileProxyMessageHtml_("Invalid or expired portal link.", dbg));
    }
    var rawValue = clean_((found.record && found.record[fieldKey]) || "");
    var fileRes = getFileBlobByUrlOrId_(rawValue, dbg, "fileProxy:" + fieldKey);
    if (!fileRes || fileRes.ok !== true || !fileRes.blob) {
      return htmlOutput_(renderFileProxyMessageHtml_("File unavailable.", dbg));
    }

    var blob = fileRes.blob;
    var bytes = blob.getBytes();
    var maxBytes = 6815744; // 6.5 MB hard guard for HTML/base64 payload
    if (!bytes || !bytes.length) {
      return htmlOutput_(renderFileProxyMessageHtml_("File is empty.", dbg));
    }
    if (bytes.length > maxBytes) {
      return htmlOutput_(renderFileProxyMessageHtml_("File too large for secure preview. Please contact admissions.", dbg));
    }
    var fileName = safeFileName_(blob.getName() || fileRes.fileName || ("document_" + fieldKey));
    var mimeType = clean_(blob.getContentType() || fileRes.mimeType || "application/octet-stream");
    var allowedMimes = ["application/pdf", "image/jpeg", "image/png", "image/gif"];
    if (allowedMimes.indexOf(mimeType) < 0) {
      return HtmlService.createHtmlOutput("<h2>Unsupported file type.</h2><p>Debug: " + esc_(dbg) + "</p>")
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }
    var b64 = Utilities.base64Encode(bytes);
    Logger.log("FILE_PROXY_OK dbg=%s applicantId=%s field=%s fileId=%s size=%s mode=%s", dbg, applicantId, fieldKey, clean_(fileRes.fileId || ""), String(bytes.length), mode);
    return htmlOutput_(renderFileProxyBlobHtml_(b64, mimeType, fileName, mode, dbg));
  } catch (err) {
    Logger.log("FILE_PROXY_FAIL dbg=%s applicantId=%s field=%s err=%s", dbg, applicantId, fieldKey, stringifyGsError_(err));
    return htmlOutput_(renderFileProxyMessageHtml_("File unavailable.", dbg));
  }
}

function renderFileProxyMessageHtml_(message, dbg) {
  var msg = String(message || "File unavailable.");
  var debugId = clean_(dbg || "");
  var html = ''
    + '<!doctype html><html><head><meta charset="utf-8"><base target="_top">'
    + '<meta name="viewport" content="width=device-width, initial-scale=1"></head>'
    + '<body style="font-family:Arial,Helvetica,sans-serif;padding:24px;max-width:760px;margin:0 auto;">'
    + '<h3 style="margin:0 0 8px 0;">FODE Document Access</h3>'
    + '<div style="padding:12px;border:1px solid #f5c26b;background:#fff6e5;border-radius:8px;">' + esc_(msg) + '</div>'
    + '<div style="margin-top:10px;color:#444;font-size:12px;">DebugId: ' + esc_(debugId || "-") + '</div>'
    + '</body></html>';
  return html;
}

function renderFileProxyBlobHtml_(b64, mimeType, fileName, mode, dbg) {
  var payload = String(b64 || "");
  var mt = clean_(mimeType || "application/octet-stream");
  var fn = safeFileName_(fileName || "document.bin");
  var openMode = clean_(mode || "open") === "download" ? "download" : "open";
  return ''
    + '<!doctype html><html><head><meta charset="utf-8"><base target="_top">'
    + '<meta name="viewport" content="width=device-width, initial-scale=1">'
    + '<style>body{margin:0;background:#0f172a;color:#e2e8f0;font-family:Arial,Helvetica,sans-serif;}'
    + '.wrap{padding:12px}.bar{font-size:12px;color:#94a3b8;margin-bottom:8px}.frame{width:100vw;height:calc(100vh - 60px);border:0;background:#111827}.img{display:block;max-width:100%;max-height:calc(100vh - 60px);margin:0 auto;}</style>'
    + '</head><body><div class="wrap"><div class="bar">Secure file proxy | DebugId: ' + esc_(clean_(dbg || "")) + '</div><div id="mount"></div></div>'
    + '<script>(function(){'
    + 'var b64=' + JSON.stringify(payload) + ';'
    + 'var mime=' + JSON.stringify(mt) + ';'
    + 'var name=' + JSON.stringify(fn) + ';'
    + 'var mode=' + JSON.stringify(openMode) + ';'
    + 'function decodeBase64(x){var bin=atob(x);var len=bin.length;var bytes=new Uint8Array(len);for(var i=0;i<len;i++){bytes[i]=bin.charCodeAt(i);}return bytes;}'
    + 'function triggerDownload(url){var a=document.createElement("a");a.href=url;a.download=name||"document.bin";document.body.appendChild(a);a.click();document.body.removeChild(a);}'
    + 'try{var bytes=decodeBase64(b64);var blob=new Blob([bytes],{type:mime||"application/octet-stream"});var url=URL.createObjectURL(blob);'
    + 'if(mode==="download"){triggerDownload(url);return;}'
    + 'var mount=document.getElementById("mount");'
    + 'if(/^application\\/pdf$/i.test(mime)){var ifr=document.createElement("iframe");ifr.className="frame";ifr.src=url;mount.appendChild(ifr);return;}'
    + 'if(/^image\\//i.test(mime)){var img=document.createElement("img");img.className="img";img.src=url;img.alt=name||"image";mount.appendChild(img);return;}'
    + 'triggerDownload(url);'
    + '}catch(e){document.body.innerHTML="<div style=\\"padding:24px;font-family:Arial,Helvetica,sans-serif;\\">Unable to open file.</div>";}'
    + '})();</script></body></html>';
}
