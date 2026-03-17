/************************************************************
ADMIN DOCUMENT VERIFICATION (MINIMAL AUDIT MODEL)
- One admin verifies docs via an admin UI (to be added next)
- We store: per-doc Status + Comment, and a single "last verification" stamp
- Roll-up: Docs_Verified becomes "Yes" only when all REQUIRED docs are VERIFIED
  (Transfer_Certificate is optional by default)
************************************************************/

var ADMIN_DETAIL_SIG = "ADMIN_DETAIL_SIG_20260220_v1";

function makeDebugId_() {
  return adminDebugId_();
}

function adminDebugId_() {
  try {
    if (typeof newDebugId_ === "function") return newDebugId_();
  } catch (_e) {}
  return "ADM-" + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd-HHmmss") + "-" + Math.floor(Math.random() * 1000000).toString().padStart(6, "0");
}

function doGet_file_(e) {
  Logger.log("PROXY CALLED - params=" + JSON.stringify(e.parameter));

  var params = (e && e.parameter && typeof e.parameter === "object") ? e.parameter : {};
  var applicantId = clean_(params.id || "");
  var secret = clean_(params.s || "");
  var fieldKey = clean_(params.field || "");
  var mode = clean_(params.mode || "open");

  var dbg = makeDebugId_();

  if (!applicantId || !secret || !fieldKey) {
    Logger.log("PROXY FAIL: missing params - id=" + applicantId + " field=" + fieldKey + " dbg=" + dbg);
    return HtmlService.createHtmlOutput("<h1>Invalid request</h1><p>Missing required parameters. Debug: " + dbg + "</p>");
  }

  var ss = getWorkingSpreadsheet_();
  var sheet = mustGetDataSheet_(ss);
  var logSheet = mustGetSheet_(ss, CONFIG.LOG_SHEET_NAME);

  var found = findPortalRowByIdSecret_(sheet, applicantId, secret);
  if (!found) {
    Logger.log("PROXY FAIL: invalid token - id=" + applicantId + " field=" + fieldKey + " dbg=" + dbg);
    return HtmlService.createHtmlOutput("<h1>Invalid token</h1><p>Debug: " + dbg + "</p>");
  }

  var row = found.rowNum;
  var record = found.record;

  var url = clean_(record[fieldKey] || "");
  if (!url) {
    Logger.log("PROXY FAIL: no URL - id=" + applicantId + " field=" + fieldKey + " dbg=" + dbg);
    return HtmlService.createHtmlOutput("<h1>No file uploaded</h1><p>Debug: " + dbg + "</p>");
  }

  var id = extractDriveFileId_(url);
  if (!id) {
    Logger.log("PROXY FAIL: invalid file URL - id=" + applicantId + " field=" + fieldKey + " dbg=" + dbg);
    return HtmlService.createHtmlOutput("<h1>Invalid file URL</h1><p>Debug: " + dbg + "</p>");
  }

  try {
    var file = DriveApp.getFileById(id);
    var blob = file.getBlob();
    var mime = blob.getContentType();
    var name = file.getName();

    Logger.log("PROXY SUCCESS: " + id + " | mime=" + mime + " | name=" + name + " | dbg=" + dbg);

    if (mode === "download") {
      return ContentService.createTextOutput(blob.getDataAsString())
        .setMimeType(mime)
        .setName(name);
    } else {
      // open/preview mode
      var base64 = Utilities.base64Encode(blob.getBytes());
      var html = "<!DOCTYPE html><html><head><title>" + name + "</title></head><body style='margin:0;'>" +
                 "<iframe src='data:" + mime + ";base64," + base64 + "' style='width:100%;height:100vh;border:none;'></iframe>" +
                 "</body></html>";
      return HtmlService.createHtmlOutput(html)
        .setTitle(name)
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }
  } catch (err) {
    Logger.log("PROXY DRIVE FAIL: " + id + " - " + err.message + " dbg=" + dbg);
    return HtmlService.createHtmlOutput("<h1>File access error</h1><p>" + err.message + "</p><p>Debug: " + dbg + "</p>");
  }
}