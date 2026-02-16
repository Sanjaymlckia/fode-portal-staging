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
  var action = clean_(payload.action || payload._action || "");

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

  return jsonOutput_({ status: "ok", ApplicantID: applicantId || "" });
}

/******************** ENTRYPOINT: GET ********************/
function doGet(e) {
  var view = String((e.parameter.view || "portal")).toLowerCase();
  var id = clean_(e.parameter.id || "");

  // accept email aliases
  var email = clean_(e.parameter.email || e.parameter.Parent_Email || e.parameter.parent_email || "").toLowerCase();

  if (!id || !email) {
    return (view === "json")
      ? jsonOutput_({ status: "error", message: "Missing id or email" })
      : htmlOutput_(renderErrorHtml_("Missing id or email"));
  }

  var ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  var sheet = mustGetSheet_(ss, CONFIG.DATA_SHEET);

  var found = findRowByIdEmail_(sheet, id, email);
  if (!found) {
    return (view === "json")
      ? jsonOutput_({ status: "error", message: "No matching record found" })
      : htmlOutput_(renderErrorHtml_("No matching record found"));
  }

  var record = found.record;
  record._PortalLocked = isPaymentVerified_(record);

  // prefill subjects from canonical OR raw Subjects_Selected
  var canonical = clean_(record.Subjects_Selected_Canonical || "");
  var fallbackCsv = subjectsToCsv_(record.Subjects_Selected || "");
  record._SubjectsCsv = canonical || fallbackCsv;

  if (view === "json") return jsonOutput_({ status: "ok", record: record });

  var examSites = getExamSites_(ss);

  return htmlOutput_(renderPortalHtml_({
    id: id,
    email: email,
    record: record,
    subjects: CONFIG.PORTAL_SUBJECTS,
    examSites: examSites,
    editFields: CONFIG.PORTAL_EDIT_FIELDS,
    docs: CONFIG.DOCS,
    visibleFields: CONFIG.PORTAL_VISIBLE_FIELDS
  }));
}

/******************** PORTAL UPDATE HANDLER ********************/
function handlePortalUpdate_(ss, dataSheet, logSheet, payload) {
  var id = clean_(payload.id || "");
  var emailFromLink = clean_(payload.email || "").toLowerCase();

  if (!id || !emailFromLink) return htmlOutput_(renderErrorHtml_("Missing id or email. Please reopen your portal link."));

  var found = findRowByIdEmail_(dataSheet, id, emailFromLink);
  if (!found) return htmlOutput_(renderErrorHtml_("No matching record found. Please reopen your portal link."));

  if (isPaymentVerified_(found.record)) {
    return htmlOutput_(renderErrorHtml_("Your record is locked because payment has been verified. No further changes are allowed."));
  }

  var updates = {};

  // Core fields from portal
  var dob = clean_(payload.Date_Of_Birth || "");
  if (dob) updates.Date_Of_Birth = dob;

  var examSite = clean_(payload.Physical_Exam_Site || "");
  if (examSite) updates.Physical_Exam_Site = examSite;

  // subjects comes from hidden packed field
  var subjectsCsv = clean_(payload.Subjects_Selected_Canonical || "");
  if (subjectsCsv) updates.Subjects_Selected_Canonical = subjectsCsv;

  // ✅ Do NOT overwrite Parent_Email (stable lookup key). Store corrected separately.
  var correctedEmail = clean_(payload.Parent_Email_Corrected || "").toLowerCase();
  if (correctedEmail) updates[CONFIG.PARENT_EMAIL_CORRECTED_HEADER] = correctedEmail;

  // (Optional extra editable fields - currently none)
  for (var i = 0; i < CONFIG.PORTAL_EDIT_FIELDS.length; i++) {
    var h = CONFIG.PORTAL_EDIT_FIELDS[i];
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

  var effectiveEmail =
    (updates[CONFIG.PARENT_EMAIL_CORRECTED_HEADER] ||
     clean_(found.record[CONFIG.PARENT_EMAIL_CORRECTED_HEADER] || "") ||
     clean_(found.record.Parent_Email || "")).toLowerCase();

  if (!effectiveEmail || !isValidEmail_(effectiveEmail)) missing.push("Valid Parent Email");

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
      email: emailFromLink,
      record: rec,
      subjects: CONFIG.PORTAL_SUBJECTS,
      examSites: examSites,
      editFields: CONFIG.PORTAL_EDIT_FIELDS,
      docs: CONFIG.DOCS,
      visibleFields: CONFIG.PORTAL_VISIBLE_FIELDS,
      error: "Please complete/fix: " + missing.join(", ")
    }));
  }

  updates.PortalLastUpdateAt = new Date().toISOString();

  // mark first submit time if empty
  if (!clean_(found.record.Portal_Submitted)) updates.Portal_Submitted = new Date().toISOString();

  writeBack_(dataSheet, found.rowNum, updates);
  log_(logSheet, "PORTAL UPDATE", "ApplicantID=" + id + " linkEmail=" + emailFromLink);

  return htmlOutput_(renderSuccessHtml_(id));
}

/******************** DRIVE UPLOAD (called via google.script.run) ********************/
function uploadPortalFile(applicantId, linkEmail, fieldName, fileName, mimeType, base64Data) {
  applicantId = clean_(applicantId);
  linkEmail = clean_(linkEmail).toLowerCase();
  fieldName = clean_(fieldName);
  fileName = clean_(fileName) || ("upload_" + Date.now());
  mimeType = clean_(mimeType) || "application/octet-stream";

  var ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  var sheet = mustGetSheet_(ss, CONFIG.DATA_SHEET);
  var logSheet = mustGetSheet_(ss, CONFIG.LOG_SHEET);

  var found = findRowByIdEmail_(sheet, applicantId, linkEmail);
  if (!found) throw new Error("Record not found.");

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

  // Create file from base64
  var bytes = Utilities.base64Decode(String(base64Data || ""));
  var blob = Utilities.newBlob(bytes, mimeType, fileName);
  var file = folder.createFile(blob);

  // Replace mode: overwrite field with new Drive URL, log old value
  var oldUrl = clean_(found.record[fieldName] || "");
  var newUrl = file.getUrl();

  var updates = {};
  updates[fieldName] = newUrl;
  updates.PortalLastUpdateAt = new Date().toISOString();
  if (!clean_(found.record.Portal_Submitted)) updates.Portal_Submitted = new Date().toISOString();

  // If status column exists, set PENDING_REVIEW
  var docMeta = docMetaByField_(fieldName);
  if (docMeta && hasHeader_(sheet, docMeta.status)) updates[docMeta.status] = "PENDING_REVIEW";

  // File_Log append
  var line = new Date().toISOString() + " | " + fieldName + " | replaced | old=" + (oldUrl || "-") + " | new=" + newUrl;
  updates.File_Log = appendLog_(clean_(found.record.File_Log || ""), line);

  writeBack_(sheet, found.rowNum, updates);
  log_(logSheet, "PORTAL UPLOAD", "ApplicantID=" + applicantId + " field=" + fieldName + " file=" + fileName);

  return { ok: true, field: fieldName, url: newUrl };
}

/******************** PORTAL HTML ********************/
function renderPortalHtml_(opts) {
  var id = opts.id, email = opts.email, record = opts.record;
  var subjects = opts.subjects || [];
  var examSites = opts.examSites || [];
  var editFields = opts.editFields || [];
  var docs = opts.docs || [];
  var visibleFields = opts.visibleFields || [];
  var error = opts.error || "";

  var locked = record._PortalLocked === true;
  var dis = locked ? "disabled" : "";

  // subject selections: canonical preferred, else fallback
  var csv = clean_(record.Subjects_Selected_Canonical || record._SubjectsCsv || "");
  var selected = parseSubjects_(csv);

  // date input expects yyyy-mm-dd; if your sheet stores dd/mm/yyyy, keep it blank rather than breaking
  var dobVal = esc_(clean_(record.Date_Of_Birth || ""));
  if (dobVal && dobVal.indexOf("/") !== -1) dobVal = ""; // avoid invalid date showing wrong

  var examVal = clean_(record.Physical_Exam_Site || "");
  var effectiveEmail = clean_(record[CONFIG.PARENT_EMAIL_CORRECTED_HEADER] || record.Parent_Email || "");
  var emailVal = esc_(effectiveEmail);

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

  // allowlist-only summary
  var summaryHtml = renderAllowlistSummary_(record, visibleFields);

  // editable fields UI
  var extraInputs = renderEditableFields_(record, editFields, dis);

  // docs upload UI
  var docsHtml = renderDocsSection_(id, email, record, docs, locked);

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
    + "<div><b>Link Email (locked for lookup):</b> " + esc_(email) + "</div>"
    + "</div>"
    + lockedBlock
    + errorBlock

    + '<div style="padding:12px;border:1px solid #ddd;border-radius:10px;margin-bottom:16px;">'
    + '<h3 style="margin-top:0;">Submitted Details (read-only)</h3>'
    + summaryHtml
    + "</div>"

    + '<div style="padding:12px;border:1px solid #ddd;border-radius:10px;margin-bottom:16px;">'
    + '<h3 style="margin-top:0;">Documents & Payment Proof</h3>'
    + docsHtml
    + "</div>"

    // ✅ hardcoded action URL to prevent blank screen / doPost not firing
    + '<form method="post" action="' + CONFIG.WEBAPP_URL + '" onsubmit="return packSubjects();"'
    + ' style="padding:12px;border:1px solid #ddd;border-radius:10px;">'
    + '<input type="hidden" name="action" value="portal_update" />'
    + '<input type="hidden" name="id" value="' + esc_(id) + '" />'
    + '<input type="hidden" name="email" value="' + esc_(email) + '" />'
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

    + '<div style="margin:12px 0;">'
    + "<label><b>Correct Parent Email (mandatory):</b></label><br/>"
    + '<input type="email" name="Parent_Email_Corrected" value="' + emailVal + '" style="padding:8px;width:520px;" ' + dis + " />"
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

function renderDocsSection_(id, email, record, docs, locked) {
  var out = "";
  for (var i = 0; i < docs.length; i++) {
    var d = docs[i];
    var cur = clean_(record[d.field] || "");
    var st = clean_(record[d.status] || "");
    var cm = clean_(record[d.comment] || "");

    var stBadge = st ? ("<b>Status:</b> " + esc_(st)) : "<b>Status:</b> (not set)";
    var cmBlock = cm ? ("<div style='margin-top:6px;'><b>Admin comment:</b> " + esc_(cm) + "</div>") : "";

    var isUrl = /^https?:\/\//i.test(cur);
    var curLink = isUrl
      ? ("<div style='margin-top:6px;'><b>Current file:</b> <a target='_blank' href='" + esc_(cur) + "'>Open</a></div>")
      : "<div style='margin-top:6px;'><b>Current file:</b> (none)</div>";

    var uploadUi = locked
      ? "<div style='margin-top:10px;color:#666;'><i>Uploads disabled (locked).</i></div>"
      : "<div style='margin-top:10px;'>"
        + "<input type='file' id='f_" + esc_(d.field) + "' /> "
        + "<button type='button' onclick=\"uploadDoc('" + esc_(d.field) + "')\">Upload / Replace</button>"
        + "<div id='msg_" + esc_(d.field) + "' style='margin-top:6px;font-size:12px;'></div>"
        + "</div>";

    out += ""
      + "<div style='padding:10px;border:1px solid #eee;border-radius:10px;margin:10px 0;'>"
      + "<div><b>" + esc_(d.label) + "</b></div>"
      + "<div style='margin-top:6px;'>" + stBadge + "</div>"
      + cmBlock
      + curLink
      + uploadUi
      + "</div>";
  }

  // uploader script
  out += ""
    + "<script>"
    + "function uploadDoc(fieldName){"
    + "  var input=document.getElementById('f_'+fieldName);"
    + "  var msg=document.getElementById('msg_'+fieldName);"
    + "  if(!input || !input.files || !input.files.length){ msg.innerHTML='Select a file first.'; return; }"
    + "  var file=input.files[0];"
    + "  msg.innerHTML='Uploading...';"
    + "  var reader=new FileReader();"
    + "  reader.onload=function(e){"
    + "    var data=e.target.result || '';"
    + "    var base64=data.split(',').pop();"
    + "    google.script.run"
    + "      .withSuccessHandler(function(res){"
    + "        msg.innerHTML='Uploaded. Refresh page to see updated link.';"
    + "      })"
    + "      .withFailureHandler(function(err){"
    + "        msg.innerHTML='Upload failed: '+(err && err.message ? err.message : err);"
    + "      })"
    + "      .uploadPortalFile('" + esc_(id) + "','" + esc_(email) + "', fieldName, file.name, file.type, base64);"
    + "  };"
    + "  reader.readAsDataURL(file);"
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
    out += "<div style='margin:10px 0;'>"
      + "<label><b>" + esc_(h) + ":</b></label><br/>"
      + "<input type='text' name='field_" + esc_(h) + "' value='" + esc_(val) + "' style='padding:8px;width:520px;' " + dis + " />"
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

function renderSuccessHtml_(applicantId) {
  return '<!doctype html><html><body style="font-family:Arial;max-width:780px;margin:24px auto;padding:0 16px;">'
    + "<h2>FODE Student Portal</h2>"
    + '<div style="background:#e8fff0;border:1px solid #9be7b2;padding:12px;border-radius:10px;">'
    + "Updates saved successfully for <b>" + esc_(applicantId) + "</b>."
    + "</div></body></html>";
}

function htmlOutput_(html) {
  return HtmlService.createHtmlOutput(html)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/******************** LOOKUP ********************/
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
    "Physical_Exam_Site",
    "Subjects_Selected_Canonical",
    CONFIG.PARENT_EMAIL_CORRECTED_HEADER,
    "File_Log"
  ];

  for (var i = 0; i < CONFIG.DOCS.length; i++) {
    meta.push(CONFIG.DOCS[i].status);
    meta.push(CONFIG.DOCS[i].comment);
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
  for (var i = 0; i < CONFIG.DOCS.length; i++) {
    if (CONFIG.DOCS[i].field === fieldName) return CONFIG.DOCS[i];
  }
  return null;
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

