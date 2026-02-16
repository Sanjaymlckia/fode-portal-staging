/******************** UTIL ********************/
function mustGetSheet_(ss, name) {
  var sh = ss.getSheetByName(name);
  if (!sh) throw new Error("Missing sheet: " + name);
  return sh;
}

function clean_(v) { return (v === null || v === undefined) ? "" : String(v).trim(); }

function normalize_(v) {
  if (v === null || v === undefined) return "";
  if (typeof v === "object") return JSON.stringify(v);
  return String(v);
}

function slug_(s) { return clean_(s).toLowerCase().replace(/[^a-z0-9]+/g, "_"); }

function log_(sheet, label, msg) { sheet.appendRow([new Date(), label, msg || ""]); }

function payloadSummary_(p) {
  return JSON.stringify({
    First_Name: p.First_Name,
    Last_Name: p.Last_Name,
    Grade: p.Grade_Applying_For,
    Intake: p.Intake_Year || p["Intake Year"] || ""
  });
}

function esc_(s) {
  return String(s === null || s === undefined ? "" : s)
    .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;").replace(/'/g, "&#039;");
}

function shallowCopy_(obj) { var out = {}; for (var k in obj) out[k] = obj[k]; return out; }

function isValidEmail_(email) { return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(clean_(email)); }
