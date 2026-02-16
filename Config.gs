/******************** CONFIG ********************/
/**
 * IMPORTANT:
 * - Keep CONFIG as a single global var.
 * - Do NOT redeclare CONFIG in any other file.
 * - Do NOT add trailing junk text after properties.
 */
var CONFIG = {
  // Versioning (change control)
  VERSION: "2026-02-17-PNG-STAGING",
  CHANGELOG_LAST: "Phase0+1: add VERSION, SCHEMA, smoke test, and Sheet IO primitives",

  // STAGING spreadsheet (FODE_Data + Webhook_Log)
  SHEET_ID: "1F_aNZGmZwI9isQ1Qj1wjxY971XFkLmLJcz_bsugcCoY",
  DATA_SHEET: "FODE_Data",
  LOG_SHEET: "Webhook_Log",

  // PORTAL LOG spreadsheet (FODE Portal Log 2026)
  LOG_SHEET_ID: "1AQbkHUafLFxqHDqwH3dVHR8gTuOZYtyUPkheby5ejhU",
  LOG_SHEET_NAME: "Submissions",

  // Drive root (baseline)
  ROOT_FOLDER_ID: "1vGD3DoOv1hlxYoTIfrNCZqAnrVKmghuB",
  YEAR_FOLDER: "2025",

  // Web app URL (staging deployment)
  WEBAPP_URL: "https://script.google.com/macros/s/AKfycbzcL4sXLW2mEPg5ADA5YS16m2Avcd4RxnLp-vKn45_sXqgtdW9AP_lsuGImyP3y1U3k/exec",

  // ApplicantID
  APPLICANT_ID_HEADER: "ApplicantID",
  APPLICANT_PREFIX: "FODE-26-",
  APPLICANT_DIGITS: 6,

  // Exam sites
  EXAM_SITES_SHEET: "Exam_Sites",

  // Subjects shown (portal checkbox list)
  PORTAL_SUBJECTS: [
    "English","Mathematics","Biology","Chemistry","Physics","History","Geography",
    "Economics","ICT","Business Studies","Personal Development","Science","Social Science",
    "Accounting","Agriculture"
  ],

  // Do not overwrite Parent_Email
  PARENT_EMAIL_CORRECTED_HEADER: "Parent_Email_Corrected",

  // Editable fields (controlled)
  PORTAL_EDIT_FIELDS: [],

  // What students can see (allowlist)
  PORTAL_VISIBLE_FIELDS: [
    "ApplicantID",
    "First_Name","Last_Name","Gender","Date_Of_Birth",
    "Grade_Applying_For","Upgrade_Grade_Stream",
    "Siblings_Name_Grade",
    "Prev_School_Name","Prev_School_Grade","Reason_For_Transfer",
    "Country_Of_Birth","Province_Of_Birth","Citizenship","Mother_Tongue",
    "Home_Address","Travel_Mode",
    "Parent_Full_Name","Relationship_To_Student","Parent_Phone","Parent_Email","Parent_Email_Corrected",
    "Program","Program_Applied_For","Intake_Year","Type",
    "Subjects_Selected_Canonical",
    "Physical_Exam_Site",
    "Birth_ID_Passport_File",
    "Latest_School_Report_File",
    "Transfer_Certificate_File",
    "Passport_Photo_File",
    "Fee_Receipt_File"
  ],

  // Document fields (Drive upload buttons)
  DOCS: [
    { label: "Birth Certificate / NID / Passport", field: "Birth_ID_Passport_File", status: "Birth_ID_Status", comment: "Birth_ID_Comment" },
    { label: "Latest School Reports / Documents", field: "Latest_School_Report_File", status: "Report_Status", comment: "Report_Comment" },
    { label: "Transfer Certificate", field: "Transfer_Certificate_File", status: "Transfer_Status", comment: "Transfer_Comment" },
    { label: "Passport Size Colour Photo", field: "Passport_Photo_File", status: "Photo_Status", comment: "Photo_Comment" },
    { label: "Admission Fee Payment Receipt", field: "Fee_Receipt_File", status: "Receipt_Status", comment: "Receipt_Comment" }
  ]
};


/******************** SCHEMA (AUTHORITATIVE HEADERS) ********************/
/**
 * Single source of truth for Sheet header names.
 * Use SCHEMA.* everywhere instead of hardcoding header strings.
 */
var SCHEMA = {
  // Identity / lookup
  APPLICANT_ID: CONFIG.APPLICANT_ID_HEADER,                 // "ApplicantID"
  PARENT_EMAIL: "Parent_Email",                             // exact header in sheet
  PARENT_EMAIL_CORRECTED: CONFIG.PARENT_EMAIL_CORRECTED_HEADER, // "Parent_Email_Corrected"

  // Portal meta / logs
  FOLDER_URL: "Folder_Url",
  FILE_LOG: "File_Log",
  PORTAL_LAST_UPDATE_AT: "PortalLastUpdateAt",
  PORTAL_SUBMITTED: "Portal_Submitted",

  // Subjects + site
  SUBJECTS_CANONICAL: "Subjects_Selected_Canonical",
  PHYSICAL_EXAM_SITE: "Physical_Exam_Site",

  // Verification / document review (reserved for Phase 3)
  DOC_VERIFIED: "Docs_Verified",              // matches your header
  DOC_LAST_VERIFIED_AT: "Doc_Last_Verified_At",
  DOC_VERIFIED_BY: "Doc_Last_Verified_By"
};
