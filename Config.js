/******************** CONFIG ********************/
/**
 * IMPORTANT:
 * - Keep CONFIG as a single global var.
 * - Do NOT redeclare CONFIG in any other file.
 * - Do NOT add trailing junk text after properties.
 */
var CONFIG = {
  // Versioning (change control)
  VERSION: "2026-03-01-PNG-STAGING-r160",
  CHANGELOG_LAST: "r160: admin UI polish – removed redundant accordion; added payment badges (UI only)",
  BUILD_LABEL: "CIS-32 runtime diag + upload timing + canonical student URL + 5MB upload safeguards",

  // STAGING/PROD data mode routing (working sheet)
  DATA_MODE: "STAGING", // "STAGING" | "PROD"
  SPREADSHEET_ID_STAGING: "1F_aNZGmZwI9isQ1Qj1wjxY971XFkLmLJcz_bsugcCoY",
  SHEET_ID_STAGING: "1F_aNZGmZwI9isQ1Qj1wjxY971XFkLmLJcz_bsugcCoY",
  SHEET_ID_PROD: "1fHmeGNmpOj9PEPQ5Fp4tUyCP4UdH70lltukraD4SalU",
  SHEET_NAME_WORKING: "FODE_Data",
  SHEET_TAB_WORKING: "FODE_Data",
  // Backward-compatible aliases (legacy callers)
  SHEET_ID: "1F_aNZGmZwI9isQ1Qj1wjxY971XFkLmLJcz_bsugcCoY",
  DATA_SHEET: "FODE_Data",
  LOG_SHEET: "Webhook_Log",

  // PORTAL LOG spreadsheet (FODE Portal Log 2026)
  LOG_SHEET_ID: "1AQbkHUafLFxqHDqwH3dVHR8gTuOZYtyUPkheby5ejhU",
  LOG_SHEET_NAME: "Submissions",

  // Drive root (baseline)
  ROOT_FOLDER_ID: "1vGD3DoOv1hlxYoTIfrNCZqAnrVKmghuB",
  APPLICANT_ROOT_FOLDER_ID_PRIMARY: "1vGD3DoOv1hlxYoTIfrNCZqAnrVKmghuB",
  APPLICANT_ROOT_FOLDER_ID_FALLBACK: "",
  SCRIPT_PROP_UPLOAD_ROOT_ID: "FODE_UPLOAD_ROOT_ID",
  AUTO_UPLOAD_ROOT_NAME: "FODE Upload Root (Auto)",
  AUTO_UPLOAD_ROOT_ENABLED: false,
  APPLICANT_ROOT_YEAR_FOLDER_NAME: "2025",
  DRIVE_REST_FALLBACK_ENABLED: true,
  PORTAL_UPLOAD_PREFER_REST: true,
  PORTAL_UPLOAD_DRIVEAPP_RETRY_ENABLED: false,
  PORTAL_UPLOAD_MAX_SERVER_MS: 20000,
  PORTAL_MAX_UPLOAD_BYTES: 5 * 1024 * 1024,
  PORTAL_ALLOWED_MIME: [
    "application/pdf",
    "image/jpeg",
    "image/png",
    "application/msword",
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
  ],
  PORTAL_ALLOWED_EXT: ["pdf", "jpg", "jpeg", "png", "doc", "docx"],
  DRIVE_API_BASE: "https://www.googleapis.com/drive/v3",
  DRIVE_UPLOAD_BASE: "https://www.googleapis.com/upload/drive/v3",
  DRIVE_FIELDS_FOLDER: "id,name,webViewLink",
  DRIVE_FIELDS_FILE: "id,name,webViewLink,parents",
  YEAR_FOLDER: "2025",
  
  // Portal Secrets File
  PORTAL_SECRETS_SHEET_ID: "1HEJPtSov-iE5YTpSWWZ89YLIQAw4Eju9DDMG46HkTRc",
  PORTAL_SECRETS_TAB: "PortalSecrets",
  


  // Web app URLs
  // - WEBAPP_URL_ADMIN: admin deployment (/exec), executeAs=USER_ACCESSING, access=DOMAIN
  // - WEBAPP_URL_STUDENT: student deployment (/exec), executeAs=ME, access=ANYONE/ANYONE_ANONYMOUS
  // WEBAPP_URL is kept for backward compatibility and mirrors WEBAPP_URL_ADMIN.
  WEBAPP_URL_ADMIN: "https://script.google.com/macros/s/AKfycbzcL4sXLW2mEPg5ADA5YS16m2Avcd4RxnLp-vKn45_sXqgtdW9AP_lsuGImyP3y1U3k/exec",
  WEBAPP_URL: "https://script.google.com/macros/s/AKfycbzcL4sXLW2mEPg5ADA5YS16m2Avcd4RxnLp-vKn45_sXqgtdW9AP_lsuGImyP3y1U3k/exec",
  WEBAPP_URL_STUDENT: "https://script.google.com/macros/s/AKfycby_AgQDFHyKxT5WV9O230By9w6R-kiTIJe_aui1a-WlZLnuJQ-I7Xh4VDFb1oe1m2LN/exec",
  WEBAPP_URL_STUDENT_EXEC: "https://script.google.com/macros/s/AKfycby_AgQDFHyKxT5WV9O230By9w6R-kiTIJe_aui1a-WlZLnuJQ-I7Xh4VDFb1oe1m2LN/exec",
  SCRIPT_ID: "1NwyKK4NV27SD7WaXlpgn_1cCMJ9ezX0erLjvMBB6DDGq9VDt12y5qtcK",
  DEPLOYMENT_ID_STUDENT: "AKfycby_AgQDFHyKxT5WV9O230By9w6R-kiTIJe_aui1a-WlZLnuJQ-I7Xh4VDFb1oe1m2LN",
  ZOHO_API_BASE: "https://www.zohoapis.com/crm/v2",
  ZOHO_OAUTH_BASE: "https://accounts.zoho.com/oauth/v2",
  DEAL_STAGE: "Qualification",
  DEAL_DUPLICATE_FIELD: "FormID",
  CRM_PIPELINE_FODE: "FODE Admissions",
  CRM_STAGE_PAYMENT_VERIFIED: "Payment Verified",

  FEE_REGISTRATION_KINA: 600,
  FEE_PER_SUBJECT_KINA: 450,
  EMAIL_ADMIN_ALERTS_TO: "fode@kundu.ac",
  EMAIL_ENABLE_PAYMENT_VERIFIED_TRIGGERS: true,
  EMAIL_RELEASE_ADMIN_TO: "sanja@minervacenters.com,ict@minervacenters.com",
  EMAIL_FROM_NAME: "FODE Admissions",
  EMAIL_PROVIDER: "MAILAPP", // "MAILAPP" | "GMAILAPP"
  EMAIL_SENDER_MODE: "ALIAS", // "DEFAULT" | "ALIAS"
  EMAIL_FROM_ADDRESS: "fode@kundu.ac",
  EMAIL_REPLY_TO: "fode@kundu.ac",
  EMAIL_STUDENT_SUBJECT_PAYMENT_VERIFIED: "FODE Application - Payment Verified | Next Steps & Bank Details",
  EMAIL_ADMIN_SUBJECT_PAYMENT_VERIFIED: "FODE - Payment Verified | Release Access for ApplicantID",
  PAYVER_STUDENT_SUBJECT: "Payment Verified - FODE",
  PAYVER_ADMIN_SUBJECT: "Payment Verified - Release Access",
  CRM_PUSH_DRY_RUN: true,
  FODE_FEE_BASE_K: 600,
  FODE_FEE_PER_SUBJECT_K: 450,
  FODE_BANK_DETAILS_TEXT: "<paste canonical bank details block here>",
  FODE_NEXT_STEPS_TEXT: "<paste canonical next steps here>",
  DOCS_FOLLOWUP_ENABLE: true,
  DOCS_FOLLOWUP_EMAIL_SUBJECT: "FODE Application - Documents Verified | Quote, Payment Instructions & Next Steps",
  DOCS_FOLLOWUP_REPLY_TO: "fode@kundu.ac",
  DOCS_FOLLOWUP_CC: "fode_kia@kundu.ac",
  DOCS_FOLLOWUP_SENDER_MODE: "ALIAS",
  DOCS_FOLLOWUP_COOLDOWN_MS: 300000,
  ADMIN_THEME_BLUE_MODE: "IMPROVED_V2",
  DOCS_FOLLOWUP_SHOW_IN_QUEUES: true,
  QUOTE_EMAIL_ENABLED: true,
  INVOICE_TRIGGER_ENABLED: false,
  INVOICE_TRIGGER_MODE: "LOG_ONLY", // "LOG_ONLY" | "WEBHOOK"
  INVOICE_WEBHOOK_URL: "",
  INTERNAL_FINANCE_EMAILS: [
    "sanjay@minervacenters.com"
  ],
  PAYMENT_INSTRUCTIONS_TEXT: "<BANK DETAILS MODE A: literal bank instructions text>",

  // Deployment model:
  // - Admin web app: executeAs=USER_ACCESSING, access=DOMAIN
  // - Student web app: executeAs=ME, access=ANYONE or ANYONE_ANONYMOUS

  // Admin allowlist (must be lowercase emails)
  ADMIN_EMAILS: [
    "sanjay@minervacenters.com",
    "enquiries@kundu.ac",
    "fode_kia@kundu.ac",
    "mlc@minervacenters.com",
    "operations@minervacenters.com",
    "mlccorporate@minervacenters.com"
  ],
  ADMIN_ROLES: {
    "sanjay@minervacenters.com": "SUPER",
    "enquiries@kundu.ac": "VERIFIER",
    "fode_kia@kundu.ac": "VERIFIER",
    "mlc@minervacenters.com": "VERIFIER",
    "operations@minervacenters.com": "VERIFIER",
    "mlccorporate@minervacenters.com": "VERIFIER"
  },
  SUPER_ADMIN_EMAILS: [
    "sanjay@minervacenters.com"
  ],
  SUPERADMIN_ALLOW_EMAIL_OVERRIDE_POST_DOCS_VERIFIED: true,
  ELEVATED_OVERRIDE_EMAILS: [
    // Super Admin may add delegate emails here
  ],

  // ApplicantID
  APPLICANT_ID_HEADER: "ApplicantID",
  APPLICANT_PREFIX: "FODE-26-",
  APPLICANT_DIGITS: 6,

  // Exam sites
  EXAM_SITES_SHEET: "Exam_Sites",

  // Subjects shown (portal checkbox list) — Agriculture removed
  PORTAL_SUBJECTS: [
    "English","Mathematics","Biology","Chemistry","Physics","History","Geography",
    "Economics","ICT","Business Studies","Personal Development","Science","Social Science",
    "Accounting"
  ],

  // Do not overwrite Parent_Email
  PARENT_EMAIL_CORRECTED_HEADER: "Parent_Email_Corrected",

  // Editable fields (controlled)
  PORTAL_EDIT_FIELDS: [
    "Home_Address",
    "Parent_Phone",
    "Travel_Mode",
    "Prev_School_Name",
    "Prev_School_Grade",
    "Reason_For_Transfer",
    "Siblings_Name_Grade"
  ],
  PORTAL_NON_EDIT_FIELDS: ["ApplicantID","First_Name","Last_Name"],
  PORTAL_EDIT_EXCLUDE_FIELDS: [
    "Travel_Mode",
    "Program",
    "Program_Applied_For",
    "Type",
    "Physical_Exam_Site",
    "Subjects_Selected_Canonical",
    "Parent_Email_Corrected",
    "Birth_ID_Passport_File",
    "Latest_School_Report_File",
    "Transfer_Certificate_File",
    "Passport_Photo_File",
    "Fee_Receipt_File"
  ],
  PORTAL_EDIT_MODE: "ALL_VISIBLE_EXCEPT_NON_EDIT",
  PORTAL_TOKEN_MAX_AGE_DAYS: 90,
  DEBUG_PORTAL_POST: true,
  DEBUG_PORTAL_SHOW_ON_PAGE: true,
  PORTAL_DEBUG_TO_LOG_SHEET: true,
  PORTAL_DEBUG_MAX_BYTES: 2500,
  PORTAL_UPLOAD_TIMEOUT_MS: 25000,
  PORTAL_UPLOAD_MAX_MB: 5,
  DIAG_RUNTIME: true,
  DIAG_UPLOAD_VERBOSE: true,
  DIAG_UPLOAD_SMOKE_ENABLED: true,
  PORTAL_UPLOAD_KEYS: [
    "Birth_ID_Passport_File",
    "Latest_School_Report_File",
    "Transfer_Certificate_File",
    "Passport_Photo_File",
    "Fee_Receipt_File"
  ],

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

  // Allowed admin verification statuses (keys, not values)
  DOC_STATUS: {
    PENDING_REVIEW: true,
    VERIFIED: true,
    REJECTED: true,
    FRAUDULENT: true
  },

  BRAND: {
    name: "Kundu FODE",
    bg: "#0b1220",
    card: "#111a2e",
    text: "#e5e7eb",
    muted: "#94a3b8",
    accent: "#0ea5a4",
    danger: "#ef4444",
    warn: "#f59e0b",
    ok: "#22c55e"
  },

  DOC_FIELDS: [
    { label: "Birth Certificate / NID / Passport", file: "Birth_ID_Passport_File", status: "Birth_ID_Status", comment: "Birth_ID_Comment", required: true, multiple: false },
    { label: "Latest School Reports / Documents", file: "Latest_School_Report_File", status: "Report_Status", comment: "Report_Comment", required: true, multiple: true },
    { label: "Transfer Certificate (optional)", file: "Transfer_Certificate_File", status: "Transfer_Status", comment: "Transfer_Comment", required: false, multiple: false },
    { label: "Passport Size Colour Photo", file: "Passport_Photo_File", status: "Photo_Status", comment: "Photo_Comment", required: true, multiple: false },
    { label: "Admission Fee Payment Receipt", file: "Fee_Receipt_File", status: "Receipt_Status", comment: "Receipt_Comment", required: true, multiple: false }
  ],


  // Optional: list of doc fields that are NOT required for Docs_Verified
  OPTIONAL_DOC_FIELDS: ["Transfer_Certificate_File"]
};


/******************** SCHEMA (AUTHORITATIVE HEADERS) ********************/
/**
 * Single source of truth for Sheet header names.
 * Use SCHEMA.* everywhere instead of hardcoding header strings.
 */
var SCHEMA = {
  // Identity / lookup
  APPLICANT_ID: CONFIG.APPLICANT_ID_HEADER,                      // "ApplicantID"
  PARENT_EMAIL: "Parent_Email",                                  // must match your sheet header
  PARENT_EMAIL_CORRECTED: CONFIG.PARENT_EMAIL_CORRECTED_HEADER,  // "Parent_Email_Corrected"

  // Portal meta / logs
  FOLDER_URL: "Folder_Url",
  FILE_LOG: "File_Log",
  PORTAL_LAST_UPDATE_AT: "PortalLastUpdateAt",
  PORTAL_SUBMITTED: "Portal_Submitted",
  PORTAL_TOKEN_HASH: "PortalTokenHash",
  PORTAL_TOKEN_ISSUED_AT: "PortalTokenIssuedAt",

  // Subjects + site
  SUBJECTS_CANONICAL: "Subjects_Selected_Canonical",
  PHYSICAL_EXAM_SITE: "Physical_Exam_Site",

  // Docs verification (minimal audit model)
  DOCS_VERIFIED: "Docs_Verified",             // "Yes" when required docs VERIFIED
  DOC_VERIFICATION_STATUS: "Doc_Verification_Status",
  PORTAL_ACCESS_STATUS: "Portal_Access_Status",
  PAYMENT_VERIFIED: "Payment_Verified",
  DOC_LAST_VERIFIED_AT: "Doc_Last_Verified_At",
  DOC_LAST_VERIFIED_BY: "Doc_Last_Verified_By",

  // Optional existing rollups
  VERIFIED_BY: "Verified_By",
  VERIFIED_AT: "Verified_At"
};
