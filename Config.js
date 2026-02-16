/******************** CONFIG ********************/
var CONFIG = {
  // STAGING SHEET ID
  SHEET_ID: "1F_aNZGmZwI9isQ1Qj1wjxY971XFkLmLJcz_bsugcCoY",
  DATA_SHEET: "FODE_Data",
  LOG_SHEET: "Webhook_Log",

  // Drive root (same as your baseline)
  ROOT_FOLDER_ID: "1vGD3DoOv1hlxYoTIfrNCZqAnrVKmghuB",
  YEAR_FOLDER: "2025",

  // ✅ Hardcode WEBAPP_URL (staging) to avoid getUrl()/deployment confusion
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
    "Accounting"
  ],

  // ✅ Do not overwrite Parent_Email; store corrected here
  PARENT_EMAIL_CORRECTED_HEADER: "Parent_Email_Corrected",

  // ✅ Keep student edits controlled (avoid “Additional Editable Fields” dumping everything)
  // Only allow these via portal:
  PORTAL_EDIT_FIELDS: [],

  // ✅ What students can SEE in Submitted Details (allowlist-only)
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

    // show human-friendly
    "Subjects_Selected_Canonical",
    "Physical_Exam_Site",

    // doc links OK for students to see
    "Birth_ID_Passport_File",
    "Latest_School_Report_File",
    "Transfer_Certificate_File",
    "Passport_Photo_File",
    "Fee_Receipt_File"
  ],

  // Document fields handled by Drive upload buttons
  DOCS: [
    { label: "Birth Certificate / NID / Passport", field: "Birth_ID_Passport_File", status: "Birth_ID_Status", comment: "Birth_ID_Comment" },
    { label: "Latest School Reports / Documents", field: "Latest_School_Report_File", status: "Report_Status", comment: "Report_Comment" },
    { label: "Transfer Certificate", field: "Transfer_Certificate_File", status: "Transfer_Status", comment: "Transfer_Comment" },
    { label: "Passport Size Colour Photo", field: "Passport_Photo_File", status: "Photo_Status", comment: "Photo_Comment" },
    { label: "Admission Fee Payment Receipt", field: "Fee_Receipt_File", status: "Receipt_Status", comment: "Receipt_Comment" }
  ]
};
