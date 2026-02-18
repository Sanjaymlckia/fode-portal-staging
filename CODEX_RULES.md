Only edit: Admin.js, AdminUI.html, Code.js, Config.js, Utils.js, Routes.js

Never edit: appsscript.json, .clasp.json (unless explicitly requested)

Never add sheet headers/columns without explicit instruction

Never store plaintext portal secrets in sheet; only PortalTokenHash + PortalTokenIssuedAt

Preserve lock rules: Payment_Verified==Yes or Portal_Access_Status==Locked

Preserve Parent_Email non-overwrite rule; use Parent_Email_Corrected

Prefer minimal diffs; avoid refactors

Always return: (1) diff summary (2) manual test checklist