MANDATORY: clasp push -> clasp version -> clasp deploy (in-place)

After implementing any CIS that changes project source files, Codex MUST:
- Run: clasp.cmd push
- Then run: clasp.cmd version "<CIS title / short description>"
- Then run: clasp.cmd deploy -i <STUDENT_DEPLOYMENT_ID> --versionNumber <NUMERIC_VERSION> --description "Student webapp (staging) - <CIS #>"

Codex MUST treat the deployment ID as stable and MUST deploy in place using -i <DEPLOYMENT_ID>.

Codex MUST NOT attempt to deploy using a text label for --versionNumber. It MUST use the numeric version returned by clasp.cmd version.

Codex MUST paste full command outputs into its completion note, including:
- push result (files pushed or "Skipping push")
- "Created version <n>"
- "Deployed <deploymentId> @<n>"

If clasp prompts:
"Manifest file has been updated. Do you want to push and overwrite?"
Codex MUST answer "Yes" and continue.

If clasp push fails (HTTP/API error), Codex MUST stop and report the exact error text and the command that failed. Codex MUST NOT proceed to version/deploy after a failed push.

If clasp push reports "Skipping push", Codex MUST still:
- Run clasp.cmd status
- Report whether there are any local diffs (git status summary if available)
- Only proceed to version/deploy if it can confirm server state already includes the intended changes (or if there are no code changes in the CIS).

Repository constants:
- STUDENT_DEPLOYMENT_ID: AKfycby_AgQDFHyKxT5WV9O230By9w6R-kiTIJe_aui1a-WlZLnuJQ-I7Xh4VDFb1oe1m2LN
- STUDENT_EXEC_BASE_URL: https://script.google.com/macros/s/AKfycby_AgQDFHyKxT5WV9O230By9w6R-kiTIJe_aui1a-WlZLnuJQ-I7Xh4VDFb1oe1m2LN/exec
- Default deploy description pattern: "Student webapp (staging) - CIS-<n>"

MANDATORY: Test URL Output
- After deploy succeeds, Codex MUST print a canonical Student Portal test URL using STUDENT_EXEC_BASE_URL.
- Format: ${STUDENT_EXEC_BASE_URL}?view=portal&id=<ApplicantID>&s=<Secret>
- If <ApplicantID> and <Secret> are provided in the CIS or user message, Codex MUST output a fully populated, ready-to-click URL.
- If they are not provided, Codex MUST output:
- Base URL: ${STUDENT_EXEC_BASE_URL}?view=portal
- A one-line note stating it requires id=<ApplicantID> and s=<Secret> from the portal token.
- Codex MUST NOT output the /a/<domain>/ URL variant as the primary test URL.

Release Checklist (must be included at end of every completion note):
- [ ] clasp.cmd push succeeded
- [ ] clasp.cmd version created numeric version
- [ ] clasp.cmd deploy updated the correct deployment ID
- [ ] Canonical Student Portal test URL printed (non-/a/ host)
- [ ] Tested URL recorded (if provided by user)

Only edit: Admin.js, AdminUI.html, Code.js, Config.js, Utils.js, Routes.js

Never edit: appsscript.json, .clasp.json (unless explicitly requested)

Never add sheet headers/columns without explicit instruction

Never store plaintext portal secrets in sheet; only PortalTokenHash + PortalTokenIssuedAt

Preserve lock rules: Payment_Verified==Yes or Portal_Access_Status==Locked

Preserve Parent_Email non-overwrite rule; use Parent_Email_Corrected

Prefer minimal diffs; avoid refactors

Always return: (1) diff summary (2) manual test checklist
