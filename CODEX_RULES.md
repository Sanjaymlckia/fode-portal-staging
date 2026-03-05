# CODEX_RULES.md — FODE Portal Codex Discipline
Last updated: 2026-03-05 (post CIS-116 / runtime self-ID, pinned @176)

---------------------------------------------------------------------

MANDATORY: clasp push -> clasp version -> clasp deploy (in-place)

After implementing any CIS that changes project source files, Codex MUST:

1) Run
clasp.cmd push

2) Then run
clasp.cmd version "<CIS title / short description>"

3) Then run
clasp.cmd deploy -i <DEPLOYMENT_ID> --versionNumber <NUMERIC_VERSION> --description "<deployment description>"

Codex MUST treat deployment IDs as stable and MUST deploy in-place using:

-i <DEPLOYMENT_ID>

Codex MUST NOT attempt to deploy using a text label for --versionNumber.
It MUST use the numeric version returned by clasp.cmd version.

---------------------------------------------------------------------

MANDATORY: Command Output Logging

Codex MUST paste full command outputs into its completion note, including:

- push result ("files pushed" OR "Skipping push")
- "Created version <n>"
- "Deployed <deploymentId> @<n>"

If clasp prompts:

"Manifest file has been updated. Do you want to push and overwrite?"

Codex MUST answer:

Yes

and continue.

---------------------------------------------------------------------

MANDATORY: Push Failure Handling

If clasp push fails (HTTP/API error):

Codex MUST
- stop immediately
- report the exact error text
- report the command that failed

Codex MUST NOT proceed to version/deploy after a failed push.

---------------------------------------------------------------------

MANDATORY: Handling "Skipping push"

If clasp push reports "Skipping push":

Codex MUST still run:

clasp.cmd status

and report:

- whether local files differ
- git status summary if available

Codex may proceed to version/deploy ONLY if:

- there are no code changes
OR
- server state already includes intended changes.

---------------------------------------------------------------------

REPOSITORY CONSTANTS

Admin Deployment ID
AKfycbwUaqGym6dZSfqltKuI3sX4V31ijl9AOy9bDTPMzqmDToZVjMZc-xezclftg1EIJ8Tx

Student Deployment ID
AKfycbx2ve4bfCEofF_pJnra-UR02BaoumJaUeDS19Amftm2con2e7ggblMfHRzcn6fYAC4g

Canonical Admin Exec Base URL
https://script.google.com/macros/s/AKfycbwUaqGym6dZSfqltKuI3sX4V31ijl9AOy9bDTPMzqmDToZVjMZc-xezclftg1EIJ8Tx/exec

Canonical Student Exec Base URL
https://script.google.com/macros/s/AKfycbx2ve4bfCEofF_pJnra-UR02BaoumJaUeDS19Amftm2con2e7ggblMfHRzcn6fYAC4g/exec

Default deploy description pattern:
Student webapp (staging) - CIS-<n>

---------------------------------------------------------------------

MANDATORY: Canonical URL Rule (Critical)

All generated portal links MUST use the domain-neutral format:

https://script.google.com/macros/s/<DEPLOYMENT_ID>/exec

Never generate:

/a/macros/<domain>/...

All URL generation MUST go through:

canonicalExecBase_(deploymentIdOrExecUrl)

which must return:

https://script.google.com/macros/s/<id>/exec

---------------------------------------------------------------------

MANDATORY: Runtime Self-Identification

Every deployment must expose runtime identity.

Required config values:

CONFIG.VERSION
CONFIG.DEPLOY_VERSION_NUMBER
CONFIG.DEPLOYMENT_ID_ADMIN
CONFIG.DEPLOYMENT_ID_STUDENT

Admin UI must display:

Runtime: <VERSION> | Deploy: <DEPLOY_VERSION_NUMBER>

If mismatch detected:

VERSION MISMATCH – STOP DEPLOY

---------------------------------------------------------------------

MANDATORY: Runtime Endpoint

Every deployment must support:

?view=whoami

This endpoint must return:

VERSION
DEPLOY_VERSION_NUMBER
ScriptApp.getService().getUrl()
canonical admin base
canonical student base
timestamp

Purpose:

- detect stale cached HTML
- confirm active deployment/version
- confirm canonical URLs (no domain rewrites)

---------------------------------------------------------------------

MANDATORY: Test URL Output (copy-paste)

After deploy succeeds, Codex MUST print:

Admin (canonical):
https://script.google.com/macros/s/<ADMIN_ID>/exec?view=admin

Admin whoami (canonical):
https://script.google.com/macros/s/<ADMIN_ID>/exec?view=whoami

Student whoami (canonical):
https://script.google.com/macros/s/<STUDENT_ID>/exec?view=whoami

Student portal (canonical pattern):
https://script.google.com/macros/s/<STUDENT_ID>/exec?view=portal&id=<ApplicantID>&s=<Secret>

Codex MUST NOT output the /a/<domain>/ variant as the primary test URL.

---------------------------------------------------------------------

ALLOWED FILES (DEFAULT)

Codex may edit ONLY:

Admin.js
AdminUI.html
Code.js
Config.js
Utils.js
Routes.js

---------------------------------------------------------------------

FORBIDDEN FILES

Codex MUST NOT edit:

appsscript.json
.clasp.json

unless explicitly requested.

---------------------------------------------------------------------

DATA SAFETY RULES

Codex MUST NOT:

- add sheet headers
- add sheet columns
- modify schema

unless explicitly instructed.

---------------------------------------------------------------------

TOKEN SECURITY RULE

Never store plaintext portal tokens.

Allowed:

PortalTokenHash
PortalTokenIssuedAt

---------------------------------------------------------------------

PAYMENT LOCK RULE

Codex MUST preserve lock logic:

Payment_Verified == Yes
OR
Portal_Access_Status == Locked

---------------------------------------------------------------------

EMAIL CORRECTION RULE

Never overwrite Parent_Email.

Use:

Parent_Email_Corrected

---------------------------------------------------------------------

CODE CHANGE PHILOSOPHY

Prefer:

- minimal diffs
- no refactors
- no structural changes

unless explicitly requested.

---------------------------------------------------------------------

MANDATORY OUTPUT FORMAT (every Codex completion note)

Codex must always return:

1) diff summary
2) manual test checklist
3) deployment command output (verbatim)
4) canonical portal test URL(s) (copy-paste)
