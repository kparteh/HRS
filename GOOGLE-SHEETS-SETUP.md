# WorkSmart HR Liberia — Google Sheets setup

This package connects the WorkSmart HR single-page app to a Google Sheet through a Google Apps Script web app. The configured personal-Gmail mode is intended only for one owner/admin (`momokparteh@gmail.com`): the script executes as the signed-in user so Apps Script can verify a consumer Google account, and that account must have direct access to the spreadsheet.

> **Single-user security boundary:** Do not add other app users or share the spreadsheet with them. In `USER_ACCESSING` mode every signed-in app user must have direct spreadsheet permission; a user with that permission can open the Sheet and bypass the app's role controls. The app's `Access` allowlist is an additional in-app check, not a way to hide spreadsheet data from people who can access the file. For multi-user access to payroll data, use the separate Workspace-owner deployment model with an appropriately domain-restricted identity, or move to a backend that securely verifies external Google identities while keeping the spreadsheet inaccessible to end users.

**Sync behavior:** browser changes are written to Sheets through authorized Apps Script methods. Open sessions check the spreadsheet-backed data every 15 seconds while visible (and immediately when a background tab becomes visible). Changes made by an authorized spreadsheet owner are detected using a SHA-256 workspace fingerprint. Concurrent app writes use a script lock and revision/fingerprint check; a conflict reloads the latest data rather than silently overwriting it. This is near-real-time polling, not a push connection.

**Payroll is disabled until configured.** New workspaces have no tax bands or FX rate. An administrator must enter and attest verified payroll assumptions before a payroll draft can be saved. The app does not provide legal payroll advice, make payments, or file returns.

## Files to add to the Apps Script project

- [`Code.gs`](./Code.gs) — web-app entry point, schemas, sync API, server-side access checks, payroll draft creation, and audit events.
- [`hrsite.html`](./hrsite.html) — the user interface; Apps Script serves this same page with `google.script.run` enabled.
- [`appsscript.json`](./appsscript.json) — V8 runtime, Monrovia time zone, web-app deployment settings, and the OAuth scopes required for Sheets and private Drive documents.

## 1. Create the private spreadsheet

1. Sign in to the personal Google account `momokparteh@gmail.com` and create a Google Sheet owned by that account.
2. Keep the spreadsheet private to `momokparteh@gmail.com`. Do not add other app users; granting them direct spreadsheet access would expose the Sheet outside the app.
3. Copy the spreadsheet ID from its URL: the value between `/d/` and `/edit`.
4. For this single-user setup, the personal Gmail account is the owner/admin. For organizational or multi-user payroll use, do not use this personal-account deployment.

The initializer creates the tabs and headers below, including `EmployeeDocuments`. Attendance also records event-specific GPS coordinates, accuracy, and private Drive photo IDs. Existing attendance headers are extended with these columns automatically when the updated app first reads the workspace; other changed or unexpected headers are rejected.

## 2. Create and configure the Apps Script project

1. Open [script.google.com](https://script.google.com/) while signed in as `momokparteh@gmail.com` and create a new standalone project.
2. Replace the Apps Script editor's `Code.gs` contents with the complete local [`Code.gs`](./Code.gs) file. This is the server-side JavaScript that validates app updates and writes the appropriate rows to the Sheet tabs.
3. Add an HTML file named exactly `hrsite` (Apps Script adds the `.html` extension) and paste in the entire current [`hrsite.html`](./hrsite.html) file. Do not paste the HTML into `Code.gs` or create a second copy of the client sync code.
4. In **Project Settings**, enable **Show "appsscript.json" manifest file in editor**. Replace the manifest with this package's `appsscript.json`.
5. In **Project Settings → Script Properties**, add:
   - `HR_SPREADSHEET_ID` — the private Sheet ID.
   - `HR_BOOTSTRAP_ADMIN_EMAIL` — `momokparteh@gmail.com` (the exact personal Google account that will be the sole app administrator).
6. In the editor, select and run `HR_initializeWorkspace_`. Review the requested Spreadsheet, Drive, trigger-management and user-email scopes and authorize them as `momokparteh@gmail.com`. The function creates/validates the schema, initializes the settings, protects tabs against routine edits, adds the bootstrap admin, and appends an audit event. Fix any schema error it reports rather than deleting existing data.

The initializer also sets `HR_SINGLE_USER_EMAIL` to `momokparteh@gmail.com`; the server uses that script property to reject every other Google account. Do not remove or change it for this personal deployment. When changing the admin email on an existing deployment, update `HR_BOOTSTRAP_ADMIN_EMAIL` in Script Properties first, ensure the new account owns or has direct access to the spreadsheet and Apps Script project, then run `HR_initializeWorkspace_` while signed in as `momokparteh@gmail.com`. This disables the previous configured admin in the app's `Access` tab, but does not remove the previous Google account's direct ownership/access to the Sheet. To make the new account the sole data owner, create a new Sheet owned by that account and migrate records deliberately; do not assume the app allowlist revokes Google Drive permissions.

The initializer deliberately ends in `_` so it is a private Apps Script helper and cannot be invoked remotely through the browser bridge. Only `HR_getWorkspace` and `HR_saveWorkspace` are callable by the app.

### How the current HTML writes to Sheets

The current HTML detects the Apps Script environment (`google.script.run`). On page load it calls the public `HR_getWorkspace()` method to read and render the authorized workspace. When the user saves an app change, its existing `persist()` function sends the updated state, expected revision, and fingerprint to `HR_saveWorkspace()`. The server validates the actor and changed data, then calls the tab-specific writers in `Code.gs` (for example `hrWriteEmployees_`, `hrWriteAttendance_`, and `hrWriteLeave_`), updates the revision, records an audit event, and returns the new workspace. Keep these client and server files in the same Apps Script project:

```javascript
// Client-side call in hrsite.html; already implemented by persist().
google.script.run
  .withSuccessHandler(response => applyWorkspace(response, false))
  .withFailureHandler(error => {
    console.error("Google Sheets save failed.", error);
    toast(error && error.message ? error.message : "Could not save to Google Sheets.");
  })
  .HR_saveWorkspace(payload, SHEETS_SYNC.revision, SHEETS_SYNC.fingerprint);
```

The corresponding server methods `HR_getWorkspace()` and `HR_saveWorkspace(clientState, expectedRevision, expectedFingerprint)` are fully implemented in `Code.gs`; copy that entire file rather than creating replacement stubs. Do not add `SpreadsheetApp` calls or spreadsheet IDs to the browser HTML. In Apps Script, `doGet()` serves the `hrsite` HTML and enables `google.script.run`; opening the HTML as a local file only uses the browser demo and does not write to Sheets.

`HR_SPREADSHEET_ID` is a locator, not an authentication token. Do not add passwords, API keys, OAuth client secrets, or service-account keys to Script Properties, HTML, or the spreadsheet. In this personal-Gmail configuration, Apps Script accesses the Sheet using the signed-in user's Google authorization; therefore only the Sheet owner/admin should use the app. The app does not need a public Google Sheets API key or the Advanced Sheets API.

### Google Cloud and OAuth

Apps Script manages OAuth consent and requests the scopes in the manifest: spreadsheet access, the signed-in user's email, private Drive file operations, and a daily time-driven trigger for photo cleanup. A personal Gmail account cannot use a Workspace-only `DOMAIN` deployment. This package sets the web app to `USER_ACCESSING` and `ANYONE`, which requires the Google user to sign in and authorize the requested scopes; the server then checks the actual Google account against `Access`. The email typed into the sign-in form is only a consistency check and is not proof of identity.

### Employee files and attendance evidence

- Employee credential and other document uploads are limited to PDF, PNG, JPG, DOC, or DOCX files up to 8 MB. Files are stored in a private Drive folder; only authorized HR/admin users (and an employee for their own records) receive file links.
- Clock events request the browser's GPS location and camera photo only after the user presses Clock in/Clock out. Live camera preview requires a supported browser, HTTPS, camera permission, and camera access allowed by the host frame. If Apps Script's embedded frame blocks live preview, use **Use device camera or choose photo**; on mobile, the file input requests the device camera, and on desktop it opens the image chooser.
- The camera flow captures a still image for supervisor review. It does **not** perform facial recognition, create face templates, or verify identity. Attendance GPS is device-reported and should not be treated as tamper-proof proof of location.
- Attendance photos are removed from application access and moved to Drive Trash after 30 days. The first photo upload installs a daily cleanup trigger and may require a one-time Apps Script consent approval. Google Drive may keep trashed files recoverable in Trash for a further period before permanent deletion.
- GPS/photo evidence is collected only for an explicit clock event. Obtain and document employees' appropriate notice and consent before enabling these features in operational use.
- The in-app Sign out button clears the WorkSmart session view; it does not sign the user out of their Google account.

`ANYONE` means users must sign in with a Google account; do not change it to `ANYONE_ANONYMOUS` (“Anyone, even anonymous”). A Google Cloud OAuth consent screen is not needed for this basic Apps Script deployment. `SpreadsheetApp` does not require an API key or a separately enabled Advanced Sheets API.

## 3. Deploy the web app

1. Choose **Deploy → New deployment → Web app**.
2. Set **Execute as** to **User accessing the web app**. This is required so the server can verify the personal Google account and access the Sheet as its owner.
3. Set **Who has access** to **Anyone with a Google account**. The manifest specifies `USER_ACCESSING` and `ANYONE`; Google still requires account sign-in.
4. Deploy, authorize, and copy the `/exec` web-app URL.
5. Open that URL in a browser signed in as `momokparteh@gmail.com` and approve Google's authorization prompt if asked. WorkSmart automatically requests the workspace; it does not use a separate email/password form. The backend verifies the actual signed-in account against the app's `Access` tab.
6. Keep `momokparteh@gmail.com` as the only active application user. Do not add other role accounts to the `Access` tab or give them spreadsheet permissions.
7. Do not use the local demo role switch as a security test; it is hidden in the Apps Script deployment.

Google, not WorkSmart, verifies the account password. The app never receives or stores a Google password. The admin account is `momokparteh@gmail.com`; there is no separate WorkSmart password for the Apps Script deployment. Password recovery goes through Google's account recovery process.

## Spreadsheet tabs and exact headers

The initializer creates one header row per tab. Dates are ISO `YYYY-MM-DD` unless a timestamp is specified; money is stored as numeric values in the row's stated currency.

| Tab | Column headers (in order) | Purpose |
| --- | --- | --- |
| `Employees` | `employee_id`, `name`, `email`, `department`, `title`, `location`, `employment_type`, `currency`, `monthly_salary`, `status`, `start_date` | Employment directory and monthly base compensation. Status is `Active` or `Inactive`; currency is `LRD` or `USD`. |
| `Attendance` | `employee_id`, `date`, `clock_in`, `clock_out`, `status`, `source`, `updated_at` | One employee/day record; timestamps are Monrovia-local `HH:mm`, `source` currently `web`, and `updated_at` is a timestamp. |
| `LeaveRequests` | `request_id`, `employee_id`, `leave_type`, `start_date`, `end_date`, `working_days`, `reason`, `status`, `created_at`, `reviewed_by`, `reviewed_at` | Employee-submitted leave and reviewer metadata. |
| `PayrollRuns` | `run_id`, `period`, `created_at`, `employee_count`, `currency`, `gross_total`, `net_total`, `status`, `rules_version`, `created_by` | Append-only draft run summary; app-created runs are always `Draft`. |
| `PayrollItems` | `run_id`, `employee_id`, `currency`, `gross`, `income_tax`, `employee_contribution`, `employer_contribution`, `work_injury`, `net_pay` | Per-employee payroll calculation snapshot in the employee's currency. |
| `TaxRules` | `rule_version`, `rule_type`, `band_number`, `annual_from_lrd`, `annual_to_lrd`, `rate_percent`, `effective_from`, `verified`, `verification_source`, `verification_reference`, `verified_by`, `verified_at` | Versioned annual LRD tax bands and contribution assumptions, including the administrator's cited source. The active version is selected in `Settings`. |
| `Access` | `email`, `role`, `active`, `employee_id`, `department` | Server-enforced allowlist. Roles: `admin`, `hr`, `payroll`, `manager`, `employee`. The last two use an employee/dept scope. |
| `Settings` | `key`, `value` | Organization, timezone/currency, exchange-rate source/effective date, payroll verification/version, and optimistic revision. |
| `AuditLog` | `event_id`, `timestamp`, `actor_email`, `actor_role`, `action`, `entity_type`, `entity_id`, `before_json`, `after_json`, `revision` | Append-only app change history. JSON snapshots may contain salary data; keep this tab private. |

## Sync API and access policy

The UI calls `HR_getWorkspace()` and `HR_saveWorkspace(state, expectedRevision, expectedFingerprint)` for workspace data, plus dedicated document listing/upload and attendance-photo methods. The server verifies the active Google email against `HR_SINGLE_USER_EMAIL` and the `Access` tab on every call; client-supplied role values are not authorization. `LockService` serializes writes; optimistic concurrency rejects stale edits. The other roles below describe the application's intended authorization model, but they are not available as separate users in this personal-Gmail deployment.

| Role | Server-side write permissions |
| --- | --- |
| Organization admin | Organization, employees, attendance, leave review, payroll drafts, payroll settings/verification, application access |
| HR manager | Employees, attendance, leave requests/reviews |
| Payroll officer | Payroll draft creation only |
| Team manager | Leave requests and approvals in the assigned department |
| Employee | Own leave requests, current-day web clock-in/out, and own documents |

The UI enforces the intended navigation permissions, but the Apps Script checks permissions again server-side. For this single-user setup, only the configured administrator can sign in; do not add other accounts or share the Sheet. At least one active administrator is required, and an administrator cannot revoke/demote their own active account from the app.

## Payroll and audit safeguards

- Tax bands and exchange rate begin **unset/unverified**. Each payroll-rules save creates a new `TaxRules.rule_version`, effective date, cited source/reference, verification flag, verifier email, and timestamp; payroll run rows capture the active rule version.
- The exchange rate requires an explicit source, effective date, and administrator attestation. Payroll draft creation is rejected server-side until both payroll rules and FX are verified.
- `PayrollRuns`/`PayrollItems` are calculated and written server-side from the current employee records and verified assumptions; client-submitted totals are not trusted. Drafts are immutable in this integration; there is no payment or filing connector.
- A script lock and revision plus SHA-256 fingerprint reject stale concurrent writes instead of silently overwriting another session's changes.
- Every app update records actor, role, entity, before/after snapshot, and revision in `AuditLog`. Keep the spreadsheet private to the sole admin, secure that Google account with two-step verification and recovery options, and make protected backups.
- The sheets are protected against routine direct edits. The spreadsheet owner can still bypass protections; direct owner edits do not create an app audit event. Prefer app workflows and review protected-tab changes regularly.
- This prototype stores payroll in Google Sheets. Sheets is not a substitute for a production payroll database, a legally reviewed payroll implementation, formal maker-checker approval, bank controls, encrypted document storage, or a statutory filing service.

## Local preview and troubleshooting

- Opening the HTML through a local file or the earlier local preview runs the browser-only demo; it does **not** connect to Google Sheets or accept production credentials.
- The local preview sign-in is `admin` / `WorkSmart#2026`. These demo credentials are embedded in `hrsite.html` and can be viewed or bypassed by anyone with browser or file access; this is only a user-interface gate, not authentication or protection for HR data. Change the constants before sharing a demo, and never put real employee/payroll data in the local preview.
- Local employee self-service includes an optional WebAuthn platform-authenticator prompt for clock-in/out. It stores only a browser-local credential ID and never reads or stores face/fingerprint data. **This is a UI demo only:** its challenge/signature is not verified by a trusted server and it can be bypassed by someone who can edit the page or browser storage. It is not enabled for Apps Script/Sheets attendance and must not be used to protect production time records. Production passkey attendance needs a server that securely verifies WebAuthn registration and assertions, plus recovery and accessible alternatives.
- The Sheets sign-in and sync flow runs only from the Apps Script `/exec` URL, which provides `google.script.run`.
- “Access denied” means the current Google email does not match `HR_SINGLE_USER_EMAIL`, is missing/inactive in `Access`, or the app is not authorized to read the account email.
- If Apps Script cannot read the signed-in email, verify deployment uses **User accessing the web app**, sign in with `momokparteh@gmail.com`, and approve the requested Google scopes. Never work around the check with an email or password supplied by the browser.
- Schema/header errors require an administrator to back up and deliberately repair the spreadsheet; do not clear tabs to silence them.
