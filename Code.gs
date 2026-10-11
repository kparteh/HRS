const HR_SCHEMA = {
  Employees: ["employee_id", "name", "email", "department", "title", "location", "employment_type", "currency", "monthly_salary", "status", "start_date"],
  Attendance: ["employee_id", "date", "clock_in", "clock_out", "status", "source", "updated_at", "clock_in_latitude", "clock_in_longitude", "clock_in_accuracy_m", "clock_in_photo_file_id", "clock_out_latitude", "clock_out_longitude", "clock_out_accuracy_m", "clock_out_photo_file_id"],
  EmployeeDocuments: ["document_id", "employee_id", "document_type", "file_name", "mime_type", "drive_file_id", "uploaded_by", "uploaded_at"],
  LeaveRequests: ["request_id", "employee_id", "leave_type", "start_date", "end_date", "working_days", "reason", "status", "created_at", "reviewed_by", "reviewed_at"],
  PayrollRuns: ["run_id", "period", "created_at", "employee_count", "currency", "gross_total", "net_total", "status", "rules_version", "created_by"],
  PayrollItems: ["run_id", "employee_id", "currency", "gross", "income_tax", "employee_contribution", "employer_contribution", "work_injury", "net_pay"],
  TaxRules: ["rule_version", "rule_type", "band_number", "annual_from_lrd", "annual_to_lrd", "rate_percent", "effective_from", "verified", "verification_source", "verification_reference", "verified_by", "verified_at"],
  Access: ["email", "role", "active", "employee_id", "department"],
  Settings: ["key", "value"],
  AuditLog: ["event_id", "timestamp", "actor_email", "actor_role", "action", "entity_type", "entity_id", "before_json", "after_json", "revision"]
};

const HR_ROLES = ["admin", "hr", "payroll", "manager", "employee"];

function hrIsDate_(value) {
  if (typeof value !== "string" || !/^\d{4}-\d{2}-\d{2}$/.test(value)) return false;
  const date = new Date(value + "T00:00:00Z");
  return !Number.isNaN(date.getTime()) && date.toISOString().slice(0, 10) === value;
}

function hrIsTime_(value) {
  return typeof value === "string" && /^(?:[01]\d|2[0-3]):[0-5]\d$/.test(value);
}

function hrMinutes_(value) {
  const [hours, minutes] = value.split(":").map(Number);
  return hours * 60 + minutes;
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile("hrsite")
    .setTitle("WorkSmart HR Liberia")
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

function HR_configureWorkspace_(spreadsheetId, bootstrapAdminEmail) {
  const id = String(spreadsheetId || "").trim();
  const email = String(bootstrapAdminEmail || "").trim().toLowerCase();
  if (!/^[a-zA-Z0-9_-]{25,}$/.test(id)) throw new Error("Enter a valid Google spreadsheet ID.");
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) throw new Error("Enter a valid Google account email.");

  const properties = PropertiesService.getScriptProperties();
  const previousSingleUserEmail = String(properties.getProperty("HR_SINGLE_USER_EMAIL") || "").trim().toLowerCase();
  properties.setProperty("HR_SPREADSHEET_ID", id);
  properties.setProperty("HR_SINGLE_USER_EMAIL", email);
  const book = SpreadsheetApp.openById(id);
  Object.keys(HR_SCHEMA).forEach(name => hrEnsureSheet_(book, name));

  const settings = hrReadRows_(book, "Settings");
  if (!settings.some(row => row.key === "revision")) {
    [
      ["revision", "1"],
      ["organization_name", "WorkSmart HR"],
      ["jurisdiction", "Liberia"],
      ["timezone", "Africa/Monrovia"],
      ["reporting_currency", "USD"],
      ["usd_to_lrd_rate", ""],
      ["usd_to_lrd_rate_verified", "false"],
      ["usd_to_lrd_source", ""],
      ["usd_to_lrd_effective_date", ""],
      ["workday_hours", "8"],
      ["payroll_rules_verified", "false"],
      ["payroll_rules_version", "0"],
      ["payroll_rules_effective_from", ""],
      ["payroll_rules_source", ""],
      ["payroll_rules_reference", ""],
      ["employee_contribution_percent", "0"],
      ["employer_contribution_percent", "0"],
      ["work_injury_contribution_percent", "0"]
    ].forEach(row => hrAppendRow_(book, "Settings", row));
  }

  const accessSheet = book.getSheetByName("Access");
  const access = hrReadRows_(book, "Access");
  if (previousSingleUserEmail && previousSingleUserEmail !== email) {
    const previousAdminIndex = access.findIndex(row => String(row.email).trim().toLowerCase() === previousSingleUserEmail);
    if (previousAdminIndex >= 0) {
      accessSheet.getRange(previousAdminIndex + 2, 3).setValue(false);
      access[previousAdminIndex].active = false;
    }
  }
  const existingAdminIndex = access.findIndex(row => String(row.email).trim().toLowerCase() === email);
  if (existingAdminIndex < 0) {
    hrAppendRow_(book, "Access", [email, "admin", true, "", ""]);
  } else {
    accessSheet.getRange(existingAdminIndex + 2, 1, 1, HR_SCHEMA.Access.length).setValues([[email, "admin", true, "", ""]]);
  }
  Object.keys(HR_SCHEMA).forEach(name => hrProtectSheet_(book.getSheetByName(name)));
  hrAudit_(book, { email, role: "admin" }, "workspace.configured", "organization", id, null, { spreadsheetId: id }, 1);
  return "Workspace initialized for the configured single Google account. The script accesses the spreadsheet as the signed-in user.";
}

function HR_initializeWorkspace_() {
  const properties = PropertiesService.getScriptProperties();
  return HR_configureWorkspace_(
    properties.getProperty("HR_SPREADSHEET_ID"),
    properties.getProperty("HR_BOOTSTRAP_ADMIN_EMAIL")
  );
}

function HR_getWorkspace() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) throw new Error("Workspace is busy. Retry shortly.");
  try {
    const book = hrBook_();
    const actor = hrActor_(book);
    return hrWorkspaceResponse_(book, actor);
  } finally {
    lock.releaseLock();
  }
}

function HR_listEmployeeDocuments(employeeId) {
  const book = hrBook_();
  const actor = hrActor_(book);
  const id = String(employeeId || "").trim();
  hrRequireEmployeeEvidenceAccess_(book, actor, id, false);
  hrProtectSheet_(hrEnsureSheet_(book, "EmployeeDocuments"));
  const attendanceOnly = actor.role === "manager";
  return hrReadRows_(book, "EmployeeDocuments")
    .filter(row => String(row.employee_id) === id && row.drive_file_id &&
      (!attendanceOnly || String(row.document_type).startsWith("Attendance photo")))
    .map(row => ({
      id: String(row.document_id),
      employeeId: String(row.employee_id),
      type: String(row.document_type),
      name: String(row.file_name),
      uploadedAt: hrDateTime_(row.uploaded_at),
      url: DriveApp.getFileById(String(row.drive_file_id)).getUrl()
    }));
}

function HR_uploadEmployeeDocument(employeeId, documentType, fileName, mimeType, base64Data) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) throw new Error("Workspace is busy. Retry shortly.");
  try {
    const book = hrBook_();
    const actor = hrActor_(book);
    const id = String(employeeId || "").trim();
    hrRequireEmployeeEvidenceAccess_(book, actor, id, true);
    const type = String(documentType || "").trim();
    if (!["Identity document", "Employment credential", "Contract", "Other"].includes(type)) {
      throw new Error("Choose a valid employee document type.");
    }
    return hrCreateEmployeeDocument_(book, actor, id, type, fileName, mimeType, base64Data);
  } finally {
    lock.releaseLock();
  }
}

function HR_uploadAttendancePhoto(employeeId, action, dataUrl) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) throw new Error("Workspace is busy. Retry shortly.");
  try {
    const book = hrBook_();
    const actor = hrActor_(book);
    const id = String(employeeId || "").trim();
    if (!["admin", "hr", "employee"].includes(actor.role)) {
      throw new Error("Your role cannot upload attendance photos.");
    }
    hrRequireEmployeeEvidenceAccess_(book, actor, id, true);
    if (!["clock-in", "clock-out"].includes(action)) throw new Error("Choose a valid attendance action.");
    const match = String(dataUrl || "").match(/^data:image\/jpeg;base64,([A-Za-z0-9+/=]+)$/);
    if (!match) throw new Error("Capture a valid JPEG camera photo before continuing.");
    HR_ensurePhotoRetentionTrigger_();
    const documentRecord = hrCreateEmployeeDocument_(
      book, actor, id, "Attendance photo · " + action, "attendance-" + action + ".jpg", "image/jpeg", match[1]
    );
    return { document: documentRecord, workspace: hrWorkspaceResponse_(book, actor) };
  } finally {
    lock.releaseLock();
  }
}

function HR_cleanupAttendancePhotos_() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) throw new Error("Workspace is busy. Retry shortly.");
  try {
    const book = hrBook_();
    const sheet = hrEnsureSheet_(book, "EmployeeDocuments");
    const attendanceSheet = hrEnsureSheet_(book, "Attendance");
    hrProtectSheet_(sheet);
    const rows = hrReadRows_(book, "EmployeeDocuments");
    const attendance = hrReadRows_(book, "Attendance");
    const cutoff = Date.now() - 30 * 24 * 60 * 60 * 1000;
    const owner = String(PropertiesService.getScriptProperties().getProperty("HR_SINGLE_USER_EMAIL") || "");
    let revision = Number(hrSetting_(book, "revision") || 1);
    rows.forEach((row, index) => {
      if (!String(row.document_type).startsWith("Attendance photo") || !row.drive_file_id) return;
      const uploadedAt = row.uploaded_at instanceof Date ? row.uploaded_at.getTime() : new Date(row.uploaded_at).getTime();
      if (!Number.isFinite(uploadedAt) || uploadedAt > cutoff) return;
      const fileId = String(row.drive_file_id);
      DriveApp.getFileById(fileId).setTrashed(true);
      attendance.forEach((record, attendanceIndex) => {
        if (String(record.employee_id) !== String(row.employee_id)) return;
        if (String(record.clock_in_photo_file_id) === fileId) attendanceSheet.getRange(attendanceIndex + 2, 11).clearContent();
        if (String(record.clock_out_photo_file_id) === fileId) attendanceSheet.getRange(attendanceIndex + 2, 15).clearContent();
      });
      sheet.getRange(index + 2, 6).clearContent();
      revision++;
      hrSetSetting_(book, "revision", String(revision));
      hrAudit_(book, { email: owner, role: "system" }, "attendance.photo.retention_expired", "employee", String(row.employee_id), null, { documentId: String(row.document_id) }, revision);
    });
  } finally {
    lock.releaseLock();
  }
}

function HR_ensurePhotoRetentionTrigger_() {
  const installed = ScriptApp.getProjectTriggers().some(trigger => trigger.getHandlerFunction() === "HR_cleanupAttendancePhotos_");
  if (!installed) {
    ScriptApp.newTrigger("HR_cleanupAttendancePhotos_").timeBased().everyDays(1).create();
  }
}

function hrRequireEmployeeEvidenceAccess_(book, actor, employeeId, upload) {
  const person = hrReadRows_(book, "Employees").find(row => String(row.employee_id) === employeeId);
  if (!employeeId || !person) throw new Error("Choose an existing employee record.");
  if (actor.role === "admin" || actor.role === "hr") return person;
  if (!upload && actor.role === "manager" && person.status === "Active" && String(person.department) === actor.department) return person;
  if (actor.role === "employee" && employeeId === actor.employeeId) return person;
  throw new Error("You do not have permission to access this employee's documents.");
}

function hrCreateEmployeeDocument_(book, actor, employeeId, documentType, fileName, mimeType, base64Data) {
  const safeName = String(fileName || "").replace(/[\\/:*?"<>|\r\n]/g, "_").trim().slice(0, 120);
  const contentType = String(mimeType || "").toLowerCase();
  const extension = safeName.split(".").pop().toLowerCase();
  const allowedTypes = {
    pdf: "application/pdf",
    png: "image/png",
    jpg: "image/jpeg",
    jpeg: "image/jpeg",
    doc: "application/msword",
    docx: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
  };
  const payload = String(base64Data || "");
  if (!safeName || allowedTypes[extension] !== contentType) throw new Error("Upload a PDF, PNG, JPG or Word document.");
  if (!payload || payload.length > 11 * 1024 * 1024) throw new Error("Files must be smaller than 8 MB.");
  const bytes = Utilities.base64Decode(payload);
  if (!bytes.length || bytes.length > 8 * 1024 * 1024) throw new Error("Files must be smaller than 8 MB.");
  const folder = hrPrivateDocumentsFolder_();
  const file = folder.createFile(Utilities.newBlob(bytes, contentType, safeName));
  try {
    file.setDescription("Private WorkSmart HR employee document. Access is controlled by the HR application.");
    const now = new Date();
    const doc = {
      id: Utilities.getUuid(),
      employeeId: employeeId,
      type: documentType,
      name: safeName,
      uploadedAt: now.toISOString(),
      driveFileId: file.getId(),
      url: file.getUrl()
    };
    hrAppendRow_(book, "EmployeeDocuments", [
      doc.id, employeeId, documentType, safeName, contentType, file.getId(), actor.email, now
    ]);
    hrProtectSheet_(book.getSheetByName("EmployeeDocuments"));
    const revision = Number(hrSetting_(book, "revision") || 1) + 1;
    hrSetSetting_(book, "revision", String(revision));
    hrAudit_(book, actor, "employee.document.uploaded", "employee", employeeId, null, {
      documentId: doc.id, documentType: documentType, fileName: safeName
    }, revision);
    return doc;
  } catch (error) {
    file.setTrashed(true);
    throw error;
  }
}

function hrPrivateDocumentsFolder_() {
  const properties = PropertiesService.getScriptProperties();
  const existingId = properties.getProperty("HR_PRIVATE_DOCUMENTS_FOLDER_ID");
  if (existingId) return DriveApp.getFolderById(existingId);
  const folder = DriveApp.createFolder("WorkSmart HR - Private Employee Documents");
  properties.setProperty("HR_PRIVATE_DOCUMENTS_FOLDER_ID", folder.getId());
  return folder;
}

function hrWorkspaceResponse_(book, actor) {
  const revision = Number(hrSetting_(book, "revision") || 1);
  const full = hrReadState_(book);
  const fingerprint = hrFingerprint_(full);
  return {
    state: hrProjectState_(full, actor),
    role: actor.role,
    email: actor.email,
    employeeId: actor.employeeId,
    department: actor.department,
    revision: revision,
    fingerprint: fingerprint,
    serverTime: new Date().toISOString()
  };
}

function HR_saveWorkspace(clientState, expectedRevision, expectedFingerprint) {
  if (!clientState || typeof clientState !== "object") throw new Error("A workspace update is required.");
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) throw new Error("Workspace is busy. Retry in a moment.");
  try {
    const book = hrBook_();
    const actor = hrActor_(book);
    const currentRevision = Number(hrSetting_(book, "revision") || 1);
    const full = hrReadState_(book);
    if (Number(expectedRevision) !== currentRevision || String(expectedFingerprint || "") !== hrFingerprint_(full)) {
      throw new Error("SYNC_CONFLICT: This workspace changed in another session. The latest sheet data will be reloaded; please review and retry.");
    }

    const baseline = hrProjectState_(full, actor);
    const incoming = hrNormalizeClientState_(clientState);
    const changed = [];
    const allow = hrAllowedChanges_(actor.role);
    const categories = ["organization", "staff", "attendance", "leave", "payrollRuns", "taxBrackets", "contributions", "payrollRulesVerified", "payrollRulesEffectiveFrom", "payrollRulesSource", "payrollRulesReference", "access"];

    categories.forEach(category => {
      if (hrSame_(incoming[category], baseline[category])) return;
      if (!allow.includes(category)) throw new Error("Permission denied: you cannot change " + category + ".");
      changed.push(category);
    });

    if (!changed.length) return hrWorkspaceResponse_(book, actor);
    changed.forEach(category => {
      if (category === "staff") hrValidateEmployees_(incoming.staff, full.staff);
      else if (category === "attendance") hrValidateAttendanceChanges_(full.attendance, incoming.attendance, actor);
      else if (category === "leave") hrValidateLeaveChanges_(full.leave, incoming.leave, actor);
      else if (category === "payrollRuns") hrValidatePayrollAppend_(full.payrollRuns, incoming.payrollRuns, actor, full);
      else if (category === "organization") hrValidateOrganization_(incoming.organization);
      else if (category === "access") hrValidateAccess_(book, incoming.access, actor);
    });

    const rulesChanged = changed.some(name => ["contributions", "taxBrackets", "payrollRulesVerified", "payrollRulesEffectiveFrom", "payrollRulesSource", "payrollRulesReference"].includes(name));
    if (rulesChanged) {
      const verified = incoming.payrollRulesVerified === true;
      const effectiveDate = String(incoming.payrollRulesEffectiveFrom || "");
      if (verified && (!hrIsDate_(effectiveDate) || effectiveDate > Utilities.formatDate(new Date(), "Africa/Monrovia", "yyyy-MM-dd"))) {
        throw new Error("Set a valid current or past effective date for the verified payroll rules.");
      }
      if (verified && (!String(incoming.payrollRulesSource || "").trim() || !String(incoming.payrollRulesReference || "").trim())) {
        throw new Error("Record the official source and reference used to verify these payroll rules.");
      }
      if (String(incoming.payrollRulesSource || "").length > 240 || String(incoming.payrollRulesReference || "").length > 500 ||
          (effectiveDate && !hrIsDate_(effectiveDate))) {
        throw new Error("Payroll rule source, reference, or effective date is invalid.");
      }
      hrValidateTaxRules_(incoming.taxBrackets, incoming.contributions);
    }

    changed.forEach(category => {
      if (category === "staff") hrWriteEmployees_(book, incoming.staff);
      else if (category === "attendance") {
        hrMergeScoped_(full.attendance, incoming.attendance, hrCanSeeEmployee_(actor), "employeeId", "date");
        hrWriteAttendance_(book, full.attendance, actor);
      } else if (category === "leave") {
        hrMergeScoped_(full.leave, incoming.leave, hrCanSeeEmployee_(actor), "id");
        hrWriteLeave_(book, full.leave);
      } else if (category === "payrollRuns") hrWritePayrollRuns_(book, incoming.payrollRuns, actor);
      else if (category === "organization") hrWriteOrganization_(book, incoming.organization);
      else if (category === "access") hrWriteAccess_(book, incoming.access);
    });

    if (rulesChanged) {
      const verified = incoming.payrollRulesVerified === true;
      const nextRuleVersion = Number(hrSetting_(book, "payroll_rules_version") || 0) + 1;
      const effectiveDate = String(incoming.payrollRulesEffectiveFrom || "");
      hrWriteTaxRules_(book, incoming.taxBrackets, incoming.contributions, actor, nextRuleVersion, effectiveDate, verified, incoming.payrollRulesSource, incoming.payrollRulesReference);
      hrSetSetting_(book, "payroll_rules_verified", verified ? "true" : "false");
      hrSetSetting_(book, "payroll_rules_version", String(nextRuleVersion));
      hrSetSetting_(book, "payroll_rules_effective_from", effectiveDate);
      hrSetSetting_(book, "payroll_rules_source", String(incoming.payrollRulesSource || "").slice(0, 240));
      hrSetSetting_(book, "payroll_rules_reference", String(incoming.payrollRulesReference || "").slice(0, 500));
    }
    const nextRevision = currentRevision + 1;
    hrSetSetting_(book, "revision", String(nextRevision));
    changed.forEach(name => hrAudit_(book, actor, "workspace.updated", name, name, baseline[name], incoming[name], nextRevision));
    return hrWorkspaceResponse_(book, actor);
  } finally {
    lock.releaseLock();
  }
}

function hrBook_() {
  const id = PropertiesService.getScriptProperties().getProperty("HR_SPREADSHEET_ID");
  if (!id) throw new Error("Workspace is not configured. Ask the system administrator to run HR_initializeWorkspace_ in the Apps Script editor.");
  return SpreadsheetApp.openById(id);
}

function hrActor_(book) {
  const email = String(Session.getActiveUser().getEmail() || "").trim().toLowerCase();
  if (!email) throw new Error("Google could not verify your account email. Sign in with the configured Google account and authorize the app.");
  const singleUserEmail = String(PropertiesService.getScriptProperties().getProperty("HR_SINGLE_USER_EMAIL") || "").trim().toLowerCase();
  if (!singleUserEmail) throw new Error("Personal-account access is not configured. Ask the app owner to rerun HR_initializeWorkspace_.");
  if (email !== singleUserEmail) throw new Error("This personal-Gmail deployment is restricted to its single configured administrator.");
  const row = hrReadRows_(book, "Access").find(item => String(item.email).trim().toLowerCase() === email && item.active === true);
  if (!row || !HR_ROLES.includes(row.role)) throw new Error("Access denied. Ask your organization admin to grant you a role.");
  const actor = { email, role: row.role, employeeId: row.employeeId || "", department: row.department || "" };
  if (actor.role === "employee" && !hrReadRows_(book, "Employees").some(person => String(person.employee_id) === actor.employeeId && person.status === "Active")) {
    throw new Error("Your employee account is not linked to an active employee record. Ask an admin to update your access.");
  }
  if (actor.role === "manager" && !hrReadRows_(book, "Employees").some(person => String(person.department) === actor.department && person.status === "Active")) {
    throw new Error("Your manager account has no active department scope. Ask an admin to update your access.");
  }
  return actor;
}

function hrAllowedChanges_(role) {
  if (role === "admin") return ["organization", "staff", "attendance", "leave", "payrollRuns", "taxBrackets", "contributions", "payrollRulesVerified", "payrollRulesEffectiveFrom", "payrollRulesSource", "payrollRulesReference", "access"];
  if (role === "hr") return ["staff", "attendance", "leave"];
  if (role === "payroll") return ["payrollRuns"];
  if (role === "employee") return ["attendance", "leave"];
  if (role === "manager") return ["leave"];
  return [];
}

function hrReadState_(book) {
  const settingsRows = hrReadRows_(book, "Settings");
  const values = Object.fromEntries(settingsRows.map(row => [row.key, row.value]));
  const employees = hrReadRows_(book, "Employees").map(row => ({
    id: String(row.employee_id), name: String(row.name), email: String(row.email), department: String(row.department),
    title: String(row.title), location: String(row.location), type: String(row.employment_type),
    currency: String(row.currency), salary: Number(row.monthly_salary) || 0, status: String(row.status), start: hrDate_(row.start_date)
  }));
  const attendance = hrReadRows_(book, "Attendance").map(row => ({
    employeeId: String(row.employee_id), date: hrDate_(row.date), in: String(row.clock_in || "—"),
    out: String(row.clock_out || ""), status: String(row.status), source: String(row.source || "web"), updatedAt: hrDateTime_(row.updated_at),
    inLatitude: hrNullableNumber_(row.clock_in_latitude), inLongitude: hrNullableNumber_(row.clock_in_longitude),
    inAccuracy: hrNullableNumber_(row.clock_in_accuracy_m), inPhotoId: String(row.clock_in_photo_file_id || ""),
    outLatitude: hrNullableNumber_(row.clock_out_latitude), outLongitude: hrNullableNumber_(row.clock_out_longitude),
    outAccuracy: hrNullableNumber_(row.clock_out_accuracy_m), outPhotoId: String(row.clock_out_photo_file_id || "")
  }));
  const leave = hrReadRows_(book, "LeaveRequests").map(row => ({
    id: String(row.request_id), employeeId: String(row.employee_id), type: String(row.leave_type),
    start: hrDate_(row.start_date), end: hrDate_(row.end_date), days: Number(row.working_days) || 0,
    reason: String(row.reason || ""), status: String(row.status), createdAt: hrDateTime_(row.created_at),
    reviewedBy: String(row.reviewed_by || ""), reviewedAt: hrDateTime_(row.reviewed_at)
  }));
  const runs = hrReadRows_(book, "PayrollRuns").map(row => ({
    id: String(row.run_id), period: String(row.period), created: hrDate_(row.created_at),
    count: Number(row.employee_count) || 0, currency: String(row.currency),
    gross: Number(row.gross_total) || 0, net: Number(row.net_total) || 0,
    status: String(row.status), rulesVersion: String(row.rules_version || ""), createdBy: String(row.created_by || "")
  }));
  const activeRuleVersion = Number(values.payroll_rules_version || 0);
  const rules = hrReadRows_(book, "TaxRules").filter(row => Number(row.rule_version) === activeRuleVersion);
  const brackets = rules.filter(row => row.rule_type === "income_tax").map(row => ({
    from: Number(row.annual_from_lrd) || 0,
    to: row.annual_to_lrd === "" || row.annual_to_lrd === null ? null : Number(row.annual_to_lrd),
    rate: Number(row.rate_percent) || 0
  }));
  const contribution = type => {
    const row = rules.find(item => item.rule_type === type);
    return row ? Number(row.rate_percent) || 0 : Number(values[type + "_percent"]) || 0;
  };
  const access = hrReadRows_(book, "Access").map(row => ({
    email: String(row.email), role: String(row.role), active: row.active === true,
    employeeId: String(row.employee_id || ""), department: String(row.department || "")
  }));
  return {
    organization: {
      name: values.organization_name || "WorkSmart HR", country: values.jurisdiction || "Liberia",
      timezone: values.timezone || "Africa/Monrovia", currency: values.reporting_currency || "USD",
      exchangeRate: Number(values.usd_to_lrd_rate) || 1, exchangeRateVerified: values.usd_to_lrd_rate_verified === "true",
      exchangeRateSource: values.usd_to_lrd_source || "", exchangeRateEffective: values.usd_to_lrd_effective_date || "",
      workdayHours: Number(values.workday_hours) || 8
    },
    staff: employees, attendance, leave, payrollRuns: runs, taxBrackets: brackets,
    contributions: { employee: contribution("employee_contribution"), employer: contribution("employer_contribution"), workInjury: contribution("work_injury_contribution") },
    payrollRulesVerified: values.payroll_rules_verified === "true",
    payrollRulesVersion: values.payroll_rules_version || "0",
    payrollRulesEffectiveFrom: values.payroll_rules_effective_from || "",
    payrollRulesSource: values.payroll_rules_source || "",
    payrollRulesReference: values.payroll_rules_reference || "",
    access, clock: {}
  };
}

function hrProjectState_(full, actor) {
  const state = JSON.parse(JSON.stringify(full));
  const today = Utilities.formatDate(new Date(), "Africa/Monrovia", "yyyy-MM-dd");
  const ownAttendance = state.attendance.find(item => item.employeeId === actor.employeeId && item.date === today);
  if (actor.employeeId && ownAttendance && ownAttendance.in && ownAttendance.in !== "—" && ownAttendance.status === "Present") {
    state.clock[actor.employeeId] = { in: ownAttendance.in, out: ownAttendance.out || "" };
  }
  if (actor.role === "manager") {
    state.staff = state.staff.filter(person => person.department === actor.department && person.status === "Active").map(hrWithoutSalary_);
    const ids = new Set(state.staff.map(person => person.id));
    state.attendance = state.attendance.filter(item => ids.has(item.employeeId));
    state.leave = state.leave.filter(item => ids.has(item.employeeId));
    state.payrollRuns = [];
    state.taxBrackets = [];
    state.contributions = { employee: 0, employer: 0, workInjury: 0 };
  } else if (actor.role === "employee") {
    state.staff = state.staff.filter(person => person.id === actor.employeeId);
    state.attendance = state.attendance.filter(item => item.employeeId === actor.employeeId);
    state.leave = state.leave.filter(item => item.employeeId === actor.employeeId);
    state.payrollRuns = [];
    state.taxBrackets = [];
    state.contributions = { employee: 0, employer: 0, workInjury: 0 };
  } else if (actor.role === "payroll") {
    state.leave = [];
    state.access = [];
    state.organization = { ...state.organization, name: "WorkSmart HR", country: "Liberia", timezone: "Africa/Monrovia" };
  } else if (actor.role !== "admin") {
    state.payrollRuns = [];
    state.taxBrackets = [];
    state.contributions = { employee: 0, employer: 0, workInjury: 0 };
    state.access = [];
  }
  if (actor.role !== "admin") state.organization = { ...state.organization };
  return state;
}

function hrWithoutSalary_(person) {
  const safe = { ...person };
  delete safe.salary;
  return safe;
}

function hrNormalizeClientState_(state) {
  const copy = JSON.parse(JSON.stringify(state));
  ["staff", "attendance", "leave", "payrollRuns", "taxBrackets", "access"].forEach(key => {
    if (!Array.isArray(copy[key])) throw new Error("Invalid workspace update: " + key + " must be a list.");
  });
  if (!copy.organization || !copy.contributions || typeof copy.payrollRulesVerified !== "boolean" ||
      typeof copy.payrollRulesEffectiveFrom !== "string" || typeof copy.payrollRulesSource !== "string" ||
      typeof copy.payrollRulesReference !== "string") {
    throw new Error("Invalid workspace update: organization and payroll settings are required.");
  }
  return copy;
}

function hrCanSeeEmployee_(actor) {
  if (actor.role === "admin" || actor.role === "hr" || actor.role === "payroll") return () => true;
  if (actor.role === "manager") return person => person.department === actor.department && person.status === "Active";
  return person => person.id === actor.employeeId;
}

function hrValidateEmployees_(staff, existing) {
  const ids = new Set(), emails = new Set();
  staff.forEach(person => {
    if (!person.id || !person.name || !person.email || !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(person.email)) throw new Error("Each employee needs an ID, name and valid unique email.");
    if (ids.has(person.id) || emails.has(person.email.toLowerCase())) throw new Error("Employee IDs and email addresses must be unique.");
    if (!["USD", "LRD"].includes(person.currency) || !Number.isFinite(Number(person.salary)) || Number(person.salary) < 0) throw new Error("Employee currency or monthly salary is invalid.");
    if (!["Active", "Inactive"].includes(person.status) || String(person.name).length > 160 || String(person.title).length > 120 || String(person.department).length > 120 || String(person.location).length > 120) throw new Error("Employee status or profile fields are invalid.");
    ids.add(person.id); emails.add(person.email.toLowerCase());
  });
  const nextIds = new Set(staff.map(person => person.id));
  existing.forEach(person => {
    if (!nextIds.has(person.id)) throw new Error("Employee records cannot be deleted; set the employment status to Inactive.");
  });
}

function hrValidateAttendanceChanges_(existing, incoming, actor) {
  const canSee = hrCanSeeEmployee_(actor);
  const employeeIndex = hrEmployeesById_(hrBook_());
  const visibleExisting = existing.filter(item => canSee({
    id: item.employeeId,
    department: employeeIndex[item.employeeId] ? String(employeeIndex[item.employeeId].department) : "",
    status: employeeIndex[item.employeeId] ? String(employeeIndex[item.employeeId].status) : ""
  }));
  const existingKeys = new Set(visibleExisting.map(item => item.employeeId + "|" + item.date));
  const incomingKeys = new Set(incoming.map(item => item.employeeId + "|" + item.date));
  if (incomingKeys.size !== incoming.length || [...existingKeys].some(key => !incomingKeys.has(key))) {
    throw new Error("Attendance records cannot be deleted or duplicated.");
  }
  if (actor.role !== "admin" && actor.role !== "hr" && actor.role !== "employee") {
    if (!hrSame_(existing, incoming)) throw new Error("You cannot modify attendance records.");
    return;
  }
  if (actor.role === "employee") {
    const oldOwn = visibleExisting.filter(item => item.employeeId === actor.employeeId);
    const newOwn = incoming.filter(item => item.employeeId === actor.employeeId);
    const today = Utilities.formatDate(new Date(), "Africa/Monrovia", "yyyy-MM-dd");
    const before = new Map(oldOwn.map(item => [item.date, item]));
    const after = new Map(newOwn.map(item => [item.date, item]));
    if (newOwn.length !== after.size || after.size < before.size) throw new Error("Attendance history cannot be deleted.");
    newOwn.forEach(item => {
      if (!hrValidAttendanceEvidence_(item) || item.date !== today || item.status !== "Present" || !hrIsTime_(item.in) ||
          (item.out && !hrIsTime_(item.out)) || (item.source && item.source !== "web")) {
        throw new Error("Employees may only use valid web clock-in/out times for today.");
      }
      const old = before.get(item.date);
      if (!old) {
        if (item.out) throw new Error("Clock in before clocking out.");
      } else if (old.in === "—" || !old.in) {
        if (item.out) throw new Error("Clock out is not valid before clocking in.");
      } else if (old.in === item.in) {
        if (old.out && item.out !== old.out) throw new Error("A recorded clock-out cannot be overwritten.");
        if (item.out && (!hrIsTime_(item.out) || hrMinutes_(item.out) <= hrMinutes_(old.in))) throw new Error("Clock-out must be later than the recorded clock-in.");
      } else if (!hrSame_(old, item)) {
        throw new Error("An existing clock-in cannot be overwritten.");
      }
    });
    if (incoming.some(item => item.employeeId !== actor.employeeId)) throw new Error("You cannot modify another employee's attendance.");
  } else {
    const before = new Map(existing.map(item => [item.employeeId + "|" + item.date, item]));
    incoming.forEach(item => {
      if (!item.employeeId || !hrIsDate_(item.date) ||
          !["Present", "Late", "Absent"].includes(item.status) ||
          (item.in !== "—" && item.in && !hrIsTime_(item.in)) ||
          (item.out && !hrIsTime_(item.out)) ||
          !hrValidAttendanceEvidence_(item) ||
          (item.out && item.in && item.in !== "—" && hrMinutes_(item.out) <= hrMinutes_(item.in))) throw new Error("Attendance record values are invalid.");
      if (!hrReadRows_(hrBook_(), "Employees").some(person => String(person.employee_id) === item.employeeId)) {
        throw new Error("Attendance must reference an existing employee.");
      }
      const old = before.get(item.employeeId + "|" + item.date);
      if (old && actor.role === "hr" && item.employeeId !== old.employeeId) throw new Error("Attendance history cannot be reassigned.");
    });
  }
}

function hrValidateLeaveChanges_(existing, incoming, actor) {
  const allowed = hrCanSeeEmployee_(actor);
  const employeeIndex = hrEmployeesById_(hrBook_());
  const before = new Map(existing.map(item => [item.id, item]));
  const after = new Map(incoming.map(item => [item.id, item]));
  if (after.size !== incoming.length) throw new Error("Leave request IDs must be unique.");
  existing.forEach(item => {
    const person = employeeIndex[item.employeeId];
    if (person && allowed({ id: item.employeeId, department: String(person.department), status: String(person.status) }) && !after.has(item.id)) {
      throw new Error("Leave requests cannot be deleted.");
    }
  });
  incoming.forEach(item => {
    const row = employeeIndex[item.employeeId];
    const person = row ? { id: item.employeeId, department: String(row.department), status: String(row.status) } : null;
    if (!person || !allowed(person)) throw new Error("Leave request is outside your access scope.");
    const old = before.get(item.id);
    if (!old) {
      if (actor.role !== "admin" && actor.role !== "hr" && actor.role !== "manager" && !(actor.role === "employee" && item.employeeId === actor.employeeId)) throw new Error("You cannot submit this leave request.");
      if (item.status !== "Pending" || !["Annual leave", "Sick leave", "Other leave"].includes(item.type) ||
          !hrIsDate_(item.start) || !hrIsDate_(item.end) ||
          item.end < item.start || !Number.isInteger(Number(item.days)) || Number(item.days) < 1 ||
          Number(item.days) > 366 || typeof item.reason !== "string" || item.reason.length > 240) throw new Error("Leave request details are invalid.");
      item.createdAt = new Date().toISOString();
    } else if (!hrSame_(old, item)) {
      const mayReview = actor.role === "admin" || actor.role === "hr" || actor.role === "manager";
      const statusOnly = { ...item, status: old.status, reviewedBy: old.reviewedBy || "", reviewedAt: old.reviewedAt || "" };
      const original = { ...old, reviewedBy: old.reviewedBy || "", reviewedAt: old.reviewedAt || "" };
      if (!mayReview || !hrSame_(statusOnly, original)) throw new Error("Only a reviewer may change the status of a leave request.");
      if ((item.status !== "Approved" && item.status !== "Rejected") || old.status !== "Pending") throw new Error("Only pending leave requests may be approved or rejected.");
      item.reviewedBy = actor.email;
      item.reviewedAt = new Date().toISOString();
    }
  });
}

function hrEmployeesById_(book) {
  return Object.fromEntries(hrReadRows_(book, "Employees").map(person => [
    String(person.employee_id), { department: String(person.department), status: String(person.status) }
  ]));
}

function hrValidatePayrollAppend_(existing, incoming, actor, full) {
  if (actor.role !== "admin" && actor.role !== "payroll") throw new Error("Only payroll officers and admins can prepare payroll.");
  if (!full.payrollRulesVerified || !full.taxBrackets.length || !full.organization.exchangeRateVerified) throw new Error("Payroll is blocked until an administrator verifies Liberia payroll rules and the LRD/USD exchange rate.");
  if (incoming.length < existing.length) throw new Error("Payroll history cannot be deleted.");
  for (let i = 0; i < existing.length; i++) if (!hrSame_(incoming[i], existing[i])) throw new Error("Recorded payroll runs are immutable.");
  const runIds = new Set(existing.map(run => run.id));
  incoming.slice(existing.length).forEach(run => {
    if (run.status !== "Draft" || !run.id || String(run.id).length > 80 || !run.period || String(run.period).length > 80 ||
        !["USD", "LRD"].includes(run.currency) || Number(run.count) < 1 ||
        !Number.isFinite(Number(run.gross)) || Number(run.gross) < 0 ||
        !Number.isFinite(Number(run.net)) || Number(run.net) < 0) {
      throw new Error("Only valid, positive-headcount payroll drafts can be recorded.");
    }
    if (runIds.has(run.id)) throw new Error("Payroll run IDs must be unique.");
    runIds.add(run.id);
  });
}

function hrMergeScoped_(existing, incoming, canSee, employeeKey, dateKey) {
  const employeeIndex = hrEmployeesById_(hrBook_());
  const visible = existing.filter(item => {
    const person = employeeIndex[String(item[employeeKey])];
    return person && canSee({ id: String(item[employeeKey]), department: person.department, status: person.status });
  });
  const visibleIds = new Set(visible.map(item => dateKey ? item[employeeKey] + "|" + item[dateKey] : item[employeeKey]));
  const kept = existing.filter(item => !visibleIds.has(dateKey ? item[employeeKey] + "|" + item[dateKey] : item[employeeKey]));
  incoming.forEach(item => {
    const key = dateKey ? item[employeeKey] + "|" + item[dateKey] : item[employeeKey];
    const person = employeeIndex[String(item[employeeKey])];
    if (person && canSee({ id: String(item[employeeKey]), department: person.department, status: person.status })) kept.push(item);
    else if (visibleIds.has(key)) throw new Error("Workspace update contains a record outside your access scope.");
  });
  existing.splice(0, existing.length, ...kept);
}

function hrWriteEmployees_(book, staff) {
  hrReplaceRows_(book, "Employees", staff.map(p => [p.id, p.name, p.email, p.department, p.title, p.location, p.type, p.currency, p.salary, p.status, p.start]));
}

function hrWriteAttendance_(book, attendance, actor) {
  const now = new Date();
  const rows = attendance.map(a => [
    a.employeeId, a.date, a.in, a.out, a.status, a.source || "web", now,
    a.inLatitude ?? "", a.inLongitude ?? "", a.inAccuracy ?? "", a.inPhotoId || "",
    a.outLatitude ?? "", a.outLongitude ?? "", a.outAccuracy ?? "", a.outPhotoId || ""
  ]);
  hrReplaceRows_(book, "Attendance", rows);
}

function hrValidAttendanceEvidence_(item) {
  const validCoordinate = (value, minimum, maximum) => value == null || value === "" ||
    (typeof value === "number" && Number.isFinite(value) && value >= minimum && value <= maximum);
  const validPhotoId = value => !value || (typeof value === "string" && /^[\w-]{20,200}$/.test(value));
  return validCoordinate(item.inLatitude, -90, 90) && validCoordinate(item.inLongitude, -180, 180) &&
    validCoordinate(item.outLatitude, -90, 90) && validCoordinate(item.outLongitude, -180, 180) &&
    validCoordinate(item.inAccuracy, 0, 100000) && validCoordinate(item.outAccuracy, 0, 100000) &&
    validPhotoId(item.inPhotoId) && validPhotoId(item.outPhotoId);
}

function hrNullableNumber_(value) {
  return value === "" || value === null || value === undefined ? null : Number(value);
}

function hrWriteLeave_(book, leave) {
  const now = new Date();
  hrReplaceRows_(book, "LeaveRequests", leave.map(item => [
    item.id, item.employeeId, item.type, item.start, item.end, item.days, item.reason, item.status,
    item.createdAt || now, item.reviewedBy || "", item.reviewedAt || ""
  ]));
}

function hrWritePayrollRuns_(book, runs, actor) {
  const existing = hrReadRows_(book, "PayrollRuns");
  const existingItems = hrReadRows_(book, "PayrollItems");
  const full = hrReadState_(book);
  const newRuns = [];
  const newItems = [];
  runs.slice(existing.length).forEach(run => {
    const eligible = full.staff.filter(person => person.status === "Active");
    const totals = eligible.reduce((sum, person) => {
      const rate = full.organization.exchangeRate;
      const grossLRD = person.currency === "USD" ? person.salary * rate : person.salary;
      const annual = grossLRD * 12;
      const taxLRD = full.taxBrackets.reduce((tax, band) => {
        const upper = band.to === null ? Infinity : band.to;
        return tax + Math.max(0, Math.min(annual, upper) - band.from) * band.rate / 100 / 12;
      }, 0);
      const employeeContributionLRD = grossLRD * full.contributions.employee / 100;
      const netLRD = Math.max(0, grossLRD - taxLRD - employeeContributionLRD);
      const asCurrency = value => person.currency === "USD" ? value / rate : value;
      newItems.push([
        run.id, person.id, person.currency, asCurrency(grossLRD), asCurrency(taxLRD),
        asCurrency(employeeContributionLRD), asCurrency(grossLRD * full.contributions.employer / 100),
        asCurrency(grossLRD * full.contributions.workInjury / 100), asCurrency(netLRD)
      ]);
      sum.grossLRD += grossLRD;
      sum.netLRD += netLRD;
      return sum;
    }, { grossLRD: 0, netLRD: 0 });
    const runCurrency = full.organization.currency;
    const convert = amount => runCurrency === "USD" ? amount / full.organization.exchangeRate : amount;
    newRuns.push([
      run.id, run.period, run.created || new Date(), eligible.length, runCurrency,
      convert(totals.grossLRD), convert(totals.netLRD), "Draft",
      hrSetting_(book, "payroll_rules_version") || "0", actor.email
    ]);
  });
  hrReplaceRows_(book, "PayrollItems", existingItems.map(item => [
    item.run_id, item.employee_id, item.currency, item.gross, item.income_tax,
    item.employee_contribution, item.employer_contribution, item.work_injury, item.net_pay
  ]).concat(newItems));
  hrReplaceRows_(book, "PayrollRuns", existing.map(item => [
    item.run_id, item.period, item.created_at, item.employee_count, item.currency,
    item.gross_total, item.net_total, item.status, item.rules_version, item.created_by
  ]).concat(newRuns));
}

function hrValidateOrganization_(organization) {
  const allowedCurrencies = ["USD", "LRD"];
  if (!organization.name || String(organization.name).length > 120 || organization.country !== "Liberia" ||
      organization.timezone !== "Africa/Monrovia" || !allowedCurrencies.includes(organization.currency) ||
      !Number.isFinite(Number(organization.exchangeRate)) || Number(organization.exchangeRate) <= 0 ||
      !Number.isFinite(Number(organization.workdayHours)) || Number(organization.workdayHours) < 1 || Number(organization.workdayHours) > 24 ||
      String(organization.exchangeRateSource || "").length > 240 ||
      (organization.exchangeRateVerified === true && (!String(organization.exchangeRateSource || "").trim() ||
        !hrIsDate_(String(organization.exchangeRateEffective || "")) ||
        organization.exchangeRateEffective > Utilities.formatDate(new Date(), "Africa/Monrovia", "yyyy-MM-dd")))) {
    throw new Error("Organization settings are invalid.");
  }
}

function hrWriteOrganization_(book, organization) {
  hrValidateOrganization_(organization);
  hrSetSetting_(book, "organization_name", organization.name);
  hrSetSetting_(book, "jurisdiction", organization.country);
  hrSetSetting_(book, "timezone", organization.timezone);
  hrSetSetting_(book, "reporting_currency", organization.currency);
  hrSetSetting_(book, "usd_to_lrd_rate", String(organization.exchangeRate));
  hrSetSetting_(book, "usd_to_lrd_rate_verified", organization.exchangeRateVerified === true ? "true" : "false");
  hrSetSetting_(book, "usd_to_lrd_source", String(organization.exchangeRateSource || "").slice(0, 240));
  hrSetSetting_(book, "usd_to_lrd_effective_date", String(organization.exchangeRateEffective || ""));
  hrSetSetting_(book, "workday_hours", String(organization.workdayHours));
}

function hrValidateTaxRules_(brackets, contributions) {
  if (!brackets.length) throw new Error("At least one income-tax band is required before payroll can be enabled.");
  for (let i = 0; i < brackets.length; i++) {
    const band = brackets[i];
    if (typeof band.from !== "number" || !Number.isFinite(band.from) || band.from < 0 ||
        (band.to !== null && (typeof band.to !== "number" || !Number.isFinite(band.to) || band.to <= band.from)) ||
        typeof band.rate !== "number" || !Number.isFinite(band.rate) || band.rate < 0 || band.rate > 100 ||
        (i > 0 && Number(band.from) !== Number(brackets[i - 1].to))) throw new Error("Progressive tax bands must be valid and contiguous.");
  }
  if (brackets[brackets.length - 1].to !== null) throw new Error("The final tax band must not have an upper limit.");
  const rates = [contributions.employee, contributions.employer, contributions.workInjury];
  if (rates.some(rate => typeof rate !== "number" || !Number.isFinite(rate) || rate < 0 || rate > 100)) throw new Error("Contribution rates must be between 0 and 100 percent.");
}

function hrWriteTaxRules_(book, brackets, contributions, actor, version, effectiveDate, verified, verificationSource, verificationReference) {
  hrValidateTaxRules_(brackets, contributions);
  if (String(verificationSource || "").length > 240 || String(verificationReference || "").length > 500) {
    throw new Error("Payroll rule source and reference are too long.");
  }
  const now = new Date();
  const rows = brackets.map((band, index) => [version, "income_tax", index + 1, band.from, band.to === null ? "" : band.to, band.rate, effectiveDate, verified, verificationSource || "", verificationReference || "", actor.email, now]);
  rows.push([version, "employee_contribution", "", "", "", contributions.employee, effectiveDate, verified, verificationSource || "", verificationReference || "", actor.email, now]);
  rows.push([version, "employer_contribution", "", "", "", contributions.employer, effectiveDate, verified, verificationSource || "", verificationReference || "", actor.email, now]);
  rows.push([version, "work_injury_contribution", "", "", "", contributions.workInjury, effectiveDate, verified, verificationSource || "", verificationReference || "", actor.email, now]);
  rows.forEach(row => hrAppendRow_(book, "TaxRules", row));
}

function hrWriteAccess_(book, access) {
  hrValidateAccess_(book, access, null);
  hrReplaceRows_(book, "Access", access.map(item => [item.email, item.role, item.active, item.employeeId || "", item.department || ""]));
}

function hrValidateAccess_(book, access, actor) {
  const seen = new Set();
  const employees = hrReadRows_(book, "Employees");
  access.forEach(item => {
    item.email = String(item.email || "").trim().toLowerCase();
    if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(item.email) || !HR_ROLES.includes(item.role) || typeof item.active !== "boolean" || seen.has(item.email)) {
      throw new Error("Access entries need unique valid email addresses, known roles and an active flag.");
    }
    if (item.role === "manager" && (!item.department || !employees.some(person => String(person.department) === item.department && person.status === "Active"))) {
      throw new Error("A manager must have a department containing at least one active employee.");
    }
    if (item.role === "employee" && !employees.some(person => String(person.employee_id) === item.employeeId && person.status === "Active")) {
      throw new Error("Employee access must link to an active employee record.");
    }
    seen.add(item.email);
  });
  if (!access.some(item => item.role === "admin" && item.active)) throw new Error("At least one active administrator must remain.");
  if (actor && !access.some(item => item.email === actor.email && item.role === "admin" && item.active)) throw new Error("You cannot revoke or demote your own active administrator account.");
}

function hrSame_(left, right) {
  return JSON.stringify(left) === JSON.stringify(right);
}

function hrFingerprint_(state) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, JSON.stringify(state), Utilities.Charset.UTF_8);
  return Utilities.base64EncodeWebSafe(bytes);
}

function hrSetting_(book, key) {
  const row = hrReadRows_(book, "Settings").find(item => item.key === key);
  return row ? String(row.value) : "";
}

function hrSetSetting_(book, key, value) {
  const sheet = book.getSheetByName("Settings");
  const rows = hrReadRows_(book, "Settings");
  const index = rows.findIndex(row => row.key === key);
  if (index < 0) hrAppendRow_(book, "Settings", [key, value]);
  else sheet.getRange(index + 2, 2).setValue(hrSafeCell_(String(value)));
}

function hrEnsureSheet_(book, name) {
  let sheet = book.getSheetByName(name);
  if (!sheet) sheet = book.insertSheet(name);
  const header = HR_SCHEMA[name];
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, header.length).setValues([header]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, header.length).setFontWeight("bold");
  } else {
    const existing = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    while (existing.length && existing[existing.length - 1] === "") existing.pop();
    const matchesPrefix = existing.length <= header.length &&
      existing.every((value, index) => value === header[index]);
    if (!matchesPrefix) throw new Error("Sheet " + name + " has an unexpected header. Back it up and fix the header before setup.");
    if (existing.length < header.length) {
      sheet.getRange(1, existing.length + 1, 1, header.length - existing.length)
        .setValues([header.slice(existing.length)]);
      sheet.getRange(1, 1, 1, header.length).setFontWeight("bold");
    }
  }
  return sheet;
}

function hrReadRows_(book, name) {
  const sheet = hrEnsureSheet_(book, name);
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0];
  return values.slice(1).filter(row => row.some(value => value !== "")).map(row =>
    Object.fromEntries(headers.map((header, index) => [header, row[index]]))
  );
}

function hrAppendRow_(book, name, row) {
  const sheet = hrEnsureSheet_(book, name);
  sheet.appendRow(row.map(hrSafeCell_));
}

function hrReplaceRows_(book, name, rows) {
  const sheet = hrEnsureSheet_(book, name);
  const columns = HR_SCHEMA[name].length;
  const oldCount = Math.max(0, sheet.getLastRow() - 1);
  if (oldCount) sheet.getRange(2, 1, oldCount, columns).clearContent();
  if (rows.length) sheet.getRange(2, 1, rows.length, columns).setValues(rows.map(row => row.map(hrSafeCell_)));
}

function hrSafeCell_(value) {
  if (typeof value === "string" && /^[\s]*[=+\-@]/.test(value)) return "'" + value;
  return value;
}

function hrProtectSheet_(sheet) {
  if (!sheet) return;
  const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
  const protection = protections.length ? protections[0] : sheet.protect();
  protection.setDescription("WorkSmart HR — edit through the authorized application only");
  protection.setWarningOnly(false);
  protection.removeEditors(protection.getEditors());
  if (protection.canDomainEdit()) protection.setDomainEdit(false);
}

function hrDate_(value) {
  if (!value) return "";
  if (value instanceof Date) return Utilities.formatDate(value, "Africa/Monrovia", "yyyy-MM-dd");
  return String(value).slice(0, 10);
}

function hrDateTime_(value) {
  if (!value) return "";
  return value instanceof Date ? value.toISOString() : String(value);
}

function hrAudit_(book, actor, action, entityType, entityId, before, after, revision) {
  hrAppendRow_(book, "AuditLog", [
    Utilities.getUuid(), new Date(), actor.email, actor.role, action, entityType, entityId,
    before == null ? "" : JSON.stringify(before).slice(0, 40000),
    after == null ? "" : JSON.stringify(after).slice(0, 40000), revision
  ]);
}
