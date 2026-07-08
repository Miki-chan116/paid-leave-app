#!/usr/bin/env node
"use strict";

const fs = require("fs");
const path = require("path");

const ROOT_DIR = path.resolve(__dirname, "../..");
const DEFAULT_INPUT_DIR = path.join(__dirname, "input");
const DEFAULT_OUTPUT_DIR = path.join(__dirname, "output");

const args = parseArgs(process.argv.slice(2));
const inputDir = path.resolve(args.input || DEFAULT_INPUT_DIR);
const outputDir = path.resolve(args.output || DEFAULT_OUTPUT_DIR);
const generatedAt = new Date().toISOString();

const INPUT_FILES = {
  employees: "employees.csv",
  leave_requests: "leave_requests.csv",
  paid_leave_grants: "paid_leave_grants.csv",
  company_calendar: "company_calendar.csv",
  usage_log: "usage_log.csv",
  admin_users: "admin_users.csv"
};

const OUTPUT_FILES = {
  employees: "employees.csv",
  leave_requests: "leave_requests.csv",
  paid_leave_grants: "paid_leave_grants.csv",
  company_calendar: "company_calendar.csv",
  usage_logs: "usage_logs.csv",
  admin_users: "admin_users.csv"
};

const OUTPUT_COLUMNS = {
  employees: [
    "employee_id",
    "display_employee_id",
    "name",
    "display_name",
    "name_kana",
    "company_code",
    "company_name",
    "department",
    "employment_type",
    "employment_status",
    "hire_date",
    "leave_date",
    "work_days_per_week",
    "fiscal_start_month",
    "leave_management_target",
    "initial_grant_check_target",
    "is_driver",
    "driver_type",
    "default_vehicle_id",
    "display_order",
    "notes",
    "created_at",
    "updated_at",
    "deleted_at"
  ],
  leave_requests: [
    "request_id",
    "employee_id",
    "request_date",
    "start_date",
    "end_date",
    "days",
    "type",
    "half_day",
    "reason",
    "reason_detail",
    "status",
    "approver_id",
    "approver_name",
    "approved_at",
    "rejected_reason",
    "cancel_reason",
    "year",
    "created_at",
    "updated_at"
  ],
  paid_leave_grants: [
    "grant_id",
    "employee_id",
    "grant_date",
    "grant_days",
    "carry_over_days",
    "valid_from",
    "valid_to",
    "grant_type",
    "year",
    "notes",
    "is_finalized",
    "finalized_at",
    "created_at",
    "updated_at"
  ],
  company_calendar: [
    "date",
    "type",
    "notes",
    "created_at",
    "updated_at"
  ],
  usage_logs: [
    "log_id",
    "legacy_request_id",
    "target_type",
    "target_id",
    "leave_request_id",
    "employee_id",
    "action_type",
    "operator_id",
    "operator_name",
    "action_date",
    "comment",
    "created_at",
    "updated_at"
  ],
  admin_users: [
    "admin_id",
    "admin_name",
    "pin",
    "role",
    "notes",
    "is_active",
    "created_at",
    "updated_at"
  ]
};

const ENUMS = {
  leaveRequestStatus: new Set(["pending", "approved", "rejected", "canceled", "canceled_by_admin"]),
  leaveRequestType: new Set(["paid_leave"]),
  halfDay: new Set(["", "am", "pm"]),
  grantType: new Set(["initial", "six_month", "six_month_processed", "six_month_skipped", "yearly"]),
  calendarType: new Set(["workday", "holiday", "no_leave"])
};

const report = {
  generated_at: generatedAt,
  input_dir: inputDir,
  output_dir: outputDir,
  dry_run: false,
  tables: {},
  excluded: [],
  warnings: [],
  errors: [],
  skipped_blank_rows: {}
};

main();

function main() {
  ensureDir(outputDir);

  const input = {};
  for (const [table, filename] of Object.entries(INPUT_FILES)) {
    input[table] = readInputCsv(table, filename);
  }

  const converted = {};
  converted.employees = convertEmployees(input.employees.rows);
  const employeeIds = new Set(converted.employees.rows.map(row => row.employee_id).filter(Boolean));

  converted.leave_requests = convertLeaveRequests(input.leave_requests.rows, employeeIds);
  const requestIds = new Set(converted.leave_requests.rows.map(row => row.request_id).filter(Boolean));

  converted.paid_leave_grants = convertPaidLeaveGrants(input.paid_leave_grants.rows, employeeIds);
  converted.company_calendar = convertCompanyCalendar(input.company_calendar.rows);
  converted.admin_users = convertAdminUsers(input.admin_users.rows);
  converted.usage_logs = convertUsageLogs(input.usage_log.rows, employeeIds, requestIds);

  for (const [table, result] of Object.entries(converted)) {
    report.tables[table] = {
      input_rows: result.inputRows,
      excluded_rows: result.excludedRows,
      output_rows: result.rows.length
    };
  }

  writeReports();

  if (report.errors.length > 0) {
    console.error("Migration conversion failed. See migration_errors.json.");
    process.exitCode = 1;
    return;
  }

  for (const [table, result] of Object.entries(converted)) {
    writeCsv(path.join(outputDir, OUTPUT_FILES[table]), OUTPUT_COLUMNS[table], result.rows);
  }

  writeReports();
  console.log("Migration CSV conversion completed.");
  console.log("Output:", outputDir);
}

function parseArgs(argv) {
  const result = {};
  for (let i = 0; i < argv.length; i++) {
    const arg = argv[i];
    if (arg === "--input") result.input = argv[++i];
    else if (arg === "--output") result.output = argv[++i];
    else if (arg === "--help" || arg === "-h") {
      printHelp();
      process.exit(0);
    }
  }
  return result;
}

function printHelp() {
  const requiredFiles = [
    "employees.csv",
    "leave_requests.csv",
    "paid_leave_grants.csv",
    "company_calendar.csv",
    "usage_log.csv",
    "admin_users.csv"
  ];
  console.log([
    "Usage:",
    "  node database/migration/convert_spreadsheet_csv_to_supabase.js",
    "  node database/migration/convert_spreadsheet_csv_to_supabase.js --input ./input --output ./output",
    "",
    "Required input files:",
    ...requiredFiles.map(name => "  - " + name)
  ].join("\n"));
}

function readInputCsv(table, filename) {
  const filePath = path.join(inputDir, filename);
  if (!fs.existsSync(filePath)) {
    addError(table, "missing_input_file", `Input file not found: ${filename}`);
    return { headers: [], rows: [] };
  }

  const text = fs.readFileSync(filePath, "utf8").replace(/^\uFEFF/, "");
  const records = parseCsv(text);
  if (records.length === 0) return { headers: [], rows: [] };

  const headers = records[0].map(normalizeHeader);
  let skippedBlankRows = 0;
  const rows = [];

  records.slice(1).forEach((record, index) => {
    if (isCsvRecordBlank(record)) {
      skippedBlankRows++;
      return;
    }

    const row = {};
    headers.forEach((header, colIndex) => {
      if (!header) return;
      row[header] = record[colIndex] == null ? "" : record[colIndex];
    });
    row.__row_number = index + 2;
    row.__raw_record = record;
    row.__headers = headers;

    if (isSourceRowEmpty(row, headers)) {
      skippedBlankRows++;
      return;
    }

    if (table === "company_calendar" && isBlankCompanyCalendarDefaultRow(row)) {
      skippedBlankRows++;
      return;
    }

    rows.push(row);
  });

  report.skipped_blank_rows[table] = skippedBlankRows;

  return { headers, rows };
}

function convertEmployees(rows) {
  const result = baseResult(rows);
  const seen = new Set();

  rows.forEach(row => {
    const employeeId = text(row.employee_id);
    if (shouldExcludeByIds("employees", row, [employeeId])) {
      result.excludedRows++;
      return;
    }

    if (!employeeId) addError("employees", "missing_employee_id", "employee_id is required", row);
    else if (seen.has(employeeId)) addError("employees", "duplicate_employee_id", `Duplicate employee_id: ${employeeId}`, row);
    else seen.add(employeeId);

    const out = {
      employee_id: employeeId,
      display_employee_id: text(row.display_employee_id),
      name: text(row.name),
      display_name: text(row.display_name),
      name_kana: text(row.name_kana),
      company_code: upper(row.company_code),
      company_name: text(row.company_name),
      department: text(row.department),
      employment_type: text(row.employment_type),
      employment_status: normalizeEmploymentStatus(row.employment_status),
      hire_date: normalizeDate(row.hire_date, "employees", "hire_date", row),
      leave_date: normalizeDate(row.leave_date, "employees", "leave_date", row),
      work_days_per_week: normalizeNumber(row.work_days_per_week, "employees", "work_days_per_week", row),
      fiscal_start_month: normalizeInteger(row.fiscal_start_month || "4", "employees", "fiscal_start_month", row),
      leave_management_target: normalizeBoolean(row.leave_management_target, "employees", "leave_management_target", row, false),
      initial_grant_check_target: normalizeBoolean(row.initial_grant_check_target, "employees", "initial_grant_check_target", row, true),
      is_driver: normalizeBoolean(row.is_driver, "employees", "is_driver", row, false),
      driver_type: text(row.driver_type),
      default_vehicle_id: text(row.default_vehicle_id),
      display_order: normalizeInteger(row.display_order, "employees", "display_order", row),
      notes: text(row.notes),
      created_at: normalizeTimestamp(row.created_at, "employees", "created_at", row) || generatedAt,
      updated_at: normalizeTimestamp(row.updated_at, "employees", "updated_at", row) || generatedAt,
      deleted_at: normalizeTimestamp(row.deleted_at, "employees", "deleted_at", row)
    };

    if (!out.name) addError("employees", "missing_name", "name is required", row);
    if (!out.company_code) addError("employees", "missing_company_code", "company_code is required", row);

    result.rows.push(out);
  });

  return result;
}

function convertLeaveRequests(rows, employeeIds) {
  const result = baseResult(rows);
  const seen = new Set();

  rows.forEach(row => {
    const requestId = text(row.request_id);
    const employeeId = text(row.employee_id);
    if (shouldExcludeByIds("leave_requests", row, [requestId, employeeId])) {
      result.excludedRows++;
      return;
    }

    if (!requestId) addError("leave_requests", "missing_request_id", "request_id is required", row);
    else if (seen.has(requestId)) addError("leave_requests", "duplicate_request_id", `Duplicate request_id: ${requestId}`, row);
    else seen.add(requestId);

    if (!employeeId || !employeeIds.has(employeeId)) {
      addError("leave_requests", "missing_employee_fk", `employee_id not found in employees: ${employeeId}`, row);
    }

    const status = lower(row.status || "pending");
    if (!ENUMS.leaveRequestStatus.has(status)) {
      addError("leave_requests", "invalid_status", `Invalid status: ${status}`, row);
    }

    const requestType = lower(row.type || "paid_leave");
    if (!ENUMS.leaveRequestType.has(requestType)) {
      addError("leave_requests", "invalid_type", `Invalid type: ${requestType}`, row);
    }

    const halfDay = lower(row.half_day);
    if (!ENUMS.halfDay.has(halfDay)) {
      addError("leave_requests", "invalid_half_day", `Invalid half_day: ${halfDay}`, row);
    }

    result.rows.push({
      request_id: requestId,
      employee_id: employeeId,
      request_date: normalizeTimestamp(row.request_date, "leave_requests", "request_date", row),
      start_date: normalizeDate(row.start_date, "leave_requests", "start_date", row),
      end_date: normalizeDate(row.end_date, "leave_requests", "end_date", row),
      days: normalizeNumber(row.days, "leave_requests", "days", row),
      type: requestType || "paid_leave",
      half_day: halfDay,
      reason: text(row.reason),
      reason_detail: text(row.reason_detail),
      status,
      approver_id: text(row.approver_id),
      approver_name: text(row.approver_name),
      approved_at: normalizeTimestamp(row.approved_at, "leave_requests", "approved_at", row),
      rejected_reason: text(row.rejected_reason),
      cancel_reason: firstText(row.cancel_reason, row.canceled_reason, row.cancelled_reason),
      year: normalizeInteger(row.year, "leave_requests", "year", row),
      created_at: normalizeTimestamp(row.created_at, "leave_requests", "created_at", row) || generatedAt,
      updated_at: normalizeTimestamp(row.updated_at, "leave_requests", "updated_at", row) || generatedAt
    });
  });

  return result;
}

function convertPaidLeaveGrants(rows, employeeIds) {
  const result = baseResult(rows);
  const seen = new Set();

  rows.forEach(row => {
    const grantId = text(row.grant_id);
    const employeeId = text(row.employee_id);
    if (shouldExcludeByIds("paid_leave_grants", row, [grantId, employeeId])) {
      result.excludedRows++;
      return;
    }

    if (!grantId) addError("paid_leave_grants", "missing_grant_id", "grant_id is required", row);
    else if (seen.has(grantId)) addError("paid_leave_grants", "duplicate_grant_id", `Duplicate grant_id: ${grantId}`, row);
    else seen.add(grantId);

    if (!employeeId || !employeeIds.has(employeeId)) {
      addError("paid_leave_grants", "missing_employee_fk", `employee_id not found in employees: ${employeeId}`, row);
    }

    const grantType = lower(row.grant_type);
    if (!ENUMS.grantType.has(grantType)) {
      addError("paid_leave_grants", "invalid_grant_type", `Invalid grant_type: ${grantType}`, row);
    }

    result.rows.push({
      grant_id: grantId,
      employee_id: employeeId,
      grant_date: normalizeDate(row.grant_date, "paid_leave_grants", "grant_date", row),
      grant_days: normalizeNumber(row.grant_days || "0", "paid_leave_grants", "grant_days", row),
      carry_over_days: normalizeNumber(row.carry_over_days || "0", "paid_leave_grants", "carry_over_days", row),
      valid_from: normalizeDate(row.valid_from, "paid_leave_grants", "valid_from", row),
      valid_to: normalizeDate(row.valid_to, "paid_leave_grants", "valid_to", row),
      grant_type: grantType,
      year: normalizeInteger(row.year, "paid_leave_grants", "year", row),
      notes: text(row.notes),
      is_finalized: normalizeBoolean(row.is_finalized, "paid_leave_grants", "is_finalized", row, true),
      finalized_at: normalizeTimestamp(row.finalized_at, "paid_leave_grants", "finalized_at", row),
      created_at: normalizeTimestamp(row.created_at, "paid_leave_grants", "created_at", row) || generatedAt,
      updated_at: normalizeTimestamp(row.updated_at, "paid_leave_grants", "updated_at", row) || generatedAt
    });
  });

  return result;
}

function convertCompanyCalendar(rows) {
  const result = baseResult(rows);
  const seen = new Set();

  rows.forEach(row => {
    const date = normalizeDate(row.date, "company_calendar", "date", row);
    if (shouldExcludeByIds("company_calendar", row, [date])) {
      result.excludedRows++;
      return;
    }
    if (!date) {
      addBlankDateDebugWarning("company_calendar", row);
      addError("company_calendar", "missing_date", "date is required", row);
      return;
    }

    if (seen.has(date)) addError("company_calendar", "duplicate_date", `Duplicate date: ${date}`, row);
    else seen.add(date);

    const type = lower(row.type || "workday");
    if (!ENUMS.calendarType.has(type)) {
      addError("company_calendar", "invalid_type", `Invalid calendar type: ${type}`, row);
    }

    result.rows.push({
      date,
      type,
      notes: text(row.notes),
      created_at: normalizeTimestamp(row.created_at, "company_calendar", "created_at", row) || generatedAt,
      updated_at: normalizeTimestamp(row.updated_at, "company_calendar", "updated_at", row) || generatedAt
    });
  });

  return result;
}

function convertUsageLogs(rows, employeeIds, requestIds) {
  const result = baseResult(rows);
  const seen = new Set();

  rows.forEach(row => {
    const logId = text(row.log_id);
    const legacyRequestId = text(row.request_id);
    if (shouldExcludeByIds("usage_logs", row, [logId, legacyRequestId])) {
      result.excludedRows++;
      return;
    }

    if (!logId) addError("usage_logs", "missing_log_id", "log_id is required", row);
    else if (seen.has(logId)) addError("usage_logs", "duplicate_log_id", `Duplicate log_id: ${logId}`, row);
    else seen.add(logId);

    let targetType = "legacy";
    let targetId = legacyRequestId;
    let leaveRequestId = "";
    let employeeId = "";

    if (legacyRequestId && requestIds.has(legacyRequestId)) {
      targetType = "leave_request";
      leaveRequestId = legacyRequestId;
    } else if (legacyRequestId && employeeIds.has(legacyRequestId)) {
      targetType = "employee";
      employeeId = legacyRequestId;
    } else if (!legacyRequestId) {
      targetType = "unknown";
      targetId = "";
    } else {
      addWarning("usage_logs", "legacy_request_id_unclassified", `Could not classify legacy request_id: ${legacyRequestId}`, row);
    }

    result.rows.push({
      log_id: logId,
      legacy_request_id: legacyRequestId,
      target_type: targetType,
      target_id: targetId,
      leave_request_id: leaveRequestId,
      employee_id: employeeId,
      action_type: text(row.action_type),
      operator_id: text(row.operator_id),
      operator_name: text(row.operator_name),
      action_date: normalizeTimestamp(row.action_date, "usage_logs", "action_date", row) || generatedAt,
      comment: text(row.comment),
      created_at: normalizeTimestamp(row.created_at, "usage_logs", "created_at", row) || generatedAt,
      updated_at: normalizeTimestamp(row.updated_at, "usage_logs", "updated_at", row) || generatedAt
    });
  });

  return result;
}

function convertAdminUsers(rows) {
  const result = baseResult(rows);
  const seen = new Set();

  rows.forEach(row => {
    const adminId = text(row.admin_id);
    if (shouldExcludeByIds("admin_users", row, [adminId])) {
      result.excludedRows++;
      return;
    }

    if (!adminId) addError("admin_users", "missing_admin_id", "admin_id is required", row);
    else if (seen.has(adminId)) addError("admin_users", "duplicate_admin_id", `Duplicate admin_id: ${adminId}`, row);
    else seen.add(adminId);

    result.rows.push({
      admin_id: adminId,
      admin_name: text(row.admin_name),
      pin: text(row.pin),
      role: text(row.role),
      notes: text(row.notes),
      is_active: normalizeBoolean(row.is_active, "admin_users", "is_active", row, true),
      created_at: normalizeTimestamp(row.created_at, "admin_users", "created_at", row) || generatedAt,
      updated_at: normalizeTimestamp(row.updated_at, "admin_users", "updated_at", row) || generatedAt
    });
  });

  return result;
}

function baseResult(rows) {
  return {
    inputRows: rows.length,
    excludedRows: 0,
    rows: []
  };
}

function shouldExcludeByIds(table, row, values) {
  const matched = values
    .map(value => text(value))
    .filter(value => isTestId(value));
  if (matched.length === 0) return false;

  report.excluded.push({
    table,
    row_number: row.__row_number,
    reason: "test_data",
    matched_values: matched
  });
  return true;
}

function isTestId(value) {
  const normalized = text(value).toUpperCase();
  return normalized === "TEST-FIFO-001" ||
    normalized === "TEST-001" ||
    normalized.startsWith("TEST-");
}

function normalizeEmploymentStatus(value) {
  const normalized = text(value) || "active";
  if (normalized === "在職") return "active";
  if (normalized === "休職") return "leave";
  if (normalized === "退職") return "retired";
  return normalized;
}

function normalizeDate(value, table, column, row) {
  const raw = clean(value);
  if (!raw) return "";

  const date = parseDateLike(raw);
  if (!date) {
    addError(table, "invalid_date", `Invalid date in ${column}: ${raw}`, row);
    return "";
  }

  return formatDate(date);
}

function normalizeTimestamp(value, table, column, row) {
  const raw = clean(value);
  if (!raw) return "";

  const date = parseDateLike(raw);
  if (!date) {
    addError(table, "invalid_timestamp", `Invalid timestamp in ${column}: ${raw}`, row);
    return "";
  }

  return date.toISOString();
}

function parseDateLike(value) {
  const raw = text(value);
  if (!raw) return null;

  const numeric = Number(raw);
  if (isFinite(numeric) && numeric > 20000 && numeric < 80000) {
    return new Date(Math.round((numeric - 25569) * 86400 * 1000));
  }

  const normalized = raw
    .replace(/[年月]/g, "/")
    .replace(/日/g, "")
    .replace(/\./g, "/")
    .replace(/\s+/g, " ")
    .trim();

  const match = normalized.match(/^(\d{4})[/-](\d{1,2})[/-](\d{1,2})(?:[ T](\d{1,2}):(\d{1,2})(?::(\d{1,2}))?)?$/);
  if (match) {
    const year = Number(match[1]);
    const month = Number(match[2]);
    const day = Number(match[3]);
    const hour = Number(match[4] || 0);
    const minute = Number(match[5] || 0);
    const second = Number(match[6] || 0);
    const date = new Date(year, month - 1, day, hour, minute, second);
    if (
      date.getFullYear() === year &&
      date.getMonth() === month - 1 &&
      date.getDate() === day
    ) {
      return date;
    }
    return null;
  }

  const parsed = new Date(raw);
  if (isNaN(parsed.getTime())) return null;
  return parsed;
}

function formatDate(date) {
  return [
    date.getFullYear(),
    String(date.getMonth() + 1).padStart(2, "0"),
    String(date.getDate()).padStart(2, "0")
  ].join("-");
}

function normalizeBoolean(value, table, column, row, defaultValue) {
  const raw = clean(value);
  if (!raw) return defaultValue ? "true" : "false";
  const upper = raw.toUpperCase();
  if (upper === "TRUE" || upper === "1" || upper === "YES" || raw === "はい") return "true";
  if (upper === "FALSE" || upper === "0" || upper === "NO" || raw === "いいえ") return "false";

  addError(table, "invalid_boolean", `Invalid boolean in ${column}: ${raw}`, row);
  return defaultValue ? "true" : "false";
}

function normalizeNumber(value, table, column, row) {
  const raw = clean(value);
  if (!raw) return "";
  const normalized = raw.replace(/,/g, "");
  const num = Number(normalized);
  if (!isFinite(num)) {
    addError(table, "invalid_number", `Invalid number in ${column}: ${raw}`, row);
    return "";
  }
  if (num < 0) {
    addError(table, "negative_number", `Negative number in ${column}: ${raw}`, row);
  }
  return String(num);
}

function normalizeInteger(value, table, column, row) {
  const raw = clean(value);
  if (!raw) return "";
  const num = Number(raw.replace(/,/g, ""));
  if (!Number.isInteger(num)) {
    addError(table, "invalid_integer", `Invalid integer in ${column}: ${raw}`, row);
    return "";
  }
  return String(num);
}

function text(value) {
  if (value == null) return "";
  return String(value).trim();
}

function firstText(...values) {
  for (const value of values) {
    const normalized = text(value);
    if (normalized) return normalized;
  }
  return "";
}

function clean(value) {
  return String(value == null ? "" : value)
    .replace(/\r/g, "")
    .replace(/^\uFEFF/, "")
    .replace(/\uFEFF/g, "")
    .replace(/[\u200B-\u200D\uFEFF]/g, "")
    .replace(/[\u00A0\u1680\u180E\u2000-\u200A\u2028\u2029\u202F\u205F\u3000]/g, " ")
    .replace(/[\u0000-\u001F\u007F]/g, "")
    .trim();
}

function isCsvRecordBlank(record) {
  return (record || []).every(value => clean(value) === "");
}

function isSourceRowEmpty(row, headers) {
  const headerCellsAreBlank = headers
    .filter(Boolean)
    .every(header => clean(row[header]) === "");
  const rawCellsAreBlank = isCsvRecordBlank(row.__raw_record || []);
  return headerCellsAreBlank && rawCellsAreBlank;
}

function isBlankCompanyCalendarDefaultRow(row) {
  const date = clean(row.date);
  if (date) return false;

  const type = clean(row.type).toLowerCase();
  const ignorableType = !type || type === "workday";
  if (!ignorableType) return false;

  return ["notes", "created_at", "updated_at"].every(column => clean(row[column]) === "");
}

function addBlankDateDebugWarning(table, row) {
  if (table !== "company_calendar") return;

  const existing = report.warnings.filter(item => {
    return item.table === table && item.code === "blank_date_row_has_values";
  }).length;
  if (existing >= 20) return;

  const values = getNonBlankRowValues(row);
  addWarning(table, "blank_date_row_has_values", "date is blank but other cells are not blank after normalization", row);
  report.warnings[report.warnings.length - 1].values = values;
}

function getNonBlankRowValues(row) {
  const values = [];
  const headers = row.__headers || [];
  const rawRecord = row.__raw_record || [];

  rawRecord.forEach((value, index) => {
    const normalized = clean(value);
    if (!normalized) return;
    values.push({
      column: headers[index] || `__extra_column_${index + 1}`,
      raw_value: String(value == null ? "" : value),
      normalized_value: normalized
    });
  });
  return values;
}

function lower(value) {
  return text(value).toLowerCase();
}

function upper(value) {
  return text(value).toUpperCase();
}

function normalizeHeader(value) {
  return text(value).replace(/^\uFEFF/, "");
}

function addError(table, code, message, row) {
  report.errors.push({
    table,
    code,
    message,
    row_number: row && row.__row_number ? row.__row_number : null
  });
}

function addWarning(table, code, message, row) {
  report.warnings.push({
    table,
    code,
    message,
    row_number: row && row.__row_number ? row.__row_number : null
  });
}

function ensureDir(dir) {
  fs.mkdirSync(dir, { recursive: true });
}

function writeReports() {
  ensureDir(outputDir);
  fs.writeFileSync(path.join(outputDir, "migration_report.json"), JSON.stringify(report, null, 2) + "\n");
  fs.writeFileSync(path.join(outputDir, "migration_errors.json"), JSON.stringify(report.errors, null, 2) + "\n");
  fs.writeFileSync(path.join(outputDir, "migration_warnings.json"), JSON.stringify(report.warnings, null, 2) + "\n");
}

function writeCsv(filePath, columns, rows) {
  const lines = [columns.map(csvEscape).join(",")];
  rows.forEach(row => {
    lines.push(columns.map(column => csvEscape(row[column] == null ? "" : row[column])).join(","));
  });
  fs.writeFileSync(filePath, lines.join("\n") + "\n");
}

function csvEscape(value) {
  const str = String(value == null ? "" : value);
  if (/[",\n\r]/.test(str)) {
    return '"' + str.replace(/"/g, '""') + '"';
  }
  return str;
}

function parseCsv(textValue) {
  const rows = [];
  let row = [];
  let cell = "";
  let inQuotes = false;

  for (let i = 0; i < textValue.length; i++) {
    const char = textValue[i];
    const next = textValue[i + 1];

    if (inQuotes) {
      if (char === '"' && next === '"') {
        cell += '"';
        i++;
      } else if (char === '"') {
        inQuotes = false;
      } else {
        cell += char;
      }
      continue;
    }

    if (char === '"') {
      inQuotes = true;
    } else if (char === ",") {
      row.push(cell);
      cell = "";
    } else if (char === "\n") {
      row.push(cell.replace(/\r$/, ""));
      rows.push(row);
      row = [];
      cell = "";
    } else {
      cell += char;
    }
  }

  if (cell.length > 0 || row.length > 0) {
    row.push(cell.replace(/\r$/, ""));
    rows.push(row);
  }

  return rows;
}
