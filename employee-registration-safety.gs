/*
 * 社員登録安全化 Phase 1: 純粋な入力・操作・採番判定。
 * Spreadsheet 等への接続と SHA-256 の実行環境アダプターは後続 Phase で実装する。
 */

const EMPLOYEE_REGISTRATION_INPUT_FIELDS_ = [
  "name", "display_name", "name_kana", "company_code", "company_name",
  "department", "employment_type", "employment_status", "hire_date",
  "leave_date", "work_days_per_week", "work_start_minute", "work_end_minute",
  "fiscal_start_month", "leave_management_target", "is_driver", "driver_type",
  "default_vehicle_id", "notes"
];
const EMPLOYEE_REGISTRATION_ID_BASELINE_ = { EMP: 82, W: 61, P: 22 };
const EMPLOYEE_REGISTRATION_HASH_VERSION_ = "v1";

function normalizeEmployeeRegistrationInput_(input) {
  if (!input || typeof input !== "object" || Array.isArray(input)) {
    throw new Error("EMPLOYEE_REGISTRATION_INPUT_INVALID");
  }
  const unknown = Object.keys(input).filter(key =>
    EMPLOYEE_REGISTRATION_INPUT_FIELDS_.indexOf(key) === -1);
  if (unknown.length) throw new Error("EMPLOYEE_REGISTRATION_UNKNOWN_FIELD: " + unknown.join(","));

  const out = {};
  EMPLOYEE_REGISTRATION_INPUT_FIELDS_.forEach(key => {
    const value = input[key];
    if (key === "leave_management_target" || key === "is_driver") {
      out[key] = normalizeEmployeeRegistrationBoolean_(value);
    } else if (key === "work_days_per_week" || key === "work_start_minute" ||
        key === "work_end_minute" || key === "fiscal_start_month") {
      out[key] = normalizeEmployeeRegistrationNumber_(value);
    } else {
      if (value != null && typeof value !== "string") {
        throw new Error("EMPLOYEE_REGISTRATION_FIELD_TYPE_INVALID: " + key);
      }
      let text = String(value == null ? "" : value).trim();
      if (key === "company_code") text = text.toUpperCase();
      if (key === "employment_type" || key === "employment_status") text = text.toLowerCase();
      if (key === "hire_date" || key === "leave_date") text = text.replace(/\//g, "-");
      out[key] = text;
    }
  });
  if (out.fiscal_start_month === "") {
    out.fiscal_start_month = out.company_code === "MAIN" ? 4 :
      out.company_code === "PARTNER" ? 6 : "";
  }
  return out;
}

function normalizeEmployeeRegistrationBoolean_(value) {
  if (value === true || value === false) return value;
  const text = String(value == null ? "" : value).trim().toUpperCase();
  if (text === "TRUE") return true;
  if (text === "FALSE") return false;
  return text;
}

function normalizeEmployeeRegistrationNumber_(value) {
  if (value === null || value === undefined || value === "") return "";
  if (typeof value === "number") return value;
  const text = String(value).trim();
  return /^\d+$/.test(text) ? Number(text) : text;
}

function isEmployeeRegistrationDateKey_(text) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(text)) return false;
  const year = Number(text.slice(0, 4));
  const month = Number(text.slice(5, 7));
  const day = Number(text.slice(8, 10));
  if (year < 1900 || month < 1 || month > 12 || day < 1 || day > 31) return false;
  const date = new Date(Date.UTC(year, month - 1, day));
  return date.getUTCFullYear() === year && date.getUTCMonth() === month - 1 &&
    date.getUTCDate() === day;
}

function employeeRegistrationWorkingMinutes_(start, end) {
  const breaks = [[600, 630], [720, 780], [900, 930]];
  return (end - start) - breaks.reduce((total, period) =>
    total + Math.max(0, Math.min(end, period[1]) - Math.max(start, period[0])), 0);
}

function validateEmployeeRegistrationInput_(input) {
  const errors = [];
  const add = (field, code) => errors.push({ field: field, code: code });
  const required = ["name", "name_kana", "company_code", "employment_type", "employment_status"];
  required.forEach(key => { if (!input[key]) add(key, "required"); });
  ["name", "display_name", "name_kana", "company_name", "department", "default_vehicle_id"]
    .forEach(key => { if (String(input[key] || "").length > 200) add(key, "too_long"); });
  if (String(input.notes || "").length > 2000) add("notes", "too_long");

  if (input.company_code && ["MAIN", "PARTNER"].indexOf(input.company_code) === -1)
    add("company_code", "invalid");
  if (input.employment_type &&
      ["regular", "partner", "part_time", "trainee", "executive", "other"]
        .indexOf(input.employment_type) === -1) add("employment_type", "invalid");
  if (input.employment_status === "retired") add("employment_status", "retired_forbidden");
  else if (input.employment_status && ["active", "leave"].indexOf(input.employment_status) === -1)
    add("employment_status", "invalid");
  if (typeof input.leave_management_target !== "boolean") add("leave_management_target", "invalid_boolean");
  if (typeof input.is_driver !== "boolean") add("is_driver", "invalid_boolean");

  if (input.leave_management_target === true && !input.hire_date) add("hire_date", "required_for_paid_leave");
  if (input.hire_date && !isEmployeeRegistrationDateKey_(input.hire_date)) add("hire_date", "invalid_date");
  if (input.leave_date) add("leave_date", "forbidden_for_new_employee");
  const workDays = input.work_days_per_week;
  if (input.leave_management_target === true && workDays === "")
    add("work_days_per_week", "required_for_paid_leave");
  else if (workDays !== "" && (!Number.isInteger(workDays) || workDays < 1 || workDays > 5))
    add("work_days_per_week", "invalid");
  const fiscalMonth = input.company_code === "MAIN" ? 4 :
    input.company_code === "PARTNER" ? 6 : null;
  if (fiscalMonth !== null && input.fiscal_start_month !== fiscalMonth)
    add("fiscal_start_month", "company_mismatch");

  if (input.is_driver === true && !input.driver_type) add("driver_type", "required_for_driver");
  if (input.driver_type && ["専任運転手", "兼任運転手"].indexOf(input.driver_type) === -1)
    add("driver_type", "invalid");
  if (input.is_driver === false && (input.driver_type || input.default_vehicle_id))
    add("is_driver", "driver_fields_without_driver");

  const start = input.work_start_minute;
  const end = input.work_end_minute;
  if ((start === "") !== (end === "")) add("work_start_minute", "work_time_pair_required");
  if (start !== "" && (!Number.isInteger(start) || start < 0 || start > 1439))
    add("work_start_minute", "invalid");
  if (end !== "" && (!Number.isInteger(end) || end < 0 || end > 1439))
    add("work_end_minute", "invalid");
  if (Number.isInteger(start) && Number.isInteger(end) &&
      start >= 0 && start <= 1439 && end >= 0 && end <= 1439) {
    if (end <= start) add("work_end_minute", "not_after_start");
    else if (input.company_code === "PARTNER")
      add("work_end_minute", "unsupported_for_partner");
    else if (input.company_code === "MAIN" && employeeRegistrationWorkingMinutes_(start, end) !== 420)
      add("work_end_minute", "scheduled_minutes_mismatch");
  }
  return { ok: errors.length === 0, errors: errors };
}

function employeeRegistrationCanonicalPayload_(normalizedInput) {
  if (!normalizedInput || EMPLOYEE_REGISTRATION_INPUT_FIELDS_.some(key =>
      !Object.prototype.hasOwnProperty.call(normalizedInput, key))) {
    throw new Error("EMPLOYEE_REGISTRATION_CANONICAL_INPUT_INVALID");
  }
  const payload = {};
  EMPLOYEE_REGISTRATION_INPUT_FIELDS_.forEach(key => { payload[key] = normalizedInput[key]; });
  return JSON.stringify(payload);
}

function employeeRegistrationInputHash_(normalizedInput, sha256Hex) {
  if (typeof sha256Hex !== "function") throw new Error("EMPLOYEE_REGISTRATION_DIGEST_REQUIRED");
  let normalized;
  try { normalized = normalizeEmployeeRegistrationInput_(normalizedInput); }
  catch (error) { throw new Error("EMPLOYEE_REGISTRATION_INPUT_NOT_NORMALIZED"); }
  if (!EMPLOYEE_REGISTRATION_INPUT_FIELDS_.every(key =>
      Object.prototype.hasOwnProperty.call(normalizedInput, key) &&
      normalizedInput[key] === normalized[key])) {
    throw new Error("EMPLOYEE_REGISTRATION_INPUT_NOT_NORMALIZED");
  }
  if (!validateEmployeeRegistrationInput_(normalizedInput).ok)
    throw new Error("EMPLOYEE_REGISTRATION_INPUT_NOT_VALIDATED");
  const digest = String(sha256Hex(employeeRegistrationCanonicalPayload_(normalizedInput)) || "").toLowerCase();
  if (!/^[a-f0-9]{64}$/.test(digest)) throw new Error("EMPLOYEE_REGISTRATION_DIGEST_INVALID");
  return EMPLOYEE_REGISTRATION_HASH_VERSION_ + ":" + digest;
}

function isEmployeeRegistrationOperationId_(value) {
  return typeof value === "string" &&
    /^[a-f0-9]{8}-[a-f0-9]{4}-4[a-f0-9]{3}-[89ab][a-f0-9]{3}-[a-f0-9]{12}$/i.test(value);
}

function isDenseEmployeeRegistrationArray_(values) {
  if (!Array.isArray(values)) return false;
  for (let index = 0; index < values.length; index++) {
    if (!Object.prototype.hasOwnProperty.call(values, index)) return false;
  }
  return true;
}

function isEmployeeRegistrationRow_(row) {
  if (!row || typeof row !== "object" || Array.isArray(row) ||
      (Object.getPrototypeOf(row) !== Object.prototype && Object.getPrototypeOf(row) !== null))
    return false;
  if (!["employee_id", "display_employee_id"].every(key =>
      Object.prototype.hasOwnProperty.call(row, key) &&
      typeof row[key] === "string" && row[key] !== "" && row[key] === row[key].trim()))
    return false;
  return Object.prototype.hasOwnProperty.call(row, "registration_operation_id") &&
    typeof row.registration_operation_id === "string";
}

function employeeRegistrationStoredInput_(employee, formatDateKey) {
  const projected = {};
  EMPLOYEE_REGISTRATION_INPUT_FIELDS_.forEach(key => {
    if (!Object.prototype.hasOwnProperty.call(employee, key))
      throw new Error("EMPLOYEE_REGISTRATION_STORED_INPUT_INCOMPLETE");
    const value = employee[key];
    if (value === undefined) throw new Error("EMPLOYEE_REGISTRATION_STORED_INPUT_INCOMPLETE");
    if ((key === "hire_date" || key === "leave_date") && value instanceof Date) {
      // 呼出側はSpreadsheetのタイムゾーンでYYYY-MM-DDへ変換する。
      if (!Number.isFinite(value.getTime()) || typeof formatDateKey !== "function")
        throw new Error("EMPLOYEE_REGISTRATION_STORED_DATE_INVALID");
      const dateKey = formatDateKey(value);
      if (typeof dateKey !== "string" || !isEmployeeRegistrationDateKey_(dateKey))
        throw new Error("EMPLOYEE_REGISTRATION_STORED_DATE_INVALID");
      projected[key] = dateKey;
    } else {
      projected[key] = value;
    }
  });
  return normalizeEmployeeRegistrationInput_(projected);
}

function classifyEmployeeRegistrationRecovery_(operation, employeeRows, request, sha256Hex, formatDateKey) {
  const operationId = request && request.operation_id;
  if (!isEmployeeRegistrationOperationId_(operationId)) return { kind: "invalid_operation_id" };
  if (!request.input_hash || !/^v1:[a-f0-9]{64}$/.test(request.input_hash))
    return { kind: "invalid_input_hash" };
  if (!request.created_by) return { kind: "invalid_admin" };
  if (!request.normalized_input) return { kind: "invalid_input" };
  if (typeof sha256Hex !== "function") return { kind: "digest_unavailable" };
  let calculatedHash;
  try {
    calculatedHash = employeeRegistrationInputHash_(request.normalized_input, sha256Hex);
  } catch (error) { return { kind: "invalid_input" }; }
  if (calculatedHash !== request.input_hash) return { kind: "input_hash_mismatch" };
  if (!isDenseEmployeeRegistrationArray_(employeeRows))
    return { kind: "employee_rows_unavailable" };
  const rows = employeeRows;
  if (rows.some(row => !isEmployeeRegistrationRow_(row)))
    return { kind: "employee_rows_invalid" };
  const employeeIds = new Set(), displayIds = new Set();
  for (const row of rows) {
    if (employeeIds.has(row.employee_id) || displayIds.has(row.display_employee_id))
      return { kind: "employee_duplicate" };
    employeeIds.add(row.employee_id);
    displayIds.add(row.display_employee_id);
  }
  if (!operation) return rows.some(row => row.registration_operation_id === operationId)
    ? { kind: "orphan_employee" } : { kind: "new" };
  if (operation.operation_id !== operationId) return { kind: "operation_mismatch" };
  if (operation.input_hash !== request.input_hash) return { kind: "hash_conflict" };
  if (operation.created_by !== request.created_by) return { kind: "admin_conflict" };
  if (["STARTED", "EMPLOYEE_SAVED", "COMPLETED"].indexOf(operation.status) === -1)
    return { kind: "operation_state_invalid" };
  if (!operation.employee_id || !operation.display_employee_id)
    return { kind: "reservation_invalid" };

  const matches = rows.filter(row => row.registration_operation_id === operationId);
  if (matches.length > 1) return { kind: "employee_duplicate" };
  if (rows.some(row => row.registration_operation_id !== operationId &&
      (row.employee_id === operation.employee_id ||
        row.display_employee_id === operation.display_employee_id))) {
    return { kind: "reserved_id_conflict" };
  }
  if (matches.length === 0) return operation.status === "STARTED"
    ? { kind: "started_without_employee" } : { kind: "employee_missing" };

  const employee = matches[0];
  if (employee.employee_id !== operation.employee_id ||
      employee.display_employee_id !== operation.display_employee_id) {
    return { kind: "employee_mismatch" };
  }
  // COMPLETED後の通常編集は再送結果を変えない。未完了時のみ保存入力を再検証する。
  if (operation.status !== "COMPLETED") {
    try {
      if (employeeRegistrationCanonicalPayload_(employeeRegistrationStoredInput_(employee, formatDateKey)) !==
          employeeRegistrationCanonicalPayload_(request.normalized_input)) {
        return { kind: "employee_mismatch" };
      }
    } catch (error) { return { kind: "employee_mismatch" }; }
  }
  return { kind: operation.status === "STARTED" ? "started_with_employee" :
    operation.status === "EMPLOYEE_SAVED" ? "employee_saved" : "completed" };
}

function nextEmployeeRegistrationId_(prefix, employeeIds, operationIds, baseline) {
  if (["EMP", "W", "P"].indexOf(prefix) === -1 ||
      !Number.isSafeInteger(baseline) || baseline < 0) {
    throw new Error("EMPLOYEE_REGISTRATION_BASELINE_INVALID");
  }
  if (!isDenseEmployeeRegistrationArray_(employeeIds) ||
      !isDenseEmployeeRegistrationArray_(operationIds))
    throw new Error("EMPLOYEE_REGISTRATION_IDS_UNAVAILABLE");
  let max = baseline;
  employeeIds.concat(operationIds).forEach(id => {
    if (typeof id !== "string") throw new Error("EMPLOYEE_REGISTRATION_ID_INVALID");
    const text = id;
    const match = /^(EMP|W|P)(\d+)$/.exec(text);
    if (!match || (match[2].length !== 4 &&
        (match[2].length < 5 || match[2][0] === "0"))) {
      throw new Error("EMPLOYEE_REGISTRATION_ID_INVALID: " + text);
    }
    const number = Number(match[2]);
    if (!Number.isSafeInteger(number) || number < 1)
      throw new Error("EMPLOYEE_REGISTRATION_ID_UNSAFE: " + text);
    if (match[1] === prefix) max = Math.max(max, number);
  });
  if (max === Number.MAX_SAFE_INTEGER) throw new Error("EMPLOYEE_REGISTRATION_ID_EXHAUSTED");
  return prefix + String(max + 1).padStart(4, "0");
}

function calculateNextEmployeeRegistrationIds_(companyCode, employeeIds, operationIds, baseline) {
  if (companyCode !== "MAIN" && companyCode !== "PARTNER")
    throw new Error("EMPLOYEE_REGISTRATION_COMPANY_INVALID");
  const floors = baseline || EMPLOYEE_REGISTRATION_ID_BASELINE_;
  const displayPrefix = companyCode === "MAIN" ? "W" : "P";
  return {
    employee_id: nextEmployeeRegistrationId_("EMP", employeeIds, operationIds, floors.EMP),
    display_employee_id: nextEmployeeRegistrationId_(displayPrefix, employeeIds, operationIds,
      floors[displayPrefix])
  };
}
