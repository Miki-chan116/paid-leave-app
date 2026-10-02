/*
 * 社員登録安全化 Phase 2A: 読取専用の Apps Script / Spreadsheet アダプター。
 * 既存登録処理には未接続。operation 予約IDの取得・書込みは後続 Phase の責務。
 */

function employeeRegistrationDigestBytesToHex_(bytes) {
  if (!Array.isArray(bytes) || bytes.length !== 32) {
    throw new Error("EMPLOYEE_REGISTRATION_ADAPTER_DIGEST_INVALID");
  }
  let hex = "";
  for (let index = 0; index < bytes.length; index++) {
    if (!Object.prototype.hasOwnProperty.call(bytes, index) ||
        !Number.isInteger(bytes[index]) || bytes[index] < -128 || bytes[index] > 255) {
      throw new Error("EMPLOYEE_REGISTRATION_ADAPTER_DIGEST_INVALID");
    }
    hex += (bytes[index] & 255).toString(16).padStart(2, "0");
  }
  return hex;
}

function employeeRegistrationSha256HexAppsScript_(text) {
  if (typeof text !== "string") throw new Error("EMPLOYEE_REGISTRATION_ADAPTER_DIGEST_INPUT_INVALID");
  // Node の crypto.update(text, "utf8") と同じ UTF-8 byte 列を SHA-256 に渡す。
  return employeeRegistrationDigestBytesToHex_(Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256, text, Utilities.Charset.UTF_8));
}

function employeeRegistrationDateKeyWithFormatter_(date, timeZone, formatDate) {
  if (!(date instanceof Date) || !Number.isFinite(date.getTime()) ||
      typeof timeZone !== "string" || !timeZone.trim() ||
      typeof formatDate !== "function") {
    throw new Error("EMPLOYEE_REGISTRATION_ADAPTER_DATE_INVALID");
  }
  let key;
  try { key = formatDate(date, timeZone, "yyyy-MM-dd"); }
  catch (error) { throw new Error("EMPLOYEE_REGISTRATION_ADAPTER_DATE_INVALID"); }
  if (typeof key !== "string" || !isEmployeeRegistrationDateKey_(key)) {
    throw new Error("EMPLOYEE_REGISTRATION_ADAPTER_DATE_INVALID");
  }
  return key;
}

function employeeRegistrationSpreadsheetDateFormatter_(spreadsheet) {
  if (!spreadsheet || typeof spreadsheet.getSpreadsheetTimeZone !== "function") {
    throw new Error("EMPLOYEE_REGISTRATION_ADAPTER_SPREADSHEET_UNAVAILABLE");
  }
  let timeZone;
  try { timeZone = spreadsheet.getSpreadsheetTimeZone(); }
  catch (error) { throw new Error("EMPLOYEE_REGISTRATION_ADAPTER_TIMEZONE_UNAVAILABLE"); }
  if (typeof timeZone !== "string" || !timeZone.trim()) {
    throw new Error("EMPLOYEE_REGISTRATION_ADAPTER_TIMEZONE_UNAVAILABLE");
  }
  return date => employeeRegistrationDateKeyWithFormatter_(
    date, timeZone, (value, zone, pattern) => Utilities.formatDate(value, zone, pattern));
}

function employeeRegistrationAdapterHeaderMap_(headers, schemaMode) {
  if (schemaMode !== "legacy_read_only" && schemaMode !== "registration_ready") {
    throw new Error("EMPLOYEE_REGISTRATION_ADAPTER_SCHEMA_MODE_REQUIRED");
  }
  if (!isDenseEmployeeRegistrationArray_(headers) || headers.length === 0) {
    throw new Error("EMPLOYEE_REGISTRATION_ADAPTER_HEADERS_INVALID");
  }
  const map = Object.create(null);
  for (let index = 0; index < headers.length; index++) {
    const header = headers[index];
    // Code.gs はヘッダーを trim するが、安全判定では曖昧な別名を許容しない。
    if (typeof header !== "string" || !header || header !== header.trim() ||
        Object.prototype.hasOwnProperty.call(map, header)) {
      throw new Error("EMPLOYEE_REGISTRATION_ADAPTER_HEADERS_INVALID");
    }
    map[header] = index;
  }
  const required = ["employee_id", "display_employee_id"]
    .concat(EMPLOYEE_REGISTRATION_INPUT_FIELDS_);
  if (schemaMode === "registration_ready") required.push("registration_operation_id");
  if (required.some(key => !Object.prototype.hasOwnProperty.call(map, key))) {
    throw new Error("EMPLOYEE_REGISTRATION_ADAPTER_HEADERS_MISSING");
  }
  return map;
}

function employeeRegistrationAdapterValidId_(value, prefixes) {
  if (typeof value !== "string" || value !== value.trim()) return false;
  const match = /^(EMP|W|P)(\d+)$/.exec(value);
  if (!match || prefixes.indexOf(match[1]) === -1 ||
      (match[2].length !== 4 && (match[2].length < 5 || match[2][0] === "0"))) return false;
  const number = Number(match[2]);
  return Number.isSafeInteger(number) && number >= 1;
}

function employeeRegistrationAdapterFieldValid_(key, value, formatDateKey) {
  if (key === "hire_date" || key === "leave_date") {
    if (value instanceof Date) {
      formatDateKey(value); // 検証だけ行い、Phase 1 の保存行比較へは Date のまま渡す。
      return true;
    }
    return typeof value === "string" &&
      (value === "" || isEmployeeRegistrationDateKey_(value));
  }
  if (key === "leave_management_target" || key === "is_driver") {
    return typeof value === "boolean" || value === "TRUE" || value === "FALSE";
  }
  if (["work_days_per_week", "work_start_minute", "work_end_minute", "fiscal_start_month"]
      .indexOf(key) !== -1) {
    return value === "" ||
      (typeof value === "number" && Number.isSafeInteger(value)) ||
      (typeof value === "string" && /^\d+$/.test(value) &&
        Number.isSafeInteger(Number(value)));
  }
  return typeof value === "string";
}

function employeeRegistrationProjectEmployees_(values, formatDateKey, schemaMode) {
  if (!isDenseEmployeeRegistrationArray_(values) || values.length === 0 ||
      typeof formatDateKey !== "function") {
    throw new Error("EMPLOYEE_REGISTRATION_ADAPTER_READ_INVALID");
  }
  const headers = values[0];
  const map = employeeRegistrationAdapterHeaderMap_(headers, schemaMode);
  const hasOperationColumn = Object.prototype.hasOwnProperty.call(map, "registration_operation_id");
  const employeeRows = [], employeeIds = [], displayEmployeeIds = [];
  const seenEmployeeIds = new Set(), seenDisplayIds = new Set();

  for (let index = 1; index < values.length; index++) {
    const cells = values[index];
    if (!isDenseEmployeeRegistrationArray_(cells) || cells.length !== headers.length) {
      throw new Error("EMPLOYEE_REGISTRATION_ADAPTER_ROW_INVALID");
    }
    const employeeId = cells[map.employee_id];
    const displayId = cells[map.display_employee_id];
    if (!employeeRegistrationAdapterValidId_(employeeId, ["EMP"]) ||
        !employeeRegistrationAdapterValidId_(displayId, ["W", "P"]) ||
        seenEmployeeIds.has(employeeId) || seenDisplayIds.has(displayId)) {
      throw new Error("EMPLOYEE_REGISTRATION_ADAPTER_ROW_INVALID");
    }
    seenEmployeeIds.add(employeeId);
    seenDisplayIds.add(displayId);
    const row = { employee_id: employeeId, display_employee_id: displayId };
    for (const key of EMPLOYEE_REGISTRATION_INPUT_FIELDS_) {
      const value = cells[map[key]];
      if (!employeeRegistrationAdapterFieldValid_(key, value, formatDateKey)) {
        throw new Error("EMPLOYEE_REGISTRATION_ADAPTER_ROW_INVALID");
      }
      row[key] = value;
    }
    for (const key of ["name", "name_kana", "company_code", "employment_type", "employment_status"]) {
      if (!row[key].trim()) throw new Error("EMPLOYEE_REGISTRATION_ADAPTER_ROW_INVALID");
    }
    // Phase 2A は新フロー未接続なので、列欠落を明示した旧データ読取モードだけで許す。
    // Phase 2B/2C の登録経路では registration_ready を指定し、書込み時に新規行の UUID を保証する。
    const operationId = hasOperationColumn ? cells[map.registration_operation_id] : "";
    if (typeof operationId !== "string" ||
        (operationId !== "" && !isEmployeeRegistrationOperationId_(operationId))) {
      throw new Error("EMPLOYEE_REGISTRATION_ADAPTER_ROW_INVALID");
    }
    row.registration_operation_id = operationId;
    employeeRows.push(row);
    employeeIds.push(employeeId);
    displayEmployeeIds.push(displayId);
  }
  return {
    employeeRows: employeeRows,
    employeeIds: employeeIds,
    displayEmployeeIds: displayEmployeeIds,
    existingIds: employeeIds.concat(displayEmployeeIds),
    hasRegistrationOperationIdColumn: hasOperationColumn
    // operationIds は返さない。予約ID一覧は Phase 2B の ledger 読取後にのみ確定する。
  };
}

function readEmployeeRegistrationEmployeesReadOnly_(schemaMode) {
  if (schemaMode !== "legacy_read_only" && schemaMode !== "registration_ready") {
    throw new Error("EMPLOYEE_REGISTRATION_ADAPTER_SCHEMA_MODE_REQUIRED");
  }
  let spreadsheet, sheet, values;
  try { spreadsheet = SpreadsheetApp.openById(SS_ID); }
  catch (error) { throw new Error("EMPLOYEE_REGISTRATION_ADAPTER_SPREADSHEET_UNAVAILABLE"); }
  if (!spreadsheet || typeof spreadsheet.getSheetByName !== "function") {
    throw new Error("EMPLOYEE_REGISTRATION_ADAPTER_SPREADSHEET_UNAVAILABLE");
  }
  const formatDateKey = employeeRegistrationSpreadsheetDateFormatter_(spreadsheet);
  try { sheet = spreadsheet.getSheetByName("employees"); }
  catch (error) { throw new Error("EMPLOYEE_REGISTRATION_ADAPTER_SHEET_UNAVAILABLE"); }
  if (!sheet || typeof sheet.getDataRange !== "function") {
    throw new Error("EMPLOYEE_REGISTRATION_ADAPTER_SHEET_UNAVAILABLE");
  }
  try { values = sheet.getDataRange().getValues(); }
  catch (error) { throw new Error("EMPLOYEE_REGISTRATION_ADAPTER_READ_UNAVAILABLE"); }
  const snapshot = employeeRegistrationProjectEmployees_(values, formatDateKey, schemaMode);
  snapshot.formatDateKey = formatDateKey;
  return snapshot;
}
