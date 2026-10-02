/*
 * 社員登録安全化 Phase 2B-1: operation 台帳の読取専用契約。
 * Sheet の作成・更新と登録経路への接続は後続 Phase の責務。
 */

const EMPLOYEE_REGISTRATION_LEDGER_HEADERS_ = [
  "operation_id", "input_hash", "created_by", "status",
  "employee_id", "display_employee_id", "created_at", "updated_at"
];

function employeeRegistrationLedgerHeaderMap_(headers) {
  if (!isDenseEmployeeRegistrationArray_(headers) || headers.length === 0) {
    throw new Error("EMPLOYEE_REGISTRATION_LEDGER_HEADERS_INVALID");
  }
  const map = Object.create(null);
  for (let index = 0; index < headers.length; index++) {
    const header = headers[index];
    if (typeof header !== "string" || !header || header !== header.trim() ||
        Object.prototype.hasOwnProperty.call(map, header)) {
      throw new Error("EMPLOYEE_REGISTRATION_LEDGER_HEADERS_INVALID");
    }
    map[header] = index;
  }
  if (EMPLOYEE_REGISTRATION_LEDGER_HEADERS_.some(key =>
      !Object.prototype.hasOwnProperty.call(map, key))) {
    throw new Error("EMPLOYEE_REGISTRATION_LEDGER_HEADERS_MISSING");
  }
  return map;
}

function employeeRegistrationLedgerOperationValid_(operation) {
  if (!operation || typeof operation !== "object" || Array.isArray(operation) ||
      (Object.getPrototypeOf(operation) !== Object.prototype &&
        Object.getPrototypeOf(operation) !== null) ||
      EMPLOYEE_REGISTRATION_LEDGER_HEADERS_.some(key =>
        !Object.prototype.hasOwnProperty.call(operation, key))) return false;
  if (!isEmployeeRegistrationOperationId_(operation.operation_id) ||
      typeof operation.input_hash !== "string" ||
      !/^v1:[a-f0-9]{64}$/.test(operation.input_hash) ||
      typeof operation.created_by !== "string" || !operation.created_by ||
      operation.created_by !== operation.created_by.trim() ||
      ["STARTED", "EMPLOYEE_SAVED", "COMPLETED"].indexOf(operation.status) === -1 ||
      !employeeRegistrationAdapterValidId_(operation.employee_id, ["EMP"]) ||
      !employeeRegistrationAdapterValidId_(operation.display_employee_id, ["W", "P"])) {
    return false;
  }
  // 台帳時刻は Spreadsheet の Date 値だけを扱い、文字列や数値を暗黙変換しない。
  const created = operation.created_at, updated = operation.updated_at;
  return created instanceof Date && Number.isFinite(created.getTime()) &&
    updated instanceof Date && Number.isFinite(updated.getTime()) &&
    created.getTime() <= updated.getTime();
}

function employeeRegistrationLedgerReservedIds_(operations) {
  if (!isDenseEmployeeRegistrationArray_(operations)) {
    throw new Error("EMPLOYEE_REGISTRATION_LEDGER_SNAPSHOT_INVALID");
  }
  const operationIds = new Set(), employeeIds = new Set(), displayIds = new Set();
  const reservedIds = [];
  for (const operation of operations) {
    if (!employeeRegistrationLedgerOperationValid_(operation)) {
      throw new Error("EMPLOYEE_REGISTRATION_LEDGER_ROW_INVALID");
    }
    // Phase 1 は UUID の大文字を許す。大小文字だけ異なる二重予約も拒否する。
    const operationKey = operation.operation_id.toLowerCase();
    if (operationIds.has(operationKey)) {
      throw new Error("EMPLOYEE_REGISTRATION_LEDGER_OPERATION_DUPLICATE");
    }
    operationIds.add(operationKey);
    if (employeeIds.has(operation.employee_id)) {
      throw new Error("EMPLOYEE_REGISTRATION_LEDGER_EMPLOYEE_ID_DUPLICATE");
    }
    employeeIds.add(operation.employee_id);
    if (displayIds.has(operation.display_employee_id)) {
      throw new Error("EMPLOYEE_REGISTRATION_LEDGER_DISPLAY_ID_DUPLICATE");
    }
    displayIds.add(operation.display_employee_id);
    reservedIds.push(operation.employee_id, operation.display_employee_id);
  }
  return reservedIds;
}

function employeeRegistrationProjectOperationLedger_(values) {
  if (!isDenseEmployeeRegistrationArray_(values) || values.length === 0) {
    throw new Error("EMPLOYEE_REGISTRATION_LEDGER_READ_INVALID");
  }
  const headers = values[0];
  const map = employeeRegistrationLedgerHeaderMap_(headers);
  const operations = [];
  for (let index = 1; index < values.length; index++) {
    const cells = values[index];
    if (!isDenseEmployeeRegistrationArray_(cells) || cells.length !== headers.length) {
      throw new Error("EMPLOYEE_REGISTRATION_LEDGER_ROW_INVALID");
    }
    const operation = {};
    for (const key of EMPLOYEE_REGISTRATION_LEDGER_HEADERS_) {
      operation[key] = cells[map[key]];
    }
    operations.push(operation);
  }
  const reservedIds = employeeRegistrationLedgerReservedIds_(operations);
  return { operations: operations, reservedIds: reservedIds };
}

function readEmployeeRegistrationOperationLedgerReadOnly_() {
  let spreadsheet, sheet, values;
  try { spreadsheet = SpreadsheetApp.openById(SS_ID); }
  catch (error) { throw new Error("EMPLOYEE_REGISTRATION_LEDGER_SPREADSHEET_UNAVAILABLE"); }
  if (!spreadsheet || typeof spreadsheet.getSheetByName !== "function") {
    throw new Error("EMPLOYEE_REGISTRATION_LEDGER_SPREADSHEET_UNAVAILABLE");
  }
  try { sheet = spreadsheet.getSheetByName("employee_registration_operations"); }
  catch (error) { throw new Error("EMPLOYEE_REGISTRATION_LEDGER_SHEET_UNAVAILABLE"); }
  if (!sheet || typeof sheet.getDataRange !== "function") {
    throw new Error("EMPLOYEE_REGISTRATION_LEDGER_SHEET_UNAVAILABLE");
  }
  try { values = sheet.getDataRange().getValues(); }
  catch (error) { throw new Error("EMPLOYEE_REGISTRATION_LEDGER_READ_UNAVAILABLE"); }
  return employeeRegistrationProjectOperationLedger_(values);
}

// snapshot 内の純粋検索。戻り値の null は Spreadsheet 読取成功を証明しない。
// 本番の不在判定には findEmployeeRegistrationOperationFromSpreadsheetReadOnly_ を使う。
function lookupEmployeeRegistrationOperation_(snapshot, operationId) {
  if (!isEmployeeRegistrationOperationId_(operationId)) {
    throw new Error("EMPLOYEE_REGISTRATION_LEDGER_LOOKUP_ID_INVALID");
  }
  if (!snapshot || typeof snapshot !== "object" ||
      !isDenseEmployeeRegistrationArray_(snapshot.reservedIds)) {
    throw new Error("EMPLOYEE_REGISTRATION_LEDGER_SNAPSHOT_INVALID");
  }
  const reservedIds = employeeRegistrationLedgerReservedIds_(snapshot.operations);
  if (reservedIds.length !== snapshot.reservedIds.length ||
      reservedIds.some((id, index) => id !== snapshot.reservedIds[index])) {
    throw new Error("EMPLOYEE_REGISTRATION_LEDGER_SNAPSHOT_INVALID");
  }
  // snapshot 内の不在だけを null にする。大小文字差は同一UUID候補として返し、
  // Phase 1 に厳密な operation_id 一致を判定させる。
  return snapshot.operations.find(operation =>
    operation.operation_id.toLowerCase() === operationId.toLowerCase()) || null;
}

// Phase 2B-2 の production lookup 入口。読取・全件検証の失敗は例外のまま伝える。
// null は正常に読めた台帳内に対象 UUID が存在しない場合に限る。
function findEmployeeRegistrationOperationFromSpreadsheetReadOnly_(operationId) {
  if (!isEmployeeRegistrationOperationId_(operationId)) {
    throw new Error("EMPLOYEE_REGISTRATION_LEDGER_LOOKUP_ID_INVALID");
  }
  const snapshot = readEmployeeRegistrationOperationLedgerReadOnly_();
  return lookupEmployeeRegistrationOperation_(snapshot, operationId);
}

function verifyEmployeeRegistrationLedgerEmployeeIntegrity_(ledgerSnapshot, employeeSnapshot) {
  if (!employeeSnapshot || employeeSnapshot.hasRegistrationOperationIdColumn !== true ||
      !isDenseEmployeeRegistrationArray_(employeeSnapshot.employeeRows) ||
      !isDenseEmployeeRegistrationArray_(employeeSnapshot.existingIds)) {
    throw new Error("EMPLOYEE_REGISTRATION_LEDGER_EMPLOYEE_SNAPSHOT_INVALID");
  }
  // lookup の検証を通し、ledger 全行と reservedIds の整合性を先に確認する。
  if (!ledgerSnapshot || !isDenseEmployeeRegistrationArray_(ledgerSnapshot.reservedIds)) {
    throw new Error("EMPLOYEE_REGISTRATION_LEDGER_SNAPSHOT_INVALID");
  }
  const reservedIds = employeeRegistrationLedgerReservedIds_(ledgerSnapshot.operations);
  if (reservedIds.length !== ledgerSnapshot.reservedIds.length ||
      reservedIds.some((id, index) => id !== ledgerSnapshot.reservedIds[index])) {
    throw new Error("EMPLOYEE_REGISTRATION_LEDGER_SNAPSHOT_INVALID");
  }
  const operationsByUuid = new Map(), employeesById = new Map(), employeesByDisplay = new Map();
  for (const operation of ledgerSnapshot.operations) {
    operationsByUuid.set(operation.operation_id.toLowerCase(), operation);
  }
  const employeeIds = [], displayIds = [];
  for (const row of employeeSnapshot.employeeRows) {
    if (!isEmployeeRegistrationRow_(row) ||
        !employeeRegistrationAdapterValidId_(row.employee_id, ["EMP"]) ||
        !employeeRegistrationAdapterValidId_(row.display_employee_id, ["W", "P"]) ||
        (row.registration_operation_id !== "" &&
          !isEmployeeRegistrationOperationId_(row.registration_operation_id)) ||
        employeesById.has(row.employee_id) || employeesByDisplay.has(row.display_employee_id)) {
      throw new Error("EMPLOYEE_REGISTRATION_LEDGER_EMPLOYEE_SNAPSHOT_INVALID");
    }
    employeesById.set(row.employee_id, row);
    employeesByDisplay.set(row.display_employee_id, row);
    employeeIds.push(row.employee_id);
    displayIds.push(row.display_employee_id);
    if (row.registration_operation_id !== "") {
      const operation = operationsByUuid.get(row.registration_operation_id.toLowerCase());
      if (!operation || operation.operation_id !== row.registration_operation_id ||
          operation.employee_id !== row.employee_id ||
          operation.display_employee_id !== row.display_employee_id) {
        throw new Error("EMPLOYEE_REGISTRATION_LEDGER_EMPLOYEE_LINK_INVALID");
      }
    }
  }
  const existingIds = employeeIds.concat(displayIds);
  if (existingIds.length !== employeeSnapshot.existingIds.length ||
      existingIds.some((id, index) => id !== employeeSnapshot.existingIds[index])) {
    throw new Error("EMPLOYEE_REGISTRATION_LEDGER_EMPLOYEE_SNAPSHOT_INVALID");
  }
  for (const operation of ledgerSnapshot.operations) {
    const employee = employeesById.get(operation.employee_id);
    const display = employeesByDisplay.get(operation.display_employee_id);
    if ((employee && employee.registration_operation_id !== operation.operation_id) ||
        (display && display.registration_operation_id !== operation.operation_id)) {
      throw new Error("EMPLOYEE_REGISTRATION_LEDGER_RESERVATION_CONFLICT");
    }
  }
  return { existingIds: existingIds, operationIds: reservedIds.slice() };
}
