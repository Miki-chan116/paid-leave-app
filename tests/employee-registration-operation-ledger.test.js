const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");
const crypto = require("node:crypto");

const root = path.join(__dirname, "..");
const files = ["employee-registration-safety.gs", "employee-registration-adapter.gs",
  "employee-registration-operation-ledger.gs"];
const uuidA = "123e4567-e89b-42d3-a456-426614174000";
const uuidB = "223e4567-e89b-42d3-a456-426614174000";
const headers = ["operation_id", "input_hash", "created_by", "status",
  "employee_id", "display_employee_id", "created_at", "updated_at"];
const inputFields = ["name", "display_name", "name_kana", "company_code", "company_name",
  "department", "employment_type", "employment_status", "hire_date", "leave_date",
  "work_days_per_week", "work_start_minute", "work_end_minute", "fiscal_start_month",
  "leave_management_target", "is_driver", "driver_type", "default_vehicle_id", "notes"];
const employeeHeaders = ["employee_id", "display_employee_id", ...inputFields,
  "registration_operation_id"];
const employee = {
  employee_id: "EMP0083", display_employee_id: "W0062", name: "例 一郎",
  display_name: "", name_kana: "れい いちろう", company_code: "MAIN",
  company_name: "", department: "", employment_type: "regular",
  employment_status: "active", hire_date: "2026-04-01", leave_date: "",
  work_days_per_week: 5, work_start_minute: "", work_end_minute: "",
  fiscal_start_month: 4, leave_management_target: true, is_driver: false,
  driver_type: "", default_vehicle_id: "", notes: "",
  registration_operation_id: uuidA
};
const operation = {
  operation_id: uuidA, input_hash: "v1:" + "a".repeat(64), created_by: "ADMIN1",
  status: "STARTED", employee_id: "EMP0083", display_employee_id: "W0062",
  created_at: new Date("2026-04-01T00:00:00Z"),
  updated_at: new Date("2026-04-01T01:00:00Z")
};
const cells = (names, data) => names.map(name => data[name]);

function fixture(options = {}) {
  const ledgerValues = options.ledgerValues === undefined ?
    [headers, cells(headers, operation)] : options.ledgerValues;
  const employeeValues = options.employeeValues === undefined ?
    [employeeHeaders, cells(employeeHeaders, employee)] : options.employeeValues;
  const spreadsheet = options.spreadsheet === undefined ? {
    getSpreadsheetTimeZone: () => "Asia/Tokyo",
    getSheetByName: name => {
      if (name === "employee_registration_operations") return {
        getDataRange: () => ({ getValues: () => ledgerValues })
      };
      if (name === "employees") return {
        getDataRange: () => ({ getValues: () => employeeValues })
      };
      return null;
    }
  } : options.spreadsheet;
  const context = {
    Date, Number, String, Set, Map, SS_ID: "test-spreadsheet",
    SpreadsheetApp: { openById: () => spreadsheet },
    Utilities: {
      DigestAlgorithm: { SHA_256: "SHA_256" }, Charset: { UTF_8: "UTF_8" },
      computeDigest: (algorithm, text, charset) => {
        assert.equal(algorithm, "SHA_256");
        assert.equal(charset, "UTF_8");
        return Array.from(crypto.createHash("sha256").update(text, "utf8").digest(),
          byte => byte > 127 ? byte - 256 : byte);
      },
      formatDate: (date, zone) => {
        const parts = new Intl.DateTimeFormat("en-US", { timeZone: zone,
          year: "numeric", month: "2-digit", day: "2-digit" }).formatToParts(date);
        const values = Object.fromEntries(parts.map(part => [part.type, part.value]));
        return `${values.year}-${values.month}-${values.day}`;
      }
    }
  };
  vm.createContext(context);
  for (const file of files) {
    const filename = path.join(root, file);
    vm.runInContext(fs.readFileSync(filename, "utf8"), context, { filename });
  }
  return context;
}

test("headerのみの正常な空台帳だけが空一覧とnot foundを返す", () => {
  const context = fixture({ ledgerValues: [headers] });
  const snapshot = context.readEmployeeRegistrationOperationLedgerReadOnly_();
  assert.equal(snapshot.operations.length, 0);
  assert.equal(snapshot.reservedIds.length, 0);
  assert.equal(context.findEmployeeRegistrationOperationFromSpreadsheetReadOnly_(uuidA), null);
});

test("STARTED・EMPLOYEE_SAVED・COMPLETEDとMAIN/PARTNERを予約IDへ全件反映", () => {
  const third = { ...operation, operation_id: "323e4567-e89b-42d3-a456-426614174000",
    status: "COMPLETED", employee_id: "EMP0085", display_employee_id: "W0063" };
  const second = { ...operation, operation_id: uuidB, status: "EMPLOYEE_SAVED",
    employee_id: "EMP0084", display_employee_id: "P0023" };
  const context = fixture({ ledgerValues: [headers, cells(headers, operation),
    cells(headers, second), cells(headers, third)] });
  const snapshot = context.readEmployeeRegistrationOperationLedgerReadOnly_();
  assert.equal(JSON.stringify(snapshot.operations.map(item => item.status)),
    JSON.stringify(["STARTED", "EMPLOYEE_SAVED", "COMPLETED"]));
  assert.equal(JSON.stringify(snapshot.reservedIds), JSON.stringify([
    "EMP0083", "W0062", "EMP0084", "P0023", "EMP0085", "W0063"]));
  assert.equal(context.lookupEmployeeRegistrationOperation_(snapshot, uuidB).employee_id, "EMP0084");
  assert.equal(context.lookupEmployeeRegistrationOperation_(snapshot,
    "323e4567-e89b-42d3-a456-426614174000").status, "COMPLETED");
});

test("header欠落・重複・空白・先頭末尾余白・空Sheetを拒否し、追加列は許可", () => {
  for (const bad of [[], [[""]], [headers.filter(key => key !== "input_hash")],
    [[...headers, "status"]], [[...headers, ""]], [[...headers, " "]],
    [[...headers, " status"]], [[...headers, "status "]]]) {
    const context = fixture({ ledgerValues: bad });
    assert.throws(() => context.readEmployeeRegistrationOperationLedgerReadOnly_(),
      /LEDGER_(READ_INVALID|HEADERS_)/);
  }
  const names = [...headers, "memo"];
  const context = fixture({ ledgerValues: [names, cells(names, { ...operation, memo: "" })] });
  assert.equal(context.readEmployeeRegistrationOperationLedgerReadOnly_().operations.length, 1);
});

test("Spreadsheet・Sheet・range読取失敗は空台帳にならない", () => {
  for (const spreadsheet of [null, {}, { getSheetByName: () => null },
    { getSheetByName: () => ({ getDataRange: () => { throw new Error("range"); } }) },
    { getSheetByName: () => ({ getDataRange: () => ({
      getValues: () => { throw new Error("values"); }
    }) }) }]) {
    const context = fixture({ spreadsheet });
    assert.throws(() => context.readEmployeeRegistrationOperationLedgerReadOnly_(),
      /UNAVAILABLE/);
  }
  const context = fixture();
  context.SpreadsheetApp.openById = () => { throw new Error("open"); };
  assert.throws(() => context.readEmployeeRegistrationOperationLedgerReadOnly_(),
    /SPREADSHEET_UNAVAILABLE/);
});

test("UUID v4の大小文字は許容し、空・不正型・別versionを拒否", () => {
  const upper = fixture({ ledgerValues: [headers, cells(headers, {
    ...operation, operation_id: uuidA.toUpperCase() })] });
  assert.equal(upper.readEmployeeRegistrationOperationLedgerReadOnly_().operations[0].operation_id,
    uuidA.toUpperCase());
  for (const id of ["", " " + uuidA, 42, {}, "bad",
    "123e4567-e89b-12d3-a456-426614174000",
    "123e4567-e89b-52d3-a456-426614174000"]) {
    const context = fixture({ ledgerValues: [headers, cells(headers, {
      ...operation, operation_id: id })] });
    assert.throws(() => context.readEmployeeRegistrationOperationLedgerReadOnly_(), /ROW_INVALID/);
  }
});

test("input_hash・created_by・status・予約ID・timestampの異常を全行拒否", () => {
  const changes = [
    { input_hash: "" }, { input_hash: " v1:" + "a".repeat(64) },
    { input_hash: "v2:" + "a".repeat(64) }, { input_hash: "v1:" + "A".repeat(64) },
    { input_hash: "v1:" + "a".repeat(63) }, { input_hash: "v1:" + "g".repeat(64) },
    { input_hash: 42 }, { created_by: "" }, { created_by: " ADMIN1" },
    { created_by: {} }, { status: "started" }, { status: "STARTED " },
    { status: "UNKNOWN" }, { status: 1 }, { employee_id: "" },
    { employee_id: "W0062" }, { employee_id: "EMP83" },
    { employee_id: " EMP0083" }, { display_employee_id: "" },
    { display_employee_id: "EMP0083" }, { display_employee_id: "P23" },
    { created_at: "2026-04-01" }, { created_at: 123 },
    { updated_at: new Date("invalid") },
    { updated_at: new Date("2026-03-31T23:59:59Z") }
  ];
  for (const change of changes) {
    const context = fixture({ ledgerValues: [headers, cells(headers, { ...operation, ...change })] });
    assert.throws(() => context.readEmployeeRegistrationOperationLedgerReadOnly_(), /ROW_INVALID/);
  }
});

test("安全整数境界のIDは許可し、超過は拒否する", () => {
  const maximum = { ...operation, employee_id: "EMP9007199254740991",
    display_employee_id: "P9007199254740991" };
  assert.equal(fixture({ ledgerValues: [headers, cells(headers, maximum)] })
    .readEmployeeRegistrationOperationLedgerReadOnly_().reservedIds.length, 2);
  const unsafe = { ...maximum, employee_id: "EMP9007199254740992" };
  assert.throws(() => fixture({ ledgerValues: [headers, cells(headers, unsafe)] })
    .readEmployeeRegistrationOperationLedgerReadOnly_(), /ROW_INVALID/);
});

test("空行・疎配列・object行は除外せず拒否する", () => {
  for (const bad of [Array(headers.length).fill(""), Array(headers.length), {},
    cells(headers, operation).slice(1)]) {
    const context = fixture({ ledgerValues: [headers, bad] });
    assert.throws(() => context.readEmployeeRegistrationOperationLedgerReadOnly_(), /ROW_INVALID/);
  }
  const sparse = [headers]; sparse.length = 3; sparse[2] = cells(headers, operation);
  assert.throws(() => fixture({ ledgerValues: sparse })
    .readEmployeeRegistrationOperationLedgerReadOnly_(), /READ_INVALID/);
});

test("UUID重複を最優先で検出し、大小文字だけ異なる重複も拒否", () => {
  for (const duplicateId of [uuidA, uuidA.toUpperCase()]) {
    const duplicate = { ...operation, operation_id: duplicateId };
    const context = fixture({ ledgerValues: [headers, cells(headers, operation),
      cells(headers, duplicate)] });
    assert.throws(() => context.readEmployeeRegistrationOperationLedgerReadOnly_(),
      /OPERATION_DUPLICATE/);
  }
});

test("異なるoperation間のemployee/display予約重複を拒否", () => {
  for (const change of [{ employee_id: operation.employee_id, display_employee_id: "P0023" },
    { employee_id: "EMP0084", display_employee_id: operation.display_employee_id }]) {
    const context = fixture({ ledgerValues: [headers, cells(headers, operation),
      cells(headers, { ...operation, operation_id: uuidB, ...change })] });
    assert.throws(() => context.readEmployeeRegistrationOperationLedgerReadOnly_(),
      /_ID_DUPLICATE/);
  }
});

test("pure lookupはsnapshot内を検索し、壊れたsnapshotを拒否", () => {
  const context = fixture();
  const snapshot = context.readEmployeeRegistrationOperationLedgerReadOnly_();
  assert.equal(context.lookupEmployeeRegistrationOperation_(snapshot, uuidA).status, "STARTED");
  assert.equal(context.lookupEmployeeRegistrationOperation_(snapshot, uuidB), null);
  assert.throws(() => context.lookupEmployeeRegistrationOperation_(null, uuidA), /SNAPSHOT_INVALID/);
  assert.throws(() => context.lookupEmployeeRegistrationOperation_({ operations: [],
    reservedIds: ["EMP0083"] }, uuidA), /SNAPSHOT_INVALID/);
  assert.throws(() => context.lookupEmployeeRegistrationOperation_(snapshot, "bad"), /LOOKUP_ID_INVALID/);
  const broken = fixture({ ledgerValues: [headers, cells(headers, { ...operation,
    input_hash: "bad" })] });
  assert.throws(() => broken.readEmployeeRegistrationOperationLedgerReadOnly_(), /ROW_INVALID/);
});

test("production lookupのnullは正常なSheet読取後の不在だけを示す", () => {
  const absent = fixture({ ledgerValues: [headers] });
  assert.equal(absent.findEmployeeRegistrationOperationFromSpreadsheetReadOnly_(uuidA), null);
  const found = fixture();
  assert.equal(found.findEmployeeRegistrationOperationFromSpreadsheetReadOnly_(uuidA).status,
    "STARTED");
  // pure lookup は手製snapshotを検索できるが、本番入口にはsnapshotを渡せない。
  assert.equal(found.lookupEmployeeRegistrationOperation_(
    { operations: [], reservedIds: [] }, uuidA), null);
  assert.equal(found.findEmployeeRegistrationOperationFromSpreadsheetReadOnly_(uuidA).operation_id,
    uuidA);
});

test("production lookupは読取失敗・壊れた台帳・不正UUIDをnullにしない", () => {
  const brokenSpreadsheets = [
    [null, /SPREADSHEET_UNAVAILABLE/],
    [{ getSheetByName: () => null }, /SHEET_UNAVAILABLE/],
    [{ getSheetByName: () => ({ getDataRange: () => {
      throw new Error("range");
    } }) }, /READ_UNAVAILABLE/],
    [{ getSheetByName: () => ({ getDataRange: () => ({
      getValues: () => { throw new Error("values"); }
    }) }) }, /READ_UNAVAILABLE/]
  ];
  for (const [spreadsheet, reason] of brokenSpreadsheets) {
    const context = fixture({ spreadsheet });
    assert.throws(() => context.findEmployeeRegistrationOperationFromSpreadsheetReadOnly_(uuidA),
      reason);
  }
  const failedOpen = fixture();
  failedOpen.SpreadsheetApp.openById = () => { throw new Error("open"); };
  assert.throws(() => failedOpen.findEmployeeRegistrationOperationFromSpreadsheetReadOnly_(uuidA),
    /SPREADSHEET_UNAVAILABLE/);
  const brokenLedgers = [
    { values: [["wrong_header"]], reason: /HEADERS_MISSING/ },
    { values: [headers, cells(headers, { ...operation, status: "BROKEN" })],
      reason: /ROW_INVALID/ },
    { values: [headers, cells(headers, operation), cells(headers, operation)],
      reason: /OPERATION_DUPLICATE/ },
    { values: [headers, cells(headers, operation), cells(headers, {
      ...operation, operation_id: uuidB, display_employee_id: "P0023" })],
      reason: /EMPLOYEE_ID_DUPLICATE/ },
    { values: [headers, cells(headers, operation), cells(headers, {
      ...operation, operation_id: uuidB, employee_id: "EMP0084" })],
      reason: /DISPLAY_ID_DUPLICATE/ }
  ];
  for (const { values, reason } of brokenLedgers) {
    const context = fixture({ ledgerValues: values });
    assert.throws(() => context.findEmployeeRegistrationOperationFromSpreadsheetReadOnly_(uuidA),
      reason);
  }
  assert.throws(() => fixture().findEmployeeRegistrationOperationFromSpreadsheetReadOnly_("bad"),
    /LOOKUP_ID_INVALID/);
});

test("社員snapshotとのcross-integrityは正しい対応と旧社員を許可する", () => {
  const old = { ...employee, employee_id: "EMP0082", display_employee_id: "W0061",
    registration_operation_id: "" };
  const context = fixture({ employeeValues: [employeeHeaders,
    cells(employeeHeaders, old), cells(employeeHeaders, employee)] });
  const ledger = context.readEmployeeRegistrationOperationLedgerReadOnly_();
  const employees = context.readEmployeeRegistrationEmployeesReadOnly_("registration_ready");
  const ids = context.verifyEmployeeRegistrationLedgerEmployeeIntegrity_(ledger, employees);
  assert.equal(JSON.stringify(ids.existingIds), JSON.stringify([
    "EMP0082", "EMP0083", "W0061", "W0062"]));
  assert.equal(JSON.stringify(ids.operationIds), JSON.stringify(["EMP0083", "W0062"]));
});

test("旧社員や別社員との予約ID衝突・orphan・列欠落を拒否", () => {
  const cases = [
    { ...employee, employee_id: "EMP0083", display_employee_id: "W0061",
      registration_operation_id: "" },
    { ...employee, employee_id: "EMP0084", display_employee_id: "W0062",
      registration_operation_id: "" },
    { ...employee, employee_id: "EMP0084", display_employee_id: "W0063",
      registration_operation_id: uuidA },
    { ...employee, employee_id: "EMP0084", display_employee_id: "W0063",
      registration_operation_id: uuidB }
  ];
  for (const badEmployee of cases) {
    const context = fixture({ employeeValues: [employeeHeaders,
      cells(employeeHeaders, badEmployee)] });
    const ledger = context.readEmployeeRegistrationOperationLedgerReadOnly_();
    const employees = context.readEmployeeRegistrationEmployeesReadOnly_("registration_ready");
    assert.throws(() => context.verifyEmployeeRegistrationLedgerEmployeeIntegrity_(ledger,
      employees), /RESERVATION_CONFLICT|EMPLOYEE_LINK_INVALID/);
  }
  const legacyHeaders = employeeHeaders.filter(key => key !== "registration_operation_id");
  const context = fixture({ employeeValues: [legacyHeaders, cells(legacyHeaders, employee)] });
  assert.throws(() => context.verifyEmployeeRegistrationLedgerEmployeeIntegrity_(
    context.readEmployeeRegistrationOperationLedgerReadOnly_(),
    context.readEmployeeRegistrationEmployeesReadOnly_("legacy_read_only")),
  /EMPLOYEE_SNAPSHOT_INVALID/);
});

test("Phase 1 recoveryへledger lookupとPhase 2A employeeRowsを渡す", () => {
  for (const [status, withEmployee, expected] of [
    [null, false, "new"], ["STARTED", false, "started_without_employee"],
    ["STARTED", true, "started_with_employee"],
    ["EMPLOYEE_SAVED", false, "employee_missing"],
    ["EMPLOYEE_SAVED", true, "employee_saved"],
    ["COMPLETED", false, "employee_missing"],
    ["COMPLETED", true, "completed"]
  ]) {
    const context = fixture({ ledgerValues: [headers], employeeValues: withEmployee ?
      [employeeHeaders, cells(employeeHeaders, employee)] : [employeeHeaders] });
    const normalized = context.normalizeEmployeeRegistrationInput_(
      Object.fromEntries(inputFields.map(key => [key, employee[key]])));
    const hash = context.employeeRegistrationInputHash_(normalized,
      context.employeeRegistrationSha256HexAppsScript_);
    const rowOperation = { ...operation, input_hash: hash, status: status || "STARTED" };
    if (status) {
      const snapshot = context.employeeRegistrationProjectOperationLedger_(
        [headers, cells(headers, rowOperation)]);
      context.SpreadsheetApp.openById = () => ({ getSpreadsheetTimeZone: () => "Asia/Tokyo",
        getSheetByName: name => ({ getDataRange: () => ({ getValues: () =>
          name === "employee_registration_operations" ?
            [headers, cells(headers, rowOperation)] : withEmployee ?
              [employeeHeaders, cells(employeeHeaders, employee)] : [employeeHeaders] }) }) });
      assert.equal(snapshot.operations.length, 1);
    }
    const employees = context.readEmployeeRegistrationEmployeesReadOnly_("registration_ready");
    const found = context.findEmployeeRegistrationOperationFromSpreadsheetReadOnly_(uuidA);
    const result = context.classifyEmployeeRegistrationRecovery_(found, employees.employeeRows,
      { operation_id: uuidA, input_hash: hash, created_by: "ADMIN1",
        normalized_input: normalized }, context.employeeRegistrationSha256HexAppsScript_,
      employees.formatDateKey);
    assert.equal(result.kind, expected);
  }
});

test("UUID大小文字差のlookup候補はPhase 1でoperation_mismatchとなる", () => {
  const context = fixture({ ledgerValues: [headers, cells(headers, {
    ...operation, operation_id: uuidA.toUpperCase() })],
  employeeValues: [employeeHeaders] });
  const normalized = context.normalizeEmployeeRegistrationInput_(
    Object.fromEntries(inputFields.map(key => [key, employee[key]])));
  const hash = context.employeeRegistrationInputHash_(normalized,
    context.employeeRegistrationSha256HexAppsScript_);
  const found = context.findEmployeeRegistrationOperationFromSpreadsheetReadOnly_(uuidA);
  assert.equal(found.operation_id, uuidA.toUpperCase());
  const employees = context.readEmployeeRegistrationEmployeesReadOnly_("registration_ready");
  const result = context.classifyEmployeeRegistrationRecovery_(found, employees.employeeRows,
    { operation_id: uuidA, input_hash: hash, created_by: "ADMIN1",
      normalized_input: normalized }, context.employeeRegistrationSha256HexAppsScript_,
    employees.formatDateKey);
  assert.equal(result.kind, "operation_mismatch");
});

test("COMPLETED予約を含む一覧はPhase 1採番へ渡せる", () => {
  const completed = { ...operation, status: "COMPLETED" };
  const context = fixture({ ledgerValues: [headers, cells(headers, completed)],
    employeeValues: [employeeHeaders] });
  const ledger = context.readEmployeeRegistrationOperationLedgerReadOnly_();
  const employees = context.readEmployeeRegistrationEmployeesReadOnly_("registration_ready");
  const ids = context.verifyEmployeeRegistrationLedgerEmployeeIntegrity_(ledger, employees);
  assert.equal(context.nextEmployeeRegistrationId_("EMP", ids.existingIds, ids.operationIds, 82),
    "EMP0084");
  assert.equal(context.nextEmployeeRegistrationId_("W", ids.existingIds, ids.operationIds, 61),
    "W0063");
});
