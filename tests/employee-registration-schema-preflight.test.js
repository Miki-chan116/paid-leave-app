const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const root = path.join(__dirname, "..");
const files = ["employee-registration-safety.gs", "employee-registration-adapter.gs",
  "employee-registration-operation-ledger.gs", "employee-registration-schema-preflight.gs"];
const fields = ["name", "display_name", "name_kana", "company_code", "company_name",
  "department", "employment_type", "employment_status", "hire_date", "leave_date",
  "work_days_per_week", "work_start_minute", "work_end_minute", "fiscal_start_month",
  "leave_management_target", "is_driver", "driver_type", "default_vehicle_id", "notes"];
const ledgerHeaders = ["operation_id", "input_hash", "created_by", "status",
  "employee_id", "display_employee_id", "created_at", "updated_at"];
const uuid = "123e4567-e89b-42d3-a456-426614174000";
const employee = {
  employee_id: "EMP0083", display_employee_id: "W0062", name: "Test User",
  display_name: "", name_kana: "てすと", company_code: "MAIN", company_name: "",
  department: "", employment_type: "regular", employment_status: "active",
  hire_date: "2026-04-01", leave_date: "", work_days_per_week: 5,
  work_start_minute: "", work_end_minute: "", fiscal_start_month: 4,
  leave_management_target: true, is_driver: false, driver_type: "",
  default_vehicle_id: "", notes: "", registration_operation_id: ""
};
const cells = (headers, object) => headers.map(header => object[header] ?? "");
const range = (column, width = 1) => ({ getColumn: () => column, getNumColumns: () => width });

function load(options = {}) {
  const headers = options.headers || ["employee_id", "display_employee_id", ...fields, ""];
  const rows = options.rows || [cells(headers, employee)];
  const employeeValues = options.employeeValues || [headers, ...rows];
  const maxRows = options.maxRows ?? 6;
  const maxColumns = options.maxColumns ?? headers.length;
  const emptyMeta = options.emptyMeta || {};
  let rangeReads = 0;
  const data = key => Array.from({ length: maxRows }, (_, index) => [
    emptyMeta[key] && Object.prototype.hasOwnProperty.call(emptyMeta[key], index)
      ? emptyMeta[key][index] : key === "validations" ? null : ""
  ]);
  const target = {
    getValues: () => options.columnValues || data("values"),
    getFormulas: () => data("formulas"), getNotes: () => data("notes"),
    getDataValidations: () => data("validations"),
    getMergedRanges: () => emptyMeta.merged || []
  };
  const employeeSheet = {
    getMaxRows: () => maxRows, getMaxColumns: () => maxColumns,
    getDataRange: () => ({ getValues: () => employeeValues }),
    getRange: (row, col, count, width) => {
      rangeReads++;
      assert.equal(row, 1); assert.equal(count, maxRows); assert.equal(width, 1);
      assert.ok(headers[col - 1] === "" ||
        (headers.indexOf("") === -1 && col === headers.length + 1 && col <= maxColumns));
      return target;
    },
    getProtections: type => type === "SHEET" ? (emptyMeta.sheetProtections || []) :
      (emptyMeta.protections || []),
    getNamedRanges: () => emptyMeta.named || [],
    getConditionalFormatRules: () => emptyMeta.rules || [],
    getFilter: () => emptyMeta.filter || null
  };
  const ledgerValues = options.ledgerValues === undefined ? null : options.ledgerValues;
  const ledgerSheet = ledgerValues === null ? null : {
    getDataRange: () => ({ getValues: () => ledgerValues })
  };
  const spreadsheet = options.spreadsheet || {
    getSheetByName: name => name === "employees" ? employeeSheet :
      name === "employee_registration_operations" ? ledgerSheet : null,
    getSpreadsheetTimeZone: () => "Asia/Tokyo"
  };
  const context = { Date, Number, String, Set, Map, SS_ID: "fixture",
    SpreadsheetApp: { openById: () => spreadsheet, ProtectionType: { RANGE: "RANGE", SHEET: "SHEET" } },
    Utilities: { formatDate: date => date.toISOString().slice(0, 10) } };
  vm.createContext(context);
  for (const file of files) vm.runInContext(fs.readFileSync(path.join(root, file), "utf8"),
    context, { filename: file });
  return { context, employeeSheet, spreadsheet, target,
    getRangeReads: () => rangeReads };
}

const inspect = options => load(options).context.inspectEmployeeRegistrationSchemaReadOnly_();
const types = result => Array.from(result.plan, item => item.type);

test("現行の空ヘッダー列を全物理行検査し、読取専用planを返す", () => {
  const result = inspect();
  assert.equal(result.state, "READY_FOR_MIGRATION");
  assert.equal(result.ok, true);
  assert.deepEqual(types(result), ["REUSE_EMPTY_COLUMN", "CREATE_OPERATION_LEDGER"]);
  assert.equal(result.plan[0].column, 22);
  assert.ok(result.issues.includes("LEGACY_REQUIRED_COLUMN_MISSING"));
  assert.equal(result.employees.legacyRequiredColumnNext, 23);
  assert.equal(result.employees.legacyRequiredColumnCapacityAvailable, false);
  assert.equal(result.employees.maxRows, 6);
  assert.equal(result.ledger.exists, false);
});

test("production入口にsnapshotを渡してもSpreadsheet読取を迂回できない", () => {
  const { context } = load();
  context.SpreadsheetApp.openById = () => { throw new Error("offline"); };
  assert.throws(() => context.inspectEmployeeRegistrationSchemaReadOnly_({
    employeeValues: [["fake"]], ledgerValues: null
  }), /SPREADSHEET_UNAVAILABLE/);
});

test("Spreadsheetとemployeesの取得失敗は例外でありmissingではない", () => {
  const missing = load({ spreadsheet: { getSheetByName: () => null } });
  assert.throws(() => missing.context.inspectEmployeeRegistrationSchemaReadOnly_(),
    /EMPLOYEES_SHEET_UNAVAILABLE/);
  const broken = load({ spreadsheet: { getSheetByName: () => { throw Error("read"); } } });
  assert.throws(() => broken.context.inspectEmployeeRegistrationSchemaReadOnly_(),
    /EMPLOYEES_SHEET_UNAVAILABLE/);
  const opened = load();
  opened.context.SpreadsheetApp.openById = () => null;
  assert.throws(() => opened.context.inspectEmployeeRegistrationSchemaReadOnly_(),
    /SPREADSHEET_UNAVAILABLE/);
});

test("物理行数・列数の取得失敗と不正値は例外", () => {
  for (const [method, expected] of [["getMaxRows", /MAX_ROWS_UNAVAILABLE/],
    ["getMaxColumns", /MAX_COLUMNS_UNAVAILABLE/]]) {
    const item = load(); item.employeeSheet[method] = () => { throw Error("read"); };
    assert.throws(() => item.context.inspectEmployeeRegistrationSchemaReadOnly_(), expected);
  }
  assert.throws(() => inspect({ maxRows: 0 }), /DIMENSIONS_INVALID/);
  assert.throws(() => inspect({ maxColumns: 0 }), /DIMENSIONS_INVALID/);
});

test("employees header/range/getValuesと列metadata読取失敗は例外", () => {
  for (const method of ["getDataRange", "getRange"]) {
    const item = load(); item.employeeSheet[method] = () => { throw Error("read"); };
    assert.throws(() => item.context.inspectEmployeeRegistrationSchemaReadOnly_(), /UNAVAILABLE/);
  }
  const item = load(); item.target.getNotes = () => { throw Error("notes"); };
  assert.throws(() => item.context.inspectEmployeeRegistrationSchemaReadOnly_(),
    /COLUMN_NOTES_UNAVAILABLE/);
  const values = load(); values.employeeSheet.getDataRange = () => ({
    getValues: () => { throw Error("values"); }
  });
  assert.throws(() => values.context.inspectEmployeeRegistrationSchemaReadOnly_(),
    /EMPLOYEES_READ_UNAVAILABLE/);
  assert.throws(() => inspect({ employeeValues: [] }), /EMPLOYEES_READ_INVALID/);
});

test("台帳missingは正常不在、lookup/read failureは例外", () => {
  assert.equal(inspect().ledger.status, "MISSING");
  const lookup = load(); lookup.spreadsheet.getSheetByName = name => {
    if (name === "employees") return lookup.employeeSheet;
    throw Error("lookup");
  };
  assert.throws(() => lookup.context.inspectEmployeeRegistrationSchemaReadOnly_(),
    /LEDGER_LOOKUP_UNAVAILABLE/);
  const absentResult = load();
  absentResult.spreadsheet.getSheetByName = name =>
    name === "employees" ? absentResult.employeeSheet : undefined;
  assert.throws(() => absentResult.context.inspectEmployeeRegistrationSchemaReadOnly_(),
    /LEDGER_LOOKUP_UNAVAILABLE/);
  const read = load({ ledgerValues: [ledgerHeaders] });
  read.spreadsheet.getSheetByName = name => name === "employees" ? read.employeeSheet :
    { getDataRange: () => { throw Error("read"); } };
  assert.throws(() => read.context.inspectEmployeeRegistrationSchemaReadOnly_(),
    /LEDGER_READ_UNAVAILABLE/);
  const values = load({ ledgerValues: [ledgerHeaders] });
  values.spreadsheet.getSheetByName = name => name === "employees" ? values.employeeSheet :
    { getDataRange: () => ({ getValues: () => { throw Error("values"); } }) };
  assert.throws(() => values.context.inspectEmployeeRegistrationSchemaReadOnly_(),
    /LEDGER_READ_UNAVAILABLE/);
});

test("値・数式・note・validationが物理行末にあれば再利用しない", () => {
  for (const [key, value, reason] of [["values", "hidden", "VALUE"],
    ["formulas", '=IF(TRUE,"","")', "FORMULA"], ["notes", "note", "NOTE"],
    ["validations", {}, "VALIDATION"]]) {
    const result = inspect({ emptyMeta: { [key]: { 5: value } } });
    assert.equal(result.state, "MANUAL_INVESTIGATION_REQUIRED", key);
    assert.deepEqual(types(result), []);
    assert.ok(result.employees.emptyColumns[0].reasons.includes(reason));
  }
});

test("merge・range/sheet protection・named range・conditional format・filterを拒否", () => {
  const cases = [
    [{ merged: [range(22)] }, "MERGE"],
    [{ protections: [{ getRange: () => range(20, 3) }] }, "PROTECTION"],
    [{ sheetProtections: [{}] }, "PROTECTION"],
    [{ named: [{ getRange: () => range(22) }] }, "NAMED_RANGE"],
    [{ rules: [{ getRanges: () => [range(22)] }] }, "CONDITIONAL_FORMAT"],
    [{ filter: { getRange: () => range(1, 22) } }, "FILTER"]
  ];
  for (const [emptyMeta, reason] of cases) {
    const result = inspect({ emptyMeta });
    assert.equal(result.state, "MANUAL_INVESTIGATION_REQUIRED", reason);
    assert.deepEqual(types(result), []);
    assert.ok(result.employees.emptyColumns[0].reasons.includes(reason));
  }
});

test("空ヘッダーの重複、必須ヘッダー欠落、重複、余白はmanual", () => {
  const base = ["employee_id", "display_employee_id", ...fields];
  for (const headers of [[...base, "", ""], [...base.filter(x => x !== "name"), ""],
    [...base, "employee_id"], [...base, " name"]]) {
    const result = inspect({ headers, maxColumns: headers.length });
    assert.equal(result.state, "MANUAL_INVESTIGATION_REQUIRED");
    assert.deepEqual(types(result), []);
  }
});

test("operation列が既に存在し、社員行が空なら台帳作成だけのpartial", () => {
  const headers = ["employee_id", "display_employee_id", ...fields,
    "registration_operation_id"];
  const result = inspect({ headers });
  assert.equal(result.state, "PARTIAL_MIGRATION");
  assert.deepEqual(types(result), ["CREATE_OPERATION_LEDGER"]);
});

test("operation列の有効UUIDでも台帳がなければmanual、不正・重複もmanual", () => {
  const headers = ["employee_id", "display_employee_id", ...fields,
    "registration_operation_id"];
  const valid = inspect({ headers, rows: [cells(headers,
    { ...employee, registration_operation_id: uuid })] });
  assert.equal(valid.state, "MANUAL_INVESTIGATION_REQUIRED");
  assert.ok(valid.issues.includes("EMPLOYEE_OPERATION_WITHOUT_LEDGER_RECORD"));
  const malformed = inspect({ headers, rows: [cells(headers,
    { ...employee, registration_operation_id: "bad" })] });
  assert.equal(malformed.state, "MANUAL_INVESTIGATION_REQUIRED");
  const duplicate = inspect({ headers, rows: [cells(headers,
    { ...employee, registration_operation_id: uuid }), cells(headers,
    { ...employee, employee_id: "EMP0084", display_employee_id: "W0063",
      registration_operation_id: uuid })] });
  assert.equal(duplicate.state, "MANUAL_INVESTIGATION_REQUIRED");
  const duplicateHeader = inspect({ headers: [...headers, "registration_operation_id"] });
  assert.equal(duplicateHeader.state, "MANUAL_INVESTIGATION_REQUIRED");
});

test("空ヘッダーなし・operation列なしは末尾追加と必要容量を提案", () => {
  const headers = ["employee_id", "display_employee_id", ...fields];
  const needed = inspect({ headers });
  assert.equal(needed.state, "READY_FOR_MIGRATION");
  assert.deepEqual(types(needed), ["ENSURE_COLUMN_CAPACITY",
    "ADD_REGISTRATION_OPERATION_ID_COLUMN", "CREATE_OPERATION_LEDGER"]);
  const enough = inspect({ headers, maxColumns: headers.length + 2 });
  assert.deepEqual(types(enough), ["ADD_REGISTRATION_OPERATION_ID_COLUMN",
    "CREATE_OPERATION_LEDGER"]);
  assert.equal(enough.employees.tailColumn.column, headers.length + 1);
  assert.equal(enough.employees.tailColumn.reusable, true);
});

test("既存physical末尾列が安全な場合だけ列追加planを返す", () => {
  const headers = ["employee_id", "display_employee_id", ...fields];
  const result = inspect({ headers, maxColumns: headers.length + 1 });
  assert.equal(result.state, "READY_FOR_MIGRATION");
  assert.deepEqual(types(result), ["ADD_REGISTRATION_OPERATION_ID_COLUMN",
    "CREATE_OPERATION_LEDGER"]);
  assert.equal(result.plan[0].column, headers.length + 1);
  assert.equal(result.employees.tailColumn.reusable, true);
});

test("既存physical末尾列の値・空表示formula・note・validationは手動調査", () => {
  const headers = ["employee_id", "display_employee_id", ...fields];
  assert.equal(headers.length, 21);
  for (const [key, value, reason] of [["values", "occupied", "VALUE"],
    ["formulas", '= ""', "FORMULA"], ["notes", "hidden note", "NOTE"],
    ["validations", {}, "VALIDATION"]]) {
    const result = inspect({ headers, maxColumns: headers.length + 1,
      emptyMeta: { [key]: { 5: value } } });
    assert.equal(result.state, "MANUAL_INVESTIGATION_REQUIRED", key);
    assert.deepEqual(types(result), []);
    assert.ok(result.issues.includes("EMPLOYEE_TARGET_COLUMN_UNSAFE"));
    assert.equal(result.employees.tailColumn.column, 22);
    assert.ok(result.employees.tailColumn.reasons.includes(reason));
  }
});

test("既存physical末尾列と交差するmetadataは手動調査", () => {
  const headers = ["employee_id", "display_employee_id", ...fields];
  const column = headers.length + 1;
  const cases = [
    [{ merged: [range(column - 1, 2)] }, "MERGE"],
    [{ protections: [{ getRange: () => range(column - 1, 2) }] }, "PROTECTION"],
    [{ named: [{ getRange: () => range(column) }] }, "NAMED_RANGE"],
    [{ rules: [{ getRanges: () => [range(1), range(column)] }] }, "CONDITIONAL_FORMAT"],
    [{ filter: { getRange: () => range(1, column) } }, "FILTER"]
  ];
  for (const [emptyMeta, reason] of cases) {
    const result = inspect({ headers, maxColumns: column, emptyMeta });
    assert.equal(result.state, "MANUAL_INVESTIGATION_REQUIRED", reason);
    assert.deepEqual(types(result), []);
    assert.ok(result.employees.tailColumn.reasons.includes(reason));
  }
});

test("物理容量外の末尾列は読まず、容量確保から提案する", () => {
  const headers = ["employee_id", "display_employee_id", ...fields, "extra"];
  const item = load({ headers, maxColumns: headers.length });
  const result = item.context.inspectEmployeeRegistrationSchemaReadOnly_();
  assert.equal(result.state, "READY_FOR_MIGRATION");
  assert.deepEqual(types(result), ["ENSURE_COLUMN_CAPACITY",
    "ADD_REGISTRATION_OPERATION_ID_COLUMN", "CREATE_OPERATION_LEDGER"]);
  assert.equal(result.plan[0].minimum, headers.length + 1);
  assert.equal(result.plan[1].column, headers.length + 1);
  assert.equal(item.getRangeReads(), 0);
});

test("既存physical末尾列の検査失敗はREADYへfallbackしない", () => {
  const headers = ["employee_id", "display_employee_id", ...fields];
  for (const method of ["getValues", "getFormulas", "getNotes",
    "getDataValidations", "getMergedRanges"]) {
    const item = load({ headers, maxColumns: headers.length + 1 });
    item.target[method] = () => { throw Error("read unavailable"); };
    assert.throws(() => item.context.inspectEmployeeRegistrationSchemaReadOnly_(),
      /UNAVAILABLE/, method);
  }
  for (const method of ["getProtections", "getNamedRanges",
    "getConditionalFormatRules", "getFilter"]) {
    const item = load({ headers, maxColumns: headers.length + 1 });
    item.employeeSheet[method] = () => { throw Error("read unavailable"); };
    assert.throws(() => item.context.inspectEmployeeRegistrationSchemaReadOnly_(),
      /UNAVAILABLE/, method);
  }
});

test("台帳の完全空状態は既存operation列と組み合わせてmigrated", () => {
  const headers = ["employee_id", "display_employee_id", ...fields,
    "registration_operation_id"];
  const result = inspect({ headers, ledgerValues: [ledgerHeaders] });
  assert.equal(result.state, "ALREADY_MIGRATED");
  assert.deepEqual(types(result), []);
});

test("正確な空prefixだけ補完候補とし、データ付きprefixはmanual", () => {
  const headers = ["employee_id", "display_employee_id", ...fields,
    "registration_operation_id"];
  const prefix = ledgerHeaders.slice(0, 3);
  const good = inspect({ headers, ledgerValues: [prefix] });
  assert.equal(good.state, "PARTIAL_MIGRATION");
  assert.deepEqual(types(good), ["COMPLETE_LEDGER_HEADERS"]);
  const bad = inspect({ headers, ledgerValues: [prefix, [uuid, "hash", "admin"]] });
  assert.equal(bad.state, "MANUAL_INVESTIGATION_REQUIRED");
  assert.deepEqual(types(bad), []);
});

test("台帳の順序違い・余分・重複・不正データはmanual", () => {
  const headers = ["employee_id", "display_employee_id", ...fields,
    "registration_operation_id"];
  const cases = [[ledgerHeaders[1], ledgerHeaders[0]],
    [...ledgerHeaders, "extra"], [...ledgerHeaders, "status"],
    ["operation_id", "input_hash", "bad"]];
  for (const names of cases) {
    const result = inspect({ headers, ledgerValues: [names] });
    assert.equal(result.state, "MANUAL_INVESTIGATION_REQUIRED");
  }
  const invalid = inspect({ headers, ledgerValues: [ledgerHeaders,
    ledgerHeaders.map(() => "bad")] });
  assert.equal(invalid.state, "MANUAL_INVESTIGATION_REQUIRED");
});

test("完全な有データ台帳は整合性が取れる場合のみmigrated", () => {
  const headers = ["employee_id", "display_employee_id", ...fields,
    "registration_operation_id"];
  const operation = { operation_id: uuid, input_hash: "v1:" + "a".repeat(64),
    created_by: "ADMIN1", status: "COMPLETED", employee_id: "EMP0083",
    display_employee_id: "W0062", created_at: new Date("2026-04-01T00:00:00Z"),
    updated_at: new Date("2026-04-01T01:00:00Z") };
  const ledgerValues = [ledgerHeaders, cells(ledgerHeaders, operation)];
  const rows = [cells(headers, { ...employee, registration_operation_id: uuid })];
  assert.equal(inspect({ headers, rows, ledgerValues }).state, "ALREADY_MIGRATED");
  assert.equal(inspect({ headers, ledgerValues }).state, "MANUAL_INVESTIGATION_REQUIRED");
});

test("新ファイルはSpreadsheet mutation APIを呼ばない", () => {
  const source = fs.readFileSync(path.join(root,
    "employee-registration-schema-preflight.gs"), "utf8");
  const executable = source.replace(/\/\*[\s\S]*?\*\/|\/\/[^\n]*/g, "");
  assert.doesNotMatch(executable,
    /\.(?:setValue|setValues|appendRow|insert\w+|delete\w+|clear\w*|setFormula|setNote|setDataValidation|setConditionalFormatRules|setNamedRange|protect|unprotect|flush)\s*\(/);
});
