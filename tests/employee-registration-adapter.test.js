const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");
const crypto = require("node:crypto");

const safetyFile = path.join(__dirname, "..", "employee-registration-safety.gs");
const adapterFile = path.join(__dirname, "..", "employee-registration-adapter.gs");
const uuid = "123e4567-e89b-42d3-a456-426614174000";
const fields = [
  "name", "display_name", "name_kana", "company_code", "company_name",
  "department", "employment_type", "employment_status", "hire_date",
  "leave_date", "work_days_per_week", "work_start_minute", "work_end_minute",
  "fiscal_start_month", "leave_management_target", "is_driver", "driver_type",
  "default_vehicle_id", "notes"
];
const headers = ["employee_id", "display_employee_id", ...fields];
const employee = {
  employee_id: "EMP0083", display_employee_id: "W0062",
  name: "例 一郎", display_name: "", name_kana: "れい いちろう",
  company_code: "MAIN", company_name: "", department: "",
  employment_type: "regular", employment_status: "active",
  hire_date: "2026-04-01", leave_date: "", work_days_per_week: 5,
  work_start_minute: "", work_end_minute: "", fiscal_start_month: 4,
  leave_management_target: true, is_driver: false, driver_type: "",
  default_vehicle_id: "", notes: ""
};
const row = (names, data) => names.map(name => data[name]);
const dateFormat = (date, zone) => {
  const parts = new Intl.DateTimeFormat("en-US", {
    timeZone: zone, year: "numeric", month: "2-digit", day: "2-digit"
  }).formatToParts(date);
  const values = Object.fromEntries(parts.map(part => [part.type, part.value]));
  return `${values.year}-${values.month}-${values.day}`;
};

function harness(options = {}) {
  const sheetValues = options.values || [headers, row(headers, employee)];
  const spreadsheet = options.spreadsheet === undefined ? {
    getSpreadsheetTimeZone: () => options.timeZone === undefined ? "Asia/Tokyo" : options.timeZone,
    getSheetByName: name => name === "employees" ? {
      getDataRange: () => ({ getValues: () => sheetValues })
    } : null
  } : options.spreadsheet;
  const calls = [];
  const context = {
    Date, Number, String, Set,
    SS_ID: "test-spreadsheet",
    SpreadsheetApp: {
      openById: id => { calls.push(["openById", id]); return spreadsheet; }
    },
    Utilities: {
      DigestAlgorithm: { SHA_256: "SHA_256" }, Charset: { UTF_8: "UTF_8" },
      computeDigest: (algorithm, text, charset) => {
        calls.push(["computeDigest", algorithm, text, charset]);
        return Array.from(crypto.createHash("sha256").update(text, "utf8").digest(),
          byte => byte > 127 ? byte - 256 : byte);
      },
      formatDate: (date, zone, pattern) => {
        calls.push(["formatDate", zone, pattern]);
        return dateFormat(date, zone);
      }
    }
  };
  vm.createContext(context);
  vm.runInContext(fs.readFileSync(safetyFile, "utf8"), context, { filename: safetyFile });
  vm.runInContext(fs.readFileSync(adapterFile, "utf8"), context, { filename: adapterFile });
  return { context, calls };
}

test("SHA-256はASCII・日本語・改行・タブをNodeのUTF-8結果と一致させる", () => {
  const { context, calls } = harness();
  for (const input of ["abc", "社員登録", "line1\nline2", "a\tb", "氏名\n\t例"]) {
    const actual = context.employeeRegistrationSha256HexAppsScript_(input);
    const expected = crypto.createHash("sha256").update(input, "utf8").digest("hex");
    assert.equal(actual, expected);
    assert.match(actual, /^[a-f0-9]{64}$/);
  }
  assert.ok(calls.filter(call => call[0] === "computeDigest")
    .every(call => call[1] === "SHA_256" && call[3] === "UTF_8"));
});

test("Phase 1 canonical payloadと入力hashがNodeと一致する", () => {
  const { context } = harness();
  const normalized = context.normalizeEmployeeRegistrationInput_({
    ...Object.fromEntries(fields.map(key => [key, employee[key]])),
    notes: "改行\n日本語\t確認"
  });
  const canonical = context.employeeRegistrationCanonicalPayload_(normalized);
  const expected = "v1:" + crypto.createHash("sha256").update(canonical, "utf8").digest("hex");
  assert.equal(context.employeeRegistrationInputHash_(
    normalized, context.employeeRegistrationSha256HexAppsScript_), expected);
});

test("signed digest byteを0〜255へ戻し、不正byte列を拒否する", () => {
  const { context } = harness();
  assert.equal(context.employeeRegistrationDigestBytesToHex_(
    Array(32).fill(-1)), "ff".repeat(32));
  for (const bytes of [null, [], Array(32), Array(32).fill(256),
    Array(32).fill(-129), Array(32).fill(1.5)]) {
    assert.throws(() => context.employeeRegistrationDigestBytesToHex_(bytes), /DIGEST_INVALID/);
  }
  assert.throws(() => context.employeeRegistrationSha256HexAppsScript_(42), /DIGEST_INPUT_INVALID/);
});

test("Spreadsheet timezoneで通常日・UTC境界・うるう日を日付キーにする", () => {
  const { context, calls } = harness();
  const format = context.employeeRegistrationSpreadsheetDateFormatter_(
    { getSpreadsheetTimeZone: () => "Asia/Tokyo" });
  assert.equal(format(new Date("2026-04-01T00:00:00+09:00")), "2026-04-01");
  assert.equal(format(new Date("2026-03-31T15:30:00Z")), "2026-04-01");
  assert.equal(format(new Date("2024-02-28T15:00:00Z")), "2024-02-29");
  assert.ok(calls.filter(call => call[0] === "formatDate")
    .every(call => call[1] === "Asia/Tokyo" && call[2] === "yyyy-MM-dd"));
});

test("timezone欠落、Invalid Date、壊れたformat結果・例外を拒否する", () => {
  const { context } = harness();
  for (const sheet of [null, {}, { getSpreadsheetTimeZone: () => "" },
    { getSpreadsheetTimeZone: () => { throw new Error("failed"); } }]) {
    assert.throws(() => context.employeeRegistrationSpreadsheetDateFormatter_(sheet), /UNAVAILABLE/);
  }
  const format = context.employeeRegistrationSpreadsheetDateFormatter_(
    { getSpreadsheetTimeZone: () => "Asia/Tokyo" });
  assert.throws(() => format(new Date("invalid")), /DATE_INVALID/);
  for (const result of ["2026-02-30", "2026/04/01", "bad", null]) {
    assert.throws(() => context.employeeRegistrationDateKeyWithFormatter_(
      new Date("2026-04-01T00:00:00Z"), "Asia/Tokyo", () => result), /DATE_INVALID/);
  }
  assert.throws(() => context.employeeRegistrationDateKeyWithFormatter_(
    new Date(), "Asia/Tokyo", () => { throw new Error("failed"); }), /DATE_INVALID/);
});

test("正常ヘッダーと0件データは空の一覧を返す", () => {
  const { context } = harness({ values: [headers] });
  const snapshot = context.readEmployeeRegistrationEmployeesReadOnly_("legacy_read_only");
  assert.equal(snapshot.employeeRows.length, 0);
  assert.equal(snapshot.employeeIds.length, 0);
  assert.equal(snapshot.displayEmployeeIds.length, 0);
  assert.equal(snapshot.existingIds.length, 0);
  assert.equal(snapshot.hasRegistrationOperationIdColumn, false);
  assert.equal(Object.hasOwn(snapshot, "operationIds"), false);
});

test("必須ヘッダー欠落・重複・空欄・余白・ヘッダー行欠落を拒否する", () => {
  for (const values of [
    [headers.filter(name => name !== "name")],
    [[...headers, "name"]],
    [[...headers, ""]],
    [[...headers, " name"]],
    [[""]],
    []
  ]) {
    const { context } = harness({ values });
    assert.throws(() => context.readEmployeeRegistrationEmployeesReadOnly_("legacy_read_only"),
      /HEADERS_|READ_INVALID/);
  }
});

test("旧社員の1行・複数行を全項目付きplain objectとID一覧へ投影する", () => {
  const second = { ...employee, employee_id: "EMP0084", display_employee_id: "P0023",
    company_code: "PARTNER", fiscal_start_month: 6,
    hire_date: new Date("2026-03-31T15:30:00Z") };
  const { context } = harness({ values: [headers, row(headers, employee), row(headers, second)] });
  const snapshot = context.readEmployeeRegistrationEmployeesReadOnly_("legacy_read_only");
  assert.equal(snapshot.employeeRows.length, 2);
  assert.equal(snapshot.employeeRows[0].registration_operation_id, "");
  assert.equal(snapshot.employeeRows[1].hire_date, second.hire_date);
  assert.equal(snapshot.formatDateKey(second.hire_date), "2026-04-01");
  assert.equal(Object.getPrototypeOf(snapshot.employeeRows[0]),
    vm.runInContext("Object.prototype", context));
  assert.equal(JSON.stringify(snapshot.existingIds),
    JSON.stringify(["EMP0083", "EMP0084", "W0062", "P0023"]));
  assert.equal(Object.hasOwn(snapshot, "operationIds"), false);
});

test("Dateを保持した投影行をPhase 1 recoveryがSpreadsheet日付として照合する", () => {
  const names = [...headers, "registration_operation_id"];
  const saved = { ...employee, hire_date: new Date("2026-03-31T15:30:00Z"),
    registration_operation_id: uuid };
  const { context } = harness({ values: [names, row(names, saved)] });
  const snapshot = context.readEmployeeRegistrationEmployeesReadOnly_("registration_ready");
  const normalized = context.normalizeEmployeeRegistrationInput_(
    Object.fromEntries(fields.map(key => [key, employee[key]])));
  const inputHash = context.employeeRegistrationInputHash_(
    normalized, context.employeeRegistrationSha256HexAppsScript_);
  const operation = { operation_id: uuid, input_hash: inputHash, created_by: "ADMIN1",
    status: "STARTED", employee_id: employee.employee_id,
    display_employee_id: employee.display_employee_id };
  const request = { operation_id: uuid, input_hash: inputHash, created_by: "ADMIN1",
    normalized_input: normalized };
  const result = context.classifyEmployeeRegistrationRecovery_(operation,
    snapshot.employeeRows, request, context.employeeRegistrationSha256HexAppsScript_,
    snapshot.formatDateKey);
  assert.equal(result.kind, "started_with_employee");
});

test("ID空欄、必須値欠落、異常型・不正日付、重複を除外せず拒否する", () => {
  for (const change of [
    { employee_id: "" }, { display_employee_id: "" }, { name: "" },
    { employee_id: 83 }, { display_employee_id: "P23" },
    { company_name: null }, { work_days_per_week: {} },
    { leave_management_target: "" }, { hire_date: new Date("invalid") },
    { hire_date: "2026-02-30" }
  ]) {
    const bad = { ...employee, ...change };
    const { context } = harness({ values: [headers, row(headers, employee), row(headers, bad)] });
    assert.throws(() => context.readEmployeeRegistrationEmployeesReadOnly_("legacy_read_only"),
      /ROW_INVALID|DATE_INVALID/);
  }
});

test("Spreadsheet・Sheet・rangeの取得失敗を空一覧にしない", () => {
  for (const spreadsheet of [null, {},
    { getSpreadsheetTimeZone: () => "Asia/Tokyo", getSheetByName: () => null },
    { getSpreadsheetTimeZone: () => "Asia/Tokyo", getSheetByName: () => ({
      getDataRange: () => ({ getValues: () => { throw new Error("read failed"); } })
    }) }]) {
    const { context } = harness({ spreadsheet });
    assert.throws(() => context.readEmployeeRegistrationEmployeesReadOnly_("legacy_read_only"),
      /UNAVAILABLE/);
  }
  const { context } = harness();
  context.SpreadsheetApp.openById = () => { throw new Error("open failed"); };
  assert.throws(() => context.readEmployeeRegistrationEmployeesReadOnly_("legacy_read_only"),
    /SPREADSHEET_UNAVAILABLE/);
});

test("operation ID列欠落は旧データモードのみ許可し、列ありでは空欄とUUIDを保持する", () => {
  const old = harness();
  assert.throws(() => old.context.readEmployeeRegistrationEmployeesReadOnly_("registration_ready"),
    /HEADERS_MISSING/);
  assert.throws(() => old.context.readEmployeeRegistrationEmployeesReadOnly_(), /SCHEMA_MODE_REQUIRED/);
  const names = [...headers, "registration_operation_id"];
  const values = [names, row(names, { ...employee, registration_operation_id: "" }),
    row(names, { ...employee, employee_id: "EMP0084", display_employee_id: "W0063",
      registration_operation_id: uuid })];
  const { context } = harness({ values });
  const snapshot = context.readEmployeeRegistrationEmployeesReadOnly_("registration_ready");
  assert.equal(snapshot.hasRegistrationOperationIdColumn, true);
  assert.equal(snapshot.employeeRows[0].registration_operation_id, "");
  assert.equal(snapshot.employeeRows[1].registration_operation_id, uuid);
  const bad = harness({ values: [names, row(names, { ...employee,
    registration_operation_id: "not-a-uuid" })] });
  assert.throws(() => bad.context.readEmployeeRegistrationEmployeesReadOnly_("registration_ready"),
    /ROW_INVALID/);
});

test("不完全な行配列は正常な空一覧として扱わない", () => {
  for (const malformed of [[], Array(headers.length), row(headers, employee).slice(1)]) {
    const { context } = harness({ values: [headers, malformed] });
    assert.throws(() => context.readEmployeeRegistrationEmployeesReadOnly_("legacy_read_only"),
      /ROW_INVALID/);
  }
});
