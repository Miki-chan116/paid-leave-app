const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const auditPath = path.join(__dirname, "..", "employee-registration-audit.gs");
const source = fs.readFileSync(auditPath, "utf8");
const context = { Date, Map, Number, Object, Set, String, isNaN };
vm.createContext(context);
vm.runInContext(source, context, { filename: auditPath });

const employeeHeaders = [
  "employee_id", "display_employee_id", "name", "name_kana", "company_code",
  "employment_status", "hire_date", "leave_date", "work_days_per_week",
  "leave_management_target", "fiscal_start_month", "initial_grant_check_target"
];
const standard = {
  employee_id: "EMP0001", display_employee_id: "W0001", name: "例 一郎",
  name_kana: "れい いちろう", company_code: "MAIN", employment_status: "active",
  hire_date: "2026-04-01", leave_date: "", work_days_per_week: 5,
  leave_management_target: true, fiscal_start_month: 4,
  initial_grant_check_target: true
};
const row = (headers, values) => headers.map(header => values[header] ?? "");
function fixture(employees, references = {}) {
  return {
    employees: { headers: employeeHeaders, rows: employees.map(item => row(employeeHeaders, item)) },
    leave_requests: { headers: ["employee_id"], rows: (references.leave_requests || []).map(id => [id]) },
    paid_leave_grants: { headers: ["employee_id"], rows: (references.paid_leave_grants || []).map(id => [id]) },
    leave_retirement_records: { headers: ["employee_id", "leave_date", "record_status"],
      rows: (references.leave_retirement_records || []).map(item =>
        [item.employee_id, item.leave_date ?? "", item.record_status ?? "completed"]) },
    time_leave_segments: { headers: ["employee_id"], rows: (references.time_leave_segments || []).map(id => [id]) },
    usage_log: { headers: ["request_id", "action_type"], rows: [] }
  };
}
const audit = (employees, references) =>
  context.auditEmployeeRegistrationSnapshot_(fixture(employees, references));
const rules = report => [...report.errors, ...report.warnings, ...report.reviews].map(issue => issue.rule);

test("正常社員と採番統計", () => {
  const report = audit([standard]);
  assert.equal(report.summary.counts.ERROR, 0);
  assert.equal(report.summary.counts.WARN, 0);
  assert.equal(report.summary.counts.REVIEW, 0);
  assert.equal(report.idStatistics.strictDecimalMax.EMP, "1");
  assert.equal(report.idStatistics.strictDecimalMax.W, "1");
  assert.equal(report.idStatistics.currentAllocatorInterpretedMax.EMP.value, "1");
});

test("社員IDと表示IDの重複", () => {
  const report = audit([standard, { ...standard }]);
  assert.ok(rules(report).includes("employee_id_duplicate"));
  assert.ok(rules(report).includes("display_employee_id_duplicate"));
});

test("紛らわしいIDと安全整数超過", () => {
  const report = audit([standard,
    { ...standard, employee_id: "emp0001", display_employee_id: "w0001" },
    { ...standard, employee_id: "EMP1", display_employee_id: "W1" },
    { ...standard, employee_id: "EMP0x10", display_employee_id: "W0x10" },
    { ...standard, employee_id: "EMP9007199254740992", display_employee_id: "P9007199254740992" }]);
  assert.ok(rules(report).includes("employee_id_case_collision"));
  assert.ok(rules(report).includes("display_employee_id_case_collision"));
  assert.ok(rules(report).includes("employee_id_unsafe_number"));
  assert.ok(rules(report).includes("display_employee_id_unsafe_number"));
  assert.ok(rules(report).includes("employee_id_same_serial"));
  assert.ok(rules(report).includes("display_employee_id_same_serial"));
  assert.ok(rules(report).includes("employee_id_number_interpretation_risk"));
  assert.equal(report.idStatistics.strictDecimalMax.EMP, "9007199254740992");
});

test("不正enumと有給対象の欠損・境界値", () => {
  const cases = [
    [{ company_code: "UNKNOWN", employment_status: "unknown" }, ["company_code_invalid", "employment_status_invalid"]],
    [{ hire_date: "" }, ["paid_leave_hire_date_missing"]],
    [{ hire_date: "2026-02-30" }, ["hire_date_invalid"]],
    [{ work_days_per_week: "" }, ["paid_leave_work_days_missing"]],
    [{ work_days_per_week: 0 }, ["paid_leave_work_days_invalid"]],
    [{ work_days_per_week: 0.5 }, ["paid_leave_work_days_invalid"]],
    [{ work_days_per_week: 6 }, ["paid_leave_work_days_invalid"]],
    [{ work_days_per_week: "0x5" }, ["paid_leave_work_days_invalid"]],
    [{ fiscal_start_month: 6 }, ["paid_leave_company_month_mismatch"]]
  ];
  cases.forEach(([change, expected]) => {
    const found = rules(audit([{ ...standard, ...change }]));
    expected.forEach(rule => assert.ok(found.includes(rule), `${rule}: ${found.join(", ")}`));
  });
});

test("退職日とstatus、旧データの初回付与フラグ", () => {
  const report = audit([
    { ...standard, leave_date: "2026-05-01" },
    { ...standard, employee_id: "EMP0002", display_employee_id: "W0002",
      employment_status: "retired", leave_date: "" },
    { ...standard, employee_id: "EMP0003", display_employee_id: "W0003",
      hire_date: "2026-05-01", leave_date: "2026-04-30",
      employment_status: "retired", initial_grant_check_target: "" }
  ]);
  const found = rules(report);
  assert.ok(found.includes("active_with_leave_date"));
  assert.ok(found.includes("retired_without_leave_date"));
  assert.ok(found.includes("leave_before_hire"));
  assert.ok(report.reviews.some(issue => issue.rule === "initial_grant_target_blank"));
  assert.ok(report.reviews.some(issue => issue.rule === "retired_without_completed_record"));
  assert.ok(!report.errors.some(issue => issue.rule === "retired_without_completed_record"));
});

test("存在しない社員への参照と退職記録の逆方向", () => {
  const report = audit([standard], {
    leave_requests: ["EMP9999"], paid_leave_grants: ["EMP9998"],
    time_leave_segments: ["EMP9997"],
    leave_retirement_records: [{ employee_id: "EMP0001", leave_date: "2026-05-01" }]
  });
  assert.equal(report.errors.filter(issue => issue.rule === "orphan_employee_reference").length, 3);
  assert.ok(rules(report).includes("completed_record_for_nonretired"));
});

test("監査入口とフラグ取得は読取APIだけを使う", () => {
  const forbidden = /\.(?:setValue|setValues|appendRow|insertRows?|deleteRows?|insertColumns?|deleteColumns?|clear(?:Content|Contents|Formats)?|sort|put|remove|setProperty|deleteProperty|fetch)\s*\(/;
  assert.equal(forbidden.test(source), false);
  assert.equal(/requireAdminSession_\s*\(/.test(source), false);
  const adminHeaders = ["admin_id", "admin_name", "pin", "is_active"];
  const data = fixture([standard]);
  const sheetData = { ...data,
    admin_users: { headers: adminHeaders, rows: [["A1", "管理者", "fixture", true]] } };
  const readCalls = [];
  const spreadsheet = {
    getSheetByName(name) {
      readCalls.push(`sheet:${name}`);
      if (!sheetData[name]) return null;
      return { getDataRange: () => ({ getValues: () =>
        [sheetData[name].headers, ...sheetData[name].rows] }) };
    },
    getSpreadsheetTimeZone: () => "Asia/Tokyo"
  };
  context.SS_ID = "fixture-only";
  context.SpreadsheetApp = { openById: () => spreadsheet };
  context.CacheService = { getScriptCache: () => ({ get: () =>
    JSON.stringify({ admin_id: "A1", expires_at: "2999-01-01T00:00:00Z" }) }) };
  context.getAdminSessionCacheKey_ = token => `fixture:${token}`;
  context.PropertiesService = { getScriptProperties: () => ({ getProperty: () => "false" }) };
  context.Utilities = { formatDate: date => date.toISOString().slice(0, 10) };
  const result = context.auditEmployeeRegistrationReadOnly("token");
  assert.equal(result.summary.counts.ERROR, 0);
  assert.equal(context.getEmployeeRegistrationAuditFlagsReadOnly("token").useSupabaseReads, false);
  assert.ok(readCalls.includes("sheet:employees"));
});

function authFixture(session, adminRows, adminHeaders = ["admin_id", "admin_name", "pin", "is_active"]) {
  context.CacheService = { getScriptCache: () => ({ get: () => session == null ? null : JSON.stringify(session) }) };
  context.getAdminSessionCacheKey_ = token => `fixture:${token}`;
  context.SpreadsheetApp = { openById: () => ({
    getSheetByName: name => name === "admin_users" ? {
      getDataRange: () => ({ getValues: () => [adminHeaders, ...adminRows] })
    } : null
  }) };
  context.SS_ID = "fixture-only";
}

test("監査認証はトークン欠落・期限切れ・不正日時・非アクティブを拒否", () => {
  const active = [["A1", "管理者", "fixture", true]];
  authFixture(null, active);
  assert.throws(() => context.auditEmployeeRegistrationReadOnly(""), /ADMIN_SESSION_REQUIRED/);
  authFixture({ admin_id: "A1", expires_at: "2000-01-01T00:00:00Z" }, active);
  assert.throws(() => context.auditEmployeeRegistrationReadOnly("token"), /ADMIN_SESSION_EXPIRED/);
  authFixture({ admin_id: "A1", expires_at: "not-a-date" }, active);
  assert.throws(() => context.auditEmployeeRegistrationReadOnly("token"), /ADMIN_SESSION_EXPIRED/);
  authFixture({ admin_id: "A1", expires_at: "2999-01-01T00:00:00Z" },
    [["A1", "管理者", "fixture", false]]);
  assert.throws(() => context.auditEmployeeRegistrationReadOnly("token"), /ADMIN_SESSION_INVALID/);
});

test("同一admin_idは通常認証と同じ先頭行を判定し、必要ヘッダーを要求", () => {
  const session = { admin_id: "A1", expires_at: "2999-01-01T00:00:00Z" };
  authFixture(session, [["A1", "先頭", "fixture", false], ["A1", "後続", "fixture", true]]);
  assert.throws(() => context.auditEmployeeRegistrationReadOnly("token"), /ADMIN_SESSION_INVALID/);
  authFixture(session, [["A1", "fixture", true]], ["admin_id", "pin", "is_active"]);
  assert.throws(() => context.auditEmployeeRegistrationReadOnly("token"), /ADMIN_SESSION_INVALID/);
});

test("退職完了の有給対象フラグ、未完了・空欄状態、記録日付、重複を検出", () => {
  const retired = { ...standard, employment_status: "retired", leave_date: "2026-05-01" };
  const report = audit([retired], { leave_retirement_records: [
    { employee_id: "EMP0001", leave_date: "2026-05-01", record_status: "completed" },
    { employee_id: "EMP0001", leave_date: "2026-05-02", record_status: "completed" },
    { employee_id: "EMP0001", leave_date: "bad", record_status: "pending" },
    { employee_id: "EMP0001", leave_date: "", record_status: "" }
  ] });
  const found = rules(report);
  ["completed_retirement_leave_target_not_false", "multiple_completed_retirement_records",
    "retirement_record_noncompleted", "retirement_record_leave_date_invalid",
    "retirement_record_leave_date_missing"].forEach(rule => assert.ok(found.includes(rule), rule));
  assert.equal(report.warnings.filter(issue => issue.rule === "retirement_record_noncompleted").length, 2);
  assert.ok(!found.includes("retired_without_completed_record"));
});

test("退職記録の日付不一致と比較不能な不正日付を検出", () => {
  const retired = { ...standard, employment_status: "retired", leave_date: "2026-05-01",
    leave_management_target: false };
  const mismatch = audit([retired], { leave_retirement_records: [
    { employee_id: "EMP0001", leave_date: "2026-05-02" }
  ] });
  assert.ok(rules(mismatch).includes("retirement_date_mismatch"));
  const invalid = audit([retired], { leave_retirement_records: [
    { employee_id: "EMP0001", leave_date: "2026-02-30" }
  ] });
  assert.ok(rules(invalid).includes("retirement_record_leave_date_invalid"));
  const stringFalse = audit([{ ...retired, leave_management_target: "FALSE" }], {
    leave_retirement_records: [{ employee_id: "EMP0001", leave_date: "2026-05-01" }]
  });
  assert.ok(rules(stringFalse).includes("completed_retirement_leave_target_not_false"));
  const pendingOnly = audit([retired], { leave_retirement_records: [
    { employee_id: "EMP0001", leave_date: "2026-05-01", record_status: "pending" }
  ] });
  assert.ok(rules(pendingOnly).includes("retirement_record_noncompleted"));
  assert.ok(!rules(pendingOnly).includes("retired_without_completed_record"));
});

test("旧employment_statusは後続互換性を警告", () => {
  for (const status of ["在職", "休職", "退職"]) {
    const report = audit([{ ...standard, employment_status: status }]);
    assert.ok(report.warnings.some(issue => issue.rule === "employment_status_legacy"));
  }
});

test("現行採番解釈と厳密な十進統計を区別", () => {
  for (const [id, expected] of [["EMP1e3", "1000"], ["EMPInfinity", "Infinity"],
    ["EMP0x10", "16"], ["EMP1.5", "1.5"]]) {
    const report = audit([{ ...standard, employee_id: id }]);
    assert.equal(report.idStatistics.strictDecimalMax.EMP, null);
    assert.equal(report.idStatistics.currentAllocatorInterpretedMax.EMP.value, expected);
    assert.ok(rules(report).includes("current_allocator_noncanonical_max"));
    if (id === "EMPInfinity" || id === "EMP1.5")
      assert.ok(rules(report).includes("current_allocator_unsafe_max"));
  }
  const boundary = audit([{ ...standard, employee_id: "EMP9007199254740991",
    display_employee_id: "P9007199254740991" }]);
  assert.equal(boundary.idStatistics.strictDecimalMax.EMP, "9007199254740991");
  assert.equal(boundary.idStatistics.strictDecimalMax.P, "9007199254740991");
  assert.ok(!rules(boundary).includes("employee_id_unsafe_number"));
  const over = audit([{ ...standard, employee_id: "EMP9007199254740992",
    display_employee_id: "P9007199254740992" }]);
  assert.ok(rules(over).includes("employee_id_unsafe_number"));
  assert.ok(rules(over).includes("display_employee_id_unsafe_number"));
  assert.equal(over.idStatistics.currentAllocatorInterpretedMax.P.safeInteger, false);
});

test("Date object、Invalid Date、うるう日、タイムゾーン境界", () => {
  const local = fixture([{ ...standard, hire_date: new Date("2026-03-31T15:00:00Z") }]);
  const converted = context.auditEmployeeRegistrationSnapshot_(local, {
    dateToKey: () => "2026-04-01"
  });
  assert.ok(!rules(converted).includes("hire_date_invalid"));
  assert.ok(!rules(audit([{ ...standard, hire_date: "2024-02-29" }])).includes("hire_date_invalid"));
  assert.ok(rules(audit([{ ...standard, hire_date: "2025-02-29" }])).includes("hire_date_invalid"));
  assert.ok(rules(audit([{ ...standard, hire_date: new Date(NaN) }])).includes("hire_date_invalid"));
});

test("ヘッダー欠落と重複はセル文字列を結果へ露出しない", () => {
  const data = fixture([standard]);
  data.employees.headers = employeeHeaders.filter(header => header !== "name_kana");
  data.employees.rows = [row(data.employees.headers, standard)];
  data.employees.headers.push("個人情報を含む任意のヘッダー");
  data.employees.headers.push("個人情報を含む任意のヘッダー");
  const report = context.auditEmployeeRegistrationSnapshot_(data);
  assert.ok(rules(report).includes("missing_header"));
  assert.ok(rules(report).includes("duplicate_header"));
  assert.ok(!JSON.stringify(report).includes("個人情報を含む任意のヘッダー"));
});

test("詳細上限を超えても総検出件数を保持", () => {
  const employees = Array.from({ length: 250 }, () => ({ ...standard, employee_id: "EMP0001" }));
  const report = audit(employees);
  assert.equal(report.summary.counts.ERROR, 498);
  assert.equal(report.errors.length, 200);
  assert.equal(report.summary.returnedCounts.ERROR, 200);
  assert.equal(report.summary.truncated, true);
});

test("サイズ確認は7シートの行列数だけを返し、セル内容・書込みAPIに触れない", () => {
  const names = ["admin_users", "employees", "leave_requests", "paid_leave_grants",
    "leave_retirement_records", "time_leave_segments", "usage_log"];
  const calls = [];
  const forbidden = () => { throw new Error("unexpected cell read or write"); };
  context.SS_ID = "fixture-only";
  context.getAdminSessionCacheKey_ = token => `fixture:${token}`;
  context.CacheService = { getScriptCache: () => ({
    get: () => JSON.stringify({ admin_id: "A1", expires_at: "2999-01-01T00:00:00Z" }),
    put: forbidden, remove: forbidden
  }) };
  context.SpreadsheetApp = { openById: () => ({
    getSheetByName(name) {
      calls.push(`sheet:${name}`);
      if (name === "admin_users") return {
        getDataRange: () => ({ getValues: () => [
          ["admin_id", "admin_name", "pin", "is_active"],
          ["A1", "非公開の氏名", "非公開のPIN", true]
        ] }),
        getLastRow: () => { calls.push("rows:admin_users"); return 2; },
        getLastColumn: () => { calls.push("columns:admin_users"); return 4; },
        setValue: forbidden
      };
      return {
        getLastRow: () => { calls.push(`rows:${name}`); return 60; },
        getLastColumn: () => { calls.push(`columns:${name}`); return 20; },
        getDataRange: forbidden, setValue: forbidden, appendRow: forbidden
      };
    }
  }) };
  const result = context.getEmployeeRegistrationAuditSheetSizesReadOnly("token");
  assert.deepEqual(Array.from(result.sheets, item => item.name), names);
  assert.deepEqual(Array.from(result.sheets, item => [item.exists, item.lastRow, item.lastColumn]),
    [[true, 2, 4], ...names.slice(1).map(() => [true, 60, 20])]);
  assert.equal(JSON.stringify(result).includes("非公開"), false);
  names.forEach(name => {
    assert.ok(calls.includes(`rows:${name}`));
    assert.ok(calls.includes(`columns:${name}`));
  });
});

test("サイズ確認は欠落シートを作成せず exists=false とする", () => {
  const missing = "time_leave_segments";
  const forbidden = () => { throw new Error("unexpected write"); };
  context.SpreadsheetApp = { openById: () => ({
    insertSheet: forbidden,
    getSheetByName(name) {
      if (name === missing) return null;
      if (name === "admin_users") return {
        getDataRange: () => ({ getValues: () => [
          ["admin_id", "admin_name", "pin", "is_active"], ["A1", "管理者", "fixture", true]
        ] }),
        getLastRow: () => 2, getLastColumn: () => 4
      };
      return { getLastRow: () => 1, getLastColumn: () => 3 };
    }
  }) };
  const result = context.getEmployeeRegistrationAuditSheetSizesReadOnly("token");
  const item = result.sheets.find(sheet => sheet.name === missing);
  assert.equal(item.exists, false);
  assert.equal(item.lastRow, null);
  assert.equal(item.lastColumn, null);
});

test("サイズ確認は無効な管理者セッションを拒否", () => {
  authFixture(null, [["A1", "管理者", "fixture", true]]);
  assert.throws(() => context.getEmployeeRegistrationAuditSheetSizesReadOnly("token"),
    /ADMIN_SESSION_EXPIRED/);
});

test("認証に必要なadmin_usersが欠落した場合は安全側に拒否", () => {
  authFixture({ admin_id: "A1", expires_at: "2999-01-01T00:00:00Z" }, []);
  context.SpreadsheetApp = { openById: () => ({
    getSheetByName: () => null,
    insertSheet: () => { throw new Error("unexpected write"); }
  }) };
  assert.throws(() => context.getEmployeeRegistrationAuditSheetSizesReadOnly("token"),
    /ADMIN_SESSION_INVALID/);
});
