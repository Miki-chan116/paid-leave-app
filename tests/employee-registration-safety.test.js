const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");
const crypto = require("node:crypto");

const file = path.join(__dirname, "..", "employee-registration-safety.gs");
const context = { Date, JSON, Math, Number, Object, String };
vm.createContext(context);
vm.runInContext(fs.readFileSync(file, "utf8"), context, { filename: file });

const normal = {
  name: "例 一郎", display_name: "", name_kana: "れい いちろう",
  company_code: "MAIN", company_name: "", department: "",
  employment_type: "regular", employment_status: "active",
  hire_date: "2026-04-01", leave_date: "", work_days_per_week: "5",
  work_start_minute: "", work_end_minute: "", fiscal_start_month: "4",
  leave_management_target: "TRUE", is_driver: "FALSE", driver_type: "",
  default_vehicle_id: "", notes: ""
};
const uuid = "123e4567-e89b-42d3-a456-426614174000";
const sha256 = value => crypto.createHash("sha256").update(value, "utf8").digest("hex");
// Phase 2ではSpreadsheetのタイムゾーンで日付キーを生成する。
const formatDateKey = date => {
  const parts = new Intl.DateTimeFormat("en-US", { timeZone: "Asia/Tokyo",
    year: "numeric", month: "2-digit", day: "2-digit" }).formatToParts(date);
  const values = Object.fromEntries(parts.map(part => [part.type, part.value]));
  return `${values.year}-${values.month}-${values.day}`;
};
const normalize = changes => context.normalizeEmployeeRegistrationInput_({ ...normal, ...changes });
const validate = changes => context.validateEmployeeRegistrationInput_(normalize(changes));
const codes = changes => Array.from(validate(changes).errors, error => `${error.field}:${error.code}`);
const hash = input => context.employeeRegistrationInputHash_(input, sha256);
const next = (prefix, employees = [], operations = [], baseline = 0) =>
  context.nextEmployeeRegistrationId_(prefix, employees, operations, baseline);

function recoveryFixture() {
  const normalized = normalize();
  const operation = { operation_id: uuid, input_hash: hash(normalized),
    created_by: "ADMIN1", status: "STARTED",
    employee_id: "EMP0083", display_employee_id: "W0062" };
  const employee = { ...normalized, registration_operation_id: uuid,
    employee_id: "EMP0083", display_employee_id: "W0062" };
  const request = { operation_id: uuid, input_hash: operation.input_hash,
    created_by: "ADMIN1", normalized_input: normalized };
  return { operation, employee, request };
}
const classify = (operation, employees, request) =>
  context.classifyEmployeeRegistrationRecovery_(operation, employees, request, sha256, formatDateKey).kind;

test("正常入力を固定型に正規化し検証する", () => {
  const value = normalize({ name: " 例 一郎 ", company_code: " main ",
    hire_date: "2026/04/01", fiscal_start_month: "" });
  assert.equal(value.name, "例 一郎");
  assert.equal(value.company_code, "MAIN");
  assert.equal(value.hire_date, "2026-04-01");
  assert.equal(value.fiscal_start_month, 4);
  assert.equal(value.work_days_per_week, 5);
  assert.equal(value.leave_management_target, true);
  assert.equal(context.validateEmployeeRegistrationInput_(value).ok, true);
});

test("canonical payloadは固定順で保存入力を含み、生成値を受け付けない", () => {
  const value = normalize();
  const canonical = context.employeeRegistrationCanonicalPayload_(value);
  assert.equal(Object.keys(JSON.parse(canonical))[0], "name");
  for (const key of ["created_at", "updated_at", "employee_id", "display_employee_id",
    "display_order", "operation_id", "operator_id", "log_id"]) {
    assert.equal(Object.hasOwn(JSON.parse(canonical), key), false);
    assert.throws(() => normalize({ [key]: "x" }), /UNKNOWN_FIELD/);
  }
  assert.notEqual(hash(normalize({ notes: "変更" })), hash(value));
  assert.equal(hash(normalize({ name: " 例 一郎 " })), hash(value));
  assert.match(hash(value), /^v1:[a-f0-9]{64}$/);
});

test("hashには正規化済み入力とSHA-256アダプターが必要", () => {
  assert.throws(() => context.employeeRegistrationInputHash_(normalize()), /DIGEST_REQUIRED/);
  assert.throws(() => context.employeeRegistrationInputHash_(normalize(), () => "bad"), /DIGEST_INVALID/);
  assert.throws(() => hash(normalize({ employment_status: "retired" })), /INPUT_NOT_VALIDATED/);
  assert.throws(() => context.employeeRegistrationCanonicalPayload_({ name: "x" }), /CANONICAL_INPUT_INVALID/);
});

test("hash生成とrecoveryは未正規化・異常型の入力を拒否", () => {
  const f = recoveryFixture();
  const invalidInputs = [
    { ...f.request.normalized_input, name: "   " },
    { ...f.request.normalized_input, name: " 例 一郎 " },
    { ...f.request.normalized_input, notes: undefined },
    { ...f.request.normalized_input, notes: () => "" },
    { ...f.request.normalized_input, display_name: 123 },
    { ...f.request.normalized_input, company_code: "main" }
  ];
  invalidInputs.forEach(input => {
    assert.throws(() => hash(input), /INPUT_NOT_NORMALIZED/);
    assert.equal(classify(f.operation, [], { ...f.request, normalized_input: input }), "invalid_input");
  });
  assert.equal(classify(f.operation, [], f.request), "started_without_employee");
});

test("必須の氏名・会社・雇用・在職・真偽値", () => {
  for (const [change, expected] of [
    [{ name: "" }, "name:required"],
    [{ name_kana: "" }, "name_kana:required"],
    [{ company_code: "" }, "company_code:required"],
    [{ employment_type: "" }, "employment_type:required"],
    [{ employment_status: "" }, "employment_status:required"],
    [{ leave_management_target: "" }, "leave_management_target:invalid_boolean"],
    [{ is_driver: "maybe" }, "is_driver:invalid_boolean"]
  ]) assert.ok(codes(change).includes(expected), expected);
});

test("不正会社コード・雇用区分・retiredを拒否", () => {
  assert.ok(codes({ company_code: "OTHER" }).includes("company_code:invalid"));
  assert.ok(codes({ employment_type: "unknown" }).includes("employment_type:invalid"));
  assert.ok(codes({ employment_status: "retired", leave_date: "2026-09-01" })
    .includes("employment_status:retired_forbidden"));
});

test("有給対象には入社日と1～5日の整数が必要", () => {
  assert.ok(codes({ hire_date: "" }).includes("hire_date:required_for_paid_leave"));
  for (const workDays of ["", "0", "6", "2.5", "0x5", "9007199254740992"])
    assert.ok(codes({ work_days_per_week: workDays }).some(code => code.startsWith("work_days_per_week:")));
  assert.equal(validate({ leave_management_target: "FALSE", hire_date: "", work_days_per_week: "" }).ok, true);
});

test("不正日付、退職日、年度開始月の矛盾を拒否", () => {
  for (const date of ["2026-02-30", "2026-13-01", "not-a-date"])
    assert.ok(codes({ hire_date: date }).includes("hire_date:invalid_date"));
  assert.ok(codes({ leave_date: "2026-09-01" }).includes("leave_date:forbidden_for_new_employee"));
  assert.ok(codes({ fiscal_start_month: "6" }).includes("fiscal_start_month:company_mismatch"));
  assert.equal(validate({ company_code: "PARTNER", fiscal_start_month: "6" }).ok, true);
});

test("運転手項目の条件を検証", () => {
  assert.ok(codes({ is_driver: "TRUE" }).includes("driver_type:required_for_driver"));
  assert.ok(codes({ is_driver: "TRUE", driver_type: "不明" }).includes("driver_type:invalid"));
  assert.ok(codes({ driver_type: "専任運転手" }).includes("is_driver:driver_fields_without_driver"));
  assert.equal(validate({ is_driver: "TRUE", driver_type: "兼任運転手" }).ok, true);
});

test("勤務時刻の片側、範囲、順序、420分、PARTNERの個別設定を検証", () => {
  assert.ok(codes({ work_start_minute: "480" }).includes("work_start_minute:work_time_pair_required"));
  assert.ok(codes({ work_start_minute: "-1", work_end_minute: "1020" }).includes("work_start_minute:invalid"));
  assert.ok(codes({ work_start_minute: "1020", work_end_minute: "480" }).includes("work_end_minute:not_after_start"));
  assert.ok(codes({ work_start_minute: "480", work_end_minute: "1000" })
    .includes("work_end_minute:scheduled_minutes_mismatch"));
  assert.equal(validate({ work_start_minute: "480", work_end_minute: "1020" }).ok, true);
  assert.ok(codes({ company_code: "PARTNER", fiscal_start_month: "6",
    work_start_minute: "480", work_end_minute: "1020" })
    .includes("work_end_minute:unsupported_for_partner"));
});

test("不正UUIDとoperationなしを分類", () => {
  const f = recoveryFixture();
  assert.equal(classify(null, [], { ...f.request, operation_id: "bad" }), "invalid_operation_id");
  assert.equal(classify(null, [], f.request), "new");
  assert.equal(classify(null, [f.employee], f.request), "orphan_employee");
});

test("社員行一覧は空配列のみ空一覧とし、非配列は全状態で拒否", () => {
  const f = recoveryFixture();
  assert.equal(classify(null, [], f.request), "new");
  assert.equal(classify(f.operation, [], f.request), "started_without_employee");
  for (const rows of [null, undefined, {}, ""]) {
    assert.equal(classify(null, rows, f.request), "employee_rows_unavailable");
    for (const status of ["STARTED", "EMPLOYEE_SAVED", "COMPLETED"]) {
      assert.equal(classify({ ...f.operation, status }, rows, f.request),
        "employee_rows_unavailable");
    }
  }
});

test("社員行一覧はdense配列と必要な行IDを要求", () => {
  const f = recoveryFixture();
  const other = { registration_operation_id: "", employee_id: "EMP0084",
    display_employee_id: "W0063" };
  assert.equal(classify(null, [], f.request), "new");
  assert.equal(classify(f.operation, [f.employee], f.request), "started_with_employee");
  assert.equal(classify(f.operation, [other, f.employee], f.request), "started_with_employee");
  const sparse = Array(1);
  const middleHole = [other];
  middleHole.length = 3;
  middleHole[2] = f.employee;
  for (const rows of [sparse, middleHole]) {
    assert.equal(classify(null, rows, f.request), "employee_rows_unavailable");
    assert.equal(classify(f.operation, rows, f.request), "employee_rows_unavailable");
  }
  for (const bad of [{}, 42, null, undefined,
    { employee_id: "EMP0084", display_employee_id: "W0063" },
    { registration_operation_id: "", employee_id: 84, display_employee_id: "W0063" }]) {
    assert.equal(classify(null, [bad], f.request), "employee_rows_invalid");
    assert.equal(classify(f.operation, [other, bad], f.request), "employee_rows_invalid");
  }
  assert.equal(classify(f.operation, [other,
    { ...f.employee, registration_operation_id: "", employee_id: other.employee_id }], f.request),
  "employee_duplicate");
  assert.equal(classify(f.operation, [{ ...f.employee, notes: undefined }], f.request),
    "employee_mismatch");
});

test("同一UUID・同一hashと異なるhash、created_by不一致", () => {
  const f = recoveryFixture();
  assert.equal(classify(f.operation, [], f.request), "started_without_employee");
  assert.equal(classify(f.operation, [], { ...f.request, input_hash: "v1:" + "0".repeat(64) }), "input_hash_mismatch");
  const changedInput = normalize({ name: "別人" });
  assert.equal(classify(f.operation, [], { ...f.request, input_hash: hash(changedInput),
    normalized_input: changedInput }), "hash_conflict");
  assert.equal(classify(f.operation, [], { ...f.request, created_by: "ADMIN2" }), "admin_conflict");
});

test("古いhashのまま氏名を変更した要求は社員行なしでも拒否", () => {
  const f = recoveryFixture();
  const request = { ...f.request, normalized_input: normalize({ name: "別人" }) };
  assert.equal(classify(f.operation, [], request), "input_hash_mismatch");
  assert.equal(classify(null, [], request), "input_hash_mismatch");
});

test("古いhashのまま氏名を変更した要求は社員保存済みでも拒否", () => {
  const f = recoveryFixture();
  const request = { ...f.request, normalized_input: normalize({ name: "別人" }) };
  assert.equal(classify({ ...f.operation, status: "EMPLOYEE_SAVED" }, [f.employee], request),
    "input_hash_mismatch");
});

test("古いhashのまま氏名を変更した要求はCOMPLETEDでも拒否", () => {
  const f = recoveryFixture();
  const request = { ...f.request, normalized_input: normalize({ name: "別人" }) };
  assert.equal(classify({ ...f.operation, status: "COMPLETED" }, [f.employee], request),
    "input_hash_mismatch");
});

test("recoveryはdigestなし・不正入力を受理しない", () => {
  const f = recoveryFixture();
  assert.equal(context.classifyEmployeeRegistrationRecovery_(f.operation, [], f.request).kind,
    "digest_unavailable");
  assert.equal(classify(f.operation, [], { ...f.request,
    normalized_input: { ...f.request.normalized_input, name: "" } }), "invalid_input");
});

test("STARTEDの社員なし・整合・重複・不一致", () => {
  const f = recoveryFixture();
  assert.equal(classify(f.operation, [], f.request), "started_without_employee");
  assert.equal(classify(f.operation, [f.employee], f.request), "started_with_employee");
  assert.equal(classify(f.operation, [f.employee, { ...f.employee }], f.request), "employee_duplicate");
  assert.equal(classify(f.operation, [{ ...f.employee, employee_id: "EMP9999" }], f.request), "employee_mismatch");
  assert.equal(classify(f.operation, [{ ...f.employee, name: "別人" }], f.request), "employee_mismatch");
  assert.equal(classify(f.operation, [{ ...f.employee, registration_operation_id: "別操作" }], f.request),
    "reserved_id_conflict");
});

test("EMPLOYEE_SAVED・COMPLETEDと社員欠落を分類", () => {
  const f = recoveryFixture();
  assert.equal(classify({ ...f.operation, status: "EMPLOYEE_SAVED" }, [f.employee], f.request), "employee_saved");
  assert.equal(classify({ ...f.operation, status: "COMPLETED" }, [f.employee], f.request), "completed");
  assert.equal(classify({ ...f.operation, status: "COMPLETED" },
    [{ ...f.employee, name: "登録後に編集" }], f.request), "completed");
  assert.equal(classify({ ...f.operation, status: "COMPLETED" }, [], f.request), "employee_missing");
});

test("保存行のDateはSpreadsheetのタイムゾーンで同じ日付キーと照合", () => {
  const f = recoveryFixture();
  const saved = { ...f.operation, status: "EMPLOYEE_SAVED" };
  const sameDay = new Date("2026-03-31T15:00:00.000Z");
  assert.equal(sameDay.toISOString().slice(0, 10), "2026-03-31");
  assert.equal(formatDateKey(sameDay), "2026-04-01");
  assert.equal(classify(saved, [{ ...f.employee, hire_date: sameDay }], f.request), "employee_saved");
  assert.equal(classify(saved, [{ ...f.employee, hire_date: "2026-04-01" }], f.request), "employee_saved");
});

test("異なる日付のDateと不正なDateは保存行不一致", () => {
  const f = recoveryFixture();
  const saved = { ...f.operation, status: "EMPLOYEE_SAVED" };
  assert.equal(classify(saved, [{ ...f.employee,
    hire_date: new Date("2026-04-01T15:00:00.000Z") }], f.request), "employee_mismatch");
  assert.equal(classify(saved, [{ ...f.employee, hire_date: new Date(NaN) }], f.request),
    "employee_mismatch");
  assert.equal(context.classifyEmployeeRegistrationRecovery_(saved,
    [{ ...f.employee, hire_date: new Date("2026-03-31T15:00:00.000Z") }],
    f.request, sha256).kind, "employee_mismatch");
});

test("baselineだけのEMP/W/P採番", () => {
  assert.equal(next("EMP", [], [], 82), "EMP0083");
  assert.equal(next("W", [], [], 61), "W0062");
  assert.equal(next("P", [], [], 22), "P0023");
  assert.equal(JSON.stringify(context.calculateNextEmployeeRegistrationIds_("PARTNER", [], [],
    { EMP: 82, W: 61, P: 22 })), JSON.stringify({ employee_id: "EMP0083", display_employee_id: "P0023" }));
});

test("現存最大値・予約最大値・欠番を尊重", () => {
  assert.equal(next("EMP", ["EMP0090", "W0062"], ["EMP0088"], 82), "EMP0091");
  assert.equal(next("W", ["W0065"], ["W0070"], 61), "W0071");
  assert.equal(next("P", ["P0025"], ["P0030"], 22), "P0031");
  assert.equal(next("EMP", ["EMP0082"], ["EMP0085"], 82), "EMP0086");
});

test("異常ID・不正baseline・安全整数上限を拒否", () => {
  for (const id of ["emp0083", " EMP0083", "EMP0x10", "EMP83", "EMP00000", "BAD0001"])
    assert.throws(() => next("EMP", [id], [], 82), /ID_INVALID/);
  assert.throws(() => next("EMP", ["EMP9007199254740992"], [], 82), /ID_UNSAFE/);
  assert.throws(() => next("EMP", ["EMP9007199254740991"], [], 82), /ID_EXHAUSTED/);
  assert.throws(() => next("EMP", [], [], -1), /BASELINE_INVALID/);
});

test("採番用の両ID一覧は空配列を許可し、片方でも非配列なら拒否", () => {
  assert.equal(context.nextEmployeeRegistrationId_("EMP", [], [], 82), "EMP0083");
  for (const [employees, operations] of [
    [null, []], [undefined, []], [{}, []], ["", []],
    [[], null], [[], undefined], [[], {}], [[], ""]
  ]) {
    assert.throws(() => context.nextEmployeeRegistrationId_("EMP", employees, operations, 82),
      /IDS_UNAVAILABLE/);
    assert.throws(() => context.calculateNextEmployeeRegistrationIds_(
      "MAIN", employees, operations), /IDS_UNAVAILABLE/);
  }
});

test("EMP/W/Pの採番はdenseな文字列ID一覧だけを受け付ける", () => {
  for (const [prefix, baseline, existing] of [
    ["EMP", 82, "EMP0083"], ["W", 61, "W0062"], ["P", 22, "P0023"]
  ]) {
    assert.equal(next(prefix, [], [], baseline), prefix + String(baseline + 1).padStart(4, "0"));
    assert.equal(next(prefix, [existing], [existing], baseline),
      prefix + String(baseline + 2).padStart(4, "0"));
    const middleHole = [existing];
    middleHole.length = 3;
    middleHole[2] = existing;
    for (const sparse of [Array(1), middleHole]) {
      assert.throws(() => context.nextEmployeeRegistrationId_(prefix, sparse, [], baseline),
        /IDS_UNAVAILABLE/);
      assert.throws(() => context.nextEmployeeRegistrationId_(prefix, [], sparse, baseline),
        /IDS_UNAVAILABLE/);
    }
    for (const bad of [null, undefined, 42, {}, { toString: () => existing }, "BAD0001"]) {
      assert.throws(() => context.nextEmployeeRegistrationId_(prefix, [bad], [], baseline),
        /ID_INVALID/);
      assert.throws(() => context.nextEmployeeRegistrationId_(prefix, [], [existing, bad], baseline),
        /ID_INVALID/);
    }
  }
});
