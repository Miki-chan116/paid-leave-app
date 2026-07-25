/* =========================
   Manual test debug utilities
   debug.gs から動作を変えずに移動
========================= */

function testDebugFifoUseRows() {
  Logger.log(JSON.stringify(
    debugFifoApprovedLeaveUseRows("EMP0046", 2026, "2026-05-23"),
    null,
    2
  ));

  Logger.log(JSON.stringify(
    debugFifoApprovedLeaveUseRows("EMP0049", 2026, "2026-05-23"),
    null,
    2
  ));
}

function testCompareFifoDiffOnly() {
  const result = compareFifoBalanceDifferencesOnly(2026, "2026-05-23");
  Logger.log(JSON.stringify(result, null, 2));
}

function testDebugYearEndFinalizedBalance() {
  const result = debugYearEndFinalizedBalance("TEST-FIFO-001", 2026);

  Logger.log(JSON.stringify(result, null, 2));
}

/* =========================
   初回有給付与の正式規則テスト
   シート・付与履歴を書き換えない純粋関数テスト
========================= */
function testCalculateInitialPaidLeaveGrantEligibility() {
  const cases = [
    {
      name: "MAIN: 6か月前の4月1日は10日付与",
      employee: { employee_id: "T-01", company_code: "MAIN", hire_date: "2026-01-15", work_days_per_week: 5 },
      as_of_date: "2026-03-01",
      expected: { next_grant_date: "2026-04-01", expected_grant_days: 10, grant_reason: "INITIAL_COMPANY_BASIS", status: "UPCOMING" }
    },
    {
      name: "PARTNER: 6か月前の6月1日は10日付与",
      employee: { employee_id: "T-02", company_code: "PARTNER", hire_date: "2026-02-10", work_days_per_week: 5 },
      as_of_date: "2026-05-01",
      expected: { next_grant_date: "2026-06-01", expected_grant_days: 10, grant_reason: "INITIAL_COMPANY_BASIS", status: "UPCOMING" }
    },
    {
      name: "週4日でも基準日前倒しは10日",
      employee: { employee_id: "T-03", company_code: "MAIN", hire_date: "2026-01-15", work_days_per_week: 4 },
      as_of_date: "2026-03-01",
      expected: { expected_grant_days: 10, grant_reason: "INITIAL_COMPANY_BASIS" }
    },
    {
      name: "週3日でも基準日前倒しは10日",
      employee: { employee_id: "T-04", company_code: "PARTNER", hire_date: "2026-02-10", work_days_per_week: 3 },
      as_of_date: "2026-05-01",
      expected: { expected_grant_days: 10, grant_reason: "INITIAL_COMPANY_BASIS" }
    },
    {
      name: "前倒しなしの週4日は6か月日に7日",
      employee: { employee_id: "T-05", company_code: "MAIN", hire_date: "2026-05-15", work_days_per_week: 4 },
      as_of_date: "2026-10-01",
      expected: { next_grant_date: "2026-11-15", expected_grant_days: 7, grant_reason: "INITIAL_SIX_MONTHS" }
    },
    {
      name: "前倒しなしの週3日は6か月日に5日",
      employee: { employee_id: "T-06", company_code: "PARTNER", hire_date: "2026-07-15", work_days_per_week: 3 },
      as_of_date: "2026-12-01",
      expected: { next_grant_date: "2027-01-15", expected_grant_days: 5, grant_reason: "INITIAL_SIX_MONTHS" }
    },
    {
      name: "基準日と6か月日が同日は通常6か月付与",
      employee: { employee_id: "T-07", company_code: "MAIN", hire_date: "2025-10-01", work_days_per_week: 4 },
      as_of_date: "2026-03-01",
      expected: { next_grant_date: "2026-04-01", expected_grant_days: 7, grant_reason: "INITIAL_SIX_MONTHS" }
    },
    {
      name: "前倒し付与済みは6か月日後も再候補化しない",
      employee: { employee_id: "T-08", company_code: "MAIN", hire_date: "2026-01-15", work_days_per_week: 5 },
      grants: [{ employee_id: "T-08", grant_type: "six_month", grant_date: "2026-04-01", grant_days: 10, notes: "会社基準日による初回付与" }],
      as_of_date: "2026-08-01",
      expected: { status: "PROCESSED", existing_grant_found: true, processed_grant_type: "six_month" }
    },
    {
      name: "前倒し付与済みはOVERDUEにしない",
      employee: { employee_id: "T-09", company_code: "PARTNER", hire_date: "2026-02-10", work_days_per_week: 5 },
      grants: [{ employee_id: "T-09", grant_type: "six_month", grant_date: "2026-06-01", grant_days: 10, notes: "会社基準日による初回付与" }],
      as_of_date: "2026-09-01",
      expected: { status: "PROCESSED" }
    },
    {
      name: "入社日と会社基準日が同日は警告付き10日",
      employee: { employee_id: "T-10", company_code: "MAIN", hire_date: "2026-04-01", work_days_per_week: 1 },
      as_of_date: "2026-04-01",
      expected: { next_grant_date: "2026-04-01", expected_grant_days: 10, grant_reason: "INITIAL_COMPANY_BASIS", is_provisional: true, status: "DUE_TODAY", warning_codes: ["HIRE_DATE_EQUALS_COMPANY_BASIS_DATE"] }
    },
    {
      name: "initial形式の履歴も処理済みと認識する",
      employee: { employee_id: "T-11", company_code: "MAIN", hire_date: "2026-01-15", work_days_per_week: 5 },
      grants: [{ employee_id: "T-11", grant_type: "initial", grant_date: "2026-04-01", grant_days: 10, notes: "旧形式" }],
      as_of_date: "2026-08-01",
      expected: { status: "PROCESSED", processed_grant_type: "initial", existing_grant_found: true }
    },
    {
      name: "他社員の前倒し履歴では処理済みにしない",
      employee: { employee_id: "T-12", company_code: "MAIN", hire_date: "2026-01-15", work_days_per_week: 5 },
      grants: [{ employee_id: "OTHER", grant_type: "six_month", grant_date: "2026-04-01", grant_days: 10, notes: "別社員" }],
      as_of_date: "2026-03-01",
      expected: { status: "UPCOMING", existing_grant_found: false, grant_reason: "INITIAL_COMPANY_BASIS" }
    }
  ];

  const results = cases.map(testCase => {
    const actual = calculateInitialPaidLeaveGrantEligibility_(
      testCase.employee,
      testCase.grants || [],
      testCase.as_of_date
    );
    const failures = [];

    Object.keys(testCase.expected).forEach(key => {
      const expected = testCase.expected[key];
      const actualValue = actual[key];
      if (JSON.stringify(actualValue) !== JSON.stringify(expected)) {
        failures.push({ key: key, expected: expected, actual: actualValue });
      }
    });

    return {
      name: testCase.name,
      ok: failures.length === 0,
      failures: failures,
      actual: actual
    };
  });

  const failed = results.filter(result => !result.ok);
  if (failed.length > 0) {
    throw new Error("初回有給付与テスト失敗: " + JSON.stringify(failed));
  }

  Logger.log(JSON.stringify(results, null, 2));
  return {
    ok: true,
    case_count: results.length,
    results: results
  };
}

/* =========================
   初回付与候補・実行判定テスト
   シート書込み・LockServiceを使用しない
========================= */
function testInitialPaidLeaveGrantCandidateAndExecutionDecision() {
  const advanceWeek4 = {
    employee_id: "C-01",
    company_code: "MAIN",
    hire_date: "2026-01-15",
    work_days_per_week: 4
  };
  const advanceWeek3 = {
    employee_id: "C-02",
    company_code: "PARTNER",
    hire_date: "2026-02-10",
    work_days_per_week: 3
  };
  const normalWeek4 = {
    employee_id: "C-03",
    company_code: "MAIN",
    hire_date: "2026-05-15",
    work_days_per_week: 4
  };

  const cases = [
    {
      name: "候補: 前倒し週4日は10日",
      actual: calculateInitialPaidLeaveGrantEligibility_(advanceWeek4, [], "2026-04-01"),
      expected: { status: "DUE_TODAY", expected_grant_days: 10, grant_reason: "INITIAL_COMPANY_BASIS" }
    },
    {
      name: "候補: 前倒し週3日は10日",
      actual: calculateInitialPaidLeaveGrantEligibility_(advanceWeek3, [], "2026-06-01"),
      expected: { status: "DUE_TODAY", expected_grant_days: 10, grant_reason: "INITIAL_COMPANY_BASIS" }
    },
    {
      name: "候補: 通常週4日は7日",
      actual: calculateInitialPaidLeaveGrantEligibility_(normalWeek4, [], "2026-11-15"),
      expected: { status: "DUE_TODAY", expected_grant_days: 7, grant_reason: "INITIAL_SIX_MONTHS" }
    },
    {
      name: "候補: 処理済み社員は候補外",
      actual: getInitialPaidLeaveGrantExecutionDecision_(
        calculateInitialPaidLeaveGrantEligibility_(advanceWeek4, [{ employee_id: "C-01", grant_type: "six_month" }], "2026-04-01")
      ),
      expected: { can_execute: false }
    },
    {
      name: "候補: initial履歴は候補外",
      actual: getInitialPaidLeaveGrantExecutionDecision_(
        calculateInitialPaidLeaveGrantEligibility_(advanceWeek3, [{ employee_id: "C-02", grant_type: "initial" }], "2026-06-01")
      ),
      expected: { can_execute: false }
    },
    {
      name: "候補: 未来予定は実付与対象外",
      actual: getInitialPaidLeaveGrantExecutionDecision_(
        calculateInitialPaidLeaveGrantEligibility_(advanceWeek4, [], "2026-03-31")
      ),
      expected: { can_execute: false }
    },
    {
      name: "候補: 入社1年超の未処理社員はOVERDUE",
      actual: calculateInitialPaidLeaveGrantEligibility_(normalWeek4, [], "2027-06-01"),
      expected: { status: "OVERDUE" }
    },
    {
      name: "実行: 画面が7日でも前倒しは10日",
      actual: calculateInitialPaidLeaveGrantEligibility_(advanceWeek4, [], "2026-04-01"),
      expected: { expected_grant_days: 10 }
    },
    {
      name: "実行: 画面が10日でも通常週4日は7日",
      actual: calculateInitialPaidLeaveGrantEligibility_(normalWeek4, [], "2026-11-15"),
      expected: { expected_grant_days: 7 }
    },
    {
      name: "実行: ロック後に付与済みなら登録しない",
      actual: getInitialPaidLeaveGrantExecutionDecision_(
        calculateInitialPaidLeaveGrantEligibility_(advanceWeek4, [{ employee_id: "C-01", grant_type: "six_month", grant_date: "2026-04-01" }], "2026-04-01")
      ),
      expected: { can_execute: false }
    },
    {
      name: "実行: UNJUDGEABLEは登録しない",
      actual: getInitialPaidLeaveGrantExecutionDecision_(
        calculateInitialPaidLeaveGrantEligibility_({ employee_id: "C-04", company_code: "OTHER", hire_date: "2026-01-01" }, [], "2026-04-01")
      ),
      expected: { can_execute: false }
    },
    {
      name: "実行: UPCOMINGは登録しない",
      actual: getInitialPaidLeaveGrantExecutionDecision_(
        calculateInitialPaidLeaveGrantEligibility_(normalWeek4, [], "2026-10-01")
      ),
      expected: { can_execute: false }
    },
    {
      name: "実行: 同じ社員の連続実行は2回目を拒否",
      actual: getInitialPaidLeaveGrantExecutionDecision_(
        calculateInitialPaidLeaveGrantEligibility_(normalWeek4, [{ employee_id: "C-03", grant_type: "six_month", grant_date: "2026-11-15" }], "2026-11-15")
      ),
      expected: { can_execute: false }
    },
    {
      name: "実行: 前倒し付与後は6か月到達日に再付与しない",
      actual: getInitialPaidLeaveGrantExecutionDecision_(
        calculateInitialPaidLeaveGrantEligibility_(advanceWeek4, [{ employee_id: "C-01", grant_type: "six_month", grant_date: "2026-04-01", grant_days: 10 }], "2026-07-15")
      ),
      expected: { can_execute: false }
    }
  ];

  const results = cases.map(testCase => {
    const failures = [];
    Object.keys(testCase.expected).forEach(key => {
      if (JSON.stringify(testCase.actual[key]) !== JSON.stringify(testCase.expected[key])) {
        failures.push({ key: key, expected: testCase.expected[key], actual: testCase.actual[key] });
      }
    });
    return { name: testCase.name, ok: failures.length === 0, failures: failures };
  });

  const failed = results.filter(result => !result.ok);
  if (failed.length > 0) {
    throw new Error("初回付与候補・実行判定テスト失敗: " + JSON.stringify(failed));
  }

  Logger.log(JSON.stringify(results, null, 2));
  return { ok: true, case_count: results.length, results: results };
}

/* =========================
   年次予定・統合予定・FIFO表示の固定データテスト
   シート・API・書込みを使用しない
========================= */
function testPaidLeaveGrantScheduleReadOnly() {
  function employee(id, hireDate, extra) {
    return Object.assign({
      employee_id: id, company_code: "MAIN", fiscal_start_month: 4,
      hire_date: hireDate, employment_status: "active",
      leave_management_target: true, work_days_per_week: 5
    }, extra || {});
  }
  function initialHistory(id) {
    return [{ employee_id: id, grant_type: "six_month", grant_date: "2025-04-01", grant_days: 10 }];
  }
  function yearly(id, hireDate, asOf, extra, rows) {
    return calculateYearlyPaidLeaveGrantEligibility_(employee(id, hireDate, extra), rows || initialHistory(id), asOf);
  }
  function fifo(lots, total, asOf) {
    return buildPaidLeaveGrantScheduleFifoView_({ grant_details: lots, current_remaining_days: total }, asOf || "2026-07-24");
  }
  const cases = [
    ["年次: 18か月は11日", yearly("Y01", "2024-10-01", "2026-04-01"), "expected_grant_days", 11],
    ["年次: 30か月は12日", yearly("Y02", "2023-10-01", "2026-04-01"), "expected_grant_days", 12],
    ["年次: 42か月は14日", yearly("Y03", "2022-10-01", "2026-04-01"), "expected_grant_days", 14],
    ["年次: 54か月は16日", yearly("Y04", "2021-10-01", "2026-04-01"), "expected_grant_days", 16],
    ["年次: 66か月は18日", yearly("Y05", "2020-10-01", "2026-04-01"), "expected_grant_days", 18],
    ["年次: 78か月は20日", yearly("Y06", "2019-10-01", "2026-04-01"), "expected_grant_days", 20],
    ["年次: PARTNER基準日", yearly("Y07", "2024-12-01", "2026-06-01", { company_code: "PARTNER", fiscal_start_month: 6 }), "next_grant_date", "2026-06-01"],
    ["年次: 基準日前はUPCOMING", yearly("Y08", "2024-10-01", "2026-03-31"), "status", "UPCOMING"],
    ["年次: 基準日当日はDUE_TODAY", yearly("Y09", "2024-10-01", "2026-04-01"), "status", "DUE_TODAY"],
    ["年次: 基準日後はOVERDUE", yearly("Y10", "2024-10-01", "2026-04-02"), "status", "OVERDUE"],
    ["年次: 同年度付与済み", yearly("Y11", "2024-10-01", "2026-04-02", null, initialHistory("Y11").concat([{ employee_id: "Y11", grant_type: "yearly", year: 2026, grant_date: "2026-04-01", grant_days: 11 }])), "status", "PROCESSED"],
    ["年次: 初回未処理", yearly("Y12", "2024-10-01", "2026-04-01", null, []), "status", "NOT_ELIGIBLE"],
    ["年次: 未知会社", yearly("Y13", "2024-10-01", "2026-04-01", { company_code: "OTHER" }), "status", "UNJUDGEABLE"],
    ["年次: 年度開始月不整合", yearly("Y14", "2024-10-01", "2026-04-01", { fiscal_start_month: 6 }), "is_provisional", true],
    ["年次: active対象", yearly("Y15", "2024-10-01", "2026-04-01"), "status", "DUE_TODAY"],
    ["年次: 在職対象", yearly("Y16", "2024-10-01", "2026-04-01", { employment_status: "在職" }), "status", "DUE_TODAY"],
    ["年次: 退職者は対象外", yearly("Y17", "2024-10-01", "2026-04-01", { employment_status: "retired" }), "status", "NOT_ELIGIBLE"],
    ["年次: 休職者は対象外", yearly("Y18", "2024-10-01", "2026-04-01", { employment_status: "休職" }), "status", "NOT_ELIGIBLE"],
    ["年次: 管理対象外", yearly("Y19", "2024-10-01", "2026-04-01", { leave_management_target: false }), "status", "NOT_ELIGIBLE"],
    ["年次: 入社日不正", yearly("Y20", "invalid", "2026-04-01"), "status", "UNJUDGEABLE"],
    ["年次: 年次履歴重複警告", yearly("Y21", "2024-10-01", "2026-04-01", null, initialHistory("Y21").concat([{ employee_id: "Y21", grant_type: "yearly", year: 2026, grant_date: "2026-04-01", grant_days: 11 }, { employee_id: "Y21", grant_type: "yearly", year: 2026, grant_date: "2026-04-01", grant_days: 11 }])), "warning_codes", ["DUPLICATE_YEARLY_GRANT_HISTORY"]],
    ["年次: yearと日付不整合警告", yearly("Y22", "2024-10-01", "2026-04-01", null, initialHistory("Y22").concat([{ employee_id: "Y22", grant_type: "yearly", year: 2025, grant_date: "2026-04-01" }])), "warning_codes", ["YEAR_AND_GRANT_DATE_MISMATCH"]],
    ["年次: 週4日は暫定", yearly("Y23", "2024-10-01", "2026-04-01", { work_days_per_week: 4 }), "expected_grant_days", null],
    ["年次: 週3日は暫定", yearly("Y24", "2024-10-01", "2026-04-01", { work_days_per_week: 3 }), "is_provisional", true],
    ["年次: 未知履歴種別を警告", yearly("Y25", "2024-10-01", "2026-04-01", null, initialHistory("Y25").concat([{ employee_id: "Y25", grant_type: "legacy_unknown", grant_date: "2024-01-01" }])), "warning_codes", ["UNKNOWN_GRANT_TYPE"]],
    ["統合: 初回未処理は初回を返す", calculateNextPaidLeaveGrantSchedule_(employee("I01", "2026-01-15"), [], "2026-04-01"), "grant_stage", "INITIAL"],
    ["統合: 初回済みは年次を返す", calculateNextPaidLeaveGrantSchedule_(employee("I02", "2024-10-01"), initialHistory("I02"), "2026-04-01"), "grant_stage", "YEARLY"],
    ["統合: 前倒し済みは6か月日に再候補化しない", calculateNextPaidLeaveGrantSchedule_(employee("I03", "2026-01-15"), [{ employee_id: "I03", grant_type: "six_month", grant_date: "2026-04-01", grant_days: 10 }], "2026-07-15"), "status", "NOT_ELIGIBLE"],
    ["統合: 年次未到来", calculateNextPaidLeaveGrantSchedule_(employee("I04", "2024-10-01"), initialHistory("I04"), "2026-03-31"), "status", "UPCOMING"],
    ["統合: 年次予定日超過", calculateNextPaidLeaveGrantSchedule_(employee("I05", "2024-10-01"), initialHistory("I05"), "2026-04-02"), "status", "OVERDUE"],
    ["統合: 年次処理済み", calculateNextPaidLeaveGrantSchedule_(employee("I06", "2024-10-01"), initialHistory("I06").concat([{ employee_id: "I06", grant_type: "yearly", year: 2026, grant_date: "2026-04-01", grant_days: 11 }]), "2026-04-02"), "status", "PROCESSED"],
    ["統合: 出勤率は未確認", calculateNextPaidLeaveGrantSchedule_(employee("I07", "2024-10-01"), initialHistory("I07"), "2026-04-01"), "attendance_status", "ATTENDANCE_UNCONFIRMED"],
    ["統合: 未来入社は判定不能", calculateNextPaidLeaveGrantSchedule_(employee("I08", "2027-01-01"), [], "2026-04-01"), "status", "UNJUDGEABLE"],
    ["統合: 会社年度不整合は暫定", calculateNextPaidLeaveGrantSchedule_(employee("I09", "2024-10-01", { fiscal_start_month: 6 }), initialHistory("I09"), "2026-04-01"), "is_provisional", true],
    ["FIFO: 古いロット優先1", fifo([{ grant_id: "G1", grant_date: "2025-04-01", valid_from: "2025-04-01", valid_to: "2027-03-31", total_days: 10, used_days: 4, active_remaining_days: 6 }], 6).lots[0], "consumption_priority", 1],
    ["FIFO: 新しいロット優先2", fifo([{ grant_id: "G1", grant_date: "2025-04-01", valid_from: "2025-04-01", valid_to: "2027-03-31", total_days: 10, used_days: 4, active_remaining_days: 6 }, { grant_id: "G2", grant_date: "2026-04-01", valid_from: "2026-04-01", valid_to: "2028-03-31", total_days: 10, used_days: 0, active_remaining_days: 10 }], 16).lots[1], "consumption_priority", 2],
    ["FIFO: 期限切れは対象外", fifo([{ grant_id: "G3", grant_date: "2024-04-01", valid_from: "2024-04-01", valid_to: "2026-03-31", total_days: 10, used_days: 0, active_remaining_days: 0, is_expired: true }], 0).lots[0], "status", "EXPIRED"],
    ["FIFO: 全消化済み", fifo([{ grant_id: "G4", grant_date: "2025-04-01", valid_from: "2025-04-01", valid_to: "2027-03-31", total_days: 10, used_days: 10, active_remaining_days: 0 }], 0).lots[0], "status", "FULLY_USED"],
    ["FIFO: 30日以内期限警告", fifo([{ grant_id: "G5", grant_date: "2025-08-01", valid_from: "2025-08-01", valid_to: "2026-08-01", total_days: 5, used_days: 0, active_remaining_days: 5 }], 5), "warning_codes", ["FIFO_LOT_EXPIRING_WITHIN_30_DAYS"]],
    ["FIFO: 未来ロット", fifo([{ grant_id: "G6", grant_date: "2026-08-01", valid_from: "2026-08-01", valid_to: "2028-07-31", total_days: 5, used_days: 0, active_remaining_days: 5 }], 0).lots[0], "status", "FUTURE"],
    ["FIFO: 合計一致", fifo([{ grant_id: "G7", grant_date: "2025-04-01", valid_from: "2025-04-01", valid_to: "2027-03-31", total_days: 8, used_days: 2, active_remaining_days: 6 }], 6).warning_codes, "length", 0],
    ["FIFO: 負残高警告", fifo([{ grant_id: "G8", grant_date: "2025-04-01", valid_from: "2025-04-01", valid_to: "2027-03-31", total_days: 8, used_days: 9, active_remaining_days: -1 }], -1), "warning_codes", ["NEGATIVE_FIFO_REMAINING", "FIFO_TOTAL_MISMATCH"]],
    ["FIFO: 複数ロット残高", fifo([{ grant_id: "G9", grant_date: "2025-04-01", valid_from: "2025-04-01", valid_to: "2027-03-31", total_days: 10, used_days: 10, active_remaining_days: 0 }, { grant_id: "G10", grant_date: "2026-04-01", valid_from: "2026-04-01", valid_to: "2028-03-31", total_days: 10, used_days: 2, active_remaining_days: 8 }], 8).total_remaining_days, "valueOf", 8],
    ["FIFO: 同日ロット順序維持", fifo([{ grant_id: "GA", grant_date: "2025-04-01", valid_from: "2025-04-01", valid_to: "2027-03-31", total_days: 1, used_days: 0, active_remaining_days: 1 }, { grant_id: "GB", grant_date: "2025-04-01", valid_from: "2025-04-01", valid_to: "2027-03-31", total_days: 1, used_days: 0, active_remaining_days: 1 }], 2).lots[0], "grant_id", "GA"],
    ["FIFO: 重複ロット警告", fifo([{ grant_id: "GD", grant_date: "2025-04-01", valid_from: "2025-04-01", valid_to: "2027-03-31", total_days: 1, used_days: 0, active_remaining_days: 1 }, { grant_id: "GD", grant_date: "2026-04-01", valid_from: "2026-04-01", valid_to: "2028-03-31", total_days: 1, used_days: 0, active_remaining_days: 1 }], 2), "warning_codes", ["DUPLICATE_GRANT_LOT"]],
    ["API整形: 優先順位ソート", comparePaidLeaveGrantScheduleRows_({ eligibility_status: "OVERDUE" }, { eligibility_status: "UPCOMING" }), "valueOf", -4],
    ["API整形: 30日以内だけ含む", shouldIncludePaidLeaveGrantScheduleRow_({ eligibility_status: "UPCOMING", next_grant_date: "2026-08-01", warning_codes: [] }, parseLocalDate("2026-08-23"), parseLocalDate("2026-06-23")), "valueOf", true],
    ["API整形: 会社フィルター", isPaidLeaveGrantScheduleCompanyMatch_({ company_code: "MAIN" }, "MAIN"), "valueOf", true],
    ["API整形: 処理済み期間", shouldIncludePaidLeaveGrantScheduleRow_({ eligibility_status: "PROCESSED", next_grant_date: "2026-07-20", warning_codes: [] }, parseLocalDate("2026-08-23"), parseLocalDate("2026-06-23")), "valueOf", true],
    ["API整形: 警告ありは要確認", shouldIncludePaidLeaveGrantScheduleRow_({ eligibility_status: "UPCOMING", next_grant_date: "2027-01-01", warning_codes: ["X"], requires_manual_confirmation: true }, parseLocalDate("2026-08-23"), parseLocalDate("2026-06-23")), "valueOf", true]
  ];

  const results = cases.map(item => {
    const value = item[2] === "valueOf" ? item[1] : item[1][item[2]];
    const ok = JSON.stringify(value) === JSON.stringify(item[3]);
    return { name: item[0], ok: ok, expected: item[3], actual: value };
  });
  const failed = results.filter(result => !result.ok);
  if (failed.length) throw new Error("年次予定・FIFO表示テスト失敗: " + JSON.stringify(failed));
  Logger.log(JSON.stringify(results, null, 2));
  return { ok: true, case_count: results.length, results: results };
}

/* =========================
   監査修正: 予定APIの安全性・境界値テスト
   固定データのみ。シート、CacheService、書込みを使用しない。
========================= */
function testPaidLeaveGrantScheduleAuditFixes() {
  const asOf = "2026-04-02";
  const employee = extra => Object.assign({
    employee_id: "AUDIT-01", company_code: "MAIN", fiscal_start_month: 4,
    hire_date: "2024-10-01", employment_status: "active",
    leave_management_target: true, work_days_per_week: 5
  }, extra || {});
  const initial = id => [{ employee_id: id || "AUDIT-01", grant_type: "six_month", grant_date: "2025-04-01", grant_days: 10 }];
  const yearly = rows => calculateYearlyPaidLeaveGrantEligibility_(employee(), initial().concat(rows || []), asOf);
  const dateKey = (date, months) => formatInitialGrantDateKey_(addMonthsClampedLocal_(date, months));
  const cases = [
    ["読取カレンダー変換", buildCompanyCalendarMapFromRows_([{ date: "2026-04-01", type: "holiday" }])["2026-04-01"], "holiday"],
    ["1月末+6か月", dateKey("2026-01-31", 6), "2026-07-31"],
    ["3月末+6か月", dateKey("2026-03-31", 6), "2026-09-30"],
    ["8月末+6か月", dateKey("2026-08-31", 6), "2027-02-28"],
    ["うるう年前8月末+6か月", dateKey("2027-08-31", 6), "2028-02-29"],
    ["10月末+6か月", dateKey("2026-10-31", 6), "2027-04-30"],
    ["2月29日+6か月", dateKey("2024-02-29", 6), "2024-08-29"],
    ["正常年次履歴", yearly([{ employee_id: "AUDIT-01", grant_type: "YEARLY", year: 2026, grant_date: "2026-04-01", grant_days: 11 }]).status, "PROCESSED"],
    ["year空欄", yearly([{ employee_id: "AUDIT-01", grant_type: "yearly", grant_date: "2026-04-01", grant_days: 11 }]).status, "UNJUDGEABLE"],
    ["年日付不一致", yearly([{ employee_id: "AUDIT-01", grant_type: "yearly", year: 2025, grant_date: "2026-04-01", grant_days: 11 }]).status, "UNJUDGEABLE"],
    ["未来年次", yearly([{ employee_id: "AUDIT-01", grant_type: "yearly", year: 2026, grant_date: "2026-05-01", grant_days: 11 }]).status, "UNJUDGEABLE"],
    ["0日年次", yearly([{ employee_id: "AUDIT-01", grant_type: "yearly", year: 2026, grant_date: "2026-04-01", grant_days: 0 }]).status, "UNJUDGEABLE"],
    ["重複年次は処理済み", yearly([{ employee_id: "AUDIT-01", grant_type: "yearly", year: 2026, grant_date: "2026-04-01", grant_days: 11 }, { employee_id: "AUDIT-01", grant_type: "yearly", year: 2026, grant_date: "2026-04-01", grant_days: 11 }]).warning_codes.indexOf("DUPLICATE_YEARLY_GRANT_HISTORY") >= 0, true],
    ["勤務日数5", normalizeWorkDaysPerWeek_("5").value, 5],
    ["勤務日数4", normalizeWorkDaysPerWeek_(4).value, 4],
    ["勤務日数字符4", normalizeWorkDaysPerWeek_("4").value, 4],
    ["勤務日数1", normalizeWorkDaysPerWeek_(1).value, 1],
    ["勤務日数0", normalizeWorkDaysPerWeek_(0).warning_code, "INVALID_WORK_DAYS_PER_WEEK"],
    ["勤務日数6", normalizeWorkDaysPerWeek_(6).warning_code, "INVALID_WORK_DAYS_PER_WEEK"],
    ["勤務日数小数", normalizeWorkDaysPerWeek_(4.5).warning_code, "INVALID_WORK_DAYS_PER_WEEK"],
    ["勤務日数空欄", normalizeWorkDaysPerWeek_("").warning_code, "WORK_DAYS_PER_WEEK_MISSING"],
    ["勤務日数null", normalizeWorkDaysPerWeek_(null).warning_code, "WORK_DAYS_PER_WEEK_MISSING"],
    ["勤務日数文字", normalizeWorkDaysPerWeek_("週4日").warning_code, "INVALID_WORK_DAYS_PER_WEEK"],
    ["勤務日数abc", normalizeWorkDaysPerWeek_("abc").warning_code, "INVALID_WORK_DAYS_PER_WEEK"],
    ["不正勤務日数は初回判定不能", calculateInitialPaidLeaveGrantEligibility_(employee({ hire_date: "2026-05-15", work_days_per_week: 0 }), [], "2026-11-15").status, "UNJUDGEABLE"],
    ["短時間勤務者は年次日数未確定", calculateYearlyPaidLeaveGrantEligibility_(employee({ work_days_per_week: 4 }), initial(), asOf).expected_grant_days, null],
    ["対象外除外", shouldIncludePaidLeaveGrantScheduleRow_({ eligibility_status: "NOT_ELIGIBLE", warning_codes: ["X"] }, parseLocalDate("2026-05-01"), parseLocalDate("2026-03-01")), false],
    ["退職者除外", shouldIncludePaidLeaveGrantScheduleRow_({ eligibility_status: "NOT_ELIGIBLE", warning_codes: ["EMPLOYMENT_STATUS_NOT_ACTIVE"] }, parseLocalDate("2026-05-01"), parseLocalDate("2026-03-01")), false],
    ["対象外除外", shouldIncludePaidLeaveGrantScheduleRow_({ eligibility_status: "NOT_ELIGIBLE", warning_codes: ["LEAVE_MANAGEMENT_TARGET_DISABLED"] }, parseLocalDate("2026-05-01"), parseLocalDate("2026-03-01")), false],
    ["会社不明は表示対象", shouldIncludePaidLeaveGrantScheduleRow_({ eligibility_status: "UNJUDGEABLE", warning_codes: ["COMPANY_CODE_UNSUPPORTED"] }, parseLocalDate("2026-05-01"), parseLocalDate("2026-03-01")), true],
    ["入社日不正は表示対象", shouldIncludePaidLeaveGrantScheduleRow_({ eligibility_status: "UNJUDGEABLE", warning_codes: ["HIRE_DATE_INVALID"] }, parseLocalDate("2026-05-01"), parseLocalDate("2026-03-01")), true],
    ["通常確認はデータ異常でない", hasPaidLeaveGrantScheduleDataIssue_(["YEARLY_PROPORTIONAL_GRANT_RULE_NOT_IMPLEMENTED"]), false],
    ["履歴不整合はデータ異常", hasPaidLeaveGrantScheduleDataIssue_(["YEAR_AND_GRANT_DATE_MISMATCH"]), true],
    ["未来six_month履歴", calculateNextPaidLeaveGrantSchedule_(employee(), [{ employee_id: "AUDIT-01", grant_type: "six_month", grant_date: "2026-05-01", grant_days: 10 }], asOf).warning_codes, ["FUTURE_INITIAL_GRANT_HISTORY"]],
    ["未来initial履歴", calculateNextPaidLeaveGrantSchedule_(employee(), [{ employee_id: "AUDIT-01", grant_type: "initial", grant_date: "2026-05-01", grant_days: 10 }], asOf).warning_codes, ["FUTURE_INITIAL_GRANT_HISTORY"]],
    ["未来processed履歴", calculateNextPaidLeaveGrantSchedule_(employee(), [{ employee_id: "AUDIT-01", grant_type: "six_month_processed", grant_date: "2026-05-01", grant_days: 0 }], asOf).warning_codes, ["FUTURE_INITIAL_GRANT_HISTORY"]],
    ["未来初回履歴", calculateNextPaidLeaveGrantSchedule_(employee({ hire_date: "2024-10-01" }), [{ employee_id: "AUDIT-01", grant_type: "six_month", grant_date: "2026-05-01", grant_days: 10 }], asOf).status, "UNJUDGEABLE"]
  ];
  const failed = cases.filter(item => JSON.stringify(item[1]) !== JSON.stringify(item[2]));
  if (failed.length) throw new Error("予定API監査修正テスト失敗: " + JSON.stringify(failed));
  return { ok: true, case_count: cases.length };
}

/* =========================
   PARTNER初期導入残高補正の固定データ安全性テスト
   シート・LockService・書込みを使用しない。
========================= */
function testPartnerOpeningBalanceRepairSafety() {
  const employeeMap = {
    EMP0060: { company_code: "PARTNER" },
    EMP0061: { company_code: "PARTNER" }
  };
  const validRows = [
    { grant_id: "G0056", employee_id: "EMP0060", grant_days: 15, carry_over_days: 20, notes: "初期導入残高" },
    { grant_id: "G0057", employee_id: "EMP0061", grant_days: 15, carry_over_days: 20, notes: "初期導入残高" }
  ];
  const expectThrow = (rows, map) => {
    try { validatePartnerOpeningBalanceRepairRows_(rows, map); return false; } catch (e) { return true; }
  };
  const asOfDate = parseLocalDate("2027-06-01");
  const context = {
    as_of_date: asOfDate,
    calendar_map: {},
    requests_by_employee: { EMP0060: [] },
    grants_by_employee: {
      EMP0060: [{
        grant_id: "G0056", employee_id: "EMP0060",
        grant_date: parseLocalDate("2025-06-01"),
        valid_from_date: parseLocalDate("2025-06-01"),
        valid_to_date: parseLocalDate("2027-05-31"),
        grant_type: "initial", year: 2025,
        grant_days: 15, carry_over_days: 20, total_days: 35,
        notes: "初期導入残高", has_recorded_valid_from: true,
        has_recorded_valid_to: true, is_finalized: true
      }]
    }
  };
  const before = calculateFifoBalanceWithOpeningBalanceFromContext_("EMP0060", asOfDate, context);
  const afterContext = buildPartnerOpeningBalanceCarryOnlyScenarioContext_("EMP0060", "G0056", context);
  const afterBalance = calculateFifoBalanceWithOpeningBalanceFromContext_("EMP0060", asOfDate, afterContext);
  const after = simulatePartnerOpeningBalanceScenario_("EMP0060", "G0056", "CARRY_OVER_DAYS_ONLY", asOfDate, context);
  const cases = [
    ["正常2件を検証", validatePartnerOpeningBalanceRepairRows_(validRows, employeeMap), true],
    ["grant_days不一致で停止", expectThrow([Object.assign({}, validRows[0], { grant_days: 14 }), validRows[1]], employeeMap), true],
    ["carry_over_days不一致で停止", expectThrow([Object.assign({}, validRows[0], { carry_over_days: 19 }), validRows[1]], employeeMap), true],
    ["notes不一致で停止", expectThrow([Object.assign({}, validRows[0], { notes: "手入力" }), validRows[1]], employeeMap), true],
    ["employee_id不一致で停止", expectThrow([Object.assign({}, validRows[0], { employee_id: "EMP9999" }), validRows[1]], employeeMap), true],
    ["片方不足で両方停止", expectThrow([validRows[0]], employeeMap), true],
    ["確認文字列不一致で本実行拒否", (() => { try { assertPartnerOpeningBalanceRepairConfirmation_(false, "invalid"); return false; } catch (e) { return true; } })(), true],
    ["旧確認文字列でも本実行拒否", (() => { try { assertPartnerOpeningBalanceRepairConfirmation_(false, PARTNER_OPENING_BALANCE_REPAIR_CONFIRMATION_); return false; } catch (e) { return true; } })(), true],
    ["補正前は二重ロット", before.grant_details.length, 2],
    ["補正後は二重ロットなし", afterBalance.grant_details.length, 1],
    ["試算後grant_daysは0", afterContext.grants_by_employee.EMP0060[0].grant_days, 0],
    ["試算後carry_over_daysは維持", afterContext.grants_by_employee.EMP0060[0].carry_over_days, 20],
    ["補正後も有効期限は不変", formatInitialGrantDateKey_(afterContext.grants_by_employee.EMP0060[0].valid_to_date), "2027-05-31"],
    ["期限切れはgrant_days分の15日減少", after.expired_days - Number(before.expired_days || 0), -15],
    ["P0004は固定対象外", PARTNER_OPENING_BALANCE_REPAIR_TARGETS_.some(row => row.grant_id === "G0058"), false]
  ];
  const failed = cases.filter(item => JSON.stringify(item[1]) !== JSON.stringify(item[2]));
  if (failed.length) throw new Error("PARTNER初期導入残高補正テスト失敗: " + JSON.stringify(failed));
  return { ok: true, case_count: cases.length };
}

/* =========================
   PARTNER繰越試算ログの純粋関数テスト
========================= */
function testPartnerCarryOverSimulationLogging() {
  const targets = [
    {
      grant_id: "G0056", employee_id: "EMP0060", display_employee_id: "P0002",
      current_fifo: { current_active_remaining_days: 19, expired_days: 34 },
      current_valid_to_maintained: { current_active_remaining_days: 19, expired_days: 19 },
      reference_valid_to_simulation: { valid_to: "2027/05/31", current_active_remaining_days: 39, expired_days: 0 },
      yearly_balance_after: { current_remaining_days: 39 },
      warnings: ["W"], required_data_corrections: ["C"], expiry_confirmation_required: ["E"]
    },
    { grant_id: "G0057", employee_id: "EMP0061", display_employee_id: "P0003" }
  ];
  const chunks = createJsonLogChunks_("[P0002][CURRENT_FIFO]", { text: Array(260).join("あ") }, 100);
  const expectThrow = () => {
    try { selectPartnerCarryOverSimulationTarget_(targets, "P9999"); return false; } catch (e) { return true; }
  };
  const original = JSON.stringify(targets[0]);
  const summary = buildPartnerCarryOverSimulationSummary_(targets[0]).join("\n");
  const cases = [
    ["P0002個別抽出", selectPartnerCarryOverSimulationTarget_(targets, "P0002").grant_id, "G0056"],
    ["P0003個別抽出", selectPartnerCarryOverSimulationTarget_(targets, "P0003").grant_id, "G0057"],
    ["存在しない表示IDは停止", expectThrow(), true],
    ["要約に現状FIFO", summary.indexOf("現状:") !== -1, true],
    ["要約に期限維持案", summary.indexOf("期限維持") !== -1, true],
    ["要約に参考期限案", summary.indexOf("参考期限") !== -1, true],
    ["要約に年度残高", summary.indexOf("年度残高") !== -1, true],
    ["JSON分割は連番", chunks.length > 1 && chunks[0].indexOf("[1/") !== -1 && chunks[1].indexOf("[2/") !== -1, true],
    ["JSON分割payloadは指定文字数以下", chunks.every(chunk => Array.from(chunk.split("\n").slice(1).join("\n")).length <= 100), true],
    ["ログ生成は元結果を変更しない", JSON.stringify(targets[0]), original],
    ["ログ要約に個人名・notesなし", summary.indexOf("氏名") === -1 && summary.indexOf("notes") === -1, true],
    ["旧補正本実行拒否", (() => { try { assertPartnerOpeningBalanceRepairConfirmation_(false, "any"); return false; } catch (e) { return true; } })(), true]
  ];
  const failed = cases.filter(item => JSON.stringify(item[1]) !== JSON.stringify(item[2]));
  if (failed.length) throw new Error("PARTNER繰越試算ログテスト失敗: " + JSON.stringify(failed));
  return { ok: true, case_count: cases.length };
}

/* =========================
   PARTNER初期導入残高: 年度開始繰越ロット試算
   固定データだけでFIFOの読取専用試算を確認する。
========================= */
function testPartnerOpeningBalanceFiscalStartSimulation() {
  const asOfDate = parseLocalDate("2026-07-25");
  const originalContext = {
    as_of_date: asOfDate,
    calendar_map: {},
    requests_by_employee: {
      EMP0060: [
        {
          request_id: "R-MAY-02", employee_id: "EMP0060",
          start_date: "2026-05-02", end_date: "2026-05-02", days: 1,
          half_day: "", type: "paid_leave", status: STATUS.APPROVED
        },
        {
          request_id: "R-JUL-18", employee_id: "EMP0060",
          start_date: "2026-07-18", end_date: "2026-07-18", days: 1,
          half_day: "", type: "paid_leave", status: STATUS.APPROVED
        }
      ],
      EMP0061: []
    },
    grants_by_employee: {
      EMP0060: [
        {
          grant_id: "G0056", employee_id: "EMP0060", grant_type: "initial", year: 2025,
          grant_date: parseLocalDate("2025-06-01"),
          valid_from_date: parseLocalDate("2025-06-01"),
          valid_to_date: parseLocalDate("2026-05-31"),
          grant_days: 15, carry_over_days: 20, total_days: 35,
          notes: "初期導入残高", has_recorded_valid_from: true,
          has_recorded_valid_to: true, is_finalized: true
        },
        {
          grant_id: "G2026-60", employee_id: "EMP0060", grant_type: "yearly", year: 2026,
          grant_date: parseLocalDate("2026-06-01"),
          valid_from_date: parseLocalDate("2026-06-01"),
          valid_to_date: parseLocalDate("2028-05-31"),
          grant_days: 20, carry_over_days: 0, total_days: 20,
          notes: "", has_recorded_valid_from: true,
          has_recorded_valid_to: true, is_finalized: true
        }
      ],
      EMP0061: [
        {
          grant_id: "G0057", employee_id: "EMP0061", grant_type: "initial", year: 2025,
          grant_date: parseLocalDate("2025-06-01"),
          valid_from_date: parseLocalDate("2025-06-01"),
          valid_to_date: parseLocalDate("2026-05-31"),
          grant_days: 15, carry_over_days: 20, total_days: 35,
          notes: "初期導入残高", has_recorded_valid_from: true,
          has_recorded_valid_to: true, is_finalized: true
        },
        {
          grant_id: "G2026-61", employee_id: "EMP0061", grant_type: "yearly", year: 2026,
          grant_date: parseLocalDate("2026-06-01"),
          valid_from_date: parseLocalDate("2026-06-01"),
          valid_to_date: parseLocalDate("2028-05-31"),
          grant_days: 20, carry_over_days: 0, total_days: 20,
          notes: "", has_recorded_valid_from: true,
          has_recorded_valid_to: true, is_finalized: true
        }
      ]
    }
  };
  const fiscalStartContext60 = buildPartnerOpeningBalanceCarryOnlyScenarioContext_(
    "EMP0060", "G0056", originalContext,
    parseLocalDate("2027-05-31"), parseLocalDate("2026-06-01")
  );
  const fiscalStartContext61 = buildPartnerOpeningBalanceCarryOnlyScenarioContext_(
    "EMP0061", "G0057", originalContext,
    parseLocalDate("2027-05-31"), parseLocalDate("2026-06-01")
  );
  const fifo60 = calculateFifoBalanceWithOpeningBalanceFromContext_(
    "EMP0060", asOfDate, fiscalStartContext60
  );
  const fifo61 = calculateFifoBalanceWithOpeningBalanceFromContext_(
    "EMP0061", asOfDate, fiscalStartContext61
  );
  const may2Allocations = getPartnerFiscalStartSimulationAllocations_(
    fifo60, "2026/05/02"
  );
  const july18Allocations = getPartnerFiscalStartSimulationAllocations_(
    fifo60, "2026/07/18"
  );
  const targetLots = fifo60.grant_details.filter(row => row.source_grant_id === "G0056");
  const cases = [
    ["valid_from=2026/06/01は2026/05/02取得を消化しない", may2Allocations.allocations.length, 0],
    ["P0002は39日", fifo60.current_remaining_days, 39],
    ["P0003は40日", fifo61.current_remaining_days, 40],
    ["2026/07/18取得は年度開始後ロットから1日消化", july18Allocations.allocations.length, 1],
    ["2026/07/18は繰越20日ロットを消化", july18Allocations.allocations[0].grant_id, "G0056#opening_balance"],
    ["grant_days=0で二重ロットが消える", targetLots.length, 1],
    ["試算コンテキストは元データを書換えない", originalContext.grants_by_employee.EMP0060[0].grant_days, 15]
  ];
  const failed = cases.filter(item => JSON.stringify(item[1]) !== JSON.stringify(item[2]));
  if (failed.length) throw new Error("PARTNER年度開始繰越試算テスト失敗: " + JSON.stringify(failed));
  return { ok: true, case_count: cases.length };
}

/* =========================
   PARTNER初期導入残高: 完了済み補正の退役テスト
   書込み関数を残さず、dry-runと再実行禁止だけを確認する。
========================= */
function testPartnerOpeningBalanceFiscalStartRepairSafety() {
  const throwsCompletedError = fn => {
    try { fn(); return ""; } catch (error) { return String(error.message || error); }
  };
  const directError = throwsCompletedError(() =>
    repairPartnerOpeningBalanceFiscalStart({ dry_run: false })
  );
  const executeError = throwsCompletedError(() => executeRepairPartnerOpeningBalanceFiscalStart());
  const source = repairPartnerOpeningBalanceFiscalStart.toString() +
    executeRepairPartnerOpeningBalanceFiscalStart.toString();
  const cases = [
    ["対象はG0056/G0057のみ", PARTNER_OPENING_BALANCE_FISCAL_START_REPAIR_TARGETS_.map(row => row.grant_id), ["G0056", "G0057"]],
    ["G0058は対象外", PARTNER_OPENING_BALANCE_FISCAL_START_REPAIR_TARGETS_.some(row => row.grant_id === "G0058"), false],
    ["本実行を直接呼んでも完了済みエラー", directError, PARTNER_OPENING_BALANCE_FISCAL_START_REPAIR_COMPLETED_ERROR_],
    ["execute入口も完了済みエラー", executeError, PARTNER_OPENING_BALANCE_FISCAL_START_REPAIR_COMPLETED_ERROR_],
    ["本実行経路にLockServiceがない", source.indexOf("LockService") === -1, true],
    ["本実行経路に書込み関数がない", source.indexOf("setValue") === -1 && source.indexOf("setValues") === -1, true],
    ["dry-runは読み取り専用分岐を維持", repairPartnerOpeningBalanceFiscalStart.toString().indexOf("readPartnerOpeningBalanceFiscalStartRepairState_") !== -1, true],
    ["dry-runラッパーを維持", typeof debugRepairPartnerOpeningBalanceFiscalStart, "function"],
    ["FIFO繰越診断を維持", typeof debugPartnerOpeningBalanceCarryOverSimulation, "function"],
    ["年度開始FIFO診断を維持", typeof debugPartnerOpeningBalanceFiscalStartSimulation, "function"],
    ["データソース診断を維持", typeof debugPaidLeaveDataSources, "function"],
    ["FIFOロット診断を維持", typeof debugFifoLots, "function"],
    ["固定FIFO試算を維持", testPartnerOpeningBalanceFiscalStartSimulation().ok, true]
  ];
  const failed = cases.filter(item => JSON.stringify(item[1]) !== JSON.stringify(item[2]));
  if (failed.length) throw new Error("PARTNER年度開始補正退役テスト失敗: " + JSON.stringify(failed));
  return { ok: true, case_count: cases.length };
}

/* =========================
   PARTNER初期導入残高: 事前条件診断の正規化テスト
   実シートを読まず、個人情報を含まない診断形式を確認する。
========================= */
function testPartnerOpeningBalanceFiscalStartRepairPreconditions() {
  const target = PARTNER_OPENING_BALANCE_FISCAL_START_REPAIR_TARGETS_[0];
  const employee = { company_code: " partner ", display_employee_id: "P0002" };
  const row = {
    grant_id: " G0056 ", employee_id: " emp0060 ", grant_days: "15", carry_over_days: "20",
    valid_from: parseLocalDate("2025-06-01"), valid_to: "2026/05/31",
    grant_type: "initial", year: "2025", notes: "初期導入残高（移行）"
  };
  const diagnostic = buildPartnerOpeningBalanceFiscalStartPreconditionDiagnostic_(row, employee, target, 1);
  const invalidNumber = buildPartnerOpeningBalanceFiscalStartNumberCheck_("", 15);
  const invalidDate = buildPartnerOpeningBalanceFiscalStartDateCheck_("not-a-date", "2025-06-01");
  const noMarker = buildPartnerOpeningBalanceFiscalStartPreconditionDiagnostic_(
    Object.assign({}, row, { notes: "移行済み" }), employee, target, 1
  );
  const duplicate = buildPartnerOpeningBalanceFiscalStartPreconditionDiagnostic_(row, employee, target, 2);
  const displayMismatch = buildPartnerOpeningBalanceFiscalStartPreconditionDiagnostic_(
    row, Object.assign({}, employee, { display_employee_id: "P9999" }), target, 1
  );
  const mismatchError = (() => {
    try {
      validatePartnerOpeningBalanceFiscalStartRepairRows_(
        [Object.assign({}, row, { grant_id: "G0056", employee_id: "EMP0060", grant_days: 14 }), {
          grant_id: "G0057", employee_id: "EMP0061", grant_days: 15, carry_over_days: 20,
          valid_from: parseLocalDate("2025-06-01"), valid_to: parseLocalDate("2026-05-31"), notes: "初期導入残高"
        }],
        { EMP0060: employee, EMP0061: { company_code: "PARTNER", display_employee_id: "P0003" } },
        "before"
      );
      return "";
    } catch (error) {
      return String(error.message || error);
    }
  })();
  const cases = [
    ["正規化後の全実行条件が一致", diagnostic.all_conditions_match, true],
    ["IDの全角半角・空白を正規化", diagnostic.checks.employee_id.actual_normalized, "EMP0060"],
    ["会社コードを正規化", diagnostic.checks.company_code.actual_normalized, "PARTNER"],
    ["文字列数値15を安全に正規化", diagnostic.checks.grant_days.matched, true],
    ["空欄数値を期待値へ補完しない", invalidNumber.matched, false],
    ["Date型の日付を同一形式へ正規化", diagnostic.checks.valid_from.actual_normalized, "2025-06-01"],
    ["文字列日付を同一形式へ正規化", diagnostic.checks.valid_to.actual_normalized, "2026-05-31"],
    ["不正日付は一致させない", invalidDate.matched, false],
    ["notes本文を診断に含めない", Object.prototype.hasOwnProperty.call(diagnostic.checks.notes_marker, "actual"), false],
    ["notesマーカー不一致を検出", noMarker.checks.notes_marker.matched, false],
    ["重複grant_idを検出", duplicate.checks.matching_grant_id_row_count.matched, false],
    ["表示ID不一致は表示し実行条件には加えない", displayMismatch.checks.display_employee_id.matched, false],
    ["不一致エラーに項目名を含める", mismatchError.indexOf("grant_days") !== -1, true]
  ];
  const failed = cases.filter(item => JSON.stringify(item[1]) !== JSON.stringify(item[2]));
  if (failed.length) throw new Error("PARTNER年度開始補正事前条件診断テスト失敗: " + JSON.stringify(failed));
  return { ok: true, case_count: cases.length };
}

/* =========================
   PARTNER年度開始補正: Apps Script実行ラッパーのログ整形テスト
   補正本体・シート・LockServiceを呼ばない。
========================= */
function testPartnerOpeningBalanceFiscalStartRepairWrappers() {
  const result = {
    ok: true,
    dry_run: true,
    warnings: ["W"],
    targets: [
      {
        grant_id: "G0056", employee_id: "EMP0060", display_employee_id: "P0002",
        before: { grant_days: 15, valid_from: "2025-06-01", valid_to: "2026-05-31" },
        after: { grant_days: 0, valid_from: "2026-06-01", valid_to: "2027-05-31" },
        difference: { expired_days: -15 }, warnings: ["T1"]
      },
      {
        grant_id: "G0057", employee_id: "EMP0061", display_employee_id: "P0003",
        before: { grant_days: 15, valid_from: "2025-06-01", valid_to: "2026-05-31" },
        after: { grant_days: 0, valid_from: "2026-06-01", valid_to: "2027-05-31" },
        difference: { expired_days: -15 }, warnings: ["T2"]
      }
    ]
  };
  const original = JSON.stringify(result);
  const dryRunSummary = buildPartnerOpeningBalanceFiscalStartRepairDryRunSummary_(result);
  const failureSummary = buildPartnerOpeningBalanceFiscalStartRepairWrapperFailureMessage_(
    new Error(PARTNER_OPENING_BALANCE_FISCAL_START_REPAIR_COMPLETED_ERROR_), "完了済み補正の再実行禁止確認"
  );
  const chunks = createJsonLogChunks_("[G0056][BEFORE]", result.targets[0].before, 100);
  logPartnerOpeningBalanceFiscalStartRepairDryRunResult_(result);
  const cases = [
    ["debugラッパーは本体をdry-runで呼ぶ", debugRepairPartnerOpeningBalanceFiscalStart.toString().indexOf("dry_run: true") !== -1, true],
    ["debugラッパーは結果をreturnする", debugRepairPartnerOpeningBalanceFiscalStart.toString().indexOf("return result") !== -1, true],
    ["executeラッパーは本実行を禁止する", executeRepairPartnerOpeningBalanceFiscalStart.toString().indexOf("dry_run: false") !== -1, true],
    ["dry-run要約に対象件数", dryRunSummary.indexOf("対象件数: 2") !== -1, true],
    ["dry-run要約に変更予定", dryRunSummary.indexOf("grant_days: 15 → 0") !== -1 && dryRunSummary.indexOf("G0057") !== -1, true],
    ["失敗要約にエラー・段階・対象を含む", failureSummary.indexOf(PARTNER_OPENING_BALANCE_FISCAL_START_REPAIR_COMPLETED_ERROR_) !== -1 && failureSummary.indexOf("完了済み補正の再実行禁止確認") !== -1 && failureSummary.indexOf("G0056") !== -1, true],
    ["詳細JSONは分割形式", chunks[0].indexOf("[1/") !== -1, true],
    ["ログ整形は戻り値を変更しない", JSON.stringify(result), original],
    ["executeラッパーは本体の例外を再throwする", executeRepairPartnerOpeningBalanceFiscalStart.toString().indexOf("throw error") !== -1, true]
  ];
  const failed = cases.filter(item => JSON.stringify(item[1]) !== JSON.stringify(item[2]));
  if (failed.length) throw new Error("PARTNER年度開始補正ラッパーテスト失敗: " + JSON.stringify(failed));
  return { ok: true, case_count: cases.length };
}
