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
   時間単位年休 Phase 1 基盤テスト
   Spreadsheet・Supabase・LockServiceを使用しない純粋関数テスト
========================= */
function testTimeLeaveBackendFoundation() {
  const mainPolicy = getCompanyLeavePolicy("MAIN");
  const partnerPolicy = getCompanyLeavePolicy("PARTNER");
  const expectError = function(callback) {
    try {
      callback();
      return false;
    } catch (error) {
      return true;
    }
  };
  const cases = [
    {
      name: "MAIN制度は時間有給有効・420分・2100分",
      actual: [mainPolicy.timeLeaveEnabled, mainPolicy.scheduledMinutesPerDay, mainPolicy.timeLeaveAnnualLimitMinutes],
      expected: [true, 420, 2100]
    },
    {
      name: "PARTNER制度は時間有給無効",
      actual: partnerPolicy.timeLeaveEnabled,
      expected: false
    },
    {
      name: "HH:mmを分へ変換",
      actual: [parseTimeToMinute("08:30"), parseTimeToMinute("00:00"), parseTimeToMinute("23:59")],
      expected: [510, 0, 1439]
    },
    {
      name: "不正時刻を拒否",
      actual: [
        expectError(() => parseTimeToMinute("8:30")),
        expectError(() => parseTimeToMinute("24:00")),
        expectError(() => parseTimeToMinute("08:60"))
      ],
      expected: [true, true, true]
    },
    {
      name: "休憩控除: 08:00-09:00は60分",
      actual: calculateTimeLeaveMinutes(480, 540, mainPolicy),
      expected: 60
    },
    {
      name: "休憩控除: 08:00-10:00は120分",
      actual: calculateTimeLeaveMinutes(480, 600, mainPolicy),
      expected: 120
    },
    {
      name: "休憩控除: 08:00-11:00は150分",
      actual: calculateTimeLeaveMinutes(480, 660, mainPolicy),
      expected: 150
    },
    {
      name: "休憩控除: 08:00-10:30は120分",
      actual: calculateTimeLeaveMinutes(480, 630, mainPolicy),
      expected: 120
    },
    {
      name: "休憩のみ: 12:00-13:00は0分",
      actual: calculateTimeLeaveMinutes(720, 780, mainPolicy),
      expected: 0
    },
    {
      name: "休憩控除: 13:00-15:00は120分",
      actual: calculateTimeLeaveMinutes(780, 900, mainPolicy),
      expected: 120
    },
    {
      name: "勤務時間外と終了時刻不正を拒否",
      actual: [
        expectError(() => calculateTimeLeaveMinutes(450, 540, mainPolicy)),
        expectError(() => calculateTimeLeaveMinutes(540, 540, mainPolicy))
      ],
      expected: [true, true]
    },
    {
      name: "60分単位は60と120を許可、150と0を拒否",
      actual: [
        !expectError(() => validateTimeLeaveUnitMinutes(60, mainPolicy)),
        !expectError(() => validateTimeLeaveUnitMinutes(120, mainPolicy)),
        expectError(() => validateTimeLeaveUnitMinutes(150, mainPolicy)),
        expectError(() => validateTimeLeaveUnitMinutes(0, mainPolicy))
      ],
      expected: [true, true, true, true]
    },
    {
      name: "v1は休憩控除後の分数を維持",
      actual: calculateTimeLeaveRequestedMinutes_(540, 660, mainPolicy, TIME_LEAVE_CALCULATION_VERSION_V1),
      expected: 90
    },
    {
      name: "v2は休憩を含む時計時間を消費分にする",
      actual: [
        calculateTimeLeaveRequestedMinutes_(540, 660, mainPolicy, TIME_LEAVE_CALCULATION_VERSION_V2),
        calculateTimeLeaveRequestedMinutes_(840, 1020, mainPolicy, TIME_LEAVE_CALCULATION_VERSION_V2)
      ],
      expected: [120, 180]
    },
    {
      name: "新規単独時間年休は180分まで許可し240分を拒否",
      actual: [
        !expectError(() => validateNewStandaloneTimeLeaveMaximum_({ requested_minutes: 180 })),
        expectError(() => validateNewStandaloneTimeLeaveMaximum_({ requested_minutes: 240 }))
      ],
      expected: [true, true]
    },
    {
      name: "時間帯の境界接続は非重複、重なりは重複",
      actual: [hasTimeOverlap(480, 540, 540, 600), hasTimeOverlap(480, 540, 510, 570)],
      expected: [false, true]
    },
    {
      name: "日次420分ちょうどを許可",
      actual: validateDailyPaidLeaveMinutes([
        { kind: "half_day", minutes: 210 },
        { kind: "time_hourly", minutes: 210 }
      ], 420).remainingMinutes,
      expected: 0
    },
    {
      name: "日次420分超過を拒否",
      actual: expectError(() => validateDailyPaidLeaveMinutes([
        { kind: "full_day", minutes: 420 },
        { kind: "time_hourly", minutes: 60 }
      ], 420)),
      expected: true
    },
    {
      name: "年間2100分ちょうどを許可",
      actual: validateAnnualTimeLeaveLimit(1980, 60, 60, 2100).remainingMinutes,
      expected: 0
    },
    {
      name: "年間2100分超過を拒否",
      actual: expectError(() => validateAnnualTimeLeaveLimit(2040, 0, 120, 2100)),
      expected: true
    },
    {
      name: "pending仮押さえを年間上限へ含める",
      actual: expectError(() => validateAnnualTimeLeaveLimit(1980, 60, 120, 2100)),
      expected: true
    },
    {
      name: "年5日義務は既存1日・半日を算入し時間休を除外",
      actual: [
        getFiveDayObligationContribution_({ request_kind: "" }, 1),
        getFiveDayObligationContribution_({ request_kind: "" }, 0.5),
        getFiveDayObligationContribution_({ request_kind: "time_hourly" }, 1)
      ],
      expected: [1, 0.5, 0]
    }
  ];

  const results = cases.map(testCase => ({
    name: testCase.name,
    ok: JSON.stringify(testCase.actual) === JSON.stringify(testCase.expected),
    actual: testCase.actual,
    expected: testCase.expected
  }));
  const failed = results.filter(result => !result.ok);
  if (failed.length > 0) {
    throw new Error("時間単位年休 Phase 1 テスト失敗: " + JSON.stringify(failed));
  }

  Logger.log(JSON.stringify(results, null, 2));
  return { ok: true, case_count: results.length, results: results };
}

/* =========================
   申請中有給予約: 表示・通常申請・取消状態の固定データ検証
========================= */
function testTimeLeaveHistoryPresentationFoundation() {
  const combined = buildTimeLeaveHistoryPresentation_({
    start_time: new Date("2026-09-14T00:00:00Z"),
    end_time: new Date("2026-09-14T02:00:00Z"),
    requested_minutes: 120
  }, "pm");
  const standalone = buildTimeLeaveHistoryPresentation_({
    start_time: "09:00",
    end_time: "11:00",
    requested_minutes: 120
  }, "");
  const cases = [
    ["Date型の09:00をHH:mmへ正規化", formatTimeLeaveClockForPresentation_(new Date("2026-09-14T00:00:00Z"), "Asia/Tokyo"), "09:00"],
    ["Date型の11:00をHH:mmへ正規化", formatTimeLeaveClockForPresentation_(new Date("2026-09-14T02:00:00Z"), "Asia/Tokyo"), "11:00"],
    ["文字列09:00を維持", formatTimeLeaveClockForPresentation_("09:00"), "09:00"],
    ["表示用文字列9:00は従来どおり09:00", formatTimeLeaveClockForPresentation_("9:00"), "09:00"],
    ["combinedの半休区分", combined.leave_type_label, "PM半休＋時間年休"],
    ["combinedの時間帯", [combined.start_time, combined.end_time], ["09:00", "11:00"]],
    ["combinedの時間年休分数", combined.requested_minutes, 120],
    ["combinedの有給使用量", combined.total_paid_leave_minutes, 330],
    ["単独時間年休の表示種別", standalone.leave_type_label, "時間年休"],
    ["単独時間年休はcombinedでない", standalone.is_combined, false],
    ["330分を5時間30分表示", formatTimeLeaveMinutesForPresentation_(330), "5時間30分"],
    ["Date文字列を時刻として表示しない", formatTimeLeaveClockForPresentation_("Fri Dec 29 1899 19:00:00 GMT-0500"), ""]
  ];
  const failed = cases.filter(item => JSON.stringify(item[1]) !== JSON.stringify(item[2]));
  if (failed.length) throw new Error("時間年休履歴表示テスト失敗: " + JSON.stringify(failed));
  return { ok: true, case_count: cases.length };
}

function testAdminPendingTimeLeavePresentationFoundation() {
  const segment = { start_time: "09:00", end_time: "11:00", requested_minutes: 120 };
  const normal = buildPendingAdminTimeLeavePresentationRow_({
    request_id: "R-NORMAL", request_kind: "", half_day: "pm", days: 0.5
  }, null);
  const single = buildPendingAdminTimeLeavePresentationRow_({
    request_id: "R-TIME", request_kind: "time_hourly", half_day: "", days: 0
  }, segment);
  const combined = buildPendingAdminTimeLeavePresentationRow_({
    request_id: "R-COMBINED", request_kind: "half_day_time_hourly", half_day: "pm", days: 0.5
  }, segment);
  const supabaseFallback = buildPendingAdminTimeLeavePresentationRow_({
    request_id: "R-SUPABASE", request_kind: "", half_day: "pm", days: 0.5
  }, segment);
  const cases = [
    ["通常PM半休はrequest_kindを空欄のまま保持", normal.request_kind, ""],
    ["通常有給のpresentationはnull", normal.time_leave_presentation, null],
    ["単独時間年休はrequest_kindを保持", single.request_kind, "time_hourly"],
    ["単独時間年休のpresentationを付与", [
      single.time_leave_presentation.leave_type_label,
      single.time_leave_presentation.start_time,
      single.time_leave_presentation.end_time,
      single.time_leave_presentation.requested_minutes
    ], ["時間年休", "09:00", "11:00", 120]],
    ["combinedはrequest_kindを保持", combined.request_kind, "half_day_time_hourly"],
    ["combinedはPM・2時間・330分を返す", [
      combined.time_leave_presentation.leave_type_label,
      combined.time_leave_presentation.start_time,
      combined.time_leave_presentation.end_time,
      combined.time_leave_presentation.requested_minutes,
      combined.time_leave_presentation.total_paid_leave_minutes
    ], ["PM半休＋時間年休", "09:00", "11:00", 120, 330]],
    ["Supabaseのrequest_kind欠落時は既存子明細から表示用に復元", [
      supabaseFallback.request_kind,
      supabaseFallback.time_leave_presentation.is_combined
    ], ["half_day_time_hourly", true]]
  ];
  const failed = cases.filter(item => JSON.stringify(item[1]) !== JSON.stringify(item[2]));
  if (failed.length) throw new Error("管理者pending時間年休表示テスト失敗: " + JSON.stringify(failed));
  return { ok: true, case_count: cases.length };
}

function testPendingPaidLeaveReservationFoundation() {
  const employeeId = "MAIN-PENDING";
  const useDate = parseLocalDate("2026-09-14");
  const grant = {
    grant_id: "G-PENDING", grant_date: parseLocalDate("2026-04-01"),
    valid_from_date: parseLocalDate("2026-04-01"), valid_to_date: parseLocalDate("2028-03-31"),
    grant_days: 16, carry_over_days: 0, carry_over_minutes: 0, is_finalized: true
  };
  const context = requests => ({
    grants_by_employee: { [employeeId]: [grant] },
    requests_by_employee: { [employeeId]: requests },
    time_leave_segments_by_request: {
      "R-COMBINED": [{ time_leave_id: "T-COMBINED", leave_date: "2026-09-14", requested_minutes: 120 }]
    },
    company_code_by_employee: { [employeeId]: "MAIN" },
    // Nodeの固定データ実行でもローカル日付の差で休日判定に依存しないよう対象日を明示する。
    calendar_map: { "2026-09-13": "workday", "2026-09-14": "workday" }
  });
  const combined = {
    request_id: "R-COMBINED", employee_id: employeeId, start_date: "2026-09-14", end_date: "2026-09-14",
    days: 0.5, half_day: "pm", status: STATUS.PENDING, type: "paid_leave", request_kind: "half_day_time_hourly"
  };
  const halfDay = {
    request_id: "R-HALF", employee_id: employeeId, start_date: "2026-09-14", end_date: "2026-09-14",
    days: 0.5, half_day: "am", status: STATUS.PENDING, type: "paid_leave"
  };
  const fullDay = Object.assign({}, halfDay, { request_id: "R-FULL", days: 1, half_day: "" });
  const canceled = Object.assign({}, combined, { request_id: "R-CANCELED", status: STATUS.CANCELED });
  const rejected = Object.assign({}, fullDay, { request_id: "R-REJECTED", status: STATUS.REJECTED });
  const combinedContext = context([combined, canceled, rejected]);
  const combinedPending = getAllPendingPaidLeaveReservationMinutes_(employeeId, combinedContext);
  const combinedBalance = calculateFifoBalanceWithOpeningBalanceFromContext_(employeeId, useDate, combinedContext);
  const oneDayCombinedContext = Object.assign({}, combinedContext, {
    grants_by_employee: { [employeeId]: [Object.assign({}, grant, { grant_days: 1 })] }
  });
  const expectError = callback => { try { callback(); return false; } catch (error) { return true; } };
  const cases = [
    ["確定残16日は6720分", combinedBalance.current_remaining_minutes, 6720],
    ["pending combinedは210分＋120分で330分", combinedPending, 330],
    ["pending combined後の申請可能残は6390分", combinedBalance.current_remaining_minutes - combinedPending, 6390],
    ["取消・否認はpending予約に含めない", getAllPendingPaidLeaveReservationMinutes_(employeeId, context([canceled, rejected])), 0],
    ["通常半休pendingは210分", getAllPendingPaidLeaveReservationMinutes_(employeeId, context([halfDay])), 210],
    ["通常1日pendingは420分", getAllPendingPaidLeaveReservationMinutes_(employeeId, context([fullDay])), 420],
    ["pending330分を含めても15日申請は可能", !expectError(() =>
      validateMainPaidLeaveRequestBalanceReservationFromContext_(employeeId, useDate, useDate, 15, "", combinedContext)), true],
    ["pending330分を含み残90分の通常半休は登録前に拒否", expectError(() =>
      validateMainPaidLeaveRequestBalanceReservationFromContext_(employeeId, useDate, useDate, 0.5, "am", oneDayCombinedContext)), true]
  ];
  const failed = cases.filter(item => JSON.stringify(item[1]) !== JSON.stringify(item[2]));
  if (failed.length) throw new Error("申請中有給予約テスト失敗: " + JSON.stringify(failed));
  return { ok: true, case_count: cases.length };
}

/* =========================
   時間単位年休 Phase 2 ライフサイクル検証テスト
   永続化は行わず、登録・編集・承認で共通利用する判定関数を検証する。
========================= */
function testTimeLeaveLifecycleValidationFoundation() {
  const policy = getCompanyLeavePolicy("MAIN");
  const candidate = {
    employee_id: "TEST-MAIN",
    leave_date: parseLocalDate("2026-04-01"),
    start_minute: 480,
    end_minute: 540,
    requested_minutes: 60,
    policy: policy
  };
  const expectError = function(callback) {
    try {
      callback();
      return false;
    } catch (error) {
      return true;
    }
  };
  const cases = [
    {
      name: "境界接触する時間休は許可",
      actual: !expectError(() => validateTimeLeaveEntriesAgainstCandidate_(candidate, [
        { kind: "time_hourly", minutes: 60, start_minute: 540, end_minute: 600 }
      ])),
      expected: true
    },
    {
      name: "重複する時間休は拒否",
      actual: expectError(() => validateTimeLeaveEntriesAgainstCandidate_(candidate, [
        { kind: "time_hourly", minutes: 60, start_minute: 510, end_minute: 570 }
      ])),
      expected: true
    },
    {
      name: "1日休との併用は拒否",
      actual: expectError(() => validateTimeLeaveEntriesAgainstCandidate_(candidate, [
        { kind: "full_day", minutes: 420 }
      ])),
      expected: true
    },
    {
      name: "AM半休との重複は拒否",
      actual: expectError(() => validateTimeLeaveEntriesAgainstCandidate_(candidate, [
        { kind: "half_day", half_day: "am", minutes: 210 }
      ])),
      expected: true
    },
    {
      name: "PM半休と午前時間休は許可",
      actual: !expectError(() => validateTimeLeaveEntriesAgainstCandidate_(candidate, [
        { kind: "half_day", half_day: "pm", minutes: 210 }
      ])),
      expected: true
    },
    {
      name: "canceled申請を除外した後の同時間帯は許可",
      actual: !expectError(() => validateTimeLeaveEntriesAgainstCandidate_(candidate, [])),
      expected: true
    },
    {
      name: "半日210分と時間休210分で日次420分ちょうどを許可",
      actual: !expectError(() => validateTimeLeaveEntriesAgainstCandidate_(Object.assign({}, candidate, {
        start_minute: 780, end_minute: 990, requested_minutes: 210
      }), [{ kind: "half_day", half_day: "am", minutes: 210 }])),
      expected: true
    },
    {
      name: "既存360分と時間休120分で日次420分超過を拒否",
      actual: expectError(() => validateTimeLeaveEntriesAgainstCandidate_(Object.assign({}, candidate, {
        start_minute: 900, end_minute: 1020, requested_minutes: 120
      }), [{ kind: "time_hourly", minutes: 360, start_minute: 480, end_minute: 840 }])),
      expected: true
    },
    {
      name: "編集時に自身を除外すれば同じ時間帯を維持できる",
      actual: !expectError(() => validateTimeLeaveEntriesAgainstCandidate_(candidate, [])),
      expected: true
    },
    {
      name: "編集後に年間上限超過なら拒否",
      actual: expectError(() => validateAnnualTimeLeaveLimit(1980, 60, 120, 2100)),
      expected: true
    },
    {
      name: "時間休明細は親statusを持たない",
      actual: TIME_LEAVE_SEGMENTS_HEADERS.indexOf("status") === -1,
      expected: true
    }
  ];
  const results = cases.map(testCase => ({
    name: testCase.name,
    ok: JSON.stringify(testCase.actual) === JSON.stringify(testCase.expected),
    actual: testCase.actual,
    expected: testCase.expected
  }));
  const failed = results.filter(result => !result.ok);
  if (failed.length > 0) {
    throw new Error("時間単位年休 Phase 2 テスト失敗: " + JSON.stringify(failed));
  }
  Logger.log(JSON.stringify(results, null, 2));
  return { ok: true, case_count: results.length, results: results };
}

/* =========================
   時間単位年休 Phase 3: 分単位 FIFO・繰越テスト
   シート/API を使わない固定データ検証
========================= */
function testTimeLeaveFifoMinutesFoundation() {
  const d = value => parseLocalDate(value);
  const grant = (id, date, days, validTo, carryDays, carryMinutes) => ({
    grant_id: id,
    grant_date: d(date),
    valid_from_date: d(date),
    valid_to_date: d(validTo || "2028-03-31"),
    grant_days: days,
    carry_over_days: carryDays || 0,
    carry_over_minutes: carryMinutes || 0,
    is_finalized: true
  });
  const request = (id, date, days, halfDay, status, kind) => ({
    request_id: id, start_date: date, end_date: date, days: days,
    half_day: halfDay || "", status: status || "approved", type: "paid_leave",
    request_kind: kind || ""
  });
  const context = (grants, requests, segments) => ({
    grants_by_employee: { M1: grants },
    requests_by_employee: { M1: requests || [] },
    time_leave_segments_by_request: segments || {},
    company_code_by_employee: { M1: "MAIN" },
    calendar_map: {}
  });
  const fifo = (grants, requests, segments, asOf) =>
    calculateFifoBalanceMinutesFromContext_("M1", d(asOf || "2026-04-10"), context(grants, requests, segments));
  const oneLot = grant("G1", "2026-04-01", 10);
  const half = fifo([oneLot], [request("R-half", "2026-04-02", 0.5, "am")]);
  const full = fifo([oneLot], [request("R-full", "2026-04-02", 1)]);
  const hourly = fifo([oneLot], [request("R-time", "2026-04-02", 0, "", "approved", "time_hourly")], {
    "R-time": [{ time_leave_id: "T1", leave_date: "2026-04-02", requested_minutes: 60 }]
  });
  const combined = fifo([oneLot], [request("R-combined", "2026-04-02", 0.5, "am", "approved", "half_day_time_hourly")], {
    "R-combined": [{ time_leave_id: "TC1", leave_date: "2026-04-02", requested_minutes: 60 }]
  });
  const split = fifo([
    grant("OLD", "2025-04-01", 10), grant("NEW", "2026-04-01", 10)
  ], [request("R-split", "2026-04-02", 0, "", "approved", "time_hourly")], {
    "R-split": [{ time_leave_id: "T2", leave_date: "2026-04-02", requested_minutes: 120 }]
  }, "2026-04-10");
  // 旧ロットの残を60分にしてロット跨ぎを固定的に作る。
  const splitLots = fifo([
    grant("OLD", "2025-04-01", 1), grant("NEW", "2026-04-01", 10)
  ], [
    request("R-old", "2026-04-01", 0, "", "approved", "time_hourly"),
    request("R-split", "2026-04-02", 0, "", "approved", "time_hourly")
  ], {
    "R-old": [{ time_leave_id: "T-old", leave_date: "2026-04-01", requested_minutes: 360 }],
    "R-split": [{ time_leave_id: "T2", leave_date: "2026-04-02", requested_minutes: 120 }]
  });
  const expiryLastDay = fifo([grant("EXP", "2024-10-01", 1, "2026-04-02")], [request("R-exp", "2026-04-02", 1)], {}, "2026-04-02");
  const expiryAfter = fifo([grant("EXP", "2024-10-01", 1, "2026-04-02")], [request("R-exp", "2026-04-03", 1)], {}, "2026-04-03");
  const mixed = fifo([grant("MIX", "2026-04-01", 10)], [
    request("R1", "2026-04-01", 1), request("R2", "2026-04-02", 0.5, "pm"),
    request("R3", "2026-04-03", 0, "", "approved", "time_hourly"),
    request("CANCEL", "2026-04-04", 0, "", "canceled", "time_hourly"),
    request("PENDING", "2026-04-04", 0, "", "pending", "time_hourly")
  ], {
    R3: [{ time_leave_id: "T3", leave_date: "2026-04-03", requested_minutes: 120 }],
    CANCEL: [{ time_leave_id: "T4", leave_date: "2026-04-04", requested_minutes: 60 }],
    PENDING: [{ time_leave_id: "T5", leave_date: "2026-04-04", requested_minutes: 60 }]
  });
  const carryFiveDaysThreeHours = calculateCarryOverMinutes_(5 * 420 + 180, 420);
  const carryCap = calculateCarryOverMinutes_(20 * 420 + 180, 420);
  const legacyCarry = getGrantCarryOverMinutes_({ carry_over_days: 1.5, carry_over_minutes: "" }, 420);
  const newCarry = getGrantCarryOverMinutes_({ carry_over_days: 5, carry_over_minutes: 180 }, 420);
  const mixedCarry = getGrantCarryOverMinutes_({ carry_over_days: 5.5, carry_over_minutes: 180 }, 420);
  const display480 = getMinuteBalanceDisplay_(480, 420);
  const expectError = callback => {
    try { callback(); return false; } catch (error) { return true; }
  };
  const cases = [
    ["10日付与は4200分", fifo([oneLot], []).total_granted_minutes, 4200],
    ["半日消化は210分", half.used_minutes, 210],
    ["1日消化は420分", full.used_minutes, 420],
    ["時間休60分を消化", hourly.used_minutes, 60],
    ["半休＋時間休は270分を消化", combined.used_minutes, 270],
    ["古いロット優先", split.allocations[0].grant_id, "OLD"],
    ["ロット跨ぎは2配賦", splitLots.allocations.filter(row => row.request_id === "R-split").length, 2],
    ["ロット跨ぎ旧60分", splitLots.allocations.filter(row => row.request_id === "R-split")[0].consumed_minutes, 60],
    ["ロット跨ぎ新60分", splitLots.allocations.filter(row => row.request_id === "R-split")[1].consumed_minutes, 60],
    ["期限最終日は使用可能", expiryLastDay.unallocated_used_minutes, 0],
    ["期限翌日は配賦しない", expiryAfter.unallocated_used_minutes, 420],
    ["混在消化は750分", mixed.used_minutes, 750],
    ["取消・pending時間休は正式FIFOに含めない", mixed.current_remaining_minutes, 3450],
    ["480分は1日1時間", [display480.remaining_full_days, display480.remaining_hours], [1, 1]],
    ["300分は5時間", getMinuteBalanceDisplay_(300, 420).remaining_hours, 5],
    ["0分残高", getMinuteBalanceDisplay_(0, 420).remaining_minutes, 0],
    ["5日3時間を正確に繰越", [carryFiveDaysThreeHours.carry_over_days, carryFiveDaysThreeHours.carry_over_minutes], [5, 180]],
    ["旧繰越1.5日は630分", legacyCarry, 630],
    ["新繰越5日180分は2280分", newCarry, 2280],
    ["旧小数繰越と分列を二重計上しない", mixedCarry, 2310],
    ["20日上限は8400分", carryCap.carry_over_candidate_minutes, 8400],
    ["上限超過分180分を繰越しない", carryCap.carry_over_limit_expired_minutes, 180],
    ["正式残300分・pending120分・新規180分は許可", !expectError(() => validateTimeLeaveFifoReservation_(300, 120, 180)), true],
    ["正式残300分・pending120分・新規240分は拒否", expectError(() => validateTimeLeaveFifoReservation_(300, 120, 240)), true],
    ["PARTNERは時間休無効", getCompanyLeavePolicy("PARTNER").timeLeaveEnabled, false]
  ];
  const results = cases.map(item => ({ name: item[0], actual: item[1], expected: item[2], ok: JSON.stringify(item[1]) === JSON.stringify(item[2]) }));
  const failed = results.filter(item => !item.ok);
  if (failed.length) throw new Error("時間単位年休 Phase 3 テスト失敗: " + JSON.stringify(failed));
  Logger.log(JSON.stringify(results, null, 2));
  return { ok: true, case_count: results.length, results: results };
}

/* =========================
   半休＋時間単位年休 Phase 1: 複合申請の固定データ検証
   Spreadsheet を操作せず、複合種別の分数・競合・上限判定を確認する。
========================= */
function testCombinedHalfDayTimeLeaveFoundation() {
  const policy = getCompanyLeavePolicy("MAIN");
  const employeeInfo = { companyCode: "MAIN", policy: policy };
  const expectError = callback => {
    try { callback(); return false; } catch (error) { return true; }
  };
  const candidate = normalizeCombinedHalfDayTimeLeavePayload_({
    employee_id: "M-COMBINED",
    leave_date: "2026-09-15",
    half_day: "am",
    start_time: "13:00",
    end_time: "15:30",
    reason: "private",
    reason_detail: ""
  }, employeeInfo);
  const requiredMinutesContext = {
    time_leave_segments_by_request: {
      "R-COMBINED": [{ time_leave_id: "TC1", leave_date: "2026-09-15", requested_minutes: 120 }]
    },
    calendar_map: {}
  };
  const combinedParent = {
    request_id: "R-COMBINED", start_date: "2026-09-15", end_date: "2026-09-15",
    days: 0.5, half_day: "am", status: "pending", type: "paid_leave",
    request_kind: "half_day_time_hourly"
  };
  const storedTimeParent = {
    employee_id: "M-COMBINED", reason: "private", reason_detail: "", request_kind: "time_hourly"
  };
  const storedV1Segment = {
    leave_date: "2026-09-15", start_time: "09:00", end_time: "11:30",
    requested_minutes: 120, calculation_version: TIME_LEAVE_CALCULATION_VERSION_V1,
    scheduled_minutes_per_day: 420, time_leave_unit_minutes: 60,
    work_start_minute: 480, work_end_minute: 1020,
    break_periods_json: JSON.stringify(policy.breakPeriods)
  };
  const storedV2Segment = {
    leave_date: "2026-09-15", start_time: "09:00", end_time: "11:00",
    requested_minutes: 120, calculation_version: TIME_LEAVE_CALCULATION_VERSION_V2,
    scheduled_minutes_per_day: 420, time_leave_unit_minutes: 60,
    work_start_minute: 480, work_end_minute: 1020,
    break_periods_json: JSON.stringify(policy.breakPeriods)
  };
  const storedV2DateSegment = Object.assign({}, storedV2Segment, {
    start_time: new Date("2026-09-15T00:00:00Z"),
    end_time: new Date("2026-09-15T02:00:00Z")
  });
  const storedV1DateSegment = Object.assign({}, storedV1Segment, {
    start_time: new Date("2026-09-15T00:00:00Z"),
    end_time: new Date("2026-09-15T02:30:00Z")
  });
  const storedV2LongSegment = Object.assign({}, storedV2Segment, {
    end_time: "13:00", requested_minutes: 240
  });
  const cases = [
    ["v1編集はv1のまま休憩控除方式を維持", (() => {
      const item = normalizeTimeLeavePayload_({
        employee_id: "M-COMBINED", leave_date: "2026-09-15", start_time: "09:00", end_time: "11:30"
      }, employeeInfo, TIME_LEAVE_CALCULATION_VERSION_V1);
      return [item.calculation_version, item.requested_minutes];
    })(), [TIME_LEAVE_CALCULATION_VERSION_V1, 120]],
    ["v2編集はv2のまま時計時間方式を維持", (() => {
      const item = normalizeTimeLeavePayload_({
        employee_id: "M-COMBINED", leave_date: "2026-09-15", start_time: "09:00", end_time: "11:00"
      }, employeeInfo, TIME_LEAVE_CALCULATION_VERSION_V2);
      return [item.calculation_version, item.requested_minutes];
    })(), [TIME_LEAVE_CALCULATION_VERSION_V2, 120]],
    ["v1承認時再検証は成功", (() => {
      const item = normalizeStoredTimeLeaveSegmentForValidation_(storedV1Segment, storedTimeParent, employeeInfo);
      return [item.calculation_version, item.requested_minutes];
    })(), [TIME_LEAVE_CALCULATION_VERSION_V1, 120]],
    ["v2承認時再検証は成功", (() => {
      const item = normalizeStoredTimeLeaveSegmentForValidation_(storedV2Segment, storedTimeParent, employeeInfo);
      return [item.calculation_version, item.requested_minutes];
    })(), [TIME_LEAVE_CALCULATION_VERSION_V2, 120]],
    ["Date型v2明細は承認時に09:00-11:00として再検証", (() => {
      const item = normalizeStoredTimeLeaveSegmentForValidation_(storedV2DateSegment, storedTimeParent, employeeInfo);
      return [item.start_time, item.end_time, item.requested_minutes, item.calculation_version];
    })(), ["09:00", "11:00", 120, TIME_LEAVE_CALCULATION_VERSION_V2]],
    ["Date型v1明細は既存の休憩控除方式で再検証", (() => {
      const item = normalizeStoredTimeLeaveSegmentForValidation_(storedV1DateSegment, storedTimeParent, employeeInfo);
      return [item.start_time, item.end_time, item.requested_minutes, item.calculation_version];
    })(), ["09:00", "11:30", 120, TIME_LEAVE_CALCULATION_VERSION_V1]],
    ["Date型combined明細はPM半休との境界接触を許可", (() => {
      const item = normalizeStoredTimeLeaveSegmentForValidation_(
        storedV2DateSegment,
        Object.assign({}, storedTimeParent, { request_kind: "half_day_time_hourly", half_day: "pm" }),
        employeeInfo
      );
      validateTimeLeaveEntriesAgainstCandidate_(item, [{ kind: "half_day", half_day: "pm", minutes: 210 }]);
      return item.half_day_minutes + item.requested_minutes;
    })(), 330],
    ["計算用時刻正規化は9:00を従来どおり拒否", expectError(() =>
      normalizeTimeLeaveClockValue_("9:00", "Asia/Tokyo")
    ), true],
    ["計算用時刻正規化は不正値を拒否", expectError(() =>
      normalizeTimeLeaveClockValue_("invalid", "Asia/Tokyo")
    ), true],
    ["既存v2の4時間明細は承認時再検証の対象外上限として維持", (() => {
      const item = normalizeStoredTimeLeaveSegmentForValidation_(storedV2LongSegment, storedTimeParent, employeeInfo);
      return [item.calculation_version, item.requested_minutes];
    })(), [TIME_LEAVE_CALCULATION_VERSION_V2, 240]],
    ["AM半休＋午後1時間は270分", (() => {
      const item = normalizeCombinedHalfDayTimeLeavePayload_({
        employee_id: "M-COMBINED", leave_date: "2026-09-15", half_day: "am",
        start_time: "13:00", end_time: "14:00"
      }, employeeInfo);
      return item.half_day_minutes + item.requested_minutes;
    })(), 270],
    ["AM半休＋13:00-15:30は330分", candidate.half_day_minutes + candidate.requested_minutes, 330],
    ["PM半休＋09:30-11:00は270分", (() => {
      const item = normalizeCombinedHalfDayTimeLeavePayload_({
        employee_id: "M-COMBINED", leave_date: "2026-09-15", half_day: "pm",
        start_time: "09:30", end_time: "11:00"
      }, employeeInfo);
      return item.half_day_minutes + item.requested_minutes;
    })(), 270],
    ["AM半休と午前時間休の重複を拒否", expectError(() => {
      const item = normalizeCombinedHalfDayTimeLeavePayload_({
        employee_id: "M-COMBINED", leave_date: "2026-09-15", half_day: "am",
        start_time: "09:00", end_time: "10:00"
      }, employeeInfo);
      validateTimeLeaveEntriesAgainstCandidate_(item, [{ kind: "half_day", half_day: "am", minutes: 210 }]);
    }), true],
    ["PM半休と午後時間休の重複を拒否", expectError(() => {
      const item = normalizeCombinedHalfDayTimeLeavePayload_({
        employee_id: "M-COMBINED", leave_date: "2026-09-15", half_day: "pm",
        start_time: "13:00", end_time: "14:00"
      }, employeeInfo);
      validateTimeLeaveEntriesAgainstCandidate_(item, [{ kind: "half_day", half_day: "pm", minutes: 210 }]);
    }), true],
    ["時間年休150分を拒否", expectError(() => normalizeCombinedHalfDayTimeLeavePayload_({
      employee_id: "M-COMBINED", leave_date: "2026-09-15", half_day: "am",
      start_time: "13:00", end_time: "15:30"
    }, Object.assign({}, employeeInfo, { policy: Object.assign({}, policy, { breakPeriods: [] }) }))), true],
    ["半休＋時間休が420分超過なら拒否", expectError(() => validateTimeLeaveEntriesAgainstCandidate_(
      Object.assign({}, candidate, { requested_minutes: 240 }),
      [{ kind: "half_day", half_day: "am", minutes: 210 }]
    )), true],
    ["年間上限は時間年休分だけで拒否", expectError(() => validateAnnualTimeLeaveLimit(2040, 0, candidate.requested_minutes, 2100)), true],
    ["FIFO残高270分で複合270分を許可", !expectError(() => validateTimeLeaveFifoReservation_(270, 0, 270)), true],
    ["FIFO残高269分で複合270分を拒否", expectError(() => validateTimeLeaveFifoReservation_(269, 0, 270)), true],
    ["既存pending時間休との重複を拒否", expectError(() => validateTimeLeaveEntriesAgainstCandidate_(candidate, [
      { kind: "time_hourly", minutes: 60, start_minute: 780, end_minute: 840 }
    ].concat([{ kind: "half_day", half_day: "am", minutes: 210 }]))), true],
    ["複合親の承認必要分は330分", getMainApprovalRequiredMinutes_(combinedParent, requiredMinutesContext), 330],
    ["複合親は年5日義務へ半日だけ算入", getFiveDayObligationContribution_(combinedParent, 0.5), 0.5]
  ];
  const results = cases.map(item => ({ name: item[0], actual: item[1], expected: item[2], ok: JSON.stringify(item[1]) === JSON.stringify(item[2]) }));
  const failed = results.filter(item => !item.ok);
  if (failed.length) throw new Error("半休＋時間単位年休 Phase 1 テスト失敗: " + JSON.stringify(failed));
  Logger.log(JSON.stringify(results, null, 2));
  return { ok: true, case_count: results.length, results: results };
}

/* =========================
   時間単位年休 Phase 2: 候補生成の固定データ検証
   Spreadsheet を操作せず、候補の時刻計算・上限・競合除外を確認する。
========================= */
function testTimeLeaveCandidatesFoundation() {
  const policy = getCompanyLeavePolicy("MAIN");
  const findCandidate = (rows, startTime, requestedMinutes) =>
    rows.find(row => row.start_time === startTime && row.requested_minutes === requestedMinutes);
  const hasStart = (rows, startTime) => rows.some(row => row.start_time === startTime);
  const startTimes = rows => Array.from(new Set(rows.map(row => row.start_time)));
  const build = options => buildTimeLeaveCandidateRows_(Object.assign({
    policy: policy,
    request_mode: "time_hourly",
    half_day: "",
    entries: [],
    annual: { approvedMinutes: 0, pendingMinutes: 0 },
    approved_remaining_minutes: 4200,
    pending_reserved_minutes: 0
  }, options || {}));
  const single = build();
  const combinedAm = build({ request_mode: "half_day_time_hourly", half_day: "am" });
  const combinedPm = build({ request_mode: "half_day_time_hourly", half_day: "pm" });
  const annualExhausted = build({ annual: { approvedMinutes: 2100, pendingMinutes: 0 } });
  const combinedFifoShort = build({
    request_mode: "half_day_time_hourly",
    half_day: "am",
    approved_remaining_minutes: 269
  });
  const conflict = build({
    entries: [{ kind: "time_hourly", minutes: 60, start_minute: 780, end_minute: 840 }]
  });
  const cases = [
    ["08:00 + 1h は09:00", findCandidate(single, "08:00", 60).end_time, "09:00"],
    ["v2 09:00 + 1h は10:00", findCandidate(single, "09:00", 60).end_time, "10:00"],
    ["v2 09:00 + 2h は11:00", findCandidate(single, "09:00", 120).end_time, "11:00"],
    ["v2 11:00 + 3h は14:00", findCandidate(single, "11:00", 180).end_time, "14:00"],
    ["v2 13:00 + 3h は16:00", findCandidate(single, "13:00", 180).end_time, "16:00"],
    ["単独09:00の4時間は候補外", !!findCandidate(single, "09:00", 240), false],
    ["標準勤務の単独14:00は勤務終了17:00までの1時間・2時間・3時間", (single.filter(row => row.start_time === "14:00").map(row => row.requested_minutes)), [60, 120, 180]],
    ["単独時間年休は4時間以上を返さない", single.some(row => row.requested_minutes > 180), false],
    ["17:00を超える候補を返さない", !!findCandidate(single, "16:00", 120), false],
    ["v1 09:00 + 2h は11:30", calculateTimeLeaveEndMinute_(540, 120, policy, TIME_LEAVE_CALCULATION_VERSION_V1), 690],
    ["休憩中の10:00開始は候補外", hasStart(single, "10:00"), false],
    ["休憩中の15:00開始は候補外", hasStart(single, "15:00"), false],
    ["有効な正時開始候補を返す", startTimes(single), ["08:00", "09:00", "11:00", "13:00", "14:00", "16:00"]],
    ["08:30開始は候補外", hasStart(single, "08:30"), false],
    ["09:30開始は候補外", hasStart(single, "09:30"), false],
    ["13:30開始は候補外", hasStart(single, "13:30"), false],
    ["14:30開始は候補外", hasStart(single, "14:30"), false],
    ["AM半休＋13:00 + 1h は有効", !!findCandidate(combinedAm, "13:00", 60), true],
    ["AM半休＋13:00 + 2h は有効", !!findCandidate(combinedAm, "13:00", 120), true],
    ["AM半休＋13:00 + 3h は既存の日次上限内で有効", !!findCandidate(combinedAm, "13:00", 180), true],
    ["半休＋時間休の4時間は合計420分超過で候補外", !!findCandidate(combinedAm, "13:00", 240), false],
    ["PM半休＋09:00 + 1h は有効", !!findCandidate(combinedPm, "09:00", 60), true],
    ["AM半休では午前候補を返さない", hasStart(combinedAm, "09:00"), false],
    ["PM半休では午後候補を返さない", hasStart(combinedPm, "13:00"), false],
    ["半休＋時間年休でも30分開始を返さない", combinedAm.concat(combinedPm).some(row => row.start_time.endsWith(":30")), false],
    ["年間上限済みなら候補を返さない", annualExhausted.length, 0],
    ["FIFO 269分では複合の最小270分を返さない", combinedFifoShort.length, 0],
    ["既存時間申請と重なる13:00-14:00を返さない", !!findCandidate(conflict, "13:00", 60), false]
  ];
  const results = cases.map(item => ({ name: item[0], actual: item[1], expected: item[2], ok: JSON.stringify(item[1]) === JSON.stringify(item[2]) }));
  const failed = results.filter(item => !item.ok);
  if (failed.length) throw new Error("時間単位年休 Phase 2 候補生成テスト失敗: " + JSON.stringify(failed));
  Logger.log(JSON.stringify(results, null, 2));
  return { ok: true, case_count: results.length, results: results };
}

/* =========================
   社員別勤務時間: 空欄fallback・候補・半休・snapshot検証
========================= */
function testEmployeeSpecificTimeLeavePolicyFoundation() {
  const standardEmployee = { employee_id: "M-STANDARD", company_code: "MAIN", work_start_minute: "", work_end_minute: "" };
  const earlyEmployee = { employee_id: "M-EARLY", company_code: "MAIN", work_start_minute: 420, work_end_minute: 960 };
  const standardPolicy = resolveEmployeeTimeLeavePolicy_(standardEmployee.employee_id, standardEmployee);
  const earlyPolicy = resolveEmployeeTimeLeavePolicy_(earlyEmployee.employee_id, earlyEmployee);
  const build = policy => buildTimeLeaveCandidateRows_({
    policy: policy, request_mode: "time_hourly", half_day: "", entries: [],
    annual: { approvedMinutes: 0, pendingMinutes: 0 }, approved_remaining_minutes: 4200,
    pending_reserved_minutes: 0, calculation_version: TIME_LEAVE_CALCULATION_VERSION_V2
  });
  const earlyCandidates = build(earlyPolicy);
  const hasStart = (rows, startTime) => rows.some(row => row.start_time === startTime);
  const findCandidate = (rows, startTime, minutes) => rows.find(row =>
    row.start_time === startTime && row.requested_minutes === minutes
  );
  const amRange = getHalfDayOccupiedRange_("am", earlyPolicy);
  const pmRange = getHalfDayOccupiedRange_("pm", earlyPolicy);
  const standardAmRange = getHalfDayOccupiedRange_("am", standardPolicy);
  const standardPmRange = getHalfDayOccupiedRange_("pm", standardPolicy);
  const snapshot = {
    scheduled_minutes_per_day: 420, time_leave_unit_minutes: 60,
    work_start_minute: 420, work_end_minute: 960,
    break_periods_json: JSON.stringify(earlyPolicy.breakPeriods)
  };
  const snapshotPolicy = resolveTimeLeavePolicyFromSegmentSnapshot_(
    { companyCode: "MAIN", policy: standardPolicy }, snapshot
  );
  const segment = buildTimeLeaveSegmentRow_(
    { headers: TIME_LEAVE_SEGMENTS_HEADERS }, "R-EARLY", {
      employee_id: "M-EARLY", company_code: "MAIN", leave_date: parseLocalDate("2026-09-15"),
      start_time: "14:00", end_time: "16:00", start_minute: 840, end_minute: 960,
      requested_minutes: 120, calculation_version: TIME_LEAVE_CALCULATION_VERSION_V2,
      policy: earlyPolicy
    }, new Date("2026-09-01T00:00:00")
  );
  const expectError = callback => {
    try { callback(); return false; } catch (error) { return true; }
  };
  const cases = [
    ["空欄は標準08:00-17:00へfallback", [standardPolicy.workStartMinute, standardPolicy.workEndMinute], [480, 1020]],
    ["標準AM/PM半休は実労働210分", [
      calculateTimeLeaveMinutes(standardAmRange.startMinute, standardAmRange.endMinute, standardPolicy),
      calculateTimeLeaveMinutes(standardPmRange.startMinute, standardPmRange.endMinute, standardPolicy)
    ], [210, 210]],
    ["早出社員は07:00-16:00", [earlyPolicy.workStartMinute, earlyPolicy.workEndMinute], [420, 960]],
    ["早出社員は07:00を開始候補に含む", hasStart(earlyCandidates, "07:00"), true],
    ["早出社員は16:00を開始候補に含まない", hasStart(earlyCandidates, "16:00"), false],
    ["早出社員の10:00/12:00/15:00は候補外", [hasStart(earlyCandidates, "10:00"), hasStart(earlyCandidates, "12:00"), hasStart(earlyCandidates, "15:00")], [false, false, false]],
    ["早出社員の14:00+2時間は16:00", findCandidate(earlyCandidates, "14:00", 120).end_time, "16:00"],
    ["早出社員の09:00は1時間・2時間・3時間", earlyCandidates.filter(row => row.start_time === "09:00").map(row => row.requested_minutes), [60, 120, 180]],
    ["早出社員の09:00+4時間は候補外", !!findCandidate(earlyCandidates, "09:00", 240), false],
    ["早出社員の14:00は1時間・2時間だけ", earlyCandidates.filter(row => row.start_time === "14:00").map(row => row.requested_minutes), [60, 120]],
    ["早出社員で14:00選択後に09:00へ再選択しても09:00+2時間は11:00", findCandidate(earlyCandidates, "09:00", 120).end_time, "11:00"],
    ["早出社員に16:00超過候補なし", !!findCandidate(earlyCandidates, "14:00", 180), false],
    ["早出AM半休は07:00-11:00", [amRange.startMinute, amRange.endMinute], [420, 660]],
    ["早出AM半休は実労働210分", calculateTimeLeaveMinutes(amRange.startMinute, amRange.endMinute, earlyPolicy), 210],
    ["早出PM半休は11:00-16:00", [pmRange.startMinute, pmRange.endMinute], [660, 960]],
    ["早出PM半休は実労働210分", calculateTimeLeaveMinutes(pmRange.startMinute, pmRange.endMinute, earlyPolicy), 210],
    ["早出AM半休と重なる時間休を拒否", expectError(() => validateTimeLeaveEntriesAgainstCandidate_({
      policy: earlyPolicy, start_minute: 540, end_minute: 600, requested_minutes: 60
    }, [{ kind: "half_day", half_day: "am", minutes: 210 }])), true],
    ["新規明細に社員別勤務時間snapshotを保存", [segment.work_start_minute, segment.work_end_minute, segment.scheduled_minutes_per_day], [420, 960, 420]],
    ["承認時は現在の標準勤務ではなくsnapshotを復元", [snapshotPolicy.workStartMinute, snapshotPolicy.workEndMinute], [420, 960]]
  ];
  const failed = cases.filter(item => JSON.stringify(item[1]) !== JSON.stringify(item[2]));
  if (failed.length) throw new Error("社員別勤務時間テスト失敗: " + JSON.stringify(failed));
  return { ok: true, case_count: cases.length };
}

/* =========================
   時間単位年休 Phase 4: 承認時残高テスト
========================= */
function testTimeLeaveApprovalBalanceFoundation() {
  const expectError = callback => {
    try { callback(); return false; } catch (error) { return true; }
  };
  const originalContextFactory = createFifoBalanceComparisonContext_;
  const twoDayBatchContext = {
    grants_by_employee: {
      M4: [{ grant_id: "G", grant_date: parseLocalDate("2026-04-01"), valid_from_date: parseLocalDate("2026-04-01"), valid_to_date: parseLocalDate("2028-03-31"), grant_days: 2, carry_over_days: 0, carry_over_minutes: 0, is_finalized: true }]
    },
    requests_by_employee: {
      M4: ["A", "B", "C"].map((id, index) => ({ request_id: id, employee_id: "M4", start_date: "2026-04-0" + (index + 1), end_date: "2026-04-0" + (index + 1), days: 1, half_day: "", status: "pending", type: "paid_leave" }))
    },
    time_leave_segments_by_request: {}, company_code_by_employee: { M4: "MAIN" }, calendar_map: {}
  };
  let batchRejected = false;
  try {
    createFifoBalanceComparisonContext_ = () => twoDayBatchContext;
    validateMainApprovalBalancesForRequests_(["A", "B", "C"]);
  } catch (error) {
    batchRejected = true;
  } finally {
    createFifoBalanceComparisonContext_ = originalContextFactory;
  }
  const validatePendingCompetition = (first, second) => {
    const requests = {
      F: { request_id: "F", employee_id: "M5", start_date: "2026-04-01", end_date: "2026-04-01", days: 1, half_day: "", status: "pending", type: "paid_leave" },
      T: { request_id: "T", employee_id: "M5", start_date: "2026-04-01", end_date: "2026-04-01", days: 0, half_day: "", status: "pending", type: "paid_leave", request_kind: "time_hourly" }
    };
    const competitionContext = {
      grants_by_employee: { M5: [{ grant_id: "G", grant_date: parseLocalDate("2026-04-01"), valid_from_date: parseLocalDate("2026-04-01"), valid_to_date: parseLocalDate("2028-03-31"), grant_days: 1, carry_over_days: 0, carry_over_minutes: 0, is_finalized: true }] },
      requests_by_employee: { M5: [requests.F, requests.T] },
      time_leave_segments_by_request: { T: [{ time_leave_id: "T1", leave_date: "2026-04-01", requested_minutes: 60 }] },
      company_code_by_employee: { M5: "MAIN" }, calendar_map: {}
    };
    const savedFactory = createFifoBalanceComparisonContext_;
    try {
      createFifoBalanceComparisonContext_ = () => competitionContext;
      validateMainApprovalBalancesForRequests_([first]);
      return expectError(() => validateMainApprovalBalancesForRequests_([second]));
    } finally {
      createFifoBalanceComparisonContext_ = savedFactory;
    }
  };
  const cases = [
    ["残420分で全日承認", !expectError(() => validateMainApprovalRemainingMinutes_(420, 420)), true],
    ["残419分で全日拒否", expectError(() => validateMainApprovalRemainingMinutes_(419, 420)), true],
    ["残210分で半日承認", !expectError(() => validateMainApprovalRemainingMinutes_(210, 210)), true],
    ["残209分で半日拒否", expectError(() => validateMainApprovalRemainingMinutes_(209, 210)), true],
    ["残60分で時間休承認", !expectError(() => validateMainApprovalRemainingMinutes_(60, 60)), true],
    ["残59分で時間休拒否", expectError(() => validateMainApprovalRemainingMinutes_(59, 60)), true],
    ["2日残から全日2件は順次承認可能", !expectError(() => {
      validateMainApprovalRemainingMinutes_(840, 420);
      validateMainApprovalRemainingMinutes_(420, 420);
    }), true],
    ["2日残から3件目の全日は拒否", expectError(() => {
      validateMainApprovalRemainingMinutes_(0, 420);
    }), true],
    ["一括承認は2日残の3全日申請を全件通過させない", batchRejected, true],
    ["pending全日を先に承認するとpending時間休は拒否", validatePendingCompetition("F", "T"), true],
    ["pending時間休を先に承認するとpending全日は拒否", validatePendingCompetition("T", "F"), true],
    ["旧繰越1.5日は630分", 1.5 * 420, 630],
    ["新繰越5日180分は2280分", 5 * 420 + 180, 2280],
    ["720分は1日300分", [calculateCarryOverMinutes_(720, 420).carry_over_days, calculateCarryOverMinutes_(720, 420).carry_over_minutes], [1, 300]]
  ];
  const results = cases.map(item => ({ name: item[0], actual: item[1], expected: item[2], ok: JSON.stringify(item[1]) === JSON.stringify(item[2]) }));
  const failed = results.filter(item => !item.ok);
  if (failed.length) throw new Error("時間単位年休 Phase 4 テスト失敗: " + JSON.stringify(failed));
  Logger.log(JSON.stringify(results, null, 2));
  return { ok: true, case_count: results.length, results: results };
}

/* =========================
   FIFO: Supabase読取時の時間有給親ソース優先順位
========================= */
function testFifoTimeLeaveParentSourcePriorityFoundation() {
  const employeeId = "FIFO-TIME-001";
  const asOfDate = parseLocalDate("2026-09-14");
  const grant = {
    grant_id: "G-16", grant_date: parseLocalDate("2026-04-01"),
    valid_from_date: parseLocalDate("2026-04-01"), valid_to_date: parseLocalDate("2028-03-31"),
    grant_days: 16, carry_over_days: 0, carry_over_minutes: 0, is_finalized: true
  };
  const combinedSupabaseParent = {
    request_id: "R-COMBINED", employee_id: employeeId,
    start_date: "2026-09-14", end_date: "2026-09-14", days: 0.5, half_day: "pm",
    status: STATUS.APPROVED, type: "paid_leave", source: "supabase"
  };
  const combinedSpreadsheetParent = Object.assign({}, combinedSupabaseParent, {
    request_kind: "half_day_time_hourly", source: "spreadsheet"
  });
  const timeSupabaseParent = {
    request_id: "R-TIME", employee_id: employeeId,
    start_date: "2026-09-14", end_date: "2026-09-14", days: 0, half_day: "",
    status: STATUS.APPROVED, type: "paid_leave", source: "supabase"
  };
  const timeSpreadsheetParent = Object.assign({}, timeSupabaseParent, {
    request_kind: "time_hourly", source: "spreadsheet"
  });
  const normalSupabaseParent = {
    request_id: "R-NORMAL", employee_id: employeeId,
    start_date: "2026-09-14", end_date: "2026-09-14", days: 1, half_day: "",
    status: STATUS.APPROVED, type: "paid_leave", source: "supabase"
  };
  const halfDaySupabaseParent = Object.assign({}, normalSupabaseParent, {
    request_id: "R-HALF", days: 0.5, half_day: "pm"
  });
  const multiDaySupabaseParent = Object.assign({}, normalSupabaseParent, {
    request_id: "R-MULTI", end_date: "2026-09-15", days: 2
  });
  const contextFor = rows => ({
    grants_by_employee: { [employeeId]: [grant] },
    requests_by_employee: { [employeeId]: rows },
    time_leave_segments_by_request: {
      "R-COMBINED": [{ time_leave_id: "S-COMBINED", request_id: "R-COMBINED", leave_date: "2026-09-14", requested_minutes: 120 }],
      "R-TIME": [{ time_leave_id: "S-TIME", request_id: "R-TIME", leave_date: "2026-09-14", requested_minutes: 120 }]
    },
    company_code_by_employee: { [employeeId]: "MAIN" }, calendar_map: {}
  });
  const sumUsed = balance => Number(balance.used_minutes || 0);
  const combinedSupabaseReadRows = buildFifoLeaveRequestRowsByEmployee_(
    [combinedSupabaseParent], [combinedSpreadsheetParent]
  )[employeeId];
  const combinedSupabaseReadBalance = calculateFifoBalanceMinutesFromContext_(
    employeeId, asOfDate, contextFor(combinedSupabaseReadRows)
  );
  const combinedSpreadsheetReadBalance = calculateFifoBalanceMinutesFromContext_(
    employeeId, asOfDate, contextFor([combinedSpreadsheetParent])
  );
  const timeSupabaseReadRows = buildFifoLeaveRequestRowsByEmployee_(
    [timeSupabaseParent], [timeSpreadsheetParent]
  )[employeeId];
  const timeSupabaseReadBalance = calculateFifoBalanceMinutesFromContext_(
    employeeId, asOfDate, contextFor(timeSupabaseReadRows)
  );
  const normalRows = buildFifoLeaveRequestRowsByEmployee_(
    [normalSupabaseParent], []
  )[employeeId];
  const normalBalance = calculateFifoBalanceMinutesFromContext_(
    employeeId, asOfDate, contextFor(normalRows)
  );
  const halfDayRows = buildFifoLeaveRequestRowsByEmployee_([halfDaySupabaseParent], [])[employeeId];
  const halfDayBalance = calculateFifoBalanceMinutesFromContext_(
    employeeId, asOfDate, contextFor(halfDayRows)
  );
  const multiDayRows = buildFifoLeaveRequestRowsByEmployee_([multiDaySupabaseParent], [])[employeeId];
  const multiDayBalance = calculateFifoBalanceMinutesFromContext_(
    employeeId, parseLocalDate("2026-09-15"), contextFor(multiDayRows)
  );
  const combinedRequiredMinutes = getMainApprovalRequiredMinutes_(
    combinedSpreadsheetParent, contextFor([combinedSpreadsheetParent])
  );
  const cases = [
    ["同一IDのcombined親はSpreadsheet版を優先", combinedSupabaseReadRows[0].request_kind, "half_day_time_hourly"],
    ["Supabase read combinedのFIFO使用量は330分", sumUsed(combinedSupabaseReadBalance), 330],
    ["Supabase read combinedの残高は6390分", combinedSupabaseReadBalance.current_remaining_minutes, 6390],
    ["Spreadsheet read combinedの残高も6390分", combinedSpreadsheetReadBalance.current_remaining_minutes, 6390],
    ["Supabase/Spreadsheet readでcombined残高は一致", combinedSupabaseReadBalance.current_remaining_minutes, combinedSpreadsheetReadBalance.current_remaining_minutes],
    ["同一IDの単独時間年休はSpreadsheet版を優先", timeSupabaseReadRows[0].request_kind, "time_hourly"],
    ["単独時間年休のFIFO使用量は120分", sumUsed(timeSupabaseReadBalance), 120],
    ["単独時間年休の残高は6600分", timeSupabaseReadBalance.current_remaining_minutes, 6600],
    ["通常1日有給はSupabase版を維持", normalRows[0].source, "supabase"],
    ["通常1日有給のFIFO使用量は420分", sumUsed(normalBalance), 420],
    ["通常1日有給の残高は6300分", normalBalance.current_remaining_minutes, 6300],
    ["通常半休はSupabase版を維持", halfDayRows[0].source, "supabase"],
    ["通常半休のFIFO使用量は210分", sumUsed(halfDayBalance), 210],
    ["通常半休の残高は6510分", halfDayBalance.current_remaining_minutes, 6510],
    ["複数日有給はSupabase版を維持", multiDayRows[0].source, "supabase"],
    ["複数日有給のFIFO使用量は840分", sumUsed(multiDayBalance), 840],
    ["複数日有給の残高は5880分", multiDayBalance.current_remaining_minutes, 5880],
    ["combined承認必要量と承認後FIFO使用量は一致", combinedRequiredMinutes, sumUsed(combinedSupabaseReadBalance)]
  ];
  const results = cases.map(item => ({ name: item[0], actual: item[1], expected: item[2], ok: JSON.stringify(item[1]) === JSON.stringify(item[2]) }));
  const failed = results.filter(item => !item.ok);
  if (failed.length) throw new Error("FIFO時間有給親ソース優先順位テスト失敗: " + JSON.stringify(failed));
  return { ok: true, case_count: results.length, results: results };
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

/* =========================
   PARTNER P0004: 初期導入残高の読み取り専用診断テスト
========================= */
function testPartnerP0004FifoDiagnosis() {
  const asOfDate = parseLocalDate("2026-07-25");
  const employee = {
    employee_id: "EMP0062", display_employee_id: "P0004", company_code: "PARTNER",
    fiscal_start_month: 6, employment_status: "active", leave_management_target: true,
    hire_date: "2020-01-01", work_days_per_week: 5
  };
  const context = {
    as_of_date: asOfDate,
    calendar_map: {},
    requests_by_employee: {
      EMP0062: [
        { request_id: "R-MAY", start_date: "2026-05-02", end_date: "2026-05-02", days: 1, half_day: "", status: STATUS.APPROVED, type: "paid_leave" },
        { request_id: "R-HALF", start_date: "2026-07-01", end_date: "2026-07-01", days: 1, half_day: "AM", status: STATUS.APPROVED, type: "paid_leave" },
        { request_id: "R-THREE", start_date: "2026-07-02", end_date: "2026-07-04", days: 3, half_day: "", status: STATUS.APPROVED, type: "paid_leave" },
        { request_id: "R-CANCEL", start_date: "2026-07-10", end_date: "2026-07-10", days: 1, half_day: "", status: "cancelled", type: "paid_leave" }
      ]
    },
    grants_by_employee: {
      EMP0062: [
        {
          grant_id: "G0058", employee_id: "EMP0062", grant_date: parseLocalDate("2025-06-01"),
          valid_from_date: parseLocalDate("2025-06-01"), valid_to_date: parseLocalDate("2026-05-31"),
          grant_type: "initial", year: 2025, grant_days: 3, carry_over_days: 0, total_days: 3,
          notes: "初期導入残高 SECRET", has_recorded_valid_from: true, has_recorded_valid_to: true, is_finalized: true
        },
        {
          grant_id: "G2026-62", employee_id: "EMP0062", grant_date: parseLocalDate("2026-06-01"),
          valid_from_date: parseLocalDate("2026-06-01"), valid_to_date: parseLocalDate("2028-05-31"),
          grant_type: "yearly", year: 2026, grant_days: 11, carry_over_days: 0, total_days: 11,
          notes: "", has_recorded_valid_from: true, has_recorded_valid_to: true, is_finalized: true
        }
      ]
    }
  };
  const original = JSON.stringify(context);
  const result = buildPartnerP0004FifoDiagnosis_(employee, context, asOfDate);
  const invalid = buildPartnerP0004FifoDiagnosis_(employee, Object.assign({}, context, {
    grants_by_employee: { EMP0062: [context.grants_by_employee.EMP0062[1]] }
  }), asOfDate);
  const cases = [
    ["P0004だけを対象", result.target.employee_id + "/" + result.target.grant_id, "EMP0062/G0058"],
    ["G0058を取得", result.grants.some(row => row.grant_id === "G0058"), true],
    ["P0002/P0003を含めない", result.grants.every(row => row.grant_id !== "G0056" && row.grant_id !== "G0057"), true],
    ["期限切れ2日の内訳", result.expired_days_breakdown[0].expired_days, 2],
    ["期限切れはG0058由来", result.expired_days_conclusion.is_g0058_derived, true],
    ["有効残高7.5日の合計", result.active_balance_total, 7.5],
    ["年度取得3.5日の合計", result.fiscal_usage.total_used_days, 3.5],
    ["半日申請を0.5日扱い", result.fiscal_usage.included_requests.find(row => row.request_id === "R-HALF").fiscal_days, 0.5],
    ["取消申請をFIFO対象外", result.requests.find(row => row.request_id === "R-CANCEL").fifo_included, false],
    ["最短期限2028/05/31の元ロット", result.nearest_expiry_lot.grant_id, "G2026-62"],
    ["案Aの試算", result.simulations.plan_a_current_is_correct.expired_days, 2],
    ["案Bの試算", result.simulations.plan_b_fiscal_start_carry_over.expired_days, 0],
    ["案Cの試算", result.simulations.plan_c_extend_initial_lot_expiry.expired_days, 0],
    ["不正データは安全な警告", invalid.warnings.some(message => message.indexOf("G0058") !== -1), true],
    ["notes本文を返さない", JSON.stringify(result).indexOf("SECRET") === -1, true],
    ["元FIFOコンテキストを変更しない", JSON.stringify(context), original],
    ["診断関数に書込み関数がない", debugPartnerCarryOverSimulationP0004.toString().indexOf("setValue") === -1, true]
  ];
  const failed = cases.filter(item => JSON.stringify(item[1]) !== JSON.stringify(item[2]));
  if (failed.length) throw new Error("PARTNER P0004 FIFO診断テスト失敗: " + JSON.stringify(failed));
  return { ok: true, case_count: cases.length };
}

/* =========================
   PARTNER P0004: G0061繰越構造の統一試算テスト
========================= */
function testPartnerP0004CarryOverStructureSimulation() {
  const asOfDate = parseLocalDate("2026-07-25");
  const employee = { employee_id: "EMP0062", display_employee_id: "P0004", company_code: "PARTNER", fiscal_start_month: 6 };
  const context = {
    as_of_date: asOfDate,
    calendar_map: {},
    requests_by_employee: { EMP0062: [
      { request_id: "R-MAY", start_date: "2026-05-02", end_date: "2026-05-02", days: 1, half_day: "", status: STATUS.APPROVED, type: "paid_leave" },
      { request_id: "R-JUN2", start_date: "2026-06-02", end_date: "2026-06-02", days: 1, half_day: "", status: STATUS.APPROVED, type: "paid_leave" },
      { request_id: "R-HALF", start_date: "2026-06-24", end_date: "2026-06-24", days: 1, half_day: "AM", status: STATUS.APPROVED, type: "paid_leave" },
      { request_id: "R-JUN26", start_date: "2026-06-26", end_date: "2026-06-26", days: 1, half_day: "", status: STATUS.APPROVED, type: "paid_leave" },
      { request_id: "R-JUL18", start_date: "2026-07-18", end_date: "2026-07-18", days: 1, half_day: "", status: STATUS.APPROVED, type: "paid_leave" }
    ] },
    grants_by_employee: { EMP0062: [
      { grant_id: "G0058", employee_id: "EMP0062", grant_date: parseLocalDate("2025-06-01"), valid_from_date: parseLocalDate("2025-06-01"), valid_to_date: parseLocalDate("2026-05-31"), grant_type: "initial", year: 2025, grant_days: 3, carry_over_days: 0, total_days: 3, notes: "初期導入残高", has_recorded_valid_from: true, has_recorded_valid_to: true, is_finalized: true },
      { grant_id: "G0061", employee_id: "EMP0062", grant_date: parseLocalDate("2026-06-01"), valid_from_date: parseLocalDate("2026-06-01"), valid_to_date: parseLocalDate("2028-05-31"), grant_type: "yearly", year: 2026, grant_days: 11, carry_over_days: 2, total_days: 13, notes: "", has_recorded_valid_from: true, has_recorded_valid_to: true, is_finalized: true }
    ] }
  };
  const original = JSON.stringify(context);
  const result = buildPartnerP0004CarryOverStructureSimulation_(employee, context, asOfDate);
  const conversion = buildPartnerP0004FiscalStartCarryOverConversion_(employee, context, asOfDate);
  const bAllocations = result.plan_b.fiscal_usage_allocations;
  const cases = [
    ["対象はP0004のみ", result.employee.employee_id, "EMP0062"],
    ["現状FIFOは7.5日", result.current.current_remaining_days, 7.5],
    ["現状FIFOはG0058由来2日失効", result.current.expired_days, 2],
    ["現状は年度残高と2日差", result.current.difference_from_yearly_balance, -2],
    ["案Aは9.5日", result.plan_a.current_remaining_days, 9.5],
    ["案Aは期限切れ0日", result.plan_a.expired_days, 0],
    ["案Bは9.5日", result.plan_b.current_remaining_days, 9.5],
    ["案BはG0058履歴の2日失効を保持", result.plan_b.expired_days, 2],
    ["案Bは年度残高と一致", result.plan_b.difference_from_yearly_balance, 0],
    ["案Bは繰越2日を先に割当", bAllocations[0].grant_id, "G0061#carry_over_simulation#opening_balance"],
    ["案Bの繰越ロットは2日消化", bAllocations.filter(row => row.grant_id === "G0061#carry_over_simulation#opening_balance").reduce((sum, row) => sum + row.consumed_days, 0), 2],
    ["案Bの次消化ロットは11日付与", result.plan_b.next_consumption_lot.grant_id, "G0061"],
    ["同一行では優先順を保証できない警告を返す", result.design_comparison.fifo_order_constraint.indexOf("同一G0061行") !== -1, true],
    ["案DはG0058のみを試算対象", conversion.target.grant_id, "G0058"],
    ["案Dのgrant_daysは0", conversion.plan_d_changes.grant_days, 0],
    ["案Dのcarry_over_daysは2", conversion.plan_d_changes.carry_over_days, 2],
    ["案Dの有効開始日は2026/06/01", conversion.plan_d_changes.valid_from, "2026-06-01"],
    ["案Dの有効期限は2027/05/31", conversion.plan_d_changes.valid_to, "2027-05-31"],
    ["案Dは5/2取得を新繰越へ割り当てない", conversion.plan_d_allocations.allocations_on_2026_05_02.length, 0],
    ["案Dは繰越2日を先に使い切る", conversion.plan_d_allocations.g0058_carry_over_consumed_days, 2],
    ["案Dは残高9.5日", conversion.plan_d.current_remaining_days, 9.5],
    ["案Dは期限切れ0日", conversion.plan_d.expired_days, 0],
    ["案Dは年度残高差0日", conversion.plan_d.difference_from_yearly_balance, 0],
    ["案DのFIFO合計は13日", conversion.plan_d.fifo_lots.reduce((sum, lot) => sum + lot.total_days, 0), 13],
    ["案DではG0061繰越を二重計上しない", conversion.checks.g0061_carry_over_not_double_counted, true],
    ["案Dでは残高のある最短期限は2028/05/31", conversion.plan_d.nearest_active_expiry_lot.valid_to, "2028-05-31"],
    ["案Dではロット全体の最短期限は2027/05/31", conversion.plan_d.earliest_lot_expiry.valid_to, "2027-05-31"],
    ["試算は元コンテキストを変更しない", JSON.stringify(context), original],
    ["診断関数に書込み関数がない", debugPartnerCarryOverStructureSimulationP0004.toString().indexOf("setValue") === -1, true]
  ];
  const failed = cases.filter(item => JSON.stringify(item[1]) !== JSON.stringify(item[2]));
  if (failed.length) throw new Error("PARTNER P0004繰越構造試算テスト失敗: " + JSON.stringify(failed));
  return { ok: true, case_count: cases.length };
}

/* =========================
   PARTNER P0004一回限り補正の固定データ安全性テスト
========================= */
function testPartnerOpeningBalanceP0004RepairSafety() {
  const throws = fn => { try { fn(); return false; } catch (error) { return true; } };
  const source = repairPartnerOpeningBalanceFiscalStartP0004.toString();
  const cases = [
    ["対象grant_idはG0058", PARTNER_OPENING_BALANCE_P0004_REPAIR_TARGET_.grant_id, "G0058"],
    ["対象employee_idはEMP0062", PARTNER_OPENING_BALANCE_P0004_REPAIR_TARGET_.employee_id, "EMP0062"],
    ["dry-runが既定", source.indexOf("opts.dry_run !== false") !== -1, true],
    ["本実行を必ず拒否", throws(() => repairPartnerOpeningBalanceFiscalStartP0004({ dry_run: false })), true],
    ["確認文字列があっても拒否", throws(() => repairPartnerOpeningBalanceFiscalStartP0004({ dry_run: false, confirmation_text: "ANY" })), true],
    ["完了済みエラーを返す", (() => { try { repairPartnerOpeningBalanceFiscalStartP0004({ dry_run: false }); } catch (error) { return String(error.message).indexOf("P0004(G0058)") !== -1; } return false; })(), true],
    ["LockServiceを削除", source.indexOf("LockService") === -1, true],
    ["シート更新を削除", source.indexOf("setValue") === -1 && source.indexOf("setValues") === -1, true],
    ["usage_log書込みを削除", source.indexOf("appendUsageLog") === -1, true],
    ["ロールバックを削除", source.indexOf("restorePartner") === -1, true],
    ["診断関数を維持", typeof debugPartnerCarryOverSimulationP0004, "function"],
    ["構造試算を維持", typeof debugPartnerCarryOverStructureSimulationP0004, "function"],
    ["変換試算を維持", typeof debugPartnerFiscalStartCarryOverConversionP0004, "function"],
    ["事前条件診断を維持", typeof debugPartnerOpeningBalanceFiscalStartRepairPreconditionsP0004, "function"]
  ];
  const failed = cases.filter(item => JSON.stringify(item[1]) !== JSON.stringify(item[2]));
  if (failed.length) throw new Error("PARTNER P0004補正安全性テスト失敗: " + JSON.stringify(failed));
  return { ok: true, case_count: cases.length };
}

/* =========================
   P0004補正ラッパーのログ整形テスト
========================= */
function testPartnerOpeningBalanceP0004RepairWrappers() {
  const result = {
    as_of_date: "2026-07-25",
    target: {
      grant_id: "G0058", employee_id: "EMP0062", display_employee_id: "P0004",
      before: { grant_days: 3, carry_over_days: 0, valid_from: "2025-06-01", valid_to: "2026-05-31" },
      after: { grant_days: 0, carry_over_days: 2, valid_from: "2026-06-01", valid_to: "2027-05-31" }
    },
    fifo_before: { current_active_remaining_days: 7.5, expired_days: 2 },
    fifo_after: { current_active_remaining_days: 9.5, expired_days: 0 },
    difference_from_yearly_balance: 0,
    warnings: []
  };
  const original = JSON.stringify(result);
  logPartnerOpeningBalanceP0004RepairDryRun_(result);
  const throws = fn => { try { fn(); return false; } catch (error) { return true; } };
  const executeSource = executeRepairPartnerOpeningBalanceFiscalStartP0004.toString();
  const cases = [
    ["dry-runラッパーを追加", typeof debugRepairPartnerOpeningBalanceFiscalStartP0004, "function"],
    ["dry-runラッパーはtrueを渡す", debugRepairPartnerOpeningBalanceFiscalStartP0004.toString().indexOf("dry_run: true") !== -1, true],
    ["dry-runラッパーは結果をreturn", debugRepairPartnerOpeningBalanceFiscalStartP0004.toString().indexOf("return result") !== -1, true],
    ["executeラッパーを追加", typeof executeRepairPartnerOpeningBalanceFiscalStartP0004, "function"],
    ["executeラッパーは必ず拒否", throws(() => executeRepairPartnerOpeningBalanceFiscalStartP0004()), true],
    ["executeラッパーに確認文字列がない", executeSource.indexOf("confirmation_text") === -1, true],
    ["executeラッパーに本実行Loggerがない", executeSource.indexOf("Logger") === -1, true],
    ["dry-runログ整形は戻り値を変更しない", JSON.stringify(result), original],
    ["詳細JSONにBASIC区分がある", logPartnerOpeningBalanceP0004RepairDryRun_.toString().indexOf("[P0004_REPAIR][BASIC]") !== -1, true],
    ["詳細JSONにWARNINGS区分がある", logPartnerOpeningBalanceP0004RepairDryRun_.toString().indexOf("[P0004_REPAIR][WARNINGS]") !== -1, true]
  ];
  const failed = cases.filter(item => JSON.stringify(item[1]) !== JSON.stringify(item[2]));
  if (failed.length) throw new Error("P0004補正ラッパーテスト失敗: " + JSON.stringify(failed));
  return { ok: true, case_count: cases.length };
}

/* =========================
   P0004補正事前条件診断の正規化テスト
========================= */
function testPartnerOpeningBalanceP0004RepairPreconditions() {
  const target = Object.assign(
    {},
    PARTNER_OPENING_BALANCE_P0004_REPAIR_TARGET_,
    PARTNER_OPENING_BALANCE_P0004_REPAIR_AFTER_
  );
  const row = {
    grant_id: " G0058 ", employee_id: " emp0062 ", grant_days: "0", carry_over_days: "2",
    valid_from: parseLocalDate("2026-06-01"), valid_to: "2027/05/31",
    grant_type: "initial", year: "2025", notes: "初期導入残高\n" + PARTNER_OPENING_BALANCE_P0004_REPAIR_MARKER_
  };
  const employee = { company_code: " partner ", display_employee_id: "P0004" };
  const matched = buildPartnerOpeningBalanceFiscalStartPreconditionDiagnostic_(row, employee, target, 1);
  const mismatched = buildPartnerOpeningBalanceFiscalStartPreconditionDiagnostic_(
    Object.assign({}, row, { carry_over_days: "1", valid_from: parseLocalDate("2026-06-02") }), employee, target, 1
  );
  const errorMessage = (() => {
    try {
      validatePartnerOpeningBalanceP0004RepairState_({
        target_row: Object.assign({}, row, { grant_id: "G0058", employee_id: "EMP0062", carry_over_days: 1 }),
        g0061_row: { grant_id: "G0061", employee_id: "EMP0062" },
        employee_map: { EMP0062: { company_code: "PARTNER", display_employee_id: "P0004" } }
      }, "after");
      return "";
    } catch (error) { return String(error.message || error); }
  })();
  const cases = [
    ["一致ケース", matched.all_conditions_match, true],
    ["日付型を表示", matched.checks.valid_from.actual_type, "Date"],
    ["数値文字列を正規化", matched.checks.grant_days.actual_normalized, 0],
    ["display_employee_idを確認", matched.checks.display_employee_id.matched, true],
    ["grant_typeを情報として返す", matched.checks.grant_type.actual, "initial"],
    ["yearを情報として返す", matched.checks.year.actual, "2025"],
    ["notes本文を返さない", Object.prototype.hasOwnProperty.call(matched.checks.notes_marker, "actual"), false],
    ["不一致carry_overを検出", mismatched.mismatch_fields.indexOf("carry_over_days") !== -1, true],
    ["不一致valid_fromを検出", mismatched.mismatch_fields.indexOf("valid_from") !== -1, true],
    ["エラーに見出しを含む", errorMessage.indexOf("不一致項目") !== -1, true],
    ["エラーに項目名を含む", errorMessage.indexOf("carry_over_days") !== -1, true],
    ["診断関数を追加", typeof debugPartnerOpeningBalanceFiscalStartRepairPreconditionsP0004, "function"]
  ];
  const failed = cases.filter(item => JSON.stringify(item[1]) !== JSON.stringify(item[2]));
  if (failed.length) throw new Error("P0004補正事前条件診断テスト失敗: " + JSON.stringify(failed));
  return { ok: true, case_count: cases.length };
}

/* =========================
   P0004補正後: 付与予定APIの固定データ診断テスト
   シート・CacheService・書込みを使用しない。
========================= */
function testPaidLeaveGrantScheduleAfterP0004Repair() {
  const fifoBalance = {
    current_remaining_days: 9.5,
    expired_days: 0,
    grant_details: [
      { grant_id: "G0058#opening_balance", valid_to: "2027/05/31", active_remaining_days: 0 },
      { grant_id: "G0061", valid_to: "2028/05/31", active_remaining_days: 9.5 }
    ]
  };
  const schedule = {
    eligibility_status: "PROCESSED",
    warning_codes: []
  };
  const diagnostic = buildPaidLeaveGrantScheduleApiAfterP0004RepairDiagnostic_(
    { success: true, rows: [{ employee_id: "EMP0062" }] },
    schedule,
    fifoBalance
  );
  const source = debugPaidLeaveGrantScheduleApiAfterP0004Repair.toString();
  const cases = [
    ["補正後FIFOを処理", diagnostic.p0004.current_remaining_days, 9.5],
    ["grant_days=0/carry_over_days=2相当の期限切れなし", diagnostic.p0004.expired_days, 0],
    ["最短期限を有効ロットから取得", diagnostic.p0004.nearest_expiry_date, "2028/05/31"],
    ["APIレスポンス成功を返す", diagnostic.ok, true],
    ["残高の期待値比較を返す", diagnostic.expected_values_check.current_remaining_days_9_5, true],
    ["期限の期待値比較を返す", diagnostic.expected_values_check.nearest_expiry_date_2028_05_31, true],
    ["診断は読み取り専用", diagnostic.read_only, true],
    ["API診断関数を追加", typeof debugPaidLeaveGrantScheduleApiAfterP0004Repair, "function"],
    ["API診断はFIFO計算を行う", source.indexOf("calculateFifoBalanceWithOpeningBalanceFromContext_") !== -1, true],
    ["API診断は書込みを行わない", source.indexOf("setValue") === -1 && source.indexOf("appendRow") === -1, true]
  ];
  const failed = cases.filter(item => JSON.stringify(item[1]) !== JSON.stringify(item[2]));
  if (failed.length) throw new Error("P0004補正後の付与予定APIテスト失敗: " + JSON.stringify(failed));
  return { ok: true, case_count: cases.length };
}
