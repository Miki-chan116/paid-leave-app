/* =========================
   FIFO debug / validation
   debug.gs から動作を変えずに移動
========================= */

/**
 * 指定社員のFIFOロットを、基準日時点の状態で全件返す読み取り専用デバッグ関数。
 * Apps Script エディタから debugFifoLots() として実行すると P0002 を確認できる。
 */
function debugFifoLots(employeeId, asOfDateValue) {
  const inputEmployeeId = String(employeeId || "P0002").trim();
  const normalizedInput = normalizePaidLeaveDebugEmployeeId_(inputEmployeeId);
  const matchedEmployee = getEmployeesForAdmin().find(emp =>
    normalizePaidLeaveDebugEmployeeId_(emp.employee_id) === normalizedInput ||
    normalizePaidLeaveDebugEmployeeId_(emp.display_employee_id) === normalizedInput
  );
  const targetEmployeeId = matchedEmployee
    ? String(matchedEmployee.employee_id || "").trim()
    : inputEmployeeId;
  const asOfDate = asOfDateValue
    ? parseLocalDate(asOfDateValue)
    : parseLocalDate(new Date());

  // 予定APIと同様に、カレンダーキャッシュを書き換えない。
  const context = createFifoBalanceComparisonContext_(asOfDate, { read_only: true });
  const fifoBalance = calculateFifoBalanceWithOpeningBalanceFromContext_(
    targetEmployeeId,
    asOfDate,
    context
  );
  const sourceRowsByGrantId = {};
  (context.grants_by_employee[targetEmployeeId] || []).forEach(row => {
    sourceRowsByGrantId[String(row.grant_id || "")] = row;
  });

  const lots = (fifoBalance.grant_details || []).map(lot => {
    const source = sourceRowsByGrantId[String(lot.source_grant_id || lot.grant_id || "")] || {};
    const remainingDays = Number(lot.remaining_days || 0);
    const activeRemainingDays = Number(lot.active_remaining_days || 0);
    const expired = lot.is_expired === true;

    return {
      grant_id: lot.grant_id,
      source_grant_id: lot.source_grant_id || lot.grant_id,
      lot_type: lot.lot_type,
      grant_date: lot.grant_date,
      grant_reason: getFifoDebugGrantReason_(lot, source),
      grant_type: lot.grant_type,
      fiscal_year: lot.year || "",
      total_days: Number(lot.total_days || 0),
      used_days: Number(lot.used_days || 0),
      remaining_days: remainingDays,
      active_remaining_days: activeRemainingDays,
      expired_days: Number(lot.expired_days || 0),
      valid_from: lot.valid_from,
      valid_to: lot.valid_to,
      is_expired: expired,
      expiry_status: expired ? "EXPIRED" : (activeRemainingDays > 0 ? "ACTIVE" : "FULLY_USED"),
      validity_basis: lot.validity_basis || "",
      validity_needs_review: lot.validity_needs_review === true
    };
  });

  const result = {
    ok: true,
    input_employee_id: inputEmployeeId,
    employee_id: targetEmployeeId,
    display_employee_id: matchedEmployee ? String(matchedEmployee.display_employee_id || "") : "",
    as_of_date: formatDateValue(asOfDate),
    calculation_mode: fifoBalance.calculation_mode,
    summary: {
      total_granted_days: Number(fifoBalance.total_granted_days || 0),
      used_days: Number(fifoBalance.used_days || 0),
      allocated_used_days: Number(fifoBalance.allocated_used_days || 0),
      unallocated_used_days: Number(fifoBalance.unallocated_used_days || 0),
      current_remaining_days: Number(fifoBalance.current_remaining_days || 0),
      expired_days: Number(fifoBalance.expired_days || 0)
    },
    lots: lots,
    allocations: fifoBalance.allocations || []
  };

  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

function getFifoDebugGrantReason_(lot, source) {
  const notes = String(source.notes || "").trim();
  if (notes) return notes;
  if (lot.lot_type === "opening_balance_virtual_lot") return "初期導入残高（仮想ロット）";

  const labels = {
    six_month: "初回付与（入社6か月または会社基準日）",
    six_month_processed: "初回付与チェック処理済み",
    six_month_skipped: "初回付与スキップ",
    initial: "初回付与（旧形式）",
    yearly: "年次付与"
  };
  return labels[String(lot.grant_type || "").trim().toLowerCase()] || "付与理由未記録";
}

/**
 * 画面表示とFIFO診断の入力データを、個人情報を最小限にして比較する。
 * employee_id と display_employee_id の取り違えも検出する。
 */
function debugPaidLeaveDataSources(employeeId, asOfDateValue) {
  const inputEmployeeId = String(employeeId || "P0002").trim();
  const asOfDate = asOfDateValue
    ? parseLocalDate(asOfDateValue)
    : parseLocalDate(new Date());
  const normalizedInput = normalizePaidLeaveDebugEmployeeId_(inputEmployeeId);
  const employees = getEmployeesForAdmin();
  const matches = employees.filter(emp =>
    normalizePaidLeaveDebugEmployeeId_(emp.employee_id) === normalizedInput ||
    normalizePaidLeaveDebugEmployeeId_(emp.display_employee_id) === normalizedInput
  );
  const resolvedEmployeeIds = matches
    .map(emp => String(emp.employee_id || "").trim())
    .filter(Boolean);
  const targetEmployeeIds = resolvedEmployeeIds.length > 0
    ? resolvedEmployeeIds
    : [inputEmployeeId];
  const context = createFifoBalanceComparisonContext_(asOfDate, { read_only: true });
  const fiscalYear = getFiscalYearFromDate(asOfDate);

  const result = {
    ok: true,
    input_employee_id: inputEmployeeId,
    normalized_input_employee_id: normalizedInput,
    as_of_date: formatDateValue(asOfDate),
    fiscal_year: fiscalYear,
    employee_master_matches: matches.map(sanitizePaidLeaveDebugEmployee_),
    resolved_employee_ids: targetEmployeeIds,
    id_resolution: matches.length === 0
      ? "社員マスターの employee_id / display_employee_id に一致しません。入力値そのものを参照しました。"
      : "display_employee_id が入力値と一致する場合、resolved_employee_ids の内部 employee_id が画面・FIFOの正式キーです。",
    employees: targetEmployeeIds.map(targetEmployeeId =>
      buildPaidLeaveDataSourceDebugForEmployee_(targetEmployeeId, asOfDate, fiscalYear, context)
    )
  };

  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

function buildPaidLeaveDataSourceDebugForEmployee_(employeeId, asOfDate, fiscalYear, context) {
  const fifoBalance = calculateFifoBalanceWithOpeningBalanceFromContext_(
    employeeId,
    asOfDate,
    context
  );
  const employee = getEmployeesForAdmin().find(row => String(row.employee_id || "").trim() === employeeId) || {};
  const fiscalStartMonth = Number(employee.fiscal_start_month || 4);
  const legacyBalance = calculateLegacyBalanceFromFifoContext_(
    employeeId,
    fiscalYear,
    fiscalStartMonth,
    context
  );
  const expiryInfo = buildPaidLeaveDashboardExpiryInfo_(fifoBalance, asOfDate);
  const rawGrantRows = getPaidLeaveDebugRawRows_("paid_leave_grants")
    .filter(row => normalizePaidLeaveDebugEmployeeId_(row.employee_id) === normalizePaidLeaveDebugEmployeeId_(employeeId))
    .map(sanitizePaidLeaveDebugGrant_);
  const rawRequestRows = getPaidLeaveDebugRawRows_("leave_requests")
    .filter(row => normalizePaidLeaveDebugEmployeeId_(row.employee_id) === normalizePaidLeaveDebugEmployeeId_(employeeId))
    .map(sanitizePaidLeaveDebugRequest_);

  return {
    employee_id: employeeId,
    employee_master: sanitizePaidLeaveDebugEmployee_(employee),
    paid_leave_grants_rows: rawGrantRows,
    opening_balance_rows: rawGrantRows.filter(row => row.is_opening_balance),
    leave_requests_rows: rawRequestRows,
    yearly_balance_fields: legacyBalance,
    fifo_result: {
      current_remaining_days: Number(fifoBalance.current_remaining_days || 0),
      expired_days: Number(fifoBalance.expired_days || 0),
      total_granted_days: Number(fifoBalance.total_granted_days || 0),
      used_days: Number(fifoBalance.used_days || 0),
      grant_details: fifoBalance.grant_details || [],
      allocations: fifoBalance.allocations || []
    },
    dashboard_result_equivalent: {
      current_remaining_days: Number(fifoBalance.current_remaining_days || 0),
      fiscal_used_days: Number(legacyBalance.used_days || 0),
      expired_days: Number(fifoBalance.expired_days || 0),
      nearest_expiry_date: expiryInfo.nearest_expiry_date,
      nearest_expiry_days: expiryInfo.nearest_expiry_days,
      expiry_status: expiryInfo.expiry_status
    }
  };
}

function getPaidLeaveDebugRawRows_(sheetName) {
  if (shouldUseSupabaseReads_()) {
    if (sheetName === "paid_leave_grants") return getPaidLeaveGrantsFromSupabase_();
    if (sheetName === "leave_requests") return getLeaveRequestsFromSupabase_();
    return [];
  }
  const sheet = getSheet(sheetName);
  const headerInfo = getHeaderMap(sheet);
  const data = sheet.getDataRange().getValues();
  return data.slice(1).map(row => rowToObject(row, headerInfo.headers));
}

function sanitizePaidLeaveDebugEmployee_(employee) {
  const row = employee || {};
  return {
    employee_id: String(row.employee_id || "").trim(),
    display_employee_id: String(row.display_employee_id || "").trim(),
    company_code: String(row.company_code || "").trim(),
    fiscal_start_month: row.fiscal_start_month || "",
    employment_status: String(row.employment_status || "").trim(),
    leave_management_target: row.leave_management_target === true,
    hire_date: row.hire_date || "",
    work_days_per_week: row.work_days_per_week || ""
  };
}

function sanitizePaidLeaveDebugGrant_(row) {
  const notes = String(row.notes || "");
  return {
    grant_id: String(row.grant_id || ""),
    employee_id: String(row.employee_id || "").trim(),
    grant_date: formatDateValue(row.grant_date),
    grant_days: Number(row.grant_days || 0),
    carry_over_days: Number(row.carry_over_days || 0),
    valid_from: formatDateValue(row.valid_from || row.grant_date),
    valid_to: formatDateValue(row.valid_to || ""),
    grant_type: String(row.grant_type || ""),
    year: row.year || "",
    is_finalized: row.is_finalized !== false && String(row.is_finalized || "").toUpperCase() !== "FALSE",
    is_opening_balance: notes.indexOf("初期導入残高") !== -1,
    notes_present: !!notes
  };
}

function sanitizePaidLeaveDebugRequest_(row) {
  return {
    request_id: String(row.request_id || ""),
    employee_id: String(row.employee_id || "").trim(),
    start_date: formatDateValue(row.start_date),
    end_date: formatDateValue(row.end_date),
    days: Number(row.days || 0),
    half_day: String(row.half_day || ""),
    status: String(row.status || ""),
    type: String(row.type || "paid_leave")
  };
}

function normalizePaidLeaveDebugEmployeeId_(value) {
  return String(value || "").trim().normalize("NFKC").toUpperCase();
}

/**
 * PARTNER の初期導入残高を、書込みなしで監査・補正案試算する。
 * 例: debugPartnerOpeningBalanceAudit("2026-07-25")
 */
function debugPartnerOpeningBalanceAudit(asOfDateValue) {
  const asOfDate = asOfDateValue
    ? parseLocalDate(asOfDateValue)
    : parseLocalDate(new Date());
  const fiscalYear = getFiscalYearFromDate(asOfDate);
  const employees = getEmployeesForAdmin();
  const employeeById = {};
  employees.forEach(employee => {
    employeeById[String(employee.employee_id || "").trim()] = employee;
  });
  const context = createFifoBalanceComparisonContext_(asOfDate, { read_only: true });
  const rawRows = getPaidLeaveDebugRawRows_("paid_leave_grants");
  const auditRows = rawRows
    .filter(row => {
      const employee = employeeById[String(row.employee_id || "").trim()];
      return employee && String(employee.company_code || "").trim().toUpperCase() === "PARTNER";
    })
    .filter(isPartnerOpeningBalanceAuditRow_)
    .map(row => buildPartnerOpeningBalanceAuditRow_(
      row,
      employeeById[String(row.employee_id || "").trim()],
      asOfDate,
      fiscalYear,
      context
    ));

  const result = {
    ok: true,
    company_code: "PARTNER",
    as_of_date: formatDateValue(asOfDate),
    fiscal_year: fiscalYear,
    row_count: auditRows.length,
    warning: "この結果は読み取り専用の試算です。シート、キャッシュ、FIFO消化ロジックは変更しません。",
    rows: auditRows
  };
  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

function isPartnerOpeningBalanceAuditRow_(row) {
  const notes = String(row && row.notes || "");
  const grantType = String(row && row.grant_type || "").trim().toLowerCase();
  const initialTypes = {
    initial: true,
    six_month: true,
    six_month_processed: true,
    six_month_skipped: true
  };
  return notes.indexOf("初期導入残高") !== -1 &&
    (initialTypes[grantType] || notes.indexOf("初期") !== -1);
}

function buildPartnerOpeningBalanceAuditRow_(row, employee, asOfDate, fiscalYear, context) {
  const employeeId = String(row.employee_id || "").trim();
  const grantId = String(row.grant_id || "");
  const grantDays = Number(row.grant_days || 0);
  const carryOverDays = Number(row.carry_over_days || 0);
  const currentFifo = calculateFifoBalanceWithOpeningBalanceFromContext_(
    employeeId,
    asOfDate,
    context
  );
  const legacyBalance = calculateLegacyBalanceFromFifoContext_(
    employeeId,
    fiscalYear,
    Number(employee.fiscal_start_month || 6),
    context
  );
  const currentLots = (currentFifo.grant_details || []).filter(lot =>
    String(lot.source_grant_id || lot.grant_id || "") === grantId
  );

  return {
    employee_id: employeeId,
    display_employee_id: String(employee.display_employee_id || ""),
    grant_id: grantId,
    grant_date: formatDateValue(row.grant_date),
    grant_days: grantDays,
    carry_over_days: carryOverDays,
    valid_from: formatDateValue(row.valid_from || row.grant_date),
    valid_to: formatDateValue(row.valid_to || ""),
    grant_type: String(row.grant_type || ""),
    year: row.year || "",
    fifo_generated_lots: currentLots,
    grant_days_derived_balance: grantDays,
    carry_over_days_derived_balance: carryOverDays,
    double_count_suspected: grantDays > 0 && carryOverDays > 0,
    warning_codes: grantDays > 0 && carryOverDays > 0
      ? ["OPENING_BALANCE_DOUBLE_COUNT_SUSPECTED"]
      : [],
    current_fifo: summarizePartnerOpeningBalanceScenario_(currentFifo, asOfDate),
    yearly_balance: legacyBalance,
    yearly_balance_difference: {
      current_remaining_days: Number(currentFifo.current_remaining_days || 0) -
        Number(legacyBalance.current_remaining_days || 0),
      expired_days: Number(currentFifo.expired_days || 0) -
        Number(legacyBalance.expired_days || 0)
    },
    correction_candidates: {
      plan_a_grant_days_only: simulatePartnerOpeningBalanceScenario_(
        employeeId, grantId, "GRANT_DAYS_ONLY", asOfDate, context
      ),
      plan_b_carry_over_days_only: simulatePartnerOpeningBalanceScenario_(
        employeeId, grantId, "CARRY_OVER_DAYS_ONLY", asOfDate, context
      ),
      plan_c_split_lots: buildPartnerOpeningBalanceSplitPlan_(grantDays, carryOverDays)
    }
  };
}

function simulatePartnerOpeningBalanceScenario_(employeeId, grantId, scenario, asOfDate, context) {
  const grantsByEmployee = {};
  Object.keys(context.grants_by_employee || {}).forEach(id => {
    grantsByEmployee[id] = (context.grants_by_employee[id] || []).map(row => {
      const copy = Object.assign({}, row);
      if (id === employeeId && String(copy.grant_id || "") === grantId) {
        if (scenario === "GRANT_DAYS_ONLY") copy.carry_over_days = 0;
        if (scenario === "CARRY_OVER_DAYS_ONLY") copy.grant_days = 0;
        copy.total_days = Number(copy.grant_days || 0) + Number(copy.carry_over_days || 0);
      }
      return copy;
    });
  });
  const scenarioContext = {
    as_of_date: context.as_of_date,
    calendar_map: context.calendar_map,
    requests_by_employee: context.requests_by_employee,
    grants_by_employee: grantsByEmployee
  };
  const fifoBalance = calculateFifoBalanceWithOpeningBalanceFromContext_(
    employeeId,
    asOfDate,
    scenarioContext
  );
  return summarizePartnerOpeningBalanceScenario_(fifoBalance, asOfDate);
}

function summarizePartnerOpeningBalanceScenario_(fifoBalance, asOfDate) {
  const fifoView = buildPaidLeaveGrantScheduleFifoView_(fifoBalance, asOfDate);
  const nextLot = (fifoView.lots || []).find(lot => lot.consumption_priority === 1) || null;
  return {
    current_active_remaining_days: Number(fifoBalance.current_remaining_days || 0),
    expired_days: Number(fifoBalance.expired_days || 0),
    next_consumption_lot: nextLot ? {
      grant_id: nextLot.grant_id,
      grant_date: nextLot.grant_date,
      valid_to: nextLot.valid_to,
      remaining_days: nextLot.remaining_days,
      consumption_priority: nextLot.consumption_priority
    } : null
  };
}

function buildPartnerOpeningBalanceSplitPlan_(grantDays, carryOverDays) {
  if (grantDays > 0 && carryOverDays > 0) {
    return {
      status: "REQUIRES_MANUAL_BREAKDOWN",
      message: "元の付与日・日数・期限の内訳が記録されていないため、安全な複数ロット試算はできません。根拠資料で内訳を確認してください。"
    };
  }
  return {
    status: "NOT_NEEDED",
    message: "片方のみ正数のため、分割する初期残高内訳はありません。"
  };
}

const PARTNER_OPENING_BALANCE_REPAIR_CONFIRMATION_ = "PARTNER_OPENING_BALANCE_REPAIR_2026_07_25";
const PARTNER_OPENING_BALANCE_REPAIR_TARGETS_ = [
  { grant_id: "G0056", employee_id: "EMP0060", grant_days: 15, carry_over_days: 20 },
  { grant_id: "G0057", employee_id: "EMP0061", grant_days: 15, carry_over_days: 20 }
];
/*
 * PARTNER初期導入残高補正
 *
 * 実施日: 2026-07-25
 *
 * 対象:
 * - G0056
 * - G0057
 *
 * 一回限りのデータ補正は完了済み。
 * 再実行禁止。
 */
const PARTNER_OPENING_BALANCE_FISCAL_START_REPAIR_COMPLETED_ERROR_ =
  "この補正は2026-07-25に完了済みです。再実行は禁止されています。";
const PARTNER_OPENING_BALANCE_FISCAL_START_REPAIR_TARGETS_ = [
  {
    grant_id: "G0056", employee_id: "EMP0060", grant_days: 15, carry_over_days: 20,
    valid_from: "2025-06-01", valid_to: "2026-05-31", display_employee_id: "P0002"
  },
  {
    grant_id: "G0057", employee_id: "EMP0061", grant_days: 15, carry_over_days: 20,
    valid_from: "2025-06-01", valid_to: "2026-05-31", display_employee_id: "P0003"
  }
];
const PARTNER_OPENING_BALANCE_FISCAL_START_REPAIR_AFTER_ = {
  grant_days: 0,
  carry_over_days: 20,
  valid_from: "2026-06-01",
  valid_to: "2027-05-31"
};
const PARTNER_OPENING_BALANCE_P0004_REPAIR_MARKER_ =
  "[一回限り補正: PARTNER初期導入残高の年度開始繰越化]";
/*
 * 2026-07-25
 * P0004(G0058)補正完了
 *
 * grant_days:      3 → 0
 * carry_over_days: 0 → 2
 * valid_from:      2025/06/01 → 2026/06/01
 * valid_to:        2026/05/31 → 2027/05/31
 *
 * 実データ補正済み。再実行禁止。
 */
const PARTNER_OPENING_BALANCE_P0004_REPAIR_COMPLETED_ERROR_ =
  "この補正は2026-07-25に完了済みです。\nP0004(G0058)への再実行は禁止されています。";
const PARTNER_OPENING_BALANCE_P0004_REPAIR_TARGET_ = {
  grant_id: "G0058", employee_id: "EMP0062", display_employee_id: "P0004",
  company_code: "PARTNER", grant_days: 3, carry_over_days: 0,
  valid_from: "2025-06-01", valid_to: "2026-05-31"
};
const PARTNER_OPENING_BALANCE_P0004_REPAIR_AFTER_ = {
  grant_days: 0, carry_over_days: 2, valid_from: "2026-06-01", valid_to: "2027-05-31"
};

/**
 * 旧補正関数。前提解釈の変更により本実行を永久禁止し、読取り試算だけを返す。
 */
function repairPartnerOpeningBalanceDoubleCount(options) {
  const opts = options || {};
  if (opts.dry_run === false) {
    throw new Error("この補正は前提解釈の変更により禁止されています。confirmation_text の値にかかわらず書込みは実行しません。");
  }
  return debugPartnerOpeningBalanceCarryOverSimulation(opts.as_of_date);
}

function assertPartnerOpeningBalanceRepairConfirmation_(dryRun, confirmationText) {
  if (!dryRun) {
    throw new Error("旧補正方針は無効です。confirmation_text の値にかかわらず本実行を拒否します。");
  }
}

/**
 * 完了済み補正の互換入口。dry-runだけを読み取り専用で残す。
 * dry_run:false は確認文字列にかかわらず、必ず拒否する。
 */
function repairPartnerOpeningBalanceFiscalStart(options) {
  const opts = options || {};
  const dryRun = opts.dry_run !== false;
  if (!dryRun) {
    throw new Error(PARTNER_OPENING_BALANCE_FISCAL_START_REPAIR_COMPLETED_ERROR_);
  }
  const asOfDate = opts.as_of_date
    ? parseLocalDate(opts.as_of_date)
    : parseLocalDate("2026-07-25");
  // 補正完了後の現値を読み取り、残高確認用のdry-runだけを提供する。
  const dryRunState = readPartnerOpeningBalanceFiscalStartRepairState_("after");
  return buildPartnerOpeningBalanceFiscalStartRepairDryRun_(dryRunState, asOfDate);
}

/** 完了済みP0004補正の互換入口。dry-run診断だけを読み取り専用で残す。 */
function repairPartnerOpeningBalanceFiscalStartP0004(options) {
  const opts = options || {};
  const dryRun = opts.dry_run !== false;
  const asOfDate = opts.as_of_date ? parseLocalDate(opts.as_of_date) : parseLocalDate("2026-07-25");
  if (!dryRun) throw new Error(PARTNER_OPENING_BALANCE_P0004_REPAIR_COMPLETED_ERROR_);
  return buildPartnerOpeningBalanceP0004RepairDryRun_(
    readPartnerOpeningBalanceP0004RepairState_("after"), asOfDate
  );
}

function readPartnerOpeningBalanceP0004RepairState_(phase) {
  const sheet = getSheet("paid_leave_grants");
  const headerInfo = requireHeaders(sheet, [
    "grant_id", "employee_id", "grant_days", "carry_over_days", "valid_from", "valid_to", "notes"
  ]);
  const rowsByGrantId = {};
  sheet.getDataRange().getValues().slice(1).forEach((values, index) => {
    const row = rowToObject(values, headerInfo.headers);
    rowsByGrantId[String(row.grant_id || "").trim()] = Object.assign({}, row, { row_number: index + 2 });
  });
  const employeeMap = {};
  getEmployeesForAdmin().forEach(employee => { employeeMap[String(employee.employee_id || "").trim()] = employee; });
  const state = {
    target_row: rowsByGrantId.G0058 || null,
    g0061_row: rowsByGrantId.G0061 || null,
    employee_map: employeeMap
  };
  validatePartnerOpeningBalanceP0004RepairState_(state, phase || "before");
  return state;
}

/** P0004補正の事前条件を可視化する読み取り専用診断。 */
function debugPartnerOpeningBalanceFiscalStartRepairPreconditionsP0004() {
  const sheet = getSheet("paid_leave_grants");
  const headerInfo = requireHeaders(sheet, [
    "grant_id", "employee_id", "grant_days", "carry_over_days", "valid_from", "valid_to",
    "grant_type", "year", "notes"
  ]);
  const employeeMap = {};
  getEmployeesForAdmin().forEach(employee => {
    employeeMap[normalizePartnerOpeningBalanceFiscalStartText_(employee.employee_id)] = employee;
  });
  const matchingRows = sheet.getDataRange().getValues().slice(1)
    .map((values, index) => Object.assign({}, rowToObject(values, headerInfo.headers), { row_number: index + 2 }))
    .filter(row => normalizePartnerOpeningBalanceFiscalStartText_(row.grant_id) === "G0058");
  const row = matchingRows.length === 1 ? matchingRows[0] : null;
  const employee = row
    ? employeeMap[normalizePartnerOpeningBalanceFiscalStartText_(row.employee_id)] || {}
    : {};
  const target = buildPartnerOpeningBalanceFiscalStartPreconditionDiagnostic_(
    row, employee, PARTNER_OPENING_BALANCE_P0004_REPAIR_TARGET_, matchingRows.length
  );
  const result = {
    ok: true,
    read_only: true,
    operation: "PARTNER_OPENING_BALANCE_P0004_REPAIR_PRECONDITIONS",
    target_grant_ids: ["G0058"],
    all_conditions_match: target.all_conditions_match,
    targets: [target]
  };
  logJsonInChunks_("[P0004_REPAIR][PRECONDITIONS]", result);
  return result;
}

function validatePartnerOpeningBalanceP0004RepairState_(state, phase) {
  const target = state.target_row;
  const expected = phase === "after" ? PARTNER_OPENING_BALANCE_P0004_REPAIR_AFTER_ : PARTNER_OPENING_BALANCE_P0004_REPAIR_TARGET_;
  const employee = (state.employee_map || {})[String(target && target.employee_id || "").trim()] || {};
  const diagnosticTarget = Object.assign({}, PARTNER_OPENING_BALANCE_P0004_REPAIR_TARGET_, expected);
  const diagnostic = buildPartnerOpeningBalanceFiscalStartPreconditionDiagnostic_(
    target, employee, diagnosticTarget, target ? 1 : 0
  );
  if (!diagnostic.all_conditions_match) {
    throw new Error(
      "P0004補正の事前条件が一致しません: G0058\n\n不一致項目\n" +
      diagnostic.mismatch_fields.join("\n")
    );
  }
  if (!state.g0061_row || String(state.g0061_row.employee_id || "") !== "EMP0062") {
    throw new Error("G0061の存在またはemployee_idを確認できません。書込みは行いませんでした。");
  }
  return true;
}

function buildPartnerOpeningBalanceP0004RepairDryRun_(state, asOfDate) {
  const employeeId = "EMP0062";
  const context = createFifoBalanceComparisonContext_(asOfDate, { read_only: true });
  const before = calculateFifoBalanceWithOpeningBalanceFromContext_(employeeId, asOfDate, context);
  const afterContext = buildPartnerP0004ScenarioContext_(employeeId, "G0058", context, PARTNER_OPENING_BALANCE_P0004_REPAIR_AFTER_);
  const after = calculateFifoBalanceWithOpeningBalanceFromContext_(employeeId, asOfDate, afterContext);
  const fiscalYear = getFiscalYearFromDateWithStart(asOfDate, 6);
  const yearly = calculateLegacyBalanceFromFifoContext_(employeeId, fiscalYear, 6, afterContext);
  const afterActive = buildPartnerP0004ActiveBalanceBreakdown_(after);
  const may2 = (after.allocations || []).filter(row => String(row.use_date || "") === "2026-05-02");
  const fiscalAllocations = (after.allocations || []).filter(row => String(row.use_date || "") >= "2026-06-01");
  const carryAllocations = fiscalAllocations.filter(row => String(row.grant_id || "") === "G0058#opening_balance");
  return {
    ok: true,
    dry_run: true,
    write_disabled: true,
    as_of_date: formatDateValue(asOfDate),
    target: {
      grant_id: "G0058", employee_id: "EMP0062", display_employee_id: "P0004",
      before: { grant_days: Number(state.target_row.grant_days || 0), carry_over_days: Number(state.target_row.carry_over_days || 0), valid_from: formatDateValue(state.target_row.valid_from), valid_to: formatDateValue(state.target_row.valid_to) },
      after: PARTNER_OPENING_BALANCE_P0004_REPAIR_AFTER_
    },
    fifo_before: summarizePartnerOpeningBalanceScenario_(before, asOfDate),
    fifo_after: summarizePartnerOpeningBalanceScenario_(after, asOfDate),
    yearly_balance_after: yearly,
    difference_from_yearly_balance: Number(after.current_remaining_days || 0) - Number(yearly.current_remaining_days || 0),
    allocations_on_2026_05_02: may2,
    fiscal_allocations: fiscalAllocations,
    g0058_carry_over_consumed_days: carryAllocations.reduce((sum, row) => sum + Number(row.consumed_days || 0), 0),
    next_consumption_lot: afterActive[0] || null,
    g0061_before: sanitizePartnerP0004Grant_(state.g0061_row),
    warnings: ["dry-runではシート、notes、usage_log、キャッシュを変更しません。"]
  };
}

/** Apps Script関数一覧から実行するP0004補正のdry-runラッパー。 */
function debugRepairPartnerOpeningBalanceFiscalStartP0004() {
  let stage = "dry-run開始";
  try {
    const result = repairPartnerOpeningBalanceFiscalStartP0004({ dry_run: true });
    stage = "dry-run結果のログ出力";
    logPartnerOpeningBalanceP0004RepairDryRun_(result);
    return result;
  } catch (error) {
    logPartnerOpeningBalanceP0004RepairWrapperFailure_(error, stage);
    throw error;
  }
}

/** 完了済み補正の旧本実行入口。互換性のため残すが、永久に拒否する。 */
function executeRepairPartnerOpeningBalanceFiscalStartP0004() {
  throw new Error(PARTNER_OPENING_BALANCE_P0004_REPAIR_COMPLETED_ERROR_);
}

function logPartnerOpeningBalanceP0004RepairDryRun_(result) {
  const target = result.target || {};
  const before = target.before || {};
  const after = target.after || {};
  Logger.log([
    "=== P0004 補正済み状態 dry-run ===",
    "対象: " + String(target.grant_id || "G0058") + " / " + String(target.employee_id || "EMP0062"),
    "確認値",
    "grant_days: " + String(before.grant_days) + " → " + String(after.grant_days),
    "carry_over_days: " + String(before.carry_over_days) + " → " + String(after.carry_over_days),
    "valid_from: " + String(before.valid_from) + " → " + String(after.valid_from),
    "valid_to: " + String(before.valid_to) + " → " + String(after.valid_to),
    "補正後FIFO",
    "現在残高: " + String((result.fifo_after || {}).current_active_remaining_days),
    "期限切れ: " + String((result.fifo_after || {}).expired_days),
    "年度残高との差: " + String(result.difference_from_yearly_balance)
  ].join("\n"));
  logJsonInChunks_("[P0004_REPAIR][BASIC]", { target: { grant_id: target.grant_id, employee_id: target.employee_id, display_employee_id: target.display_employee_id }, as_of_date: result.as_of_date });
  logJsonInChunks_("[P0004_REPAIR][BEFORE]", before);
  logJsonInChunks_("[P0004_REPAIR][AFTER]", after);
  logJsonInChunks_("[P0004_REPAIR][DIFFERENCE]", { fifo_before: result.fifo_before, fifo_after: result.fifo_after, difference_from_yearly_balance: result.difference_from_yearly_balance });
  logJsonInChunks_("[P0004_REPAIR][WARNINGS]", result.warnings || []);
}

function logPartnerOpeningBalanceP0004RepairWrapperFailure_(error, stage) {
  Logger.log([
    "=== 補正失敗 ===",
    "エラー: " + String(error && error.message || error),
    "grant_id: G0058",
    "処理段階: " + String(stage || "不明")
  ].join("\n"));
}

function readPartnerOpeningBalanceFiscalStartRepairState_(phase) {
  const sheet = getSheet("paid_leave_grants");
  const headerInfo = requireHeaders(sheet, [
    "grant_id", "employee_id", "grant_days", "carry_over_days", "valid_from", "valid_to", "notes"
  ]);
  const data = sheet.getDataRange().getValues();
  const employeeMap = {};
  getEmployeesForAdmin().forEach(employee => {
    employeeMap[String(employee.employee_id || "").trim()] = employee;
  });
  const targetByGrantId = {};
  PARTNER_OPENING_BALANCE_FISCAL_START_REPAIR_TARGETS_.forEach(target => {
    targetByGrantId[target.grant_id] = true;
  });
  const rows = data.slice(1).reduce((result, values, index) => {
    const row = rowToObject(values, headerInfo.headers);
    if (targetByGrantId[String(row.grant_id || "").trim()]) {
      result.push(Object.assign({}, row, { row_number: index + 2 }));
    }
    return result;
  }, []);
  validatePartnerOpeningBalanceFiscalStartRepairRows_(rows, employeeMap, phase || "before");
  return {
    rows: rows,
    employee_map: employeeMap
  };
}

/**
 * 本補正の事前条件だけを診断する読み取り専用関数。
 * notes本文は返さず、初期導入残高マーカーの有無だけを返す。
 */
function debugPartnerOpeningBalanceFiscalStartRepairPreconditions() {
  const sheet = getSheet("paid_leave_grants");
  const headerInfo = requireHeaders(sheet, [
    "grant_id", "employee_id", "grant_days", "carry_over_days", "valid_from", "valid_to",
    "grant_type", "year", "notes"
  ]);
  const employeeMap = {};
  getEmployeesForAdmin().forEach(employee => {
    employeeMap[normalizePartnerOpeningBalanceFiscalStartText_(employee.employee_id)] = employee;
  });
  const rowsByGrantId = {};
  sheet.getDataRange().getValues().slice(1).forEach((values, index) => {
    const row = rowToObject(values, headerInfo.headers);
    const grantId = normalizePartnerOpeningBalanceFiscalStartText_(row.grant_id);
    if (grantId) {
      if (!rowsByGrantId[grantId]) rowsByGrantId[grantId] = [];
      rowsByGrantId[grantId].push(Object.assign({}, row, { row_number: index + 2 }));
    }
  });
  const targets = PARTNER_OPENING_BALANCE_FISCAL_START_REPAIR_TARGETS_.map(target => {
    const matchingRows = rowsByGrantId[normalizePartnerOpeningBalanceFiscalStartText_(target.grant_id)] || [];
    const row = matchingRows.length === 1 ? matchingRows[0] : null;
    const employee = row
      ? employeeMap[normalizePartnerOpeningBalanceFiscalStartText_(row.employee_id)] || {}
      : {};
    return buildPartnerOpeningBalanceFiscalStartPreconditionDiagnostic_(row, employee, target, matchingRows.length);
  });
  const result = {
    ok: true,
    read_only: true,
    operation: "PARTNER_OPENING_BALANCE_FISCAL_START_REPAIR_PRECONDITIONS",
    target_grant_ids: PARTNER_OPENING_BALANCE_FISCAL_START_REPAIR_TARGETS_.map(target => target.grant_id),
    all_conditions_match: targets.every(target => target.all_conditions_match),
    targets: targets
  };
  logJsonInChunks_("[PARTNER_OPENING_BALANCE_FISCAL_START_REPAIR][PRECONDITIONS]", result);
  return result;
}

function buildPartnerOpeningBalanceFiscalStartPreconditionDiagnostic_(row, employee, target, matchingRowCount) {
  const source = row || {};
  const employeeRow = employee || {};
  const checks = {
    grant_id: buildPartnerOpeningBalanceFiscalStartTextCheck_(source.grant_id, target.grant_id),
    employee_id: buildPartnerOpeningBalanceFiscalStartTextCheck_(source.employee_id, target.employee_id),
    display_employee_id: buildPartnerOpeningBalanceFiscalStartTextCheck_(
      employeeRow.display_employee_id, target.display_employee_id
    ),
    company_code: buildPartnerOpeningBalanceFiscalStartTextCheck_(employeeRow.company_code, "PARTNER"),
    grant_days: buildPartnerOpeningBalanceFiscalStartNumberCheck_(source.grant_days, target.grant_days),
    carry_over_days: buildPartnerOpeningBalanceFiscalStartNumberCheck_(source.carry_over_days, target.carry_over_days),
    valid_from: buildPartnerOpeningBalanceFiscalStartDateCheck_(source.valid_from, target.valid_from),
    valid_to: buildPartnerOpeningBalanceFiscalStartDateCheck_(source.valid_to, target.valid_to),
    // grant_type/year are read-only reference information, not execution preconditions.
    grant_type: buildPartnerOpeningBalanceFiscalStartInformationalCheck_(source.grant_type),
    year: buildPartnerOpeningBalanceFiscalStartInformationalCheck_(source.year),
    notes_marker: buildPartnerOpeningBalanceFiscalStartNotesMarkerCheck_(source.notes),
    matching_grant_id_row_count: {
      actual: Number(matchingRowCount || 0),
      expected: 1,
      actual_normalized: Number(matchingRowCount || 0),
      expected_normalized: 1,
      matched: Number(matchingRowCount || 0) === 1
    }
  };
  checks.display_employee_id.required_for_execution = false;
  const requiredKeys = [
    "grant_id", "employee_id", "company_code", "grant_days",
    "carry_over_days", "valid_from", "valid_to", "notes_marker", "matching_grant_id_row_count"
  ];
  const mismatchFields = requiredKeys.filter(key => checks[key].matched !== true);
  return {
    grant_id: String(target.grant_id || ""),
    employee_id: String(target.employee_id || ""),
    display_employee_id: String(target.display_employee_id || ""),
    all_conditions_match: mismatchFields.length === 0,
    mismatch_fields: mismatchFields,
    checks: checks
  };
}

function buildPartnerOpeningBalanceFiscalStartTextCheck_(actual, expected) {
  const actualText = actual == null ? "" : String(actual);
  const expectedText = expected == null ? "" : String(expected);
  const actualNormalized = normalizePartnerOpeningBalanceFiscalStartText_(actualText);
  const expectedNormalized = normalizePartnerOpeningBalanceFiscalStartText_(expectedText);
  return {
    actual: actualText,
    expected: expectedText,
    actual_normalized: actualNormalized,
    expected_normalized: expectedNormalized,
    matched: actualNormalized !== "" && actualNormalized === expectedNormalized
  };
}

function buildPartnerOpeningBalanceFiscalStartNumberCheck_(actual, expected) {
  const actualNormalized = normalizePartnerOpeningBalanceFiscalStartNumber_(actual);
  const expectedNormalized = normalizePartnerOpeningBalanceFiscalStartNumber_(expected);
  return {
    actual: actual == null ? null : actual,
    expected: expected == null ? null : expected,
    actual_normalized: actualNormalized.value,
    expected_normalized: expectedNormalized.value,
    actual_valid_number: actualNormalized.valid,
    expected_valid_number: expectedNormalized.valid,
    matched: actualNormalized.valid && expectedNormalized.valid && actualNormalized.value === expectedNormalized.value
  };
}

function buildPartnerOpeningBalanceFiscalStartDateCheck_(actual, expected) {
  const actualNormalized = normalizePartnerOpeningBalanceFiscalStartDate_(actual);
  const expectedNormalized = normalizePartnerOpeningBalanceFiscalStartDate_(expected);
  return {
    actual: actualNormalized.display_value,
    expected: expectedNormalized.display_value,
    actual_type: actualNormalized.type,
    expected_type: expectedNormalized.type,
    actual_normalized: actualNormalized.value,
    expected_normalized: expectedNormalized.value,
    matched: actualNormalized.valid && expectedNormalized.valid && actualNormalized.value === expectedNormalized.value
  };
}

function buildPartnerOpeningBalanceFiscalStartInformationalCheck_(actual) {
  const actualText = actual == null ? "" : String(actual);
  return {
    actual: actualText,
    expected: null,
    actual_normalized: normalizePartnerOpeningBalanceFiscalStartText_(actualText),
    expected_normalized: null,
    matched: null,
    required_for_execution: false
  };
}

function buildPartnerOpeningBalanceFiscalStartNotesMarkerCheck_(notes) {
  const raw = notes == null ? "" : String(notes);
  const normalized = normalizePartnerOpeningBalanceFiscalStartText_(raw);
  const marker = normalizePartnerOpeningBalanceFiscalStartText_("初期導入残高");
  const contains = raw.indexOf("初期導入残高") !== -1;
  return {
    notes_present: raw.trim() !== "",
    contains_initial_opening_balance_marker: contains,
    normalized_marker_matches: normalized.indexOf(marker) !== -1,
    matched: contains && normalized.indexOf(marker) !== -1
  };
}

function normalizePartnerOpeningBalanceFiscalStartText_(value) {
  return String(value == null ? "" : value).trim().normalize("NFKC").toUpperCase();
}

function normalizePartnerOpeningBalanceFiscalStartNumber_(value) {
  if (value === null || value === undefined || String(value).trim() === "") {
    return { valid: false, value: null };
  }
  const numberValue = Number(value);
  return Number.isFinite(numberValue)
    ? { valid: true, value: numberValue }
    : { valid: false, value: null };
}

function normalizePartnerOpeningBalanceFiscalStartDate_(value) {
  const type = value instanceof Date ? "Date" : typeof value;
  if (value === null || value === undefined || String(value).trim() === "") {
    return { valid: false, value: null, display_value: "", type: type };
  }
  try {
    const parsed = parseLocalDate(value);
    if (!(parsed instanceof Date) || isNaN(parsed.getTime())) {
      return { valid: false, value: null, display_value: String(value), type: type };
    }
    const normalized = formatDateValue(parsed);
    return { valid: !!normalized, value: normalized || null, display_value: normalized || String(value), type: type };
  } catch (error) {
    return { valid: false, value: null, display_value: String(value), type: type };
  }
}

// シート非依存の全件検証。before / after とも対象外行を更新しない前提を守る。
function validatePartnerOpeningBalanceFiscalStartRepairRows_(rows, employeeMap, phase) {
  const actualRows = Array.isArray(rows) ? rows : [];
  const expectedAfter = phase === "after";
  if (actualRows.length !== PARTNER_OPENING_BALANCE_FISCAL_START_REPAIR_TARGETS_.length) {
    throw new Error("補正対象行数が想定と一致しません。全体を中止しました。");
  }
  const seen = {};
  PARTNER_OPENING_BALANCE_FISCAL_START_REPAIR_TARGETS_.forEach(target => {
    const row = actualRows.find(item => String(item.grant_id || "").trim() === target.grant_id);
    if (!row || seen[target.grant_id]) {
      throw new Error("対象grant_idが不足または重複しています: " + target.grant_id);
    }
    seen[target.grant_id] = true;
    const employeeId = String(row.employee_id || "").trim();
    const employee = (employeeMap || {})[employeeId] || {};
    const expected = expectedAfter ? PARTNER_OPENING_BALANCE_FISCAL_START_REPAIR_AFTER_ : target;
    const diagnosticTarget = Object.assign({}, target, expected);
    const diagnostic = buildPartnerOpeningBalanceFiscalStartPreconditionDiagnostic_(
      row, employee, diagnosticTarget, 1
    );
    if (!diagnostic.all_conditions_match) {
      throw new Error(
        "対象行の事前条件が一致しません: " + target.grant_id +
        " / 不一致項目: " + diagnostic.mismatch_fields.join(", ")
      );
    }
  });
  return true;
}

function buildPartnerOpeningBalanceFiscalStartRepairDryRun_(state, asOfDate) {
  const context = createFifoBalanceComparisonContext_(asOfDate, { read_only: true });
  const fiscalStart = parseLocalDate(PARTNER_OPENING_BALANCE_FISCAL_START_REPAIR_AFTER_.valid_from);
  const fiscalEnd = parseLocalDate(PARTNER_OPENING_BALANCE_FISCAL_START_REPAIR_AFTER_.valid_to);
  const fiscalYear = getFiscalYearFromDateWithStart(asOfDate, 6);
  const targets = PARTNER_OPENING_BALANCE_FISCAL_START_REPAIR_TARGETS_.map(target => {
    const row = state.rows.find(item => String(item.grant_id || "") === target.grant_id);
    const employeeId = target.employee_id;
    const beforeFifo = calculateFifoBalanceWithOpeningBalanceFromContext_(employeeId, asOfDate, context);
    const afterContext = buildPartnerOpeningBalanceCarryOnlyScenarioContext_(
      employeeId, target.grant_id, context, fiscalEnd, fiscalStart
    );
    const afterFifo = calculateFifoBalanceWithOpeningBalanceFromContext_(employeeId, asOfDate, afterContext);
    const employee = state.employee_map[employeeId] || {};
    const yearlyBalance = calculateLegacyBalanceFromFifoContext_(
      employeeId, fiscalYear, Number(employee.fiscal_start_month || 6), afterContext
    );
    const afterSummary = summarizePartnerOpeningBalanceScenario_(afterFifo, asOfDate);
    return {
      grant_id: target.grant_id,
      employee_id: employeeId,
      display_employee_id: String(employee.display_employee_id || ""),
      before: {
        grant_days: Number(row.grant_days || 0),
        carry_over_days: Number(row.carry_over_days || 0),
        valid_from: formatDateValue(row.valid_from),
        valid_to: formatDateValue(row.valid_to)
      },
      after: Object.assign({}, PARTNER_OPENING_BALANCE_FISCAL_START_REPAIR_AFTER_),
      fifo_before: summarizePartnerOpeningBalanceScenario_(beforeFifo, asOfDate),
      fifo_after: afterSummary,
      difference: {
        current_remaining_days: afterSummary.current_active_remaining_days - Number(beforeFifo.current_remaining_days || 0),
        expired_days: afterSummary.expired_days - Number(beforeFifo.expired_days || 0)
      },
      allocations_on_2026_05_02: getPartnerFiscalStartSimulationAllocations_(afterFifo, "2026/05/02"),
      allocations_on_2026_07_18: getPartnerFiscalStartSimulationAllocations_(afterFifo, "2026/07/18"),
      next_consumption_lot: afterSummary.next_consumption_lot,
      yearly_balance: yearlyBalance,
      difference_from_yearly_balance: {
        current_remaining_days: afterSummary.current_active_remaining_days -
          Number(yearlyBalance.current_remaining_days || 0),
        expired_days: afterSummary.expired_days - Number(yearlyBalance.expired_days || 0)
      },
      warnings: [
        "2026/05/02の未割当は、2025年度の取得として2026/06/01開始ロットから再消化しない意図的な結果です。",
        "P0004 / EMP0062 / G0058 は対象外です。"
      ]
    };
  });
  return {
    ok: true,
    dry_run: true,
    write_disabled: true,
    as_of_date: formatDateValue(asOfDate),
    operation: "PARTNER_OPENING_BALANCE_FISCAL_START_REPAIR",
    target_grant_ids: PARTNER_OPENING_BALANCE_FISCAL_START_REPAIR_TARGETS_.map(row => row.grant_id),
    targets: targets,
    warnings: ["dry-runではシート、キャッシュ、FIFOロジックを変更しません。"]
  };
}

/**
 * Apps Script関数一覧から実行する、年度開始補正のdry-run専用ラッパー。
 * 補正本体の戻り値を変えず、ログ整形だけを追加する。
 */
function debugRepairPartnerOpeningBalanceFiscalStart() {
  let stage = "dry-run開始";
  try {
    const result = repairPartnerOpeningBalanceFiscalStart({ dry_run: true });
    stage = "dry-run結果のログ出力";
    logPartnerOpeningBalanceFiscalStartRepairDryRunResult_(result);
    return result;
  } catch (error) {
    logPartnerOpeningBalanceFiscalStartRepairWrapperFailure_(error, stage);
    throw error;
  }
}

/**
 * 完了済み補正の旧本実行入口。常に再実行禁止エラーを返す。
 */
function executeRepairPartnerOpeningBalanceFiscalStart() {
  const stage = "完了済み補正の再実行禁止確認";
  try {
    return repairPartnerOpeningBalanceFiscalStart({ dry_run: false });
  } catch (error) {
    logPartnerOpeningBalanceFiscalStartRepairWrapperFailure_(error, stage);
    throw error;
  }
}

function logPartnerOpeningBalanceFiscalStartRepairDryRunResult_(result) {
  Logger.log(buildPartnerOpeningBalanceFiscalStartRepairDryRunSummary_(result));
  (result.targets || []).forEach(target => {
    const label = "[" + String(target.grant_id || "UNKNOWN") + "]";
    logJsonInChunks_(label + "[BASIC]", {
      grant_id: target.grant_id,
      employee_id: target.employee_id,
      display_employee_id: target.display_employee_id
    });
    logJsonInChunks_(label + "[BEFORE]", target.before);
    logJsonInChunks_(label + "[AFTER]", target.after);
    logJsonInChunks_(label + "[DIFFERENCE]", target.difference);
    logJsonInChunks_(label + "[WARNINGS]", target.warnings || []);
  });
  logJsonInChunks_("[PARTNER_REPAIR][WARNINGS]", result.warnings || []);
}

function buildPartnerOpeningBalanceFiscalStartRepairDryRunSummary_(result) {
  const targets = result && result.targets || [];
  const lines = [
    "=== PARTNER補正 dry-run ===",
    "対象件数: " + String(targets.length),
    "対象grant_id: " + targets.map(target => String(target.grant_id || "")).join(", "),
    "変更予定:"
  ];
  targets.forEach(target => {
    const before = target.before || {};
    const after = target.after || {};
    lines.push(
      "-------------------",
      String(target.grant_id || "UNKNOWN"),
      "grant_days: " + String(before.grant_days == null ? "-" : before.grant_days) +
        " → " + String(after.grant_days == null ? "-" : after.grant_days),
      "valid_from: " + String(before.valid_from || "-") + " ↓ " + String(after.valid_from || "-"),
      "valid_to: " + String(before.valid_to || "-") + " ↓ " + String(after.valid_to || "-")
    );
  });
  return lines.join("\n");
}

function logPartnerOpeningBalanceFiscalStartRepairWrapperFailure_(error, stage) {
  Logger.log(buildPartnerOpeningBalanceFiscalStartRepairWrapperFailureMessage_(error, stage));
}

function buildPartnerOpeningBalanceFiscalStartRepairWrapperFailureMessage_(error, stage) {
  return [
    "=== 補正失敗 ===",
    "エラー内容: " + String(error && error.message || error),
    "対象grant_id: " + PARTNER_OPENING_BALANCE_FISCAL_START_REPAIR_TARGETS_
      .map(row => row.grant_id).join(", "),
    "処理段階: " + String(stage || "不明"),
    "=================="
  ].join("\n");
}

function readPartnerOpeningBalanceRepairState_() {
  const sheet = getSheet("paid_leave_grants");
  const headerInfo = requireHeaders(sheet, [
    "grant_id", "employee_id", "grant_days", "carry_over_days", "notes"
  ]);
  const data = sheet.getDataRange().getValues();
  const employeeMap = {};
  getEmployeesForAdmin().forEach(employee => {
    employeeMap[String(employee.employee_id || "").trim()] = employee;
  });
  const targetByGrantId = {};
  PARTNER_OPENING_BALANCE_REPAIR_TARGETS_.forEach(target => { targetByGrantId[target.grant_id] = target; });
  const rows = [];
  data.slice(1).forEach((values, index) => {
    const row = rowToObject(values, headerInfo.headers);
    const grantId = String(row.grant_id || "").trim();
    if (!targetByGrantId[grantId]) return;
    rows.push(Object.assign({}, row, { row_number: index + 2 }));
  });
  validatePartnerOpeningBalanceRepairRows_(rows, employeeMap);
  return {
    sheet: sheet,
    carry_over_column: headerInfo.map.carry_over_days + 1,
    rows: rows,
    employee_map: employeeMap
  };
}

// シート非依存の全件検証。テストでも利用する。
function validatePartnerOpeningBalanceRepairRows_(rows, employeeMap) {
  const actualRows = Array.isArray(rows) ? rows : [];
  if (actualRows.length !== PARTNER_OPENING_BALANCE_REPAIR_TARGETS_.length) {
    throw new Error("補正対象行数が想定と一致しません。全体を中止しました。");
  }
  const seen = {};
  PARTNER_OPENING_BALANCE_REPAIR_TARGETS_.forEach(target => {
    const row = actualRows.find(item => String(item.grant_id || "").trim() === target.grant_id);
    if (!row || seen[target.grant_id]) {
      throw new Error("対象grant_idが不足または重複しています: " + target.grant_id);
    }
    seen[target.grant_id] = true;
    const employeeId = String(row.employee_id || "").trim();
    const employee = (employeeMap || {})[employeeId] || {};
    if (employeeId !== target.employee_id ||
      String(employee.company_code || "").trim().toUpperCase() !== "PARTNER" ||
      Number(row.grant_days || 0) !== target.grant_days ||
      Number(row.carry_over_days || 0) !== target.carry_over_days ||
      String(row.notes || "").indexOf("初期導入残高") === -1) {
      throw new Error("対象行の事前条件が一致しません: " + target.grant_id);
    }
  });
  return true;
}

function debugPartnerOpeningBalanceCarryOverSimulation(asOfDateValue) {
  const result = buildPartnerOpeningBalanceCarryOverSimulation_(asOfDateValue);
  (result.targets || []).forEach(logPartnerCarryOverSimulationTarget_);
  return result;
}

function debugPartnerOpeningBalanceCarryOverSimulationForEmployee_(displayEmployeeId, asOfDateValue) {
  const result = buildPartnerOpeningBalanceCarryOverSimulation_(asOfDateValue);
  const target = selectPartnerCarryOverSimulationTarget_(result.targets, displayEmployeeId);
  logPartnerCarryOverSimulationTarget_(target);
  return Object.assign({}, result, { targets: [target] });
}

function selectPartnerCarryOverSimulationTarget_(targets, displayEmployeeId) {
  const normalizedId = normalizePaidLeaveDebugEmployeeId_(displayEmployeeId);
  const target = (targets || []).find(row =>
    normalizePaidLeaveDebugEmployeeId_(row.display_employee_id) === normalizedId ||
    normalizePaidLeaveDebugEmployeeId_(row.employee_id) === normalizedId
  );
  if (!target) throw new Error("対象の表示employee_idが見つかりません: " + String(displayEmployeeId || ""));
  return target;
}

function debugPartnerCarryOverSimulationP0002() {
  return debugPartnerOpeningBalanceCarryOverSimulationForEmployee_("P0002", "2026-07-25");
}

function debugPartnerCarryOverSimulationP0003() {
  return debugPartnerOpeningBalanceCarryOverSimulationForEmployee_("P0003", "2026-07-25");
}

/**
 * PARTNER P0004（EMP0062 / G0058）の初期導入残高を調査する読み取り専用診断。
 * 実データ、キャッシュ、FIFOロジックは変更しない。
 */
function debugPartnerCarryOverSimulationP0004() {
  const asOfDate = parseLocalDate("2026-07-25");
  const employeeId = "EMP0062";
  const employee = getEmployeesForAdmin().find(row =>
    String(row.employee_id || "").trim() === employeeId &&
    String(row.display_employee_id || "").trim() === "P0004"
  ) || {};
  const context = createFifoBalanceComparisonContext_(asOfDate, { read_only: true });
  const result = buildPartnerP0004FifoDiagnosis_(employee, context, asOfDate);
  logPartnerP0004FifoDiagnosis_(result);
  return result;
}

function buildPartnerP0004FifoDiagnosis_(employee, context, asOfDate) {
  const employeeId = "EMP0062";
  const grantId = "G0058";
  const fiscalStartMonth = Number((employee || {}).fiscal_start_month || 6);
  const fiscalYear = getFiscalYearFromDateWithStart(asOfDate, fiscalStartMonth);
  const grants = context.grants_by_employee[employeeId] || [];
  const targetGrant = grants.find(row => String(row.grant_id || "") === grantId) || null;
  const currentFifo = calculateFifoBalanceWithOpeningBalanceFromContext_(employeeId, asOfDate, context);
  const fiscalUsage = buildPartnerP0004FiscalUsage_(employeeId, context, asOfDate, fiscalYear, fiscalStartMonth);
  const yearlyBalance = calculateLegacyBalanceFromFifoContext_(
    employeeId, fiscalYear, fiscalStartMonth, context
  );
  const expiredBreakdown = buildPartnerP0004ExpiredBreakdown_(currentFifo);
  const activeBalanceBreakdown = buildPartnerP0004ActiveBalanceBreakdown_(currentFifo);
  const nearestExpiryLot = activeBalanceBreakdown.length > 0 ? activeBalanceBreakdown[0] : null;
  const simulations = {
    plan_a_current_is_correct: buildPartnerP0004Scenario_(
      employeeId, grantId, context, asOfDate, fiscalYear, fiscalStartMonth,
      { grant_days: 3, carry_over_days: 0, valid_from: "2025-06-01", valid_to: "2026-05-31" }
    ),
    plan_b_fiscal_start_carry_over: buildPartnerP0004Scenario_(
      employeeId, grantId, context, asOfDate, fiscalYear, fiscalStartMonth,
      { grant_days: 0, carry_over_days: 3, valid_from: "2026-06-01", valid_to: "2027-05-31" }
    ),
    plan_c_extend_initial_lot_expiry: buildPartnerP0004Scenario_(
      employeeId, grantId, context, asOfDate, fiscalYear, fiscalStartMonth,
      { grant_days: 3, carry_over_days: 0, valid_from: "2025-06-01", valid_to: "2027-05-31" }
    )
  };
  const warnings = [];
  if (!targetGrant) warnings.push("G0058がpaid_leave_grantsに見つかりません。");
  if (String((employee || {}).display_employee_id || "") !== "P0004") {
    warnings.push("社員マスターのdisplay_employee_idがP0004と一致しません。");
  }
  if (Number(currentFifo.current_remaining_days || 0) !== sumPartnerP0004Days_(activeBalanceBreakdown, "remaining_days")) {
    warnings.push("有効残高内訳の合計とFIFO現在残高が一致しません。");
  }
  return {
    ok: warnings.length === 0,
    read_only: true,
    as_of_date: formatDateValue(asOfDate),
    employee: sanitizePaidLeaveDebugEmployee_(employee),
    target: { employee_id: employeeId, display_employee_id: "P0004", grant_id: grantId },
    grants: grants.map(sanitizePartnerP0004Grant_),
    requests: fiscalUsage.requests,
    current_fifo: sanitizePartnerP0004Fifo_(currentFifo),
    expired_days_breakdown: expiredBreakdown,
    expired_days_conclusion: {
      expired_days: Number(currentFifo.expired_days || 0),
      is_g0058_derived: expiredBreakdown.some(row => row.source_grant_id === grantId),
      source_grant_ids: expiredBreakdown.map(row => row.source_grant_id)
    },
    active_balance_breakdown: activeBalanceBreakdown,
    active_balance_total: sumPartnerP0004Days_(activeBalanceBreakdown, "remaining_days"),
    fiscal_usage: fiscalUsage,
    yearly_balance: yearlyBalance,
    nearest_expiry_lot: nearestExpiryLot,
    g0058_evaluation: buildPartnerP0004G0058Evaluation_(targetGrant, currentFifo),
    simulations: simulations,
    conclusion: buildPartnerP0004Conclusion_(currentFifo, expiredBreakdown, simulations, yearlyBalance),
    warnings: warnings
  };
}

function sanitizePartnerP0004Grant_(row) {
  const source = row || {};
  const notes = String(source.notes || "");
  return {
    grant_id: String(source.grant_id || ""),
    grant_date: formatDateValue(source.grant_date),
    grant_days: Number(source.grant_days || 0),
    carry_over_days: Number(source.carry_over_days || 0),
    valid_from: formatDateValue(source.valid_from_date || source.valid_from || source.grant_date),
    valid_to: formatDateValue(source.valid_to_date || source.valid_to || ""),
    grant_type: String(source.grant_type || ""),
    year: source.year || "",
    is_finalized: source.is_finalized !== false,
    is_opening_balance: notes.indexOf("初期導入残高") !== -1,
    notes_present: !!notes
  };
}

function sanitizePartnerP0004Fifo_(fifoBalance) {
  const fifo = fifoBalance || {};
  return {
    current_remaining_days: Number(fifo.current_remaining_days || 0),
    expired_days: Number(fifo.expired_days || 0),
    total_granted_days: Number(fifo.total_granted_days || 0),
    used_days: Number(fifo.used_days || 0),
    unallocated_used_days: Number(fifo.unallocated_used_days || 0),
    grant_details: (fifo.grant_details || []).map(lot => ({
      grant_id: lot.grant_id,
      source_grant_id: lot.source_grant_id || lot.grant_id,
      lot_type: lot.lot_type,
      grant_date: lot.grant_date,
      valid_from: lot.valid_from,
      valid_to: lot.valid_to,
      grant_days: Number(lot.grant_days || 0),
      opening_balance_days: Number(lot.opening_balance_days || 0),
      total_days: Number(lot.total_days || 0),
      used_days: Number(lot.used_days || 0),
      remaining_days: Number(lot.remaining_days || 0),
      active_remaining_days: Number(lot.active_remaining_days || 0),
      expired_days: Number(lot.expired_days || 0),
      is_expired: lot.is_expired === true,
      validity_basis: lot.validity_basis || ""
    })),
    allocations: fifo.allocations || []
  };
}

function buildPartnerP0004ExpiredBreakdown_(fifoBalance) {
  return (fifoBalance.grant_details || [])
    .filter(lot => Number(lot.expired_days || 0) > 0)
    .map(lot => ({
      grant_id: lot.grant_id,
      source_grant_id: lot.source_grant_id || lot.grant_id,
      lot_type: lot.lot_type,
      original_days: Number(lot.total_days || 0),
      used_days_before_expiry: Number(lot.used_days || 0),
      expired_days: Number(lot.expired_days || 0),
      valid_to: lot.valid_to
    }));
}

function buildPartnerP0004ActiveBalanceBreakdown_(fifoBalance) {
  return (fifoBalance.grant_details || [])
    .filter(lot => Number(lot.active_remaining_days || 0) > 0)
    .map(lot => ({
      grant_id: lot.grant_id,
      source_grant_id: lot.source_grant_id || lot.grant_id,
      grant_date: lot.grant_date,
      valid_from: lot.valid_from,
      valid_to: lot.valid_to,
      original_days: Number(lot.total_days || 0),
      used_days: Number(lot.used_days || 0),
      remaining_days: Number(lot.active_remaining_days || 0)
    }))
    .sort((a, b) => String(a.valid_to).localeCompare(String(b.valid_to)) ||
      String(a.grant_date).localeCompare(String(b.grant_date)) || String(a.grant_id).localeCompare(String(b.grant_id)))
    .map((lot, index) => Object.assign({}, lot, { consumption_priority: index + 1 }));
}

function buildPartnerP0004FiscalUsage_(employeeId, context, asOfDate, fiscalYear, fiscalStartMonth) {
  const range = getFiscalYearRangeWithStart(fiscalYear, fiscalStartMonth);
  const requests = context.requests_by_employee[employeeId] || [];
  const requestRows = requests.map(row => {
    const status = norm(row.status);
    const type = String(row.type || "paid_leave").trim();
    const hasDates = !!(row.start_date && row.end_date);
    const included = status === STATUS.APPROVED && (!type || type === "paid_leave") && hasDates;
    const dailyRows = included ? expandLeaveRequestToDailyRows(
      row.start_date, row.end_date, row.days, row.half_day, context.calendar_map
    ) : [];
    const fiscalRows = dailyRows.filter(item =>
      isDateInRange(item.date, range.start, range.end) && parseLocalDate(item.date) <= asOfDate
    );
    return {
      request_id: String(row.request_id || ""),
      start_date: formatDateValue(row.start_date),
      end_date: formatDateValue(row.end_date),
      days: Number(row.days || 0),
      half_day: String(row.half_day || ""),
      status: String(row.status || ""),
      type: type,
      fifo_included: included,
      fifo_excluded_reason: included ? "" : getPartnerP0004RequestExclusionReason_(status, type, hasDates),
      fiscal_days: fiscalRows.reduce((sum, item) => sum + Number(item.days || 0), 0)
    };
  });
  const includedFiscalRows = requestRows.filter(row => row.fifo_included && Number(row.fiscal_days || 0) > 0);
  return {
    fiscal_year: fiscalYear,
    fiscal_start: formatDateValue(range.start),
    fiscal_end: formatDateValue(range.end),
    requests: requestRows,
    included_requests: includedFiscalRows,
    total_used_days: sumPartnerP0004Days_(includedFiscalRows, "fiscal_days"),
    half_day_rule: "half_dayがある申請はFIFO日別展開で0.5日、通常申請は営業日ごとに1日として計上します。"
  };
}

function getPartnerP0004RequestExclusionReason_(status, type, hasDates) {
  if (status !== STATUS.APPROVED) return "statusがapprovedではないためFIFO対象外";
  if (type && type !== "paid_leave") return "typeがpaid_leaveではないためFIFO対象外";
  if (!hasDates) return "start_date/end_date不足のためFIFO対象外";
  return "FIFO対象外";
}

function buildPartnerP0004Scenario_(employeeId, grantId, context, asOfDate, fiscalYear, fiscalStartMonth, changes) {
  const scenarioContext = buildPartnerP0004ScenarioContext_(employeeId, grantId, context, changes);
  const fifo = calculateFifoBalanceWithOpeningBalanceFromContext_(employeeId, asOfDate, scenarioContext);
  const yearly = calculateLegacyBalanceFromFifoContext_(employeeId, fiscalYear, fiscalStartMonth, scenarioContext);
  const fiscalUsage = buildPartnerP0004FiscalUsage_(employeeId, scenarioContext, asOfDate, fiscalYear, fiscalStartMonth);
  return {
    assumptions: changes,
    current_remaining_days: Number(fifo.current_remaining_days || 0),
    expired_days: Number(fifo.expired_days || 0),
    next_consumption_lot: summarizePartnerOpeningBalanceScenario_(fifo, asOfDate).next_consumption_lot,
    fiscal_usage_allocations: (fifo.allocations || []).filter(row =>
      String(row.use_date || "") >= "2026-06-01" && String(row.use_date || "") <= "2027-05-31"
    ),
    fiscal_usage: fiscalUsage,
    yearly_balance: yearly,
    difference_from_yearly_balance: Number(fifo.current_remaining_days || 0) -
      Number(yearly.current_remaining_days || 0),
    warnings: []
  };
}

function buildPartnerP0004ScenarioContext_(employeeId, grantId, context, changes) {
  const grantsByEmployee = {};
  Object.keys(context.grants_by_employee || {}).forEach(id => {
    grantsByEmployee[id] = (context.grants_by_employee[id] || []).map(row => {
      const copy = Object.assign({}, row);
      if (id === employeeId && String(copy.grant_id || "") === grantId) {
        copy.grant_days = Number(changes.grant_days);
        copy.carry_over_days = Number(changes.carry_over_days);
        copy.total_days = copy.grant_days + copy.carry_over_days;
        copy.valid_from_date = parseLocalDate(changes.valid_from);
        copy.valid_to_date = parseLocalDate(changes.valid_to);
      }
      return copy;
    });
  });
  return {
    as_of_date: context.as_of_date,
    calendar_map: context.calendar_map,
    requests_by_employee: context.requests_by_employee,
    grants_by_employee: grantsByEmployee
  };
}

function buildPartnerP0004G0058Evaluation_(targetGrant, fifoBalance) {
  const targetLots = (fifoBalance.grant_details || []).filter(lot =>
    String(lot.source_grant_id || lot.grant_id || "") === "G0058"
  );
  return {
    confirmed_facts: targetGrant ? {
      grant_id: String(targetGrant.grant_id || ""),
      grant_days: Number(targetGrant.grant_days || 0),
      carry_over_days: Number(targetGrant.carry_over_days || 0),
      grant_date: formatDateValue(targetGrant.grant_date),
      valid_from: formatDateValue(targetGrant.valid_from_date),
      valid_to: formatDateValue(targetGrant.valid_to_date),
      grant_type: String(targetGrant.grant_type || ""),
      fifo_lot_count: targetLots.length
    } : null,
    interpretations: [
      { hypothesis: "2025/06/01時点の初期残高3日", status: "POSSIBLE", reason: "initial行かつgrant_days=3である事実とは整合しますが、notes本文・元資料なしでは断定できません。" },
      { hypothesis: "2025/06/01の新規付与3日", status: "POSSIBLE", reason: "FIFOはgrant_daysを通常付与ロットとして扱いますが、付与根拠資料が必要です。" },
      { hypothesis: "前年度付与分の残り3日", status: "POSSIBLE", reason: "valid_to=2026/05/31は繰越・初期残高の期限としても解釈できます。" },
      { hypothesis: "初期導入時に総残日数として入力した3日", status: "POSSIBLE", reason: "P0002/P0003とは入力構造が異なるため、同一意図とは断定できません。" }
    ],
    conclusion: "コードと登録値だけではG0058の業務上の意味は確定できません。元資料または当時の入力意図の確認が必要です。"
  };
}

function buildPartnerP0004Conclusion_(currentFifo, expiredBreakdown, simulations, yearlyBalance) {
  const g0058Expired = expiredBreakdown.some(row => row.source_grant_id === "G0058");
  return {
    fifo_against_registered_data: "FIFOは登録済みのgrant_days / carry_over_days / valid_from / valid_toに従って計算しています。",
    expired_days_assessment: g0058Expired
      ? "期限切れ日数にはG0058由来のロットが含まれます。登録期限を正とする限りFIFO結果は整合します。"
      : "期限切れ日数にG0058由来ロットは含まれません。",
    correction_needed: "UNDETERMINED",
    correction_decision_required: "G0058の3日が2025年度末までの初期残高か、2026/06/01時点の確定繰越残高かを元資料で確認してから判断してください。",
    current_vs_yearly_difference: Number(currentFifo.current_remaining_days || 0) -
      Number((yearlyBalance || {}).current_remaining_days || 0),
    reference_scenarios_available: Object.keys(simulations || {})
  };
}

function sumPartnerP0004Days_(rows, key) {
  return (rows || []).reduce((sum, row) => sum + Number(row[key] || 0), 0);
}

function logPartnerP0004FifoDiagnosis_(result) {
  const fifo = result.current_fifo || {};
  const usage = result.fiscal_usage || {};
  const nearest = result.nearest_expiry_lot || {};
  Logger.log([
    "=== PARTNER P0004 FIFO診断 ===",
    "現在有効残高: " + String(fifo.current_remaining_days || 0) + "日",
    "期限切れ日数: " + String(fifo.expired_days || 0) + "日",
    "年度取得日数: " + String(usage.total_used_days || 0) + "日",
    "最短期限: " + String(nearest.valid_to || "-"),
    "期限切れ内訳: " + (result.expired_days_breakdown || []).map(row =>
      row.grant_id + " " + row.original_days + "日 - 使用" + row.used_days_before_expiry + "日 = " + row.expired_days + "日失効"
    ).join(" / ")
  ].join("\n"));
  logJsonInChunks_("[P0004][BASIC]", { employee: result.employee, target: result.target, as_of_date: result.as_of_date });
  logJsonInChunks_("[P0004][GRANTS]", result.grants);
  logJsonInChunks_("[P0004][REQUESTS]", result.requests);
  logJsonInChunks_("[P0004][CURRENT_FIFO]", result.current_fifo);
  logJsonInChunks_("[P0004][EXPIRED_BREAKDOWN]", result.expired_days_breakdown);
  logJsonInChunks_("[P0004][ACTIVE_BALANCE_BREAKDOWN]", result.active_balance_breakdown);
  logJsonInChunks_("[P0004][FISCAL_USAGE]", result.fiscal_usage);
  logJsonInChunks_("[P0004][SIMULATION_A]", result.simulations.plan_a_current_is_correct);
  logJsonInChunks_("[P0004][SIMULATION_B]", result.simulations.plan_b_fiscal_start_carry_over);
  logJsonInChunks_("[P0004][SIMULATION_C]", result.simulations.plan_c_extend_initial_lot_expiry);
  logJsonInChunks_("[P0004][CONCLUSION]", { conclusion: result.conclusion, warnings: result.warnings });
}

/**
 * P0004のG0058/G0061を、年度開始繰越構造へ統一できるか確認する読み取り専用試算。
 */
function debugPartnerCarryOverStructureSimulationP0004() {
  const asOfDate = parseLocalDate("2026-07-25");
  const employeeId = "EMP0062";
  const employee = getEmployeesForAdmin().find(row => String(row.employee_id || "").trim() === employeeId) || {};
  const context = createFifoBalanceComparisonContext_(asOfDate, { read_only: true });
  const result = buildPartnerP0004CarryOverStructureSimulation_(employee, context, asOfDate);
  logPartnerP0004CarryOverStructureSimulation_(result);
  return result;
}

function buildPartnerP0004CarryOverStructureSimulation_(employee, context, asOfDate) {
  const employeeId = "EMP0062";
  const fiscalStartMonth = Number((employee || {}).fiscal_start_month || 6);
  const fiscalYear = getFiscalYearFromDateWithStart(asOfDate, fiscalStartMonth);
  const baseYearlyBalance = calculateLegacyBalanceFromFifoContext_(
    employeeId, fiscalYear, fiscalStartMonth, context
  );
  const current = buildPartnerP0004StructureScenario_(
    "CURRENT", employeeId, context, asOfDate, fiscalYear, fiscalStartMonth, baseYearlyBalance
  );
  const planAContext = buildPartnerP0004ScenarioContext_(employeeId, "G0058", context, {
    grant_days: 3, carry_over_days: 0, valid_from: "2025-06-01", valid_to: "2027-05-31"
  });
  const planA = buildPartnerP0004StructureScenario_(
    "PLAN_A_EXTEND_G0058", employeeId, planAContext, asOfDate, fiscalYear, fiscalStartMonth, baseYearlyBalance
  );
  const planBContext = buildPartnerP0004G0061CarryOverLotScenarioContext_(context);
  const planB = buildPartnerP0004StructureScenario_(
    "PLAN_B_G0061_CARRY_OVER_LOT", employeeId, planBContext, asOfDate, fiscalYear, fiscalStartMonth, baseYearlyBalance
  );
  const g0061 = (context.grants_by_employee[employeeId] || []).find(row => String(row.grant_id || "") === "G0061");
  const warnings = [];
  if (!g0061) warnings.push("G0061が見つかりません。案Bは参考試算のみで、実データの構造確認が必要です。");
  if (g0061 && Number(g0061.carry_over_days || 0) !== 2) warnings.push("G0061のcarry_over_daysが期待値2と一致しません。");
  return {
    ok: warnings.length === 0,
    read_only: true,
    as_of_date: formatDateValue(asOfDate),
    employee: sanitizePaidLeaveDebugEmployee_(employee),
    target_grant_ids: ["G0058", "G0061"],
    current: current,
    plan_a: planA,
    plan_b: planB,
    design_comparison: {
      plan_a: "G0058の期限延長で有効残高を維持する案。G0061のcarry_over_days=2は現行FIFOでは通常の繰越として除外されます。",
      plan_b: "G0058は期限切れ履歴として残し、G0061のcarry_over_days=2を独立した年度開始繰越ロットとして先に消化する案。",
      fifo_order_constraint: "現行FIFOは同一G0061行のgrant_daysロットとcarry_over_daysロットに優先順を持たせられないため、案Bでは読み取りコピー内で繰越2日を独立ロット化して順序を検証しています。",
      preferred_when_source_confirms_carry_over: "元資料が『G0061のcarry_over_days=2は2026/06/01時点の確定繰越』を示す場合、P0002/P0003と同じ設計思想には案Bがより近いです。"
    },
    warnings: warnings
  };
}

function buildPartnerP0004StructureScenario_(name, employeeId, context, asOfDate, fiscalYear, fiscalStartMonth, referenceYearlyBalance) {
  const fifo = calculateFifoBalanceWithOpeningBalanceFromContext_(employeeId, asOfDate, context);
  const activeLots = buildPartnerP0004ActiveBalanceBreakdown_(fifo);
  const allLots = sanitizePartnerP0004Fifo_(fifo).grant_details;
  const fiscalUsage = buildPartnerP0004FiscalUsage_(employeeId, context, asOfDate, fiscalYear, fiscalStartMonth);
  return {
    name: name,
    current_remaining_days: Number(fifo.current_remaining_days || 0),
    expired_days: Number(fifo.expired_days || 0),
    difference_from_yearly_balance: Number(fifo.current_remaining_days || 0) -
      Number((referenceYearlyBalance || {}).current_remaining_days || 0),
    next_consumption_lot: activeLots.length > 0 ? activeLots[0] : null,
    nearest_active_expiry_lot: activeLots.length > 0 ? activeLots[0] : null,
    earliest_lot_expiry: allLots.slice().sort((a, b) =>
      String(a.valid_to).localeCompare(String(b.valid_to)) || String(a.grant_id).localeCompare(String(b.grant_id))
    )[0] || null,
    fiscal_usage_allocations: (fifo.allocations || []).filter(row =>
      String(row.use_date || "") >= "2026-06-01" && String(row.use_date || "") <= "2027-05-31"
    ),
    fiscal_usage: fiscalUsage,
    fifo_lots: allLots,
    yearly_balance: referenceYearlyBalance
  };
}

function buildPartnerP0004G0061CarryOverLotScenarioContext_(context) {
  const employeeId = "EMP0062";
  const grantsByEmployee = {};
  Object.keys(context.grants_by_employee || {}).forEach(id => {
    grantsByEmployee[id] = (context.grants_by_employee[id] || []).reduce((rows, row) => {
      if (id !== employeeId || String(row.grant_id || "") !== "G0061") {
        rows.push(Object.assign({}, row));
        return rows;
      }
      // 読み取りコピーでのみ、2日繰越を11日新規付与から独立させる。
      const yearlyGrant = Object.assign({}, row, { carry_over_days: 0, total_days: Number(row.grant_days || 0) });
      const carryOverLot = Object.assign({}, row, {
        grant_id: "G0061#carry_over_simulation",
        grant_date: parseLocalDate("2026-05-31"),
        valid_from_date: parseLocalDate("2026-06-01"),
        valid_to_date: parseLocalDate("2027-05-31"),
        grant_days: 0,
        carry_over_days: Number(row.carry_over_days || 0),
        total_days: Number(row.carry_over_days || 0),
        notes: "初期導入残高（読み取り専用試算）"
      });
      rows.push(carryOverLot, yearlyGrant);
      return rows;
    }, []);
  });
  return {
    as_of_date: context.as_of_date,
    calendar_map: context.calendar_map,
    requests_by_employee: context.requests_by_employee,
    grants_by_employee: grantsByEmployee
  };
}

/**
 * G0058をP0002/P0003と同じ年度開始繰越残高へ変換する案Dの最終・読み取り専用試算。
 */
function debugPartnerFiscalStartCarryOverConversionP0004() {
  const asOfDate = parseLocalDate("2026-07-25");
  const employee = getEmployeesForAdmin().find(row => String(row.employee_id || "").trim() === "EMP0062") || {};
  const context = createFifoBalanceComparisonContext_(asOfDate, { read_only: true });
  const result = buildPartnerP0004FiscalStartCarryOverConversion_(employee, context, asOfDate);
  logPartnerP0004FiscalStartCarryOverConversion_(result);
  return result;
}

function buildPartnerP0004FiscalStartCarryOverConversion_(employee, context, asOfDate) {
  const employeeId = "EMP0062";
  const fiscalStartMonth = Number((employee || {}).fiscal_start_month || 6);
  const fiscalYear = getFiscalYearFromDateWithStart(asOfDate, fiscalStartMonth);
  const yearlyBalance = calculateLegacyBalanceFromFifoContext_(employeeId, fiscalYear, fiscalStartMonth, context);
  const structure = buildPartnerP0004CarryOverStructureSimulation_(employee, context, asOfDate);
  const planDChanges = { grant_days: 0, carry_over_days: 2, valid_from: "2026-06-01", valid_to: "2027-05-31" };
  const planDContext = buildPartnerP0004ScenarioContext_(employeeId, "G0058", context, planDChanges);
  const planD = buildPartnerP0004StructureScenario_(
    "PLAN_D_G0058_FISCAL_START_CARRY_OVER", employeeId, planDContext,
    asOfDate, fiscalYear, fiscalStartMonth, yearlyBalance
  );
  const planDTarget = (planDContext.grants_by_employee[employeeId] || []).find(row => String(row.grant_id || "") === "G0058") || {};
  const originalG0061 = (context.grants_by_employee[employeeId] || []).find(row => String(row.grant_id || "") === "G0061") || {};
  const simulatedG0061 = (planDContext.grants_by_employee[employeeId] || []).find(row => String(row.grant_id || "") === "G0061") || {};
  const may2Allocations = (planD.fifo_lots && calculateFifoBalanceWithOpeningBalanceFromContext_(employeeId, asOfDate, planDContext).allocations || [])
    .filter(row => String(row.use_date || "") === "2026-05-02");
  const fiscalAllocations = planD.fiscal_usage_allocations;
  const g0058CarryAllocations = fiscalAllocations.filter(row =>
    String(row.grant_id || "") === "G0058#opening_balance"
  );
  const g0058CarryUsed = g0058CarryAllocations.reduce((sum, row) => sum + Number(row.consumed_days || 0), 0);
  const comparison = [
    buildPartnerP0004ConversionComparisonRow_("現状", structure.current),
    buildPartnerP0004ConversionComparisonRow_("案A：G0058期限延長", structure.plan_a),
    buildPartnerP0004ConversionComparisonRow_("案B：G0061繰越独立ロット", structure.plan_b),
    buildPartnerP0004ConversionComparisonRow_("案D：G0058年度開始繰越2日", planD)
  ];
  const checks = {
    target_is_g0058_only: String(planDTarget.grant_id || "") === "G0058",
    grant_days_is_zero: Number(planDTarget.grant_days || 0) === 0,
    carry_over_days_is_two: Number(planDTarget.carry_over_days || 0) === 2,
    valid_from_is_fiscal_start: formatDateValue(planDTarget.valid_from_date) === "2026-06-01",
    valid_to_is_2027_05_31: formatDateValue(planDTarget.valid_to_date) === "2027-05-31",
    may_2_not_allocated_to_new_carry_over: may2Allocations.length === 0,
    fiscal_carry_over_used_first: g0058CarryUsed === 2 &&
      fiscalAllocations.length > 0 && String(fiscalAllocations[0].grant_id || "") === "G0058#opening_balance",
    current_remaining_is_9_5: Number(planD.current_remaining_days || 0) === 9.5,
    expired_days_is_zero: Number(planD.expired_days || 0) === 0,
    difference_from_yearly_is_zero: Number(planD.difference_from_yearly_balance || 0) === 0,
    total_fifo_days_is_13: planD.fifo_lots.reduce((sum, lot) => sum + Number(lot.total_days || 0), 0) === 13,
    g0061_unchanged: Number(originalG0061.grant_days || 0) === Number(simulatedG0061.grant_days || 0) &&
      Number(originalG0061.carry_over_days || 0) === Number(simulatedG0061.carry_over_days || 0) &&
      formatDateValue(originalG0061.valid_from_date) === formatDateValue(simulatedG0061.valid_from_date) &&
      formatDateValue(originalG0061.valid_to_date) === formatDateValue(simulatedG0061.valid_to_date),
    g0061_carry_over_not_double_counted: planD.fifo_lots.filter(lot =>
      String(lot.source_grant_id || "") === "G0061"
    ).reduce((sum, lot) => sum + Number(lot.total_days || 0), 0) === 11,
    nearest_active_expiry_is_2027_05_31: String((planD.nearest_active_expiry_lot || {}).valid_to || "") === "2027-05-31",
    earliest_lot_expiry_is_2027_05_31: String((planD.earliest_lot_expiry || {}).valid_to || "") === "2027-05-31"
  };
  return {
    ok: checks.target_is_g0058_only && checks.grant_days_is_zero && checks.carry_over_days_is_two &&
      checks.valid_from_is_fiscal_start && checks.valid_to_is_2027_05_31 &&
      checks.may_2_not_allocated_to_new_carry_over && checks.fiscal_carry_over_used_first &&
      checks.current_remaining_is_9_5 && checks.expired_days_is_zero &&
      checks.difference_from_yearly_is_zero && checks.total_fifo_days_is_13 &&
      checks.g0061_unchanged && checks.g0061_carry_over_not_double_counted,
    read_only: true,
    as_of_date: formatDateValue(asOfDate),
    target: { employee_id: employeeId, display_employee_id: "P0004", grant_id: "G0058" },
    plan_d_changes: planDChanges,
    current: structure.current,
    plan_a: structure.plan_a,
    plan_b: structure.plan_b,
    plan_d: planD,
    plan_d_allocations: {
      allocations_on_2026_05_02: may2Allocations,
      fiscal_allocations: fiscalAllocations,
      g0058_carry_over_consumed_days: g0058CarryUsed
    },
    comparison: comparison,
    checks: checks,
    recommendation: {
      plan_d_recommended: checks.current_remaining_is_9_5 && checks.expired_days_is_zero &&
        checks.difference_from_yearly_is_zero && checks.nearest_active_expiry_is_2027_05_31,
      summary: checks.nearest_active_expiry_is_2027_05_31
        ? "案Dは指定された残高・期限切れ・年度残高・最短期限の条件を満たします。"
        : "案Dは残高9.5日・期限切れ0日・年度残高差0日を満たしますが、繰越2日を使い切るため、残高のある最短期限はG0061の2028/05/31です。2027/05/31はロット全体の最短期限であり、残高のある最短期限ではありません。"
    }
  };
}

function buildPartnerP0004ConversionComparisonRow_(label, scenario) {
  const row = scenario || {};
  return {
    label: label,
    current_remaining_days: Number(row.current_remaining_days || 0),
    expired_days: Number(row.expired_days || 0),
    nearest_active_expiry: (row.nearest_active_expiry_lot || {}).valid_to || "",
    fifo_lot_count: (row.fifo_lots || []).length,
    next_consumption_lot: (row.next_consumption_lot || {}).grant_id || "",
    difference_from_yearly_balance: Number(row.difference_from_yearly_balance || 0)
  };
}

function logPartnerP0004FiscalStartCarryOverConversion_(result) {
  Logger.log([
    "=== PARTNER P0004 年度開始繰越変換 試算 ===",
    "案D 現在残: " + String(result.plan_d.current_remaining_days) + "日",
    "案D 期限切れ: " + String(result.plan_d.expired_days) + "日",
    "案D 年度残高との差: " + String(result.plan_d.difference_from_yearly_balance) + "日",
    "案D 残高のある最短期限: " + String((result.plan_d.nearest_active_expiry_lot || {}).valid_to || "-"),
    "評価: " + result.recommendation.summary
  ].join("\n"));
  logJsonInChunks_("[P0004][CONVERSION][COMPARISON]", result.comparison);
  logJsonInChunks_("[P0004][CONVERSION][CURRENT]", result.current);
  logJsonInChunks_("[P0004][CONVERSION][PLAN_A]", result.plan_a);
  logJsonInChunks_("[P0004][CONVERSION][PLAN_B]", result.plan_b);
  logJsonInChunks_("[P0004][CONVERSION][PLAN_D]", {
    changes: result.plan_d_changes,
    scenario: result.plan_d,
    allocations: result.plan_d_allocations,
    checks: result.checks,
    recommendation: result.recommendation
  });
}

function logPartnerP0004CarryOverStructureSimulation_(result) {
  const printSummary = (label, scenario) => [
    "=== P0004 " + label + " ===",
    "現在残: " + String(scenario.current_remaining_days) + "日",
    "期限切れ: " + String(scenario.expired_days) + "日",
    "年度残高との差: " + String(scenario.difference_from_yearly_balance) + "日",
    "次消化ロット: " + String((scenario.next_consumption_lot || {}).grant_id || "-")
  ].join("\n");
  Logger.log(printSummary("現状", result.current));
  Logger.log(printSummary("案A", result.plan_a));
  Logger.log(printSummary("案B", result.plan_b));
  [
    ["CURRENT", result.current],
    ["PLAN_A", result.plan_a],
    ["PLAN_B", result.plan_b]
  ].forEach(item => {
    const label = "[P0004][" + item[0] + "]";
    logJsonInChunks_(label + "[RESULT]", item[1]);
  });
  logJsonInChunks_("[P0004][DESIGN_COMPARISON]", {
    design_comparison: result.design_comparison,
    warnings: result.warnings
  });
}

function buildPartnerOpeningBalanceCarryOverSimulation_(asOfDateValue) {
  const state = readPartnerOpeningBalanceRepairState_();
  const asOfDate = asOfDateValue ? parseLocalDate(asOfDateValue) : parseLocalDate(new Date());
  const context = createFifoBalanceComparisonContext_(asOfDate, { read_only: true });
  return {
    ok: true,
    dry_run: true,
    write_disabled: true,
    warning: "grant_days=15を失効日数、carry_over_days=20を有効繰越残高と解釈する参考試算です。期限は書き換えません。",
    as_of_date: formatDateValue(asOfDate),
    targets: state.rows.map(row => {
      const employeeId = String(row.employee_id || "").trim();
      const employee = state.employee_map[employeeId] || {};
      const before = calculateFifoBalanceWithOpeningBalanceFromContext_(employeeId, asOfDate, context);
      const carryOnly = simulatePartnerOpeningBalanceScenario_(
        employeeId, String(row.grant_id || ""), "CARRY_OVER_DAYS_ONLY", asOfDate, context
      );
      const fiscalYear = getFiscalYearFromDate(asOfDate);
      const fiscalStartMonth = Number(employee.fiscal_start_month || 6);
      const yearlyBefore = calculateLegacyBalanceFromFifoContext_(employeeId, fiscalYear, fiscalStartMonth, context);
      const carryOnlyContext = buildPartnerOpeningBalanceCarryOnlyScenarioContext_(
        employeeId, String(row.grant_id || ""), context
      );
      const yearlyAfter = calculateLegacyBalanceFromFifoContext_(employeeId, fiscalYear, fiscalStartMonth, carryOnlyContext);
      const currentValidTo = row.valid_to ? parseLocalDate(row.valid_to) : null;
      const referenceValidTo = currentValidTo
        ? addDaysLocal_(addYearsLocal_(parseLocalDate(row.valid_from || row.grant_date), 2), -1)
        : null;
      const referenceContext = referenceValidTo
        ? buildPartnerOpeningBalanceCarryOnlyScenarioContext_(employeeId, String(row.grant_id || ""), context, referenceValidTo)
        : null;
      const referenceFifo = referenceContext
        ? calculateFifoBalanceWithOpeningBalanceFromContext_(employeeId, asOfDate, referenceContext)
        : null;
      return {
        grant_id: String(row.grant_id || ""),
        employee_id: employeeId,
        display_employee_id: String(employee.display_employee_id || ""),
        current_grant_days: Number(row.grant_days || 0),
        current_carry_over_days: Number(row.carry_over_days || 0),
        current_valid_from: formatDateValue(row.valid_from || row.grant_date),
        current_valid_to: formatDateValue(row.valid_to || ""),
        current_fifo: summarizePartnerOpeningBalanceScenario_(before, asOfDate),
        grant_days_zero_carry_over_maintained: carryOnly,
        current_valid_to_maintained: carryOnly,
        reference_valid_to_simulation: referenceFifo ? Object.assign(
          { valid_from: formatDateValue(row.valid_from || row.grant_date), valid_to: formatDateValue(referenceValidTo) },
          summarizePartnerOpeningBalanceScenario_(referenceFifo, asOfDate)
        ) : { status: "NOT_CALCULABLE", reason: "valid_from / grant_date が不足しています。" },
        yearly_balance_before: yearlyBefore,
        yearly_balance_after: yearlyAfter,
        difference: {
          current_remaining_days: carryOnly.current_active_remaining_days - Number(before.current_remaining_days || 0),
          expired_days: carryOnly.expired_days - Number(before.expired_days || 0),
          yearly_current_remaining_days: Number(yearlyAfter.current_remaining_days || 0) - Number(yearlyBefore.current_remaining_days || 0)
        },
        comparison_to_reference_balances: [39, 40].map(days => ({
          reference_remaining_days: days,
          difference_from_carry_only: carryOnly.current_active_remaining_days - days
        })),
        required_data_corrections: ["grant_daysを0とする案は参考試算のみ。期限確定までは書込み不可。"],
        expiry_confirmation_required: [
          "carry_over_days 20の元になった付与ロットの実際の付与日と有効期限",
          "2025/06/01が元付与日か、単なる初期導入基準日か",
          "繰越20日が2026/06/01時点の残高であることを示す根拠資料"
        ],
        warnings: ["旧carry_over_days→0補正は無効です。P0004 / EMP0062 / G0058 は対象外です。"]
      };
    })
  };
}

function logPartnerCarryOverSimulationTarget_(target) {
  const summary = buildPartnerCarryOverSimulationSummary_(target);
  Logger.log(summary.join("\n"));
  const label = "[" + String(target.display_employee_id || target.employee_id || "UNKNOWN") + "]";
  logJsonInChunks_(label + "[BASIC]", {
    grant_id: target.grant_id,
    employee_id: target.employee_id,
    display_employee_id: target.display_employee_id,
    current_grant_days: target.current_grant_days,
    current_carry_over_days: target.current_carry_over_days,
    current_valid_from: target.current_valid_from,
    current_valid_to: target.current_valid_to
  });
  logJsonInChunks_(label + "[CURRENT_FIFO]", target.current_fifo);
  logJsonInChunks_(label + "[CARRY_ONLY_CURRENT_VALID_TO]", target.current_valid_to_maintained);
  logJsonInChunks_(label + "[REFERENCE_VALID_TO]", target.reference_valid_to_simulation);
  logJsonInChunks_(label + "[YEARLY_BALANCE]", {
    before: target.yearly_balance_before,
    after: target.yearly_balance_after,
    comparison_to_reference_balances: target.comparison_to_reference_balances
  });
  logJsonInChunks_(label + "[WARNINGS]", {
    required_data_corrections: target.required_data_corrections,
    expiry_confirmation_required: target.expiry_confirmation_required,
    warnings: target.warnings
  });
}

function buildPartnerCarryOverSimulationSummary_(target) {
  const current = target.current_fifo || {};
  const carryOnly = target.current_valid_to_maintained || {};
  const reference = target.reference_valid_to_simulation || {};
  const yearly = target.yearly_balance_after || {};
  const displayId = String(target.display_employee_id || target.employee_id || "UNKNOWN");
  return [
    "=== PARTNER繰越試算 " + displayId + " / " + String(target.employee_id || "") + " ===",
    "現状:",
    "  現在有効残高: " + String(current.current_active_remaining_days == null ? "-" : current.current_active_remaining_days),
    "  期限切れ日数: " + String(current.expired_days == null ? "-" : current.expired_days),
    "grant_days=0、期限維持:",
    "  現在有効残高: " + String(carryOnly.current_active_remaining_days == null ? "-" : carryOnly.current_active_remaining_days),
    "  期限切れ日数: " + String(carryOnly.expired_days == null ? "-" : carryOnly.expired_days),
    "grant_days=0、参考期限 " + String(reference.valid_to || "未算出") + ":",
    "  現在有効残高: " + String(reference.current_active_remaining_days == null ? "-" : reference.current_active_remaining_days),
    "  期限切れ日数: " + String(reference.expired_days == null ? "-" : reference.expired_days),
    "年度残高:",
    "  現在残高: " + String(yearly.current_remaining_days == null ? "-" : yearly.current_remaining_days),
    "差:",
    "  参考期限案と年度残高の差: " + String(
      reference.current_active_remaining_days == null || yearly.current_remaining_days == null
        ? "-"
        : Number(reference.current_active_remaining_days) - Number(yearly.current_remaining_days)
    )
  ];
}

function createJsonLogChunks_(label, value, maxChars) {
  const limit = Math.max(100, Number(maxChars || 5000));
  const text = JSON.stringify(value == null ? null : value, null, 2);
  const characters = Array.from(text);
  const total = Math.max(1, Math.ceil(characters.length / limit));
  const chunks = [];
  for (let index = 0; index < total; index++) {
    chunks.push(label + " [" + (index + 1) + "/" + total + "]\n" +
      characters.slice(index * limit, (index + 1) * limit).join(""));
  }
  return chunks;
}

function logJsonInChunks_(label, value, maxChars) {
  createJsonLogChunks_(label, value, maxChars).forEach(message => Logger.log(message));
}

/**
 * PARTNER初期導入残高を、2026年度開始の繰越ロットとして扱う読み取り専用試算。
 */
function debugPartnerOpeningBalanceFiscalStartSimulation(asOfDateValue) {
  const asOfDate = asOfDateValue ? parseLocalDate(asOfDateValue) : parseLocalDate(new Date());
  const base = buildPartnerOpeningBalanceCarryOverSimulation_(asOfDate);
  const state = readPartnerOpeningBalanceRepairState_();
  const context = createFifoBalanceComparisonContext_(asOfDate, { read_only: true });
  const fiscalStart = parseLocalDate("2026-06-01");
  const fiscalEnd = parseLocalDate("2027-05-31");
  const targets = (base.targets || []).map(baseTarget => {
    const employeeId = String(baseTarget.employee_id || "");
    const scenarioContext = buildPartnerOpeningBalanceCarryOnlyScenarioContext_(
      employeeId,
      String(baseTarget.grant_id || ""),
      context,
      fiscalEnd,
      fiscalStart
    );
    const fifo = calculateFifoBalanceWithOpeningBalanceFromContext_(employeeId, asOfDate, scenarioContext);
    const fifoSummary = summarizePartnerOpeningBalanceScenario_(fifo, asOfDate);
    const employee = state.employee_map[employeeId] || {};
    const fiscalYear = getFiscalYearFromDate(asOfDate);
    const yearlyBalance = calculateLegacyBalanceFromFifoContext_(
      employeeId, fiscalYear, Number(employee.fiscal_start_month || 6), scenarioContext
    );
    return Object.assign({}, baseTarget, {
      valid_to_only_simulation: baseTarget.reference_valid_to_simulation,
      valid_from_and_valid_to_simulation: Object.assign({
        valid_from: "2026/06/01",
        valid_to: "2027/05/31",
        grant_days: 0,
        carry_over_days: 20
      }, fifoSummary),
      fiscal_start_yearly_balance: yearlyBalance,
      fiscal_start_difference_from_yearly: {
        current_remaining_days: fifoSummary.current_active_remaining_days -
          Number(yearlyBalance.current_remaining_days || 0),
        expired_days: fifoSummary.expired_days - Number(yearlyBalance.expired_days || 0)
      },
      allocations_on_2026_05_02: getPartnerFiscalStartSimulationAllocations_(fifo, "2026/05/02"),
      allocations_on_2026_07_18: getPartnerFiscalStartSimulationAllocations_(fifo, "2026/07/18"),
      next_consumption_lot: fifoSummary.next_consumption_lot,
      warnings: (baseTarget.warnings || []).concat([
        "これは valid_from=2026/06/01、valid_to=2027/05/31 の読み取り専用試算です。"
      ])
    });
  });
  const result = Object.assign({}, base, {
    scenario: "FISCAL_START_CARRY_OVER_LOT",
    targets: targets
  });
  targets.forEach(logPartnerFiscalStartSimulationTarget_);
  return result;
}

function getPartnerFiscalStartSimulationAllocations_(fifoBalance, dateKey) {
  const normalizedDateKey = formatDateValue(parseLocalDate(dateKey));
  const allocations = (fifoBalance.allocations || [])
    .filter(row => formatDateValue(row.use_date) === normalizedDateKey)
    .map(row => ({
      request_id: String(row.request_id || ""),
      grant_id: String(row.grant_id || ""),
      lot_type: String(row.lot_type || ""),
      consumed_days: Number(row.consumed_days || 0)
    }));
  const unallocatedDays = (fifoBalance.used_details || [])
    .filter(row => formatDateValue(row.use_date) === normalizedDateKey)
    .reduce((sum, row) => sum + Number(row.unallocated_days || 0), 0);
  return {
    use_date: normalizedDateKey,
    allocations: allocations,
    unallocated_days: unallocatedDays
  };
}

function logPartnerFiscalStartSimulationTarget_(target) {
  const fiscalStart = target.valid_from_and_valid_to_simulation || {};
  const yearly = target.fiscal_start_yearly_balance || {};
  const label = "[" + String(target.display_employee_id || target.employee_id || "UNKNOWN") + "][FISCAL_START]";
  Logger.log([
    "=== PARTNER年度開始繰越試算 " + String(target.display_employee_id || "") + " / " + String(target.employee_id || "") + " ===",
    "現状FIFO 有効残高: " + String((target.current_fifo || {}).current_active_remaining_days),
    "期限のみ変更案 有効残高: " + String((target.valid_to_only_simulation || {}).current_active_remaining_days),
    "valid_from/valid_to変更案 有効残高: " + String(fiscalStart.current_active_remaining_days),
    "年度残高: " + String(yearly.current_remaining_days),
    "2026/05/02割当件数: " + String(((target.allocations_on_2026_05_02 || {}).allocations || []).length),
    "2026/07/18割当件数: " + String(((target.allocations_on_2026_07_18 || {}).allocations || []).length)
  ].join("\n"));
  logJsonInChunks_(label + "[CURRENT_FIFO]", target.current_fifo);
  logJsonInChunks_(label + "[VALID_TO_ONLY]", target.valid_to_only_simulation);
  logJsonInChunks_(label + "[VALID_FROM_AND_TO]", target.valid_from_and_valid_to_simulation);
  logJsonInChunks_(label + "[ALLOCATIONS_2026_05_02]", target.allocations_on_2026_05_02);
  logJsonInChunks_(label + "[ALLOCATIONS_2026_07_18]", target.allocations_on_2026_07_18);
  logJsonInChunks_(label + "[YEARLY_BALANCE]", {
    yearly_balance: yearly,
    difference: target.fiscal_start_difference_from_yearly,
    next_consumption_lot: target.next_consumption_lot,
    warnings: target.warnings
  });
}

function buildPartnerOpeningBalanceCarryOnlyScenarioContext_(employeeId, grantId, context, validToOverride, validFromOverride) {
  const grantsByEmployee = {};
  Object.keys(context.grants_by_employee || {}).forEach(id => {
    grantsByEmployee[id] = (context.grants_by_employee[id] || []).map(source => {
      const copy = Object.assign({}, source);
      if (id === employeeId && String(copy.grant_id || "") === grantId) {
        copy.grant_days = 0;
        copy.total_days = Number(copy.carry_over_days || 0);
        if (validToOverride) copy.valid_to_date = validToOverride;
        if (validFromOverride) copy.valid_from_date = validFromOverride;
      }
      return copy;
    });
  });
  return {
    as_of_date: context.as_of_date,
    calendar_map: context.calendar_map,
    requests_by_employee: context.requests_by_employee,
    grants_by_employee: grantsByEmployee
  };
}

function debugYearEndFinalizedBalance(employeeId, fiscalYear) {
  const targetEmployeeId = String(employeeId || "").trim();
  const targetFiscalYear = Number(fiscalYear || 0);

  if (!targetEmployeeId) throw new Error("employeeId がありません");
  if (!targetFiscalYear) throw new Error("対象年度がありません");

  const employeeMap = getEmployeeDetailMap();
  const fiscalStartMonth = getFiscalStartMonthByEmployeeId(targetEmployeeId, employeeMap);
  const fiscalRange = getFiscalYearRangeWithStart(targetFiscalYear, fiscalStartMonth);
  const grantRecords = getPaidLeaveGrantDebugRecordsForFiscalYear_(
    targetEmployeeId,
    targetFiscalYear,
    fiscalStartMonth
  );
  const grantDaysTotal = grantRecords.reduce((sum, row) => sum + Number(row.grant_days || 0), 0);
  const carryOverDaysTotal = grantRecords.reduce((sum, row) => sum + Number(row.carry_over_days || 0), 0);
  const usedMap = getApprovedUsedDaysByFiscalYearForEmployeeIds(targetFiscalYear, [targetEmployeeId]);
  const approvedUsedDays = Number(usedMap[targetEmployeeId] || 0);
  const legacyBalance = buildBalance(
    targetEmployeeId,
    {
      employee_id: targetEmployeeId,
      grant_days: grantDaysTotal,
      carry_over_days: carryOverDaysTotal
    },
    approvedUsedDays
  );
  const fifoContext = createFifoBalanceComparisonContext_(fiscalRange.end);
  const fifoBalance = calculateFifoBalanceFromContext_(
    targetEmployeeId,
    fiscalRange.end,
    fifoContext
  );
  const suspectedYearEndRecords = grantRecords.filter(row =>
    String(row.grant_type || "") === "yearly" &&
    (
      String(row.notes || "").indexOf("年跨ぎ確定") !== -1 ||
      Number(row.carry_over_days || 0) > 0
    )
  );
  const result = {
    employee_id: targetEmployeeId,
    fiscal_year: targetFiscalYear,
    fiscal_start_month: fiscalStartMonth,
    fiscal_year_start: formatDateValue(fiscalRange.start),
    fiscal_year_end: formatDateValue(fiscalRange.end),
    paid_leave_grants: grantRecords,
    grant_days_total: grantDaysTotal,
    carry_over_days_total: carryOverDaysTotal,
    approved_used_days: approvedUsedDays,
    build_balance: legacyBalance,
    fifo_balance: fifoBalance,
    difference: {
      current_remaining_days:
        Number(fifoBalance.current_remaining_days || 0) -
        Number(legacyBalance.current_remaining_days || 0),
      used_days:
        Number(fifoBalance.used_days || 0) -
        Number(legacyBalance.used_days || 0),
      expired_days:
        Number(fifoBalance.expired_days || 0) -
        Number(legacyBalance.expired_days || 0)
    },
    suspected_year_end_finalized_records: suspectedYearEndRecords.map(row => ({
      grant_id: row.grant_id,
      notes: row.notes,
      grant_date: row.grant_date,
      valid_from: row.valid_from,
      valid_to: row.valid_to,
      grant_days: row.grant_days,
      carry_over_days: row.carry_over_days
    })),
    carry_over_days_handling: {
      calculation: "buildBalance は対象年度の carry_over_days 合計を前年度繰越として扱い、grant_days 合計と足して used_days を差し引きます。",
      formula: "current_remaining_days = carry_over_days_total + grant_days_total - approved_used_days",
      carry_over_days_total: carryOverDaysTotal,
      grant_days_total: grantDaysTotal,
      approved_used_days: approvedUsedDays
    },
    valid_period_handling: {
      fifo_as_of_date: formatDateValue(fiscalRange.end),
      note: "FIFO試算では grant_days + carry_over_days を同一付与レコードの total_days として扱い、同じ valid_from / valid_to を適用します。",
      warning: suspectedYearEndRecords.length > 0
        ? "年跨ぎ確定レコードの繰越分と新規付与分が同じ有効期限になっている可能性があります。繰越分の元付与期限を厳密に維持する運用とは差が出る可能性があります。"
        : ""
    },
    annual_report_alignment: "年間一覧・CSVは getGrantMapByFiscalYear と buildBalance を使うため、build_balance と同じ grant_days/carry_over_days/used_days/next_carry_over_days/expired_days になります。"
  };

  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

function getPaidLeaveGrantDebugRecordsForFiscalYear_(employeeId, fiscalYear, fiscalStartMonth) {
  const sheet = getSheet("paid_leave_grants");
  const headerInfo = requireHeaders(sheet, [
    "grant_id",
    "employee_id",
    "grant_date",
    "grant_days",
    "carry_over_days",
    "valid_from",
    "valid_to",
    "grant_type",
    "year",
    "notes"
  ]);
  const data = sheet.getDataRange().getValues();
  const targetEmployeeId = String(employeeId || "").trim();

  if (data.length <= 1) return [];

  return data.slice(1)
    .map(row => rowToObject(row, headerInfo.headers))
    .filter(rowObj => String(rowObj.employee_id || "").trim() === targetEmployeeId)
    .filter(rowObj => rowObj.grant_date)
    .map(rowObj => {
      const grantDate = parseLocalDate(rowObj.grant_date);
      const recordFiscalYear = getFiscalYearFromDateWithStart(grantDate, fiscalStartMonth);

      return {
        grant_id: String(rowObj.grant_id || ""),
        employee_id: targetEmployeeId,
        grant_date: formatDateValue(grantDate),
        grant_days: Number(rowObj.grant_days || 0),
        carry_over_days: Number(rowObj.carry_over_days || 0),
        total_days: Number(rowObj.grant_days || 0) + Number(rowObj.carry_over_days || 0),
        valid_from: formatDateValue(rowObj.valid_from || rowObj.grant_date),
        valid_to: formatDateValue(
          rowObj.valid_to || addDaysLocal_(addYearsLocal_(grantDate, 2), -1)
        ),
        grant_type: String(rowObj.grant_type || ""),
        year: rowObj.year || "",
        fiscal_year_by_grant_date: recordFiscalYear,
        notes: String(rowObj.notes || "")
      };
    })
    .filter(row => Number(row.fiscal_year_by_grant_date) === Number(fiscalYear))
    .sort((a, b) => {
      if (a.grant_date !== b.grant_date) return a.grant_date < b.grant_date ? -1 : 1;
      return String(a.grant_id).localeCompare(String(b.grant_id));
    });
}

function debugFifoBalanceWithoutCarryOver(employeeId, asOfDateValue) {
  const targetEmployeeId = String(employeeId || "").trim();
  if (!targetEmployeeId) throw new Error("employeeId がありません");

  const asOfDate = asOfDateValue ? parseLocalDate(asOfDateValue) : parseLocalDate(new Date());
  const context = createFifoBalanceComparisonContext_(asOfDate);
  const result = calculateFifoBalanceWithoutCarryOverFromContext_(
    targetEmployeeId,
    asOfDate,
    context
  );

  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

function compareYearEndFinalizedBalanceModes(employeeId, fiscalYear) {
  const targetEmployeeId = String(employeeId || "").trim();
  const targetFiscalYear = Number(fiscalYear || 0);

  if (!targetEmployeeId) throw new Error("employeeId がありません");
  if (!targetFiscalYear) throw new Error("対象年度がありません");

  const employeeMap = getEmployeeDetailMap();
  const fiscalStartMonth = getFiscalStartMonthByEmployeeId(targetEmployeeId, employeeMap);
  const fiscalRange = getFiscalYearRangeWithStart(targetFiscalYear, fiscalStartMonth);
  const fiscalYearGrantRecords = getPaidLeaveGrantDebugRecordsForFiscalYear_(
    targetEmployeeId,
    targetFiscalYear,
    fiscalStartMonth
  );
  const grantDaysTotal = fiscalYearGrantRecords.reduce(
    (sum, row) => sum + Number(row.grant_days || 0),
    0
  );
  const carryOverDaysTotal = fiscalYearGrantRecords.reduce(
    (sum, row) => sum + Number(row.carry_over_days || 0),
    0
  );
  const usedMap = getApprovedUsedDaysByFiscalYearForEmployeeIds(
    targetFiscalYear,
    [targetEmployeeId]
  );
  const approvedUsedDays = Number(usedMap[targetEmployeeId] || 0);
  const legacyBalance = buildBalance(
    targetEmployeeId,
    {
      employee_id: targetEmployeeId,
      grant_days: grantDaysTotal,
      carry_over_days: carryOverDaysTotal
    },
    approvedUsedDays
  );
  const context = createFifoBalanceComparisonContext_(fiscalRange.end);
  const currentFifoBalance = calculateFifoBalanceFromContext_(
    targetEmployeeId,
    fiscalRange.end,
    context
  );
  const fifoWithoutCarryOverBalance = calculateFifoBalanceWithoutCarryOverFromContext_(
    targetEmployeeId,
    fiscalRange.end,
    context
  );
  const yearEndFinalizedRecords = fiscalYearGrantRecords.filter(row =>
    String(row.grant_type || "") === "yearly" &&
    (
      String(row.notes || "").indexOf("年跨ぎ確定") !== -1 ||
      Number(row.carry_over_days || 0) > 0
    )
  );
  const yearEndGrantIds = {};
  yearEndFinalizedRecords.forEach(row => {
    yearEndGrantIds[String(row.grant_id || "")] = true;
  });
  const originalGrantRecords = (currentFifoBalance.grant_details || [])
    .filter(row => !yearEndGrantIds[String(row.grant_id || "")])
    .filter(row => parseLocalDate(row.grant_date) < fiscalRange.start)
    .map(row => ({
      grant_id: row.grant_id,
      grant_date: row.grant_date,
      grant_type: row.grant_type,
      year: row.year,
      grant_days: row.grant_days,
      carry_over_days: row.carry_over_days,
      total_days_in_current_fifo: row.total_days,
      valid_from: row.valid_from,
      valid_to: row.valid_to,
      active_remaining_days_in_current_fifo: row.active_remaining_days,
      expired_days_in_current_fifo: row.expired_days
    }));
  const currentRemainingDays = Number(currentFifoBalance.current_remaining_days || 0);
  const withoutCarryOverRemainingDays = Number(
    fifoWithoutCarryOverBalance.current_remaining_days || 0
  );
  const legacyRemainingDays = Number(legacyBalance.current_remaining_days || 0);
  const suspectedDuplicateDays = Math.max(
    currentRemainingDays - withoutCarryOverRemainingDays,
    0
  );
  const result = {
    employee_id: targetEmployeeId,
    fiscal_year: targetFiscalYear,
    fiscal_start_month: fiscalStartMonth,
    as_of_date: formatDateValue(fiscalRange.end),
    legacy_build_balance_remaining_days: legacyRemainingDays,
    current_fifo_remaining_days: currentRemainingDays,
    fifo_without_carry_over_remaining_days: withoutCarryOverRemainingDays,
    carry_over_days_total: carryOverDaysTotal,
    suspected_duplicate_days: suspectedDuplicateDays,
    approved_used_days: approvedUsedDays,
    legacy_build_balance: legacyBalance,
    current_fifo_balance: currentFifoBalance,
    fifo_without_carry_over_balance: fifoWithoutCarryOverBalance,
    differences: {
      current_fifo_minus_legacy: currentRemainingDays - legacyRemainingDays,
      fifo_without_carry_over_minus_legacy:
        withoutCarryOverRemainingDays - legacyRemainingDays,
      removed_by_excluding_carry_over:
        currentRemainingDays - withoutCarryOverRemainingDays
    },
    year_end_finalized_records: yearEndFinalizedRecords.map(row => ({
      grant_id: row.grant_id,
      grant_date: row.grant_date,
      grant_type: row.grant_type,
      grant_days: row.grant_days,
      carry_over_days: row.carry_over_days,
      valid_from: row.valid_from,
      valid_to: row.valid_to,
      notes: row.notes
    })),
    original_grant_records: originalGrantRecords,
    valid_period_note: "carry_over除外FIFOでは元付与レコードの valid_from / valid_to を維持し、年跨ぎ確定行の carry_over_days は権利日数に加算しません。",
    difference_reason: getYearEndFinalizedBalanceModeDifferenceReason_({
      legacy_remaining_days: legacyRemainingDays,
      current_fifo_remaining_days: currentRemainingDays,
      fifo_without_carry_over_remaining_days: withoutCarryOverRemainingDays,
      carry_over_days_total: carryOverDaysTotal,
      suspected_duplicate_days: suspectedDuplicateDays
    })
  };

  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

function calculateFifoBalanceWithoutCarryOverFromContext_(employeeId, asOfDate, context) {
  const grants = (context.grants_by_employee[employeeId] || [])
    .filter(grant => grant.is_finalized)
    .filter(grant => grant.valid_from_date <= asOfDate)
    .map(grant => ({
      grant_id: grant.grant_id,
      grant_date: grant.grant_date,
      valid_from_date: grant.valid_from_date,
      valid_to_date: grant.valid_to_date,
      grant_type: grant.grant_type,
      year: grant.year,
      grant_days: grant.grant_days,
      excluded_carry_over_days: grant.carry_over_days,
      total_days: Number(grant.grant_days || 0),
      used_days: 0,
      remaining_days: Number(grant.grant_days || 0),
      active_remaining_days: 0,
      expired_days: 0,
      is_expired: false
    }))
    .sort((a, b) => {
      if (a.grant_date.getTime() !== b.grant_date.getTime()) {
        return a.grant_date - b.grant_date;
      }
      return String(a.grant_id).localeCompare(String(b.grant_id));
    });
  const usedRows = getFifoApprovedLeaveUseRowsFromContext_(employeeId, asOfDate, context);
  const allocations = [];

  usedRows.forEach(useRow => {
    let remainingUseDays = Number(useRow.days || 0);

    grants.forEach(grant => {
      if (remainingUseDays <= 0) return;
      if (grant.remaining_days <= 0) return;
      if (useRow.use_date < grant.valid_from_date) return;
      if (useRow.use_date > grant.valid_to_date) return;

      const consumedDays = Math.min(grant.remaining_days, remainingUseDays);
      grant.remaining_days -= consumedDays;
      grant.used_days += consumedDays;
      remainingUseDays -= consumedDays;

      allocations.push({
        request_id: useRow.request_id,
        use_date: formatDateValue(useRow.use_date),
        grant_id: grant.grant_id,
        consumed_days: consumedDays
      });
    });

    useRow.unallocated_days = remainingUseDays > 0 ? remainingUseDays : 0;
  });

  grants.forEach(grant => {
    const isExpired = grant.valid_to_date < asOfDate;
    grant.is_expired = isExpired;
    grant.expired_days = isExpired ? grant.remaining_days : 0;
    grant.active_remaining_days = isExpired ? 0 : grant.remaining_days;
  });

  return {
    employee_id: employeeId,
    as_of_date: formatDateValue(asOfDate),
    calculation_mode: "grant_days_only_carry_over_excluded",
    current_remaining_days: grants.reduce((sum, grant) => sum + grant.active_remaining_days, 0),
    total_granted_days: grants.reduce((sum, grant) => sum + grant.total_days, 0),
    excluded_carry_over_days_total: grants.reduce(
      (sum, grant) => sum + Number(grant.excluded_carry_over_days || 0),
      0
    ),
    used_days: usedRows.reduce((sum, row) => sum + Number(row.days || 0), 0),
    allocated_used_days: allocations.reduce((sum, row) => sum + Number(row.consumed_days || 0), 0),
    unallocated_used_days: usedRows.reduce((sum, row) => sum + Number(row.unallocated_days || 0), 0),
    expired_days: grants.reduce((sum, grant) => sum + grant.expired_days, 0),
    grant_details: grants.map(grant => ({
      grant_id: grant.grant_id,
      grant_date: formatDateValue(grant.grant_date),
      valid_from: formatDateValue(grant.valid_from_date),
      valid_to: formatDateValue(grant.valid_to_date),
      grant_type: grant.grant_type,
      year: grant.year,
      grant_days: grant.grant_days,
      excluded_carry_over_days: grant.excluded_carry_over_days,
      total_days: grant.total_days,
      used_days: grant.used_days,
      remaining_days: grant.remaining_days,
      active_remaining_days: grant.active_remaining_days,
      expired_days: grant.expired_days,
      is_expired: grant.is_expired
    })),
    used_details: usedRows.map(row => ({
      request_id: row.request_id,
      use_date: formatDateValue(row.use_date),
      days: row.days,
      unallocated_days: row.unallocated_days || 0
    })),
    allocations: allocations
  };
}

function getYearEndFinalizedBalanceModeDifferenceReason_(info) {
  const legacyDays = Number(info.legacy_remaining_days || 0);
  const currentFifoDays = Number(info.current_fifo_remaining_days || 0);
  const withoutCarryOverDays = Number(info.fifo_without_carry_over_remaining_days || 0);
  const carryOverDays = Number(info.carry_over_days_total || 0);
  const suspectedDuplicateDays = Number(info.suspected_duplicate_days || 0);

  if (currentFifoDays === legacyDays && withoutCarryOverDays === legacyDays) {
    return "差分はありません";
  }

  if (suspectedDuplicateDays > 0 && withoutCarryOverDays === legacyDays) {
    return "carry_over_days をFIFO権利日数から除外すると旧計算と一致します。繰越分の二重計上が疑われます。";
  }

  if (
    suspectedDuplicateDays > 0 &&
    Math.abs(withoutCarryOverDays - legacyDays) < Math.abs(currentFifoDays - legacyDays)
  ) {
    return "carry_over_days の除外で差分が縮小します。残る差分は元付与残・有効期限・使用割当の確認が必要です。";
  }

  if (carryOverDays > 0 && suspectedDuplicateDays === 0) {
    return "carry_over_days はありますが、試算日時点では残数差に現れていません。期限切れまたは消化状況を確認してください。";
  }

  return "carry_over_days 以外にも、元付与残・有効期限・使用割当による差分がある可能性があります。";
}

function debugFifoBalanceWithOpeningBalance(employeeId, asOfDateValue) {
  const targetEmployeeId = String(employeeId || "").trim();
  if (!targetEmployeeId) throw new Error("employeeId がありません");

  const asOfDate = asOfDateValue ? parseLocalDate(asOfDateValue) : parseLocalDate(new Date());
  const result = calculateFifoBalanceWithOpeningBalance_(
    targetEmployeeId,
    asOfDate
  );

  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

function compareFifoOpeningBalanceModes(employeeId, fiscalYear) {
  const targetEmployeeId = String(employeeId || "").trim();
  const targetFiscalYear = Number(fiscalYear || 0);

  if (!targetEmployeeId) throw new Error("employeeId がありません");
  if (!targetFiscalYear) throw new Error("対象年度がありません");

  const employeeMap = getEmployeeDetailMap();
  const fiscalStartMonth = getFiscalStartMonthByEmployeeId(targetEmployeeId, employeeMap);
  const fiscalRange = getFiscalYearRangeWithStart(targetFiscalYear, fiscalStartMonth);
  const legacyBalance = calculateYearlyBalanceByEmployee(targetEmployeeId, targetFiscalYear);
  const context = createFifoBalanceComparisonContext_(fiscalRange.end);
  const fifoWithoutCarryOver = calculateFifoBalanceWithoutCarryOverFromContext_(
    targetEmployeeId,
    fiscalRange.end,
    context
  );
  const fifoWithOpeningBalance = calculateFifoBalanceWithOpeningBalance_(
    targetEmployeeId,
    fiscalRange.end
  );
  const result = {
    employee_id: targetEmployeeId,
    fiscal_year: targetFiscalYear,
    fiscal_start_month: fiscalStartMonth,
    as_of_date: formatDateValue(fiscalRange.end),
    legacy_build_balance_remaining_days: Number(legacyBalance.current_remaining_days || 0),
    fifo_without_carry_over_remaining_days: Number(
      fifoWithoutCarryOver.current_remaining_days || 0
    ),
    fifo_with_opening_balance_remaining_days: Number(
      fifoWithOpeningBalance.current_remaining_days || 0
    ),
    opening_balance_days_total: Number(
      fifoWithOpeningBalance.opening_balance_days_total || 0
    ),
    excluded_year_end_carry_over_days_total: Number(
      fifoWithOpeningBalance.excluded_non_opening_carry_over_days_total || 0
    ),
    expiry_unconfirmed_days_total: Number(
      fifoWithOpeningBalance.expiry_unconfirmed_opening_balance_days_total || 0
    ),
    differences: {
      fifo_without_carry_over_minus_legacy:
        Number(fifoWithoutCarryOver.current_remaining_days || 0) -
        Number(legacyBalance.current_remaining_days || 0),
      fifo_with_opening_balance_minus_legacy:
        Number(fifoWithOpeningBalance.current_remaining_days || 0) -
        Number(legacyBalance.current_remaining_days || 0),
      restored_by_opening_balance:
        Number(fifoWithOpeningBalance.current_remaining_days || 0) -
        Number(fifoWithoutCarryOver.current_remaining_days || 0)
    },
    legacy_build_balance: legacyBalance,
    fifo_without_carry_over_balance: fifoWithoutCarryOver,
    fifo_with_opening_balance: fifoWithOpeningBalance,
    opening_balance_records: fifoWithOpeningBalance.opening_balance_records || [],
    excluded_carry_over_records: fifoWithOpeningBalance.excluded_carry_over_records || [],
    difference_reason: getFifoOpeningBalanceDifferenceReason_({
      legacy_remaining_days: legacyBalance.current_remaining_days,
      without_carry_over_remaining_days: fifoWithoutCarryOver.current_remaining_days,
      with_opening_balance_remaining_days: fifoWithOpeningBalance.current_remaining_days,
      opening_balance_days_total: fifoWithOpeningBalance.opening_balance_days_total,
      expiry_unconfirmed_days_total:
        fifoWithOpeningBalance.expiry_unconfirmed_opening_balance_days_total
    })
  };

  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

function debugFifoApprovedLeaveUseRows(employeeId, fiscalYear, asOfDateValue) {
  const targetEmployeeId = String(employeeId || "").trim();
  if (!targetEmployeeId) throw new Error("employeeId がありません");

  const asOfDate = asOfDateValue ? parseLocalDate(asOfDateValue) : parseLocalDate(new Date());
  const employeeMap = getEmployeeDetailMap();
  const fiscalStartMonth = getFiscalStartMonthByEmployeeId(targetEmployeeId, employeeMap);
  const targetFiscalYear = Number(
    fiscalYear || getFiscalYearFromDateWithStart(asOfDate, fiscalStartMonth)
  );
  const rows = getFifoApprovedLeaveUseDebugRows_(
    targetEmployeeId,
    targetFiscalYear,
    asOfDate
  );

  logFifoApprovedLeaveUseDebugRows_(
    targetEmployeeId,
    targetFiscalYear,
    asOfDate,
    rows
  );

  return rows;
}

function getFifoApprovedLeaveUseDebugRows_(employeeId, fiscalYear, asOfDate) {
  const targetEmployeeId = String(employeeId || "").trim();
  const employeeMap = getEmployeeDetailMap();
  const fiscalStartMonth = getFiscalStartMonthByEmployeeId(targetEmployeeId, employeeMap);
  const targetFiscalYear = Number(
    fiscalYear || getFiscalYearFromDateWithStart(asOfDate, fiscalStartMonth)
  );
  const fiscalRange = getFiscalYearRangeWithStart(targetFiscalYear, fiscalStartMonth);
  const sheet = getSheet("leave_requests");
  const headerInfo = requireHeaders(sheet, [
    "request_id",
    "employee_id",
    "start_date",
    "end_date",
    "days",
    "half_day",
    "status"
  ]);
  const data = sheet.getDataRange().getValues();
  const calendarMap = getCompanyCalendarMap();

  if (data.length <= 1) return [];

  return data.slice(1)
    .map(row => {
      const rowObj = rowToObject(row, headerInfo.headers);
      const rowEmployeeId = String(rowObj.employee_id || "").trim();
      const status = norm(rowObj.status);
      const requestType = String(rowObj.type || "paid_leave").trim();
      const hasDates = !!(rowObj.start_date && rowObj.end_date);
      const isSameEmployee = rowEmployeeId === targetEmployeeId;
      const isApproved = status === STATUS.APPROVED;
      const isPaidLeaveType = !requestType || requestType === "paid_leave";
      let dailyRows = [];
      let isInFiscalYear = false;
      let isBeforeAsOfDate = false;

      if (hasDates) {
        dailyRows = expandLeaveRequestToDailyRows(
          rowObj.start_date,
          rowObj.end_date,
          rowObj.days,
          rowObj.half_day,
          calendarMap
        ).map(item => {
          const useDate = parseLocalDate(item.date);
          const inFiscalYear = isDateInRange(useDate, fiscalRange.start, fiscalRange.end);
          const beforeAsOfDate = useDate <= asOfDate;

          if (inFiscalYear) isInFiscalYear = true;
          if (beforeAsOfDate) isBeforeAsOfDate = true;

          return {
            use_date: formatDateValue(useDate),
            days: Number(item.days || 0),
            is_in_fiscal_year: inFiscalYear,
            is_before_as_of_date: beforeAsOfDate
          };
        });
      }

      return {
        request_id: String(rowObj.request_id || ""),
        employee_id: rowEmployeeId,
        start_date: formatDateValue(rowObj.start_date),
        end_date: formatDateValue(rowObj.end_date),
        days: rowObj.days || 0,
        half_day: String(rowObj.half_day || ""),
        type: requestType || "",
        status: String(rowObj.status || ""),
        fiscal_year: targetFiscalYear,
        is_same_employee: isSameEmployee,
        is_approved: isApproved,
        is_paid_leave_type: isPaidLeaveType,
        is_in_fiscal_year: isInFiscalYear,
        is_before_as_of_date: isBeforeAsOfDate,
        daily_rows: dailyRows,
        excluded_reason: getFifoDebugExcludedReason_({
          has_dates: hasDates,
          is_same_employee: isSameEmployee,
          is_approved: isApproved,
          is_paid_leave_type: isPaidLeaveType,
          is_in_fiscal_year: isInFiscalYear,
          is_before_as_of_date: isBeforeAsOfDate
        })
      };
    })
    .filter(row => row.is_same_employee || row.employee_id === targetEmployeeId);
}

function getFifoDebugExcludedReason_(flags) {
  const reasons = [];

  if (!flags.is_same_employee) reasons.push("employee_id不一致");
  if (!flags.is_approved) reasons.push("statusがapprovedではない");
  if (!flags.is_paid_leave_type) reasons.push("typeがpaid_leaveではない");
  if (!flags.has_dates) reasons.push("start_date/end_date不足");
  if (!flags.is_in_fiscal_year) reasons.push("年度範囲外");
  if (!flags.is_before_as_of_date) reasons.push("asOfDateより後");

  return reasons.length > 0 ? reasons.join(" / ") : "";
}

function logFifoApprovedLeaveUseDebugRows_(employeeId, fiscalYear, asOfDate, rows) {
  const header = [
    "request_id",
    "employee_id",
    "start_date",
    "end_date",
    "days",
    "half_day",
    "type",
    "status",
    "fiscal_year",
    "is_same_employee",
    "is_approved",
    "is_paid_leave_type",
    "is_in_fiscal_year",
    "is_before_as_of_date",
    "excluded_reason"
  ];

  Logger.log(
    "FIFO使用日数取得デバッグ" +
    " / employee_id=" + employeeId +
    " / fiscalYear=" + fiscalYear +
    " / asOfDate=" + formatDateValue(asOfDate) +
    " / count=" + rows.length
  );
  Logger.log(header.join("\t"));

  rows.forEach(row => {
    Logger.log(header.map(key => row[key]).join("\t"));
  });
}
