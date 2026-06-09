/* =========================
   FIFO残日数比較（管理画面）
   debug.gs から動作を変えずに移動
========================= */

function compareFifoBalanceWithBuildBalance(employeeId, fiscalYear, asOfDateValue) {
  const targetEmployeeId = String(employeeId || "").trim();
  if (!targetEmployeeId) throw new Error("employeeId がありません");

  const asOfDate = asOfDateValue ? parseLocalDate(asOfDateValue) : parseLocalDate(new Date());
  const employeeMap = getEmployeeDetailMap();
  const fiscalStartMonth = getFiscalStartMonthByEmployeeId(targetEmployeeId, employeeMap);
  const targetFiscalYear = Number(
    fiscalYear || getFiscalYearFromDateWithStart(asOfDate, fiscalStartMonth)
  );
  const legacyBalance = calculateYearlyBalanceByEmployee(targetEmployeeId, targetFiscalYear);
  const fifoBalance = calculateFifoPaidLeaveBalance(targetEmployeeId, asOfDate);

  return {
    employee_id: targetEmployeeId,
    fiscal_year: targetFiscalYear,
    fiscal_start_month: fiscalStartMonth,
    as_of_date: formatDateValue(asOfDate),
    legacy_balance: legacyBalance,
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
    }
  };
}

function compareFifoBalanceForAllEmployees(fiscalYear, asOfDateValue) {
  const asOfDate = asOfDateValue ? parseLocalDate(asOfDateValue) : parseLocalDate(new Date());
  const rows = getFifoBalanceComparisonRows_(fiscalYear, asOfDate);

  logFifoBalanceComparisonRows_(
    "FIFO残日数比較（全対象社員）",
    rows,
    fiscalYear,
    asOfDate
  );

  return rows;
}

function compareFifoBalanceDifferencesOnly(fiscalYear, asOfDateValue) {
  const asOfDate = asOfDateValue ? parseLocalDate(asOfDateValue) : parseLocalDate(new Date());
  const rows = getFifoBalanceComparisonRows_(fiscalYear, asOfDate)
    .filter(row => row.has_difference === true);

  logFifoBalanceComparisonRows_(
    "FIFO残日数比較（差分ありのみ）",
    rows,
    fiscalYear,
    asOfDate
  );

  return rows;
}

function compareFifoBalanceDifferencesForAdmin(fiscalYear, asOfDateValue, employeeId, limit, offset) {
  const asOfDate = asOfDateValue ? parseLocalDate(asOfDateValue) : parseLocalDate(new Date());
  const targetEmployeeId = String(employeeId || "").trim();
  const page = normalizePagingOptions_({
    limit: targetEmployeeId ? 1 : limit,
    offset: targetEmployeeId ? 0 : offset
  });
  const comparison = getFifoBalanceComparisonRows_(fiscalYear, asOfDate, {
    employee_id: targetEmployeeId,
    limit: page.limit,
    offset: page.offset,
    include_paging: true
  });
  const rows = comparison.rows || [];
  const differenceRows = rows.filter(row => row.has_difference === true);
  const displayRows = rows.filter(row =>
    row.has_difference === true ||
    row.has_validity_warning === true
  );

  return {
    ok: true,
    fiscal_year: Number(fiscalYear || 0),
    as_of_date: formatDateValue(asOfDate),
    employee_id: targetEmployeeId,
    limit: page.limit,
    offset: page.offset,
    target_limited: !targetEmployeeId,
    scanned_count: rows.length,
    difference_count: differenceRows.length,
    warning_count: rows.filter(row => row.has_validity_warning === true).length,
    total_count: comparison.total_count || rows.length,
    has_prev: page.offset > 0,
    has_next: page.offset + page.limit < Number(comparison.total_count || rows.length),
    rows: displayRows
  };
}

function getFifoBalanceComparisonRows_(fiscalYear, asOfDate, options) {
  options = options || {};
  const targetEmployeeId = String(options.employee_id || "").trim();
  const limit = Number(options.limit || 0);
  const offset = Math.max(0, Number(options.offset || 0));
  const context = createFifoBalanceComparisonContext_(asOfDate);
  const employees = getEmployeesForAdmin()
    .filter(emp => isFifoBalanceCompareTargetEmployee_(emp))
    .filter(emp => {
      if (!targetEmployeeId) return true;
      return String(emp.employee_id || "").trim() === targetEmployeeId;
    });

  const targetEmployees = limit > 0
    ? employees.slice(offset, offset + limit)
    : employees;

  const rows = targetEmployees.map(emp => {
    const fiscalStartMonth = Number(emp.fiscal_start_month || 4);
    const targetFiscalYear = Number(
      fiscalYear || getFiscalYearFromDateWithStart(asOfDate, fiscalStartMonth)
    );

    return buildFifoBalanceComparisonRow_(emp, targetFiscalYear, asOfDate, context);
  });

  if (options.include_paging === true) {
    return {
      total_count: employees.length,
      offset: offset,
      limit: limit,
      rows: rows
    };
  }

  return rows;
}

function buildFifoBalanceComparisonRow_(emp, fiscalYear, asOfDate, context) {
  const employeeId = String(emp.employee_id || "").trim();
  const comparison = context
    ? compareFifoBalanceWithBuildBalanceFromContext_(emp, fiscalYear, asOfDate, context)
    : compareFifoBalanceWithBuildBalance(
    employeeId,
    fiscalYear,
    asOfDate
  );
  const legacyBalance = comparison.legacy_balance || {};
  const fifoBalance = comparison.fifo_balance || {};
  const remainingDifference = Number(comparison.difference.current_remaining_days || 0);
  const usedDifference = Number(comparison.difference.used_days || 0);
  const expiredDifference = Number(comparison.difference.expired_days || 0);
  const futureInfo = context
    ? getFutureApprovedUsedInfoForFifoComparisonFromContext_(
      employeeId,
      fiscalYear,
      asOfDate,
      Number(emp.fiscal_start_month || 4),
      context
    )
    : getFutureApprovedUsedInfoForFifoComparison_(
    employeeId,
    fiscalYear,
    asOfDate
  );
  const futureApprovedUsedDays = Number(futureInfo.future_approved_used_days || 0);
  const adjustedLegacyUsedDays = Number(legacyBalance.used_days || 0) - futureApprovedUsedDays;
  const adjustedLegacyRemainingDays =
    Number(legacyBalance.current_remaining_days || 0) + futureApprovedUsedDays;
  const adjustedRemainingDifference =
    Number(fifoBalance.current_remaining_days || 0) - adjustedLegacyRemainingDays;
  const approvedRequestIds = {};

  (fifoBalance.used_details || []).forEach(row => {
    if (row.request_id) approvedRequestIds[String(row.request_id)] = true;
  });

  return {
    employee_id: employeeId,
    name: String(emp.name || ""),
    display_name: String(emp.display_name || ""),
    company_code: String(emp.company_code || ""),
    department: String(emp.department || ""),
    fiscal_start_month: Number(emp.fiscal_start_month || 4),
    legacy_current_remaining_days: Number(legacyBalance.current_remaining_days || 0),
    fifo_current_remaining_days: Number(fifoBalance.current_remaining_days || 0),
    remaining_difference: remainingDifference,
    legacy_used_days: Number(legacyBalance.used_days || 0),
    fifo_used_days: Number(fifoBalance.used_days || 0),
    used_difference: usedDifference,
    future_approved_used_days: futureApprovedUsedDays,
    future_approved_request_count: futureInfo.future_approved_request_count,
    adjusted_legacy_used_days: adjustedLegacyUsedDays,
    adjusted_legacy_remaining_days: adjustedLegacyRemainingDays,
    adjusted_remaining_difference: adjustedRemainingDifference,
    legacy_expired_days: Number(legacyBalance.expired_days || 0),
    fifo_expired_days: Number(fifoBalance.expired_days || 0),
    expired_difference: expiredDifference,
    opening_balance_days_total: Number(fifoBalance.opening_balance_days_total || 0),
    expiry_unconfirmed_days_total: Number(
      fifoBalance.expiry_unconfirmed_opening_balance_days_total || 0
    ),
    validity_warning: String(fifoBalance.validity_warning || ""),
    has_validity_warning: !!fifoBalance.validity_warning,
    grant_count: (fifoBalance.grant_details || []).length,
    approved_request_count: Object.keys(approvedRequestIds).length,
    difference_reason: getFifoComparisonDifferenceReason_({
      remaining_difference: remainingDifference,
      adjusted_remaining_difference: adjustedRemainingDifference,
      future_approved_used_days: futureApprovedUsedDays,
      expired_difference: expiredDifference,
      fiscal_start_month: Number(emp.fiscal_start_month || 4)
    }),
    has_difference:
      remainingDifference !== 0 ||
      usedDifference !== 0 ||
      expiredDifference !== 0
  };
}

function compareFifoBalanceWithBuildBalanceFromContext_(emp, fiscalYear, asOfDate, context) {
  const employeeId = String(emp.employee_id || "").trim();
  const fiscalStartMonth = Number(emp.fiscal_start_month || 4);
  const legacyBalance = calculateLegacyBalanceFromFifoContext_(
    employeeId,
    fiscalYear,
    fiscalStartMonth,
    context
  );
  const fifoBalance = calculateFifoBalanceWithOpeningBalanceFromContext_(
    employeeId,
    asOfDate,
    context
  );

  return {
    employee_id: employeeId,
    fiscal_year: fiscalYear,
    fiscal_start_month: fiscalStartMonth,
    as_of_date: formatDateValue(asOfDate),
    legacy_balance: legacyBalance,
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
    }
  };
}

function calculateLegacyBalanceFromFifoContext_(employeeId, fiscalYear, fiscalStartMonth, context) {
  const grants = context.grants_by_employee[employeeId] || [];
  const requests = context.requests_by_employee[employeeId] || [];
  const grantInfo = {
    employee_id: employeeId,
    grant_days: 0,
    carry_over_days: 0
  };
  const range = getFiscalYearRangeWithStart(fiscalYear, fiscalStartMonth);
  let usedDays = 0;

  grants.forEach(grant => {
    const rowYear = getFiscalYearFromDateWithStart(grant.grant_date, fiscalStartMonth);
    if (rowYear !== Number(fiscalYear)) return;

    grantInfo.grant_days += Number(grant.grant_days || 0);
    grantInfo.carry_over_days += Number(grant.carry_over_days || 0);
  });

  requests.forEach(rowObj => {
    const status = norm(rowObj.status);
    if (status !== STATUS.APPROVED) return;
    if (!rowObj.start_date || !rowObj.end_date) return;

    const dailyRows = expandLeaveRequestToDailyRows(
      rowObj.start_date,
      rowObj.end_date,
      rowObj.days,
      rowObj.half_day,
      context.calendar_map
    );

    dailyRows.forEach(item => {
      if (!isDateInRange(item.date, range.start, range.end)) return;
      usedDays += Number(item.days || 0);
    });
  });

  return buildBalance(employeeId, grantInfo, usedDays);
}

function getFutureApprovedUsedInfoForFifoComparisonFromContext_(
  employeeId,
  fiscalYear,
  asOfDate,
  fiscalStartMonth,
  context
) {
  const requests = context.requests_by_employee[employeeId] || [];
  const fiscalRange = getFiscalYearRangeWithStart(fiscalYear, fiscalStartMonth);
  const futureRequestIds = {};
  let futureUsedDays = 0;

  requests.forEach(rowObj => {
    const status = norm(rowObj.status);
    const requestType = String(rowObj.type || "paid_leave").trim();

    if (status !== STATUS.APPROVED) return;
    if (requestType && requestType !== "paid_leave") return;
    if (!rowObj.start_date || !rowObj.end_date) return;

    const dailyRows = expandLeaveRequestToDailyRows(
      rowObj.start_date,
      rowObj.end_date,
      rowObj.days,
      rowObj.half_day,
      context.calendar_map
    );

    dailyRows.forEach(item => {
      const useDate = parseLocalDate(item.date);
      if (!isDateInRange(useDate, fiscalRange.start, fiscalRange.end)) return;
      if (useDate <= asOfDate) return;

      futureUsedDays += Number(item.days || 0);
      if (rowObj.request_id) futureRequestIds[String(rowObj.request_id)] = true;
    });
  });

  return {
    future_approved_used_days: futureUsedDays,
    future_approved_request_count: Object.keys(futureRequestIds).length
  };
}

function getFutureApprovedUsedInfoForFifoComparison_(employeeId, fiscalYear, asOfDate) {
  const debugRows = getFifoApprovedLeaveUseDebugRows_(
    employeeId,
    fiscalYear,
    asOfDate
  );
  const futureRequestIds = {};
  let futureUsedDays = 0;

  debugRows.forEach(row => {
    if (!row.is_same_employee) return;
    if (!row.is_approved) return;
    if (!row.is_paid_leave_type) return;

    (row.daily_rows || []).forEach(dailyRow => {
      if (!dailyRow.is_in_fiscal_year) return;
      if (dailyRow.is_before_as_of_date) return;

      futureUsedDays += Number(dailyRow.days || 0);
      if (row.request_id) futureRequestIds[String(row.request_id)] = true;
    });
  });

  return {
    future_approved_used_days: futureUsedDays,
    future_approved_request_count: Object.keys(futureRequestIds).length
  };
}

function getFifoComparisonDifferenceReason_(info) {
  const remainingDifference = Number(info.remaining_difference || 0);
  const adjustedRemainingDifference = Number(info.adjusted_remaining_difference || 0);
  const futureApprovedUsedDays = Number(info.future_approved_used_days || 0);
  const expiredDifference = Number(info.expired_difference || 0);
  const fiscalStartMonth = Number(info.fiscal_start_month || 4);

  if (remainingDifference === 0 && expiredDifference === 0) return "";

  if (
    futureApprovedUsedDays > 0 &&
    Math.abs(adjustedRemainingDifference) < Math.abs(remainingDifference)
  ) {
    return "未来の承認済み申請による差分";
  }

  if (expiredDifference !== 0) {
    return "期限切れ計算方式の違い";
  }

  if (fiscalStartMonth !== 4) {
    return "fiscal_start_month / 年度範囲の違い";
  }

  return "要確認";
}

function logFifoBalanceComparisonRows_(title, rows, fiscalYear, asOfDate) {
  const header = [
    "employee_id",
    "name",
    "display_name",
    "company_code",
    "department",
    "fiscal_start_month",
    "legacy_current_remaining_days",
    "fifo_current_remaining_days",
    "remaining_difference",
    "legacy_used_days",
    "fifo_used_days",
    "used_difference",
    "future_approved_used_days",
    "future_approved_request_count",
    "adjusted_legacy_used_days",
    "adjusted_legacy_remaining_days",
    "adjusted_remaining_difference",
    "legacy_expired_days",
    "fifo_expired_days",
    "expired_difference",
    "grant_count",
    "approved_request_count",
    "difference_reason",
    "has_difference"
  ];

  Logger.log(
    title +
    " / fiscalYear=" + (fiscalYear || "社員ごとの基準日") +
    " / asOfDate=" + formatDateValue(asOfDate) +
    " / count=" + rows.length
  );
  Logger.log(header.join("\t"));

  rows.slice(0, 20).forEach(row => {
    Logger.log(header.map(key => row[key]).join("\t"));
  });

  if (rows.length > 20) {
    Logger.log("ログ出力は先頭20件までに制限しました。残り " + (rows.length - 20) + " 件");
  }
}
