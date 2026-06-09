/* =========================
   FIFO debug / validation
   debug.gs から動作を変えずに移動
========================= */

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

