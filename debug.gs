/* =========================
   debug / audit / compare 関連
   Code.gs から動作を変えずに移動
========================= */

function debugPartnerLeaveGrantDays2026_EMP0062() {
  const employeeId = "EMP0062";
  const fiscalYearStartDate = parseLocalDate("2026-06-01");
  const emp = getEmployeesForAdmin().find(row =>
    String(row.employee_id || "").trim() === employeeId
  );

  Logger.log("=== EMP0062 2026年度新規付与日数 debug ===");
  Logger.log("データ変更: なし");

  if (!emp) {
    Logger.log("employee_id: " + employeeId);
    Logger.log("エラー: employees シートに EMP0062 が見つかりません");
    Logger.log("=== debug 完了 ===");
    return {
      ok: false,
      employee_id: employeeId,
      error: "employees シートに EMP0062 が見つかりません"
    };
  }

  const monthsWorked = emp.hire_date
    ? getMonthsWorked_(parseLocalDate(emp.hire_date), fiscalYearStartDate)
    : 0;
  const yearlyGrantDays = getYearlyGrantDays_(monthsWorked);
  const zeroReason = getPartnerLeaveGrantDaysZeroReason_(
    emp.hire_date,
    monthsWorked,
    yearlyGrantDays
  );
  const initialGrantInfo = emp.hire_date
    ? getInitialPaidLeaveGrantInfo_(emp)
    : null;
  const companyBasisGrantInfo = emp.hire_date
    ? getCompanyBasisYearlyGrantInfo_(
        emp.hire_date,
        fiscalYearStartDate,
        Number(emp.fiscal_start_month || 0)
      )
    : null;
  const result = {
    ok: true,
    employee_id: String(emp.employee_id || ""),
    name: String(emp.name || ""),
    display_name: String(emp.display_name || ""),
    hire_date: String(emp.hire_date || ""),
    fiscal_start_month: Number(emp.fiscal_start_month || 0),
    leave_management_target: emp.leave_management_target === true,
    employment_status: String(emp.employment_status || ""),
    company_code: String(emp.company_code || ""),
    company_name: String(emp.company_name || ""),
    fiscalYearStartDate: formatDateValue(fiscalYearStartDate),
    monthsWorked: monthsWorked,
    simpleMonthsWorkedGrantDays: yearlyGrantDays,
    yearlyGrantDaysZeroReason: zeroReason,
    firstCompanyBasisDate: companyBasisGrantInfo
      ? formatDateValue(companyBasisGrantInfo.first_company_basis_date)
      : "",
    companyBasisGrantNumber: companyBasisGrantInfo
      ? Number(companyBasisGrantInfo.company_basis_grant_number || 0)
      : 0,
    companyBasisEquivalentMonths: companyBasisGrantInfo
      ? Number(companyBasisGrantInfo.equivalent_months || 0)
      : 0,
    yearlyGrantDays: companyBasisGrantInfo
      ? Number(companyBasisGrantInfo.grant_days || 0)
      : 0,
    initialGrantDate: initialGrantInfo ? formatDateValue(initialGrantInfo.grant_date) : "",
    sixMonthDate: initialGrantInfo ? formatDateValue(initialGrantInfo.six_month_date) : "",
    companyBasisDate: initialGrantInfo ? formatDateValue(initialGrantInfo.company_basis_date) : "",
    initialGrantReason: initialGrantInfo ? String(initialGrantInfo.grant_reason || "") : ""
  };

  Logger.log("employee_id: " + result.employee_id);
  Logger.log("name: " + result.name);
  Logger.log("display_name: " + result.display_name);
  Logger.log("hire_date: " + result.hire_date);
  Logger.log("fiscal_start_month: " + result.fiscal_start_month);
  Logger.log("leave_management_target: " + result.leave_management_target);
  Logger.log("employment_status: " + result.employment_status);
  Logger.log("company_code: " + result.company_code);
  Logger.log("company_name: " + result.company_name);
  Logger.log("fiscalYearStartDate: " + result.fiscalYearStartDate);
  Logger.log("monthsWorked: " + result.monthsWorked);
  Logger.log("単純勤続月数での getYearlyGrantDays_ 戻り値: " + result.simpleMonthsWorkedGrantDays);
  Logger.log(
    "単純勤続月数方式で yearlyGrantDays が0になる理由: " +
    (result.yearlyGrantDaysZeroReason || "0日ではありません")
  );
  Logger.log("最初の会社基準日: " + result.firstCompanyBasisDate);
  Logger.log("会社基準日付与回数: " + result.companyBasisGrantNumber);
  Logger.log("会社基準日方式の換算月数: " + result.companyBasisEquivalentMonths);
  Logger.log("dry-run で適用する新規付与日数: " + result.yearlyGrantDays);
  Logger.log("初回付与予定日: " + result.initialGrantDate);
  Logger.log("入社6か月到達日: " + result.sixMonthDate);
  Logger.log("会社基準日: " + result.companyBasisDate);
  Logger.log("初回付与判定理由: " + result.initialGrantReason);
  Logger.log("=== debug 完了 ===");

  return result;
}

function debugMainLeaveGrantMethodDiff2026() {
  const previousFiscalYear = 2025;
  const nextFiscalYear = 2026;
  const fiscalStartMonth = 4;
  const nextFiscalYearStartDate = parseLocalDate("2026-04-01");
  const previousFiscalYearEndDate = addDaysLocal_(nextFiscalYearStartDate, -1);
  const context = createFifoBalanceComparisonContext_(previousFiscalYearEndDate);
  const finalizedMap = getYearlyGrantFinalizedMap_(nextFiscalYear);

  const rows = getEmployeesForAdmin()
    .filter(emp => {
      const status = String(emp.employment_status || "").trim().toLowerCase();
      return (
        String(emp.company_code || "").trim().toUpperCase() === "MAIN" &&
        Number(emp.fiscal_start_month || 0) === fiscalStartMonth &&
        emp.leave_management_target === true &&
        (status === "active" || status === "在職")
      );
    })
    .sort((a, b) =>
      String(a.employee_id || "").localeCompare(String(b.employee_id || ""))
    )
    .map(emp => {
      const errors = [];
      let candidate = null;
      let companyBasisInfo = null;

      if (!emp.hire_date) {
        errors.push("入社日が未入力です");
      }

      try {
        candidate = buildYearEndCarryOverCandidate_(
          emp,
          previousFiscalYear,
          context,
          finalizedMap
        );
        companyBasisInfo = getCompanyBasisYearlyGrantInfo_(
          emp.hire_date,
          nextFiscalYearStartDate,
          fiscalStartMonth
        );
      } catch (e) {
        errors.push(e && e.message ? e.message : String(e));
      }

      const currentMethodDays = candidate
        ? Number(candidate.new_grant_days || 0)
        : "";
      const companyBasisDays = companyBasisInfo
        ? Number(companyBasisInfo.grant_days || 0)
        : "";

      return {
        employee_id: String(emp.employee_id || ""),
        name: String(emp.name || ""),
        display_name: String(emp.display_name || ""),
        hire_date: String(emp.hire_date || ""),
        previous_remaining_days: candidate
          ? Number(candidate.previous_remaining_days || 0)
          : "",
        current_method_new_grant_days: currentMethodDays,
        company_basis_new_grant_days: companyBasisDays,
        difference_days:
          currentMethodDays === "" || companyBasisDays === ""
            ? ""
            : companyBasisDays - currentMethodDays,
        company_basis_grant_number: companyBasisInfo
          ? Number(companyBasisInfo.company_basis_grant_number || 0)
          : "",
        has_2026_yearly_record: !!finalizedMap[
          String(emp.employee_id || "").trim()
        ],
        errors: errors
      };
    });

  const differenceRows = rows.filter(
    row => row.difference_days !== "" && Number(row.difference_days) !== 0
  );
  const errorRows = rows.filter(row => row.errors.length > 0);
  const result = {
    ok: errorRows.length === 0,
    debug_only: true,
    data_changed: false,
    company_code: "MAIN",
    fiscal_start_month: fiscalStartMonth,
    previous_fiscal_year: previousFiscalYear,
    next_fiscal_year: nextFiscalYear,
    previous_fiscal_year_end_date: formatDateValue(previousFiscalYearEndDate),
    next_fiscal_year_start_date: formatDateValue(nextFiscalYearStartDate),
    target_employee_count: rows.length,
    difference_count: differenceRows.length,
    error_count: errorRows.length,
    has_difference: differenceRows.length > 0,
    rows: rows
  };

  Logger.log("=== MAIN 2026年度 付与判定方式差分 debug ===");
  Logger.log("データ変更: なし");
  Logger.log(
    "対象: company_code=MAIN / fiscal_start_month=4 / " +
      "leave_management_target=TRUE / employment_status=active または 在職"
  );
  Logger.log(
    "期間: 2025年度末 " +
      result.previous_fiscal_year_end_date +
      " → 2026年度開始 " +
      result.next_fiscal_year_start_date
  );
  Logger.log(
    [
      "employee_id",
      "name",
      "display_name",
      "hire_date",
      "2025年度末残日数",
      "現在方式の新規付与日数",
      "会社基準日方式の新規付与日数",
      "差分",
      "会社基準日付与回数",
      "2026年度レコード有無",
      "注意・エラー"
    ].join("\t")
  );
  rows.forEach(row => {
    Logger.log(
      [
        row.employee_id,
        row.name,
        row.display_name,
        row.hire_date,
        row.previous_remaining_days,
        row.current_method_new_grant_days,
        row.company_basis_new_grant_days,
        row.difference_days,
        row.company_basis_grant_number,
        row.has_2026_yearly_record ? "あり" : "なし",
        row.errors.length > 0 ? row.errors.join(" / ") : "なし"
      ].join("\t")
    );
  });
  Logger.log("対象人数: " + result.target_employee_count);
  Logger.log("差分あり人数: " + result.difference_count);
  Logger.log("エラー人数: " + result.error_count);
  Logger.log(
    "結論: " +
      (result.has_difference
        ? "MAINにも会社基準日方式との差分があります。統一UI実装前に対象社員を確認してください。"
        : "MAINでは両方式の差分はありません。")
  );
  Logger.log("=== debug 完了 ===");

  return result;
}

function getPartnerLeaveGrantDaysZeroReason_(hireDate, monthsWorked, yearlyGrantDays) {
  if (Number(yearlyGrantDays || 0) > 0) return "";
  if (!hireDate) return "employees シートの hire_date が空です";

  return (
    "2026-06-01 時点の勤続月数が " +
    Number(monthsWorked || 0) +
    "か月で、11日付与の条件である18か月以上を満たしていません"
  );
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

function auditCarryOverGrantRowsForFifoMigration() {
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
  const employeeMap = getEmployeeDetailMap();

  if (data.length <= 1) {
    return {
      row_count: 0,
      needs_review_count: 0,
      carry_over_days_total: 0,
      needs_review_carry_over_days_total: 0,
      rows: []
    };
  }

  const allRows = data.slice(1)
    .map(row => rowToObject(row, headerInfo.headers))
    .filter(rowObj => rowObj.employee_id && rowObj.grant_date)
    .map(rowObj => ({
      grant_id: String(rowObj.grant_id || ""),
      employee_id: String(rowObj.employee_id || "").trim(),
      grant_date_value: parseLocalDate(rowObj.grant_date),
      grant_date: formatDateValue(rowObj.grant_date),
      grant_days: Number(rowObj.grant_days || 0),
      carry_over_days: Number(rowObj.carry_over_days || 0),
      valid_from: formatDateValue(rowObj.valid_from || rowObj.grant_date),
      valid_to: formatDateValue(rowObj.valid_to || ""),
      grant_type: String(rowObj.grant_type || ""),
      year: rowObj.year || "",
      notes: String(rowObj.notes || "")
    }));
  const carryOverRows = allRows.filter(row => row.carry_over_days > 0);
  const resultRows = carryOverRows.map(row => {
    const earlierGrantRows = allRows.filter(candidate =>
      candidate.employee_id === row.employee_id &&
      candidate.grant_id !== row.grant_id &&
      candidate.grant_days > 0 &&
      candidate.grant_date_value < row.grant_date_value
    );
    const hasEarlierGrantDaysRecord = earlierGrantRows.length > 0;
    const isCarryOverOnlyRow = row.grant_days <= 0;
    const needsReview = isCarryOverOnlyRow || !hasEarlierGrantDaysRecord;
    const employee = employeeMap[row.employee_id] || {};

    return {
      employee_id: row.employee_id,
      name: String(employee.name || ""),
      display_name: String(employee.display_name || ""),
      grant_id: row.grant_id,
      grant_date: row.grant_date,
      grant_type: row.grant_type,
      year: row.year,
      grant_days: row.grant_days,
      carry_over_days: row.carry_over_days,
      valid_from: row.valid_from,
      valid_to: row.valid_to,
      notes: row.notes,
      is_carry_over_only_row: isCarryOverOnlyRow,
      has_earlier_grant_days_record: hasEarlierGrantDaysRecord,
      earlier_grant_ids: earlierGrantRows.map(candidate => candidate.grant_id),
      needs_review: needsReview,
      review_reason: isCarryOverOnlyRow
        ? "carry_over_days のみで残を持つ行です。FIFO除外前に権利の元データを確認してください。"
        : !hasEarlierGrantDaysRecord
          ? "繰越元となる過去の grant_days 行が確認できません。初期移行残または手入力残の可能性があります。"
          : "過去の grant_days 行があり、年度集計用 carry_over_days の可能性があります。"
    };
  });
  const needsReviewRows = resultRows.filter(row => row.needs_review);
  const result = {
    row_count: resultRows.length,
    needs_review_count: needsReviewRows.length,
    carry_over_days_total: resultRows.reduce(
      (sum, row) => sum + Number(row.carry_over_days || 0),
      0
    ),
    needs_review_carry_over_days_total: needsReviewRows.reduce(
      (sum, row) => sum + Number(row.carry_over_days || 0),
      0
    ),
    rows: resultRows
  };

  Logger.log(JSON.stringify(result, null, 2));
  return result;
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
