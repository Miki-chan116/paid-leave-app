/* =========================
   Leave rollover debug
   debug.gs から動作を変えずに移動
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

