/* =========================
   年度切替関連
   Code.gs から動作を変えずに移動
========================= */

function getYearEndCarryOverCandidates(fiscalYear, options) {
  const targetFiscalYear = Number(fiscalYear || getFiscalYearFromDate(new Date()));
  const opts = options || {};
  const page = normalizePagingOptions_(opts);
  const employeeIdFilter = String(opts.employee_id || "").trim();
  const companyCodeFilter = String(opts.company_code || "").trim().toUpperCase();
  const departmentFilter = String(opts.department || "").trim();
  const includeFinalized = opts.include_finalized === true;
  const context = createFifoBalanceComparisonContext_(new Date());
  const finalizedMap = getYearlyGrantFinalizedMap_(targetFiscalYear + 1);

  const employees = getEmployeesForAdmin()
    .filter(emp => {
      const status = String(emp.employment_status || "").trim().toLowerCase();
      const isActive = status === "active" || status === "在職";
      if (!isActive) return false;
      if (emp.leave_management_target !== true) return false;
      if (employeeIdFilter && String(emp.employee_id || "").trim() !== employeeIdFilter) return false;
      if (companyCodeFilter && String(emp.company_code || "").trim().toUpperCase() !== companyCodeFilter) return false;
      if (departmentFilter && String(emp.department || "").trim() !== departmentFilter) return false;
      return true;
    })
    .filter(emp => {
      if (includeFinalized) return true;
      return !finalizedMap[String(emp.employee_id || "").trim()];
    })
    .sort((a, b) => {
      if (String(a.company_code || "") !== String(b.company_code || "")) {
        return String(a.company_code || "").localeCompare(String(b.company_code || ""));
      }
      return String(a.employee_id || "").localeCompare(String(b.employee_id || ""));
    });
  const pageEmployees = employees.slice(page.offset, page.offset + page.limit);
  const rows = pageEmployees
    .map(emp => buildYearEndCarryOverCandidate_(emp, targetFiscalYear, context, finalizedMap));

  return {
    ok: true,
    fiscal_year: targetFiscalYear,
    row_count: rows.length,
    total_count: employees.length,
    offset: page.offset,
    limit: page.limit,
    has_prev: page.offset > 0,
    has_next: page.offset + page.limit < employees.length,
    warning_count: rows.filter(row => row.validity_warning).length,
    rows: rows
  };
}

function buildYearEndCarryOverCandidate_(emp, fiscalYear, context, finalizedMap) {
  const employeeId = String(emp.employee_id || "").trim();
  const fiscalStartMonth = Number(emp.fiscal_start_month || 4);
  const fiscalRange = getFiscalYearRangeWithStart(fiscalYear, fiscalStartMonth);
  const fiscalYearEndDate = fiscalRange.end;
  const nextFiscalYearStartDate = addDaysLocal_(fiscalYearEndDate, 1);
  const fifoBalance = calculateFifoBalanceWithOpeningBalanceFromContext_(
    employeeId,
    fiscalYearEndDate,
    context
  );
  const previousRemainingDays = Number(fifoBalance.current_remaining_days || 0);
  const carryOverCandidateDays = Math.min(previousRemainingDays, 20);
  const carryOverLimitExpiredDays = Math.max(previousRemainingDays - 20, 0);
  const expiredDays =
    Number(fifoBalance.expired_days || 0) +
    carryOverLimitExpiredDays;
  const months = emp.hire_date
    ? getMonthsWorked_(parseLocalDate(emp.hire_date), nextFiscalYearStartDate)
    : 0;
  const newGrantDays = getYearlyGrantDays_(months);
  const estimatedAfterGrantDays = carryOverCandidateDays + newGrantDays;
  const isFinalized = !!(finalizedMap && finalizedMap[employeeId]);

  return {
    employee_id: employeeId,
    name: String(emp.name || ""),
    display_name: String(emp.display_name || ""),
    company_code: String(emp.company_code || ""),
    company_name: String(emp.company_name || ""),
    department: String(emp.department || ""),
    fiscal_start_month: fiscalStartMonth,
    fiscal_year: Number(fiscalYear),
    fiscal_year_end_date: formatDateValue(fiscalYearEndDate),
    previous_remaining_days: previousRemainingDays,
    carry_over_candidate_days: carryOverCandidateDays,
    expired_days: expiredDays,
    new_grant_days: newGrantDays,
    estimated_after_grant_days: estimatedAfterGrantDays,
    opening_balance_days_total: Number(fifoBalance.opening_balance_days_total || 0),
    expiry_unconfirmed_days_total: Number(
      fifoBalance.expiry_unconfirmed_opening_balance_days_total || 0
    ),
    validity_warning: String(fifoBalance.validity_warning || ""),
    is_finalized: isFinalized
  };
}

function buildYearEndCarryOverFinalizedNotes_(candidate, nextFiscalYear) {
  return [
    "年跨ぎ確定",
    "前年度: " + candidate.fiscal_year,
    "次年度: " + nextFiscalYear,
    "前年度残: " + formatGrantDaysForNote_(candidate.previous_remaining_days) + "日",
    "繰越: " + formatGrantDaysForNote_(candidate.carry_over_candidate_days) + "日",
    "消滅見込み: " + formatGrantDaysForNote_(candidate.expired_days) + "日",
    "新規付与: " + formatGrantDaysForNote_(candidate.new_grant_days) + "日"
  ].join(" / ");
}

function getLeaveRolloverCompanyConfig_(companyCode) {
  const code = String(companyCode || "").trim().toUpperCase();
  const configs = {
    MAIN: {
      company_code: "MAIN",
      company_name: "",
      company_display_name: "正社員",
      fiscal_start_month: 4
    },
    PARTNER: {
      company_code: "PARTNER",
      company_name: "（有）友尚建設",
      company_display_name: "友尚建設",
      fiscal_start_month: 6
    }
  };

  if (!configs[code]) {
    throw new Error("年度切替に対応していない company_code です: " + code);
  }

  return Object.assign({}, configs[code]);
}

function getLeaveRolloverFiscalYearDates_(config, fiscalYear) {
  const nextFiscalYear = Number(fiscalYear || 0);
  const fiscalStartMonth = Number(config && config.fiscal_start_month || 0);

  if (!nextFiscalYear) {
    throw new Error("対象年度が不正です");
  }

  if (fiscalStartMonth < 1 || fiscalStartMonth > 12) {
    throw new Error("fiscal_start_month が不正です");
  }

  const nextFiscalYearStartDate = new Date(nextFiscalYear, fiscalStartMonth - 1, 1);
  const previousFiscalYearEndDate = addDaysLocal_(nextFiscalYearStartDate, -1);

  return {
    previous_fiscal_year: nextFiscalYear - 1,
    next_fiscal_year: nextFiscalYear,
    previous_fiscal_year_end_date: formatDateValue(previousFiscalYearEndDate),
    next_fiscal_year_start_date: formatDateValue(nextFiscalYearStartDate)
  };
}

function buildCompanyLeaveYearRolloverCandidate_(emp, fiscalYear, context, finalizedMap, config) {
  const dates = getLeaveRolloverFiscalYearDates_(config, fiscalYear);
  const errors = [];
  const warnings = [];
  let candidate = null;

  if (!emp.hire_date) {
    errors.push("入社日が未入力です");
  }

  try {
    candidate = buildYearEndCarryOverCandidate_(
      emp,
      dates.previous_fiscal_year,
      context,
      finalizedMap
    );

    const companyBasisGrantInfo = getCompanyBasisYearlyGrantInfo_(
      emp.hire_date,
      parseLocalDate(dates.next_fiscal_year_start_date),
      config.fiscal_start_month
    );
    candidate.new_grant_days = companyBasisGrantInfo.grant_days;
    candidate.estimated_after_grant_days =
      Number(candidate.carry_over_candidate_days || 0) +
      Number(companyBasisGrantInfo.grant_days || 0);
    candidate.company_basis_grant_number =
      companyBasisGrantInfo.company_basis_grant_number;
    candidate.company_basis_equivalent_months =
      companyBasisGrantInfo.equivalent_months;
  } catch (e) {
    errors.push("候補計算エラー: " + (e && e.message ? e.message : String(e)));
  }

  const employeeId = String(emp.employee_id || "").trim();
  const hasNextFiscalYearRecord = !!finalizedMap[employeeId];
  if (hasNextFiscalYearRecord) {
    warnings.push(
      dates.next_fiscal_year +
      "年度の yearly レコードがすでにあります。重複作成しません"
    );
  }

  if (candidate && candidate.validity_warning) {
    warnings.push(candidate.validity_warning);
  }

  return {
    employee_id: employeeId,
    name: String(emp.name || ""),
    display_name: String(emp.display_name || ""),
    company_code: String(emp.company_code || ""),
    company_name: String(emp.company_name || ""),
    fiscal_start_month: Number(emp.fiscal_start_month || 0),
    fiscal_year: dates.previous_fiscal_year,
    fiscal_year_end_date: candidate ? candidate.fiscal_year_end_date : "",
    previous_remaining_days: candidate ? Number(candidate.previous_remaining_days || 0) : "",
    carry_over_candidate_days: candidate ? Number(candidate.carry_over_candidate_days || 0) : "",
    expired_days: candidate ? Number(candidate.expired_days || 0) : "",
    new_grant_days: candidate ? Number(candidate.new_grant_days || 0) : "",
    estimated_after_grant_days: candidate ? Number(candidate.estimated_after_grant_days || 0) : "",
    company_basis_grant_number: candidate ? Number(candidate.company_basis_grant_number || 0) : "",
    company_basis_equivalent_months: candidate ? Number(candidate.company_basis_equivalent_months || 0) : "",
    has_next_fiscal_year_record: hasNextFiscalYearRecord,
    has_2026_yearly_record: dates.next_fiscal_year === 2026 && hasNextFiscalYearRecord,
    will_create_grant_record: !hasNextFiscalYearRecord && errors.length === 0 && warnings.length === 0,
    can_execute: !hasNextFiscalYearRecord && errors.length === 0 && warnings.length === 0,
    errors: errors,
    warnings: warnings,
    messages: errors.concat(warnings)
  };
}

function getCompanyLeaveYearRolloverCandidates_(companyCode, fiscalYear) {
  const config = getLeaveRolloverCompanyConfig_(companyCode);
  const dates = getLeaveRolloverFiscalYearDates_(config, fiscalYear);
  const employees = getEmployeesForAdmin();
  const context = createFifoBalanceComparisonContext_(
    parseLocalDate(dates.previous_fiscal_year_end_date)
  );
  const finalizedMap = getYearlyGrantFinalizedMap_(dates.next_fiscal_year);
  const globalErrors = [];
  const globalWarnings = [];

  const relatedEmployees = employees.filter(emp =>
    isCompanyLeaveYearRolloverRelatedEmployee_(emp, config)
  );
  const targetEmployees = employees.filter(emp =>
    isCompanyLeaveYearRolloverTarget_(emp, config)
  );

  relatedEmployees.forEach(emp => {
    const reasons = getCompanyLeaveYearRolloverTargetMismatchReasons_(emp, config);
    if (reasons.length > 0) {
      globalErrors.push(
        String(emp.employee_id || "(社員IDなし)") +
        " は " +
        config.company_display_name +
        " の社員ですが対象条件と一致しません: " +
        reasons.join("、")
      );
    }
  });

  const rows = targetEmployees
    .sort((a, b) => String(a.employee_id || "").localeCompare(String(b.employee_id || "")))
    .map(emp =>
      buildCompanyLeaveYearRolloverCandidate_(
        emp,
        dates.next_fiscal_year,
        context,
        finalizedMap,
        config
      )
    );
  const rowErrorCount = rows.reduce((sum, row) => sum + row.errors.length, 0);
  const rowWarningCount = rows.reduce((sum, row) => sum + row.warnings.length, 0);
  return {
    ok: true,
    dry_run: true,
    data_changed: false,
    company_code: config.company_code,
    company_name: config.company_name || config.company_display_name,
    company_display_name: config.company_display_name,
    fiscal_start_month: config.fiscal_start_month,
    previous_fiscal_year: dates.previous_fiscal_year,
    next_fiscal_year: dates.next_fiscal_year,
    previous_fiscal_year_end_date: dates.previous_fiscal_year_end_date,
    next_fiscal_year_start_date: dates.next_fiscal_year_start_date,
    target_count_basis: "社員マスターから自動判定",
    target_condition_label: buildCompanyLeaveYearRolloverTargetConditionLabel_(config),
    target_employee_count: rows.length,
    error_count: globalErrors.length + rowErrorCount,
    warning_count: globalWarnings.length + rowWarningCount,
    can_execute:
      rows.length > 0 &&
      globalErrors.length === 0 &&
      globalWarnings.length === 0 &&
      rowErrorCount === 0 &&
      rowWarningCount === 0,
    global_errors: globalErrors,
    global_warnings: globalWarnings,
    rows: rows
  };
}

function dryRunCompanyLeaveYearRollover(companyCode, fiscalYear) {
  const result = getCompanyLeaveYearRolloverCandidates_(companyCode, fiscalYear);
  logCompanyLeaveYearRolloverDryRun_(result);
  return result;
}

function dryRunMainLeaveYearRollover2026() {
  return dryRunCompanyLeaveYearRollover("MAIN", 2026);
}

function executeCompanyLeaveYearRollover(companyCode, fiscalYear, options) {
  const config = getLeaveRolloverCompanyConfig_(companyCode);
  const dates = getLeaveRolloverFiscalYearDates_(config, fiscalYear);
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    validateCompanyLeaveYearRolloverConfirmation_(
      config,
      dates.next_fiscal_year,
      options || {}
    );

    Logger.log(
      "=== " +
      config.company_display_name +
      " " +
      dates.next_fiscal_year +
      "年度切替 本処理: 開始 ==="
    );
    Logger.log("本処理前に dry-run 相当の検証を再実行します。");

    const dryRun = getCompanyLeaveYearRolloverCandidates_(
      config.company_code,
      dates.next_fiscal_year
    );
    logCompanyLeaveYearRolloverDryRun_(dryRun);
    validateCompanyLeaveYearRolloverExecution_(dryRun, config);

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
      "notes",
      "created_at",
      "updated_at"
    ]);

    // 付与行の追加後にログ不足で止まらないよう、先にログシートも検証する。
    requireHeaders(getSheet("usage_log"), [
      "log_id",
      "request_id",
      "action_type",
      "operator_id",
      "operator_name",
      "action_date",
      "comment"
    ]);

    dryRun.rows.forEach(row => {
      if (hasYearlyGrantForFiscalYear_(row.employee_id, dates.next_fiscal_year)) {
        throw new Error(
          row.employee_id +
          " は " +
          dates.next_fiscal_year +
          "年度の yearly レコードがすでにあります。本処理を停止しました。"
        );
      }
    });

    const grantIds = getNextGrantIds_(dryRun.rows.length);
    const now = new Date();
    const grantDate = parseLocalDate(dates.next_fiscal_year_start_date);
    const validTo = addDaysLocal_(addYearsLocal_(grantDate, 2), -1);
    const newRows = dryRun.rows.map((row, index) => {
      const rowObj = createEmptyRowObject(headerInfo.headers);

      rowObj.grant_id = grantIds[index];
      rowObj.employee_id = row.employee_id;
      rowObj.grant_date = grantDate;
      rowObj.grant_days = Number(row.new_grant_days || 0);
      rowObj.carry_over_days = Number(row.carry_over_candidate_days || 0);
      rowObj.valid_from = grantDate;
      rowObj.valid_to = validTo;
      rowObj.grant_type = "yearly";
      rowObj.year = dates.next_fiscal_year;
      rowObj.notes = buildYearEndCarryOverFinalizedNotes_(
        row,
        dates.next_fiscal_year
      );
      rowObj.created_at = now;
      rowObj.updated_at = now;

      if ("is_finalized" in headerInfo.map) {
        rowObj.is_finalized = true;
      }

      if ("finalized_at" in headerInfo.map) {
        rowObj.finalized_at = now;
      }

      return objectToRow(rowObj, headerInfo.headers);
    });

    if (newRows.length > 0) {
      sheet
        .getRange(sheet.getLastRow() + 1, 1, newRows.length, headerInfo.headers.length)
        .setValues(newRows);
    }

    dryRun.rows.forEach((row, index) => {
      appendUsageLog({
        request_id: row.employee_id,
        action_type: "year_end_carry_over_finalized",
        operator_id: "admin",
        operator_name: "管理者",
        comment:
          (row.display_name || row.name || row.employee_id) +
          " さんの " +
          config.company_display_name +
          dates.next_fiscal_year +
          "年度切替を確定しました: " +
          "grant_id=" +
          grantIds[index] +
          " / " +
          buildYearEndCarryOverFinalizedNotes_(row, dates.next_fiscal_year)
      });
    });

    clearAppCache();

    const result = {
      ok: true,
      executed: true,
      company_code: config.company_code,
      company_name: config.company_name || config.company_display_name,
      company_display_name: config.company_display_name,
      next_fiscal_year: dates.next_fiscal_year,
      grant_date: formatDateValue(grantDate),
      created_count: newRows.length,
      employee_ids: dryRun.rows.map(row => row.employee_id),
      grant_ids: grantIds
    };

    Logger.log(config.company_display_name + " 年度切替 本処理が完了しました。");
    Logger.log("追加件数: " + result.created_count);
    Logger.log("社員ID: " + result.employee_ids.join(", "));
    Logger.log("grant_id: " + result.grant_ids.join(", "));
    Logger.log("=== 年度切替 本処理: 完了 ===");

    return result;
  } finally {
    lock.releaseLock();
  }
}

function validateCompanyLeaveYearRolloverConfirmation_(config, fiscalYear, options) {
  const opts = options || {};
  const expectedConfirmText = buildCompanyLeaveYearRolloverConfirmText_(
    config,
    fiscalYear
  );

  if (opts.backup_confirmed !== true) {
    throw new Error("年度切替前のバックアップ確認が完了していません。本処理を停止しました。");
  }

  if (String(opts.confirm_text || "").trim() !== expectedConfirmText) {
    throw new Error(
      "確認文字列が一致しません。本処理を停止しました。確認文字列: " +
      expectedConfirmText
    );
  }
}

function buildCompanyLeaveYearRolloverConfirmText_(config, fiscalYear) {
  const code = String(config && config.company_code || "").trim().toUpperCase();
  const year = Number(fiscalYear || 0);
  const companyLabel = code === "PARTNER" ? "友尚建設" : "MAIN";
  return companyLabel + year + "年度切替";
}

function validateCompanyLeaveYearRolloverExecution_(dryRun, config) {
  const rows = dryRun && Array.isArray(dryRun.rows) ? dryRun.rows : [];

  if (!dryRun || dryRun.dry_run !== true) {
    throw new Error("dry-run 相当の検証結果を確認できません。本処理を停止しました。");
  }

  if (
    dryRun.company_code !== config.company_code ||
    Number(dryRun.fiscal_start_month || 0) !== Number(config.fiscal_start_month)
  ) {
    throw new Error("会社条件が一致しません。本処理を停止しました。");
  }

  if (
    !dryRun.can_execute ||
    Number(dryRun.error_count || 0) !== 0 ||
    Number(dryRun.warning_count || 0) !== 0 ||
    rows.length === 0 ||
    !rows.every(row => row.can_execute === true)
  ) {
    throw new Error(
      config.company_display_name +
      " " +
      dryRun.next_fiscal_year +
      "年度切替を停止しました。事前確認の注意・エラーを確認してください。"
    );
  }
}

function isCompanyLeaveYearRolloverRelatedEmployee_(emp, config) {
  if (config.company_code === "PARTNER") {
    return (
      String(emp.company_code || "").trim().toUpperCase() === config.company_code ||
      String(emp.company_name || "").trim() === config.company_name
    );
  }

  const status = String(emp.employment_status || "").trim().toLowerCase();
  const isActive = status === "active" || status === "在職";

  if (emp.leave_management_target !== true || !isActive) {
    return false;
  }

  if (config.company_name) {
    return String(emp.company_name || "").trim() === config.company_name;
  }

  return (
    String(emp.company_code || "").trim().toUpperCase() === config.company_code
  );
}

function isCompanyLeaveYearRolloverTarget_(emp, config) {
  return getCompanyLeaveYearRolloverTargetMismatchReasons_(emp, config).length === 0;
}

function getCompanyLeaveYearRolloverTargetMismatchReasons_(emp, config) {
  const reasons = [];
  const status = String(emp.employment_status || "").trim().toLowerCase();

  if (String(emp.company_code || "").trim().toUpperCase() !== config.company_code) {
    reasons.push("company_code が " + config.company_code + " ではありません");
  }

  if (
    config.company_name &&
    String(emp.company_name || "").trim() !== config.company_name
  ) {
    reasons.push("company_name が " + config.company_name + " ではありません");
  }

  if (Number(emp.fiscal_start_month || 0) !== config.fiscal_start_month) {
    reasons.push("fiscal_start_month が " + config.fiscal_start_month + " ではありません");
  }

  if (emp.leave_management_target !== true) {
    reasons.push("leave_management_target が TRUE ではありません");
  }

  if (status !== "active" && status !== "在職") {
    reasons.push("employment_status が active または 在職 ではありません");
  }

  return reasons;
}

function buildCompanyLeaveYearRolloverTargetConditionLabel_(config) {
  const parts = [];

  parts.push(config.company_display_name || config.company_code);
  parts.push(config.fiscal_start_month + "月開始");
  parts.push("有給管理対象");
  parts.push("在職");

  return parts.join(" / ");
}

function logCompanyLeaveYearRolloverDryRun_(result) {
  Logger.log(
    "=== " +
    result.company_display_name +
    " " +
    result.next_fiscal_year +
    "年度切替 dry-run ==="
  );
  Logger.log("データ変更: なし");
  Logger.log(
    "対象: " +
    result.company_name +
    " / company_code=" +
    result.company_code +
    " / fiscal_start_month=" +
    result.fiscal_start_month
  );
  Logger.log(
    "年度: " +
    result.previous_fiscal_year +
    "年度末 " +
    result.previous_fiscal_year_end_date +
    " → " +
    result.next_fiscal_year +
    "年度開始 " +
    result.next_fiscal_year_start_date
  );
  Logger.log(
    "対象人数: 社員マスターから自動判定 " +
      result.target_employee_count +
      "名"
  );
  Logger.log("対象条件: " + result.target_condition_label);

  result.global_errors.forEach(message => Logger.log("[全体エラー] " + message));
  result.global_warnings.forEach(message => Logger.log("[全体注意] " + message));

  Logger.log([
    "employee_id",
    "name",
    "display_name",
    "company_name",
    "fiscal_start_month",
    result.previous_fiscal_year + "年度の残日数",
    "繰越予定日数",
    result.next_fiscal_year + "年度の新規付与予定日数",
    result.next_fiscal_year + "年度開始後の予定残日数",
    "会社基準日付与回数",
    result.next_fiscal_year + "年度レコードが既にあるか",
    "付与レコードを新規作成するか",
    "本処理可能か",
    "注意・エラー内容"
  ].join("\t"));

  result.rows.forEach(row => {
    Logger.log([
      row.employee_id,
      row.name,
      row.display_name,
      row.company_name,
      row.fiscal_start_month,
      row.previous_remaining_days,
      row.carry_over_candidate_days,
      row.new_grant_days,
      row.estimated_after_grant_days,
      row.company_basis_grant_number,
      row.has_next_fiscal_year_record ? "はい" : "いいえ",
      row.will_create_grant_record ? "はい" : "いいえ",
      row.can_execute ? "はい" : "いいえ",
      row.messages.length > 0 ? row.messages.join(" / ") : "なし"
    ].join("\t"));
  });

  Logger.log("エラー件数: " + result.error_count + " / 注意件数: " + result.warning_count);
  Logger.log("本処理可能か: " + (result.can_execute ? "はい" : "いいえ"));
  Logger.log("=== dry-run 完了 ===");
}

function dryRunPartnerLeaveYearRollover2026() {
  return dryRunCompanyLeaveYearRollover("PARTNER", 2026);
}

function getPartnerLeaveYearRollover2026Config_() {
  const config = getLeaveRolloverCompanyConfig_("PARTNER");
  return Object.assign(
    {},
    config,
    getLeaveRolloverFiscalYearDates_(config, 2026)
  );
}

function buildPartnerLeaveYearRollover2026DryRun_() {
  return getCompanyLeaveYearRolloverCandidates_("PARTNER", 2026);
}

function isPartnerLeaveYearRollover2026Target_(emp, config) {
  const status = String(emp.employment_status || "").trim().toLowerCase();
  const isActive = status === "active" || status === "在職";

  return (
    String(emp.company_code || "").trim().toUpperCase() === config.company_code &&
    String(emp.company_name || "").trim() === config.company_name &&
    Number(emp.fiscal_start_month || 0) === config.fiscal_start_month &&
    emp.leave_management_target === true &&
    isActive
  );
}

function getPartnerLeaveYearRollover2026TargetMismatchReasons_(emp, config) {
  const reasons = [];
  const status = String(emp.employment_status || "").trim().toLowerCase();

  if (String(emp.company_code || "").trim().toUpperCase() !== config.company_code) {
    reasons.push("company_code が PARTNER ではありません");
  }

  if (Number(emp.fiscal_start_month || 0) !== config.fiscal_start_month) {
    reasons.push("fiscal_start_month が 6 ではありません");
  }

  if (emp.leave_management_target !== true) {
    reasons.push("leave_management_target が TRUE ではありません");
  }

  if (status !== "active" && status !== "在職") {
    reasons.push("employment_status が active または 在職 ではありません");
  }

  return reasons;
}

function getCompanyBasisYearlyGrantInfo_(hireDateValue, fiscalYearStartDateValue, fiscalStartMonth) {
  if (!hireDateValue) {
    throw new Error("入社日が未入力です");
  }

  const hireDate = parseLocalDate(hireDateValue);
  const fiscalYearStartDate = parseLocalDate(fiscalYearStartDateValue);
  const startMonth = Number(fiscalStartMonth || 0);

  if (startMonth < 1 || startMonth > 12) {
    throw new Error("fiscal_start_month が不正です");
  }

  if (
    fiscalYearStartDate.getMonth() + 1 !== startMonth ||
    fiscalYearStartDate.getDate() !== 1
  ) {
    throw new Error("年度開始日と fiscal_start_month が一致しません");
  }

  let firstCompanyBasisDate = new Date(
    hireDate.getFullYear(),
    startMonth - 1,
    1
  );

  if (firstCompanyBasisDate < hireDate) {
    firstCompanyBasisDate = new Date(
      hireDate.getFullYear() + 1,
      startMonth - 1,
      1
    );
  }

  if (fiscalYearStartDate < firstCompanyBasisDate) {
    return {
      first_company_basis_date: firstCompanyBasisDate,
      company_basis_grant_number: 0,
      equivalent_months: 0,
      grant_days: 0
    };
  }

  const companyBasisGrantNumber =
    fiscalYearStartDate.getFullYear() -
    firstCompanyBasisDate.getFullYear() +
    1;
  const equivalentMonths = 6 + (companyBasisGrantNumber - 1) * 12;

  return {
    first_company_basis_date: firstCompanyBasisDate,
    company_basis_grant_number: companyBasisGrantNumber,
    equivalent_months: equivalentMonths,
    grant_days: getYearlyGrantDays_(equivalentMonths)
  };
}

function getNextGrantIds_(count) {
  const total = Number(count || 0);
  if (total <= 0) return [];

  const firstId = getNextGrantId_();
  const match = String(firstId || "").match(/^G(\d+)$/);

  if (!match) {
    throw new Error("grant_id の採番形式が不正です: " + firstId);
  }

  const firstNumber = Number(match[1]);
  const width = Math.max(4, match[1].length);
  const result = [];

  for (let index = 0; index < total; index++) {
    result.push("G" + String(firstNumber + index).padStart(width, "0"));
  }

  return result;
}

function logPartnerLeaveYearRollover2026DryRun_(result) {
  Logger.log("=== 友尚建設 2026年度切替 dry-run ===");
  Logger.log("データ変更: なし");
  Logger.log(
    "対象: " +
    result.company_name +
    " / company_code=" +
    result.company_code +
    " / fiscal_start_month=" +
    result.fiscal_start_month
  );
  Logger.log(
    "年度: " +
    result.previous_fiscal_year +
    "年度末 " +
    result.previous_fiscal_year_end_date +
    " → " +
    result.next_fiscal_year +
    "年度開始 " +
    result.next_fiscal_year_start_date
  );
  Logger.log(
    "対象人数: 社員マスターから自動判定 " +
    result.target_employee_count +
    "名"
  );
  Logger.log("対象条件: " + result.target_condition_label);

  result.global_errors.forEach(message => Logger.log("[全体エラー] " + message));
  result.global_warnings.forEach(message => Logger.log("[全体注意] " + message));

  Logger.log([
    "employee_id",
    "name",
    "display_name",
    "company_name",
    "fiscal_start_month",
    "2025年度の残日数",
    "繰越予定日数",
    "2026年度の新規付与予定日数",
    "2026年度開始後の予定残日数",
    "会社基準日付与回数",
    "2026年度レコードが既にあるか",
    "付与レコードを新規作成するか",
    "本処理可能か",
    "注意・エラー内容"
  ].join("\t"));

  result.rows.forEach(row => {
    Logger.log([
      row.employee_id,
      row.name,
      row.display_name,
      row.company_name,
      row.fiscal_start_month,
      row.previous_remaining_days,
      row.carry_over_candidate_days,
      row.new_grant_days,
      row.estimated_after_grant_days,
      row.company_basis_grant_number,
      row.has_2026_yearly_record ? "はい" : "いいえ",
      row.will_create_grant_record ? "はい" : "いいえ",
      row.can_execute ? "はい" : "いいえ",
      row.messages.length > 0 ? row.messages.join(" / ") : "なし"
    ].join("\t"));
  });

  Logger.log("エラー件数: " + result.error_count + " / 注意件数: " + result.warning_count);
  Logger.log("本処理可能か: " + (result.can_execute ? "はい" : "いいえ"));
  Logger.log("=== dry-run 完了 ===");
}
