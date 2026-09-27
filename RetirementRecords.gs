/* =========================
   退職時有給記録（Phase 1）
   - 既に退職済みの社員へ後追いで記録を追加する専用機能。
   - 付与・取得履歴および通常の退職処理は変更しない。
========================= */

const LEAVE_RETIREMENT_RECORDS_SHEET = "leave_retirement_records";
const LEAVE_RETIREMENT_RECORD_HEADERS = [
  "retirement_record_id",
  "employee_id",
  "leave_date",
  "balance_as_of_leave_date",
  "adjustment_days",
  "balance_after_adjustment",
  "adjustment_type",
  "reason",
  "notes",
  "calculation_version",
  "fifo_grant_details_json",
  "fifo_allocations_json",
  "calculated_at",
  "operator_id",
  "operator_name",
  "record_status",
  "revision",
  "supersedes_record_id",
  "created_at",
  "updated_at"
];
const LEAVE_RETIREMENT_RECORD_MINUTE_HEADERS = [
  "remaining_minutes",
  "remaining_full_days",
  "remaining_hours",
  "remaining_remainder_minutes",
  "adjustment_minutes",
  "adjusted_remaining_minutes"
];

function ensureLeaveRetirementRecordsSheet_() {
  const ss = getAppSpreadsheet();
  let sheet = ss.getSheetByName(LEAVE_RETIREMENT_RECORDS_SHEET);

  if (!sheet) {
    sheet = ss.insertSheet(LEAVE_RETIREMENT_RECORDS_SHEET);
    sheet.getRange(1, 1, 1, LEAVE_RETIREMENT_RECORD_HEADERS.length)
      .setValues([LEAVE_RETIREMENT_RECORD_HEADERS]);
    sheet.setFrozenRows(1);
  }

  LEAVE_RETIREMENT_RECORD_MINUTE_HEADERS.forEach(header => ensureSheetColumn_(sheet, header));
  requireHeaders(sheet, LEAVE_RETIREMENT_RECORD_HEADERS.concat(LEAVE_RETIREMENT_RECORD_MINUTE_HEADERS));
  return sheet;
}

function getSpreadsheetEmployeeForRetirementRecord_(employeeId) {
  const targetEmployeeId = String(employeeId || "").trim();
  const sheet = getSheet("employees");
  const headerInfo = requireHeaders(sheet, [
    "employee_id",
    "name",
    "display_name",
    "company_code",
    "employment_status",
    "leave_date"
  ]);
  const data = sheet.getDataRange().getValues();

  const row = data.slice(1).find(values => {
    const rowObj = rowToObject(values, headerInfo.headers);
    return String(rowObj.employee_id || "").trim() === targetEmployeeId;
  });

  return row ? rowToObject(row, headerInfo.headers) : null;
}

function assertRetiredEmployeeForRetirementRecord_(employeeId) {
  const employee = getSpreadsheetEmployeeForRetirementRecord_(employeeId);
  if (!employee) throw new Error("対象社員が見つかりません");

  const status = String(employee.employment_status || "").trim().toLowerCase();
  if (status !== "retired") {
    throw new Error("退職済み社員だけ退職時有給記録を登録できます");
  }
  if (!employee.leave_date) throw new Error("社員マスターに退職日がありません");

  return employee;
}

// FIFO入力の保護判定用。記録シートが未作成なら、確定済み記録は存在しない。
// ensureLeaveRetirementRecordsSheet_ は書込み得るため、ここでは呼び出さない。
function getCompletedRetirementLeaveRecords_() {
  const sheet = getAppSpreadsheet().getSheetByName(LEAVE_RETIREMENT_RECORDS_SHEET);
  if (!sheet) return [];
  const headerInfo = requireHeaders(sheet, ["employee_id", "leave_date", "record_status"]);
  return sheet.getDataRange().getValues().slice(1)
    .map(row => rowToObject(row, headerInfo.headers))
    .filter(rowObj =>
      String(rowObj.employee_id || "").trim() &&
      String(rowObj.record_status || "").trim().toLowerCase() === "completed"
    );
}

function hasCompletedRetirementLeaveRecord_(employeeId) {
  const targetEmployeeId = String(employeeId || "").trim();
  if (!targetEmployeeId) return false;
  return getCompletedRetirementLeaveRecords_().some(rowObj =>
    String(rowObj.employee_id || "").trim() === targetEmployeeId
  );
}

function assertNoCompletedRetirementLeaveRecordForFifoMutation_(employeeId, actionLabel) {
  if (!hasCompletedRetirementLeaveRecord_(employeeId)) return;
  throw new Error("RETIREMENT_RECORD_FINALIZED: 退職時有給記録の確定後は" +
    String(actionLabel || "FIFO入力の変更") + "を実行できません");
}

/*
 * 退職時記録はSpreadsheetを正DBとして算出する。
 * 既存ダッシュボード用のFIFO本体は変更せず、同じFIFO関数へ渡すコンテキストだけを
 * Spreadsheetから構成する（Supabaseの読取フラグには依存しない）。
 */
function createSpreadsheetFifoBalanceContext_(asOfDate) {
  const calendarSheet = getSheet("company_calendar");
  const calendarHeaders = requireHeaders(calendarSheet, ["date", "type"]);
  const calendarData = calendarSheet.getDataRange().getValues();

  const grantSheet = getSheet("paid_leave_grants");
  const grantHeaders = requireHeaders(grantSheet, [
    "grant_id", "employee_id", "grant_date", "grant_days", "carry_over_days",
    "valid_from", "valid_to", "grant_type", "year", "notes"
  ]);
  const grantData = grantSheet.getDataRange().getValues();
  const grantsByEmployee = {};

  grantData.slice(1).forEach(row => {
    const source = rowToObject(row, grantHeaders.headers);
    const employeeId = String(source.employee_id || "").trim();
    if (!employeeId || !source.grant_date) return;

    const grantDate = parseLocalDate(source.grant_date);
    const validFrom = source.valid_from ? parseLocalDate(source.valid_from) : grantDate;
    const validTo = source.valid_to
      ? parseLocalDate(source.valid_to)
      : addDaysLocal_(addYearsLocal_(grantDate, 2), -1);
    const finalized = String(source.is_finalized == null ? "" : source.is_finalized)
      .trim().toUpperCase() !== "FALSE";

    if (!grantsByEmployee[employeeId]) grantsByEmployee[employeeId] = [];
    grantsByEmployee[employeeId].push({
      grant_id: String(source.grant_id || ""), employee_id: employeeId,
      grant_date: grantDate, valid_from_date: validFrom, valid_to_date: validTo,
      grant_type: String(source.grant_type || ""), year: source.year || "",
      grant_days: Number(source.grant_days || 0),
      carry_over_days: Number(source.carry_over_days || 0),
      carry_over_minutes: Number(source.carry_over_minutes || 0),
      total_days: Number(source.grant_days || 0) + Number(source.carry_over_days || 0),
      notes: String(source.notes || ""),
      has_recorded_valid_from: !!source.valid_from,
      has_recorded_valid_to: !!source.valid_to,
      is_finalized: finalized
    });
  });

  const requestSheet = getSheet("leave_requests");
  const requestHeaders = requireHeaders(requestSheet, [
    "request_id", "employee_id", "start_date", "end_date", "days", "half_day", "status"
  ]);
  const requestData = requestSheet.getDataRange().getValues();
  const requestsByEmployee = {};
  requestData.slice(1).forEach(row => {
    const source = rowToObject(row, requestHeaders.headers);
    const employeeId = String(source.employee_id || "").trim();
    if (!employeeId) return;
    if (!requestsByEmployee[employeeId]) requestsByEmployee[employeeId] = [];
    requestsByEmployee[employeeId].push(source);
  });

  const timeSegmentsByRequest = {};
  const segmentSheet = getAppSpreadsheet().getSheetByName(TIME_LEAVE_SEGMENTS_SHEET);
  if (segmentSheet) {
    const segmentHeaders = requireHeaders(segmentSheet, [
      "time_leave_id", "request_id", "leave_date", "requested_minutes"
    ]);
    segmentSheet.getDataRange().getValues().slice(1).forEach(row => {
      const segment = rowToObject(row, segmentHeaders.headers);
      const requestId = String(segment.request_id || "").trim();
      if (!requestId) return;
      if (!timeSegmentsByRequest[requestId]) timeSegmentsByRequest[requestId] = [];
      timeSegmentsByRequest[requestId].push(segment);
    });
  }

  const companyCodeByEmployee = {};
  const employeeSheet = getSheet("employees");
  const employeeHeaders = requireHeaders(employeeSheet, ["employee_id", "company_code"]);
  employeeSheet.getDataRange().getValues().slice(1).forEach(row => {
    const employee = rowToObject(row, employeeHeaders.headers);
    const employeeId = String(employee.employee_id || "").trim();
    if (employeeId) companyCodeByEmployee[employeeId] = String(employee.company_code || "").trim().toUpperCase();
  });

  return {
    as_of_date: asOfDate,
    calendar_map: buildCompanyCalendarMapFromRows_(
      calendarData.slice(1).map(row => rowToObject(row, calendarHeaders.headers))
    ),
    grants_by_employee: grantsByEmployee,
    requests_by_employee: requestsByEmployee,
    time_leave_segments_by_request: timeSegmentsByRequest,
    company_code_by_employee: companyCodeByEmployee
  };
}

/* 退職前の試算。既存の退職済み向けプレビューと異なり、記録シートには触れない。 */
function previewRetirementForAdmin(employeeId, plannedLeaveDate, token) {
  requireAdminSession_(token);
  const targetId = String(employeeId || "").trim();
  if (!targetId) throw new Error("employee_id がありません");
  const dateText = String(plannedLeaveDate || "").trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(dateText) && !/^\d{4}\/\d{2}\/\d{2}$/.test(dateText)) {
    throw new Error("退職日はYYYY-MM-DD形式で入力してください");
  }
  const leaveDate = parseLocalDate(dateText);
  // HTML の業務日を Spreadsheet timezone で再解釈しない。
  const dateKey = dateText.replace(/\//g, "-");
  const employee = getSpreadsheetEmployeeForRetirementRecord_(targetId);
  if (!employee) throw new Error("対象社員が見つかりません");
  const status = String(employee.employment_status || "").trim().toLowerCase();
  if (status !== "active" && status !== "在職") {
    throw new Error("在職中の社員だけ退職前の残高を確認できます");
  }

  return buildRetirementPreviewResponse_(targetId, dateKey, leaveDate, employee);
}

/* 中核更新の途中状態またはC状態だけ、再読込後の再確認を許可する。 */
function previewRetirementRecoveryForAdmin(employeeId, plannedLeaveDate, token) {
  requireAdminSession_(token);
  const targetId = String(employeeId || "").trim();
  const dateKey = String(plannedLeaveDate || "").trim();
  if (!targetId) throw new Error("RETIREMENT_INVALID_EMPLOYEE");
  if (!/^\d{4}-\d{2}-\d{2}$/.test(dateKey)) throw new Error("RETIREMENT_INVALID_DATE");
  let leaveDate;
  try { leaveDate = parseLocalDate(dateKey); } catch (err) { throw new Error("RETIREMENT_INVALID_DATE"); }
  const source = getRetirementFinalizeEmployeeRow_(targetId);
  const records = getRetirementFinalizeRecordRows_(targetId).records;
  const state = classifyRetirementFinalizeState_(source.employee, records, dateKey);
  if (state.kind !== "recovery" && state.kind !== "partial_recovery") {
    throw new Error("RETIREMENT_RECOVERY_REQUIRED");
  }
  return buildRetirementPreviewResponse_(targetId, dateKey, leaveDate, source.employee, true);
}

function buildRetirementPreviewResponse_(targetId, dateKey, leaveDate, employee, verifyCompany) {
  const context = createSpreadsheetFifoBalanceContext_(leaveDate);
  const companyCode = String(context.company_code_by_employee[targetId] || "").trim().toUpperCase();
  if (verifyCompany && companyCode !== String(employee.company_code || "").trim().toUpperCase()) {
    throw new Error("RETIREMENT_STATE_CONFLICT");
  }
  const policy = getCompanyLeavePolicy(companyCode); // 対応する会社制度がない場合は試算しない。
  const fifoBalance = calculateFifoBalanceWithOpeningBalanceFromContext_(targetId, leaveDate, context);
  const isMain = isMainTimeLeaveEmployeeForFifo_(targetId, context);
  const fingerprint = createRetirementPreviewFingerprint_(targetId, dateKey, companyCode, policy, fifoBalance, isMain);

  return {
    ok: true,
    employee: {
      employee_id: targetId,
      name: getDisplayName(employee) || String(employee.name || "").trim(),
      company_code: companyCode
    },
    planned_leave_date: dateKey,
    balance: {
      unit: isMain ? "minutes" : "days",
      remaining_days: Number(fifoBalance.current_remaining_days || 0),
      remaining_minutes: isMain ? Number(fifoBalance.current_remaining_minutes || 0) : null
    },
    grant_details: (fifoBalance.grant_details || []).map(lot => ({
      grant_date: lot.grant_date ? retirementFingerprintDate_(lot.grant_date) : "",
      valid_to: lot.valid_to ? retirementFingerprintDate_(lot.valid_to) : "",
      total_days: Number(lot.total_days || 0),
      used_days: Number(lot.used_days || 0),
      active_remaining_days: Number(lot.active_remaining_days || 0),
      is_expired: lot.is_expired === true
    })),
    fingerprint: fingerprint
  };
}

// プレビューと確定処理で共有する、業務上の FIFO 事実からの fingerprint。
function createRetirementPreviewFingerprint_(employeeId, dateKey, companyCode, policy, balance, isMain) {
  const canonical = buildRetirementPreviewFingerprintData_(employeeId, dateKey, companyCode, policy, balance, isMain);
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, JSON.stringify(canonical));
  return bytes.map(byte => (byte + 256).toString(16).slice(-2)).join("");
}

// 保存用証跡の文字数制限は適用せず、FIFO結果の計算関連フィールドだけを固定順に投影する。
function buildRetirementPreviewFingerprintData_(employeeId, dateKey, companyCode, policy, balance, isMain) {
  const fields = isMain
    ? ["calculation_mode", "scheduled_minutes_per_day", "current_remaining_minutes",
      "total_granted_minutes", "used_minutes", "allocated_used_minutes",
      "unallocated_used_minutes", "expired_minutes"]
    : ["calculation_mode", "current_remaining_days", "total_granted_days", "used_days",
      "allocated_used_days", "unallocated_used_days", "expired_days",
      "opening_balance_days_total", "excluded_non_opening_carry_over_days_total",
      "expiry_unconfirmed_opening_balance_days_total"];
  const grantFields = isMain
    ? ["grant_id", "grant_date", "valid_from", "valid_to", "grant_type", "year",
      "granted_minutes", "carryover_minutes", "total_minutes", "used_minutes",
      "remaining_minutes", "active_remaining_minutes",
      "expired_minutes", "is_expired"]
    : ["grant_id", "source_grant_id", "lot_type", "grant_date", "valid_from", "valid_to",
      "grant_type", "year", "validity_basis", "validity_needs_review", "grant_days",
      "opening_balance_days", "total_days", "used_days", "remaining_days",
      "active_remaining_days", "expired_days", "is_expired"];
  const useFields = isMain
    ? ["request_id", "time_leave_id", "use_date", "leave_kind", "consumed_minutes", "unallocated_minutes"]
    : ["request_id", "use_date", "days", "unallocated_days"];
  const allocationFields = isMain
    ? ["request_id", "time_leave_id", "use_date", "leave_kind", "grant_id",
      "consumed_minutes", "grant_valid_to", "calculation_version"]
    : ["request_id", "use_date", "grant_id", "lot_type", "consumed_days"];
  const carryFields = ["grant_id", "grant_date", "grant_type", "year", "carry_over_days",
    "valid_from", "valid_to", "validity_basis", "validity_needs_review"];
  return {
    employee_id: retirementFingerprintId_(employeeId, "employee_id"),
    planned_leave_date: retirementFingerprintDate_(dateKey),
    company_code: retirementFingerprintId_(companyCode, "company_code"),
    policy_version: retirementFingerprintId_(policy.policyVersion, "policy_version"),
    unit: isMain ? "minutes" : "days",
    balance: retirementFingerprintFields_(balance, fields, isMain),
    grant_details: retirementFingerprintRows_(
      isMain ? retirementFingerprintMainGrantRows_(balance.grant_details, balance.scheduled_minutes_per_day)
        : balance.grant_details, grantFields, isMain),
    used_details: retirementFingerprintRows_(balance.used_details, useFields, isMain),
    allocations: retirementFingerprintRows_(balance.allocations, allocationFields, isMain),
    opening_balance_records: retirementFingerprintRows_(balance.opening_balance_records, isMain ? [] : carryFields, isMain),
    excluded_carry_over_records: retirementFingerprintRows_(balance.excluded_carry_over_records, isMain ? [] : carryFields, isMain)
  };
}

// MAIN FIFO の grantMinutes / carryOverMinutes と同じ意味の分数をロット別に固定する。
function retirementFingerprintMainGrantRows_(rows, minutesPerDay) {
  return (rows || []).map(lot => {
    const grantedMinutes = lot.grant_days * minutesPerDay;
    const carryoverMinutes = getGrantCarryOverMinutes_(lot, minutesPerDay);
    if (!Number.isInteger(grantedMinutes) || !Number.isFinite(grantedMinutes) ||
        !Number.isInteger(carryoverMinutes) || !Number.isFinite(carryoverMinutes) ||
        grantedMinutes + carryoverMinutes !== lot.total_minutes) {
      throw new Error("FIFO付与ロットの分数内訳が不正です");
    }
    return Object.assign({}, lot, {
      granted_minutes: grantedMinutes === 0 ? 0 : grantedMinutes,
      carryover_minutes: carryoverMinutes === 0 ? 0 : carryoverMinutes
    });
  });
}

function retirementFingerprintRows_(rows, fields, isMain) {
  if (rows != null && !Array.isArray(rows)) throw new Error("FIFO明細の形式が不正です");
  return (rows || []).map(row => retirementFingerprintFields_(row, fields, isMain))
    .sort((a, b) => {
      const left = JSON.stringify(a);
      const right = JSON.stringify(b);
      return left < right ? -1 : left > right ? 1 : 0;
    });
}

function retirementFingerprintFields_(source, fields, isMain) {
  if (!source || typeof source !== "object") throw new Error("FIFO計算値の形式が不正です");
  const result = {};
  fields.forEach(key => {
    const value = source[key];
    if (/^(grant_id|source_grant_id|request_id)$/.test(key)) {
      result[key] = retirementFingerprintId_(value, key);
    } else if (key === "time_leave_id") {
      // 日・半休では空、時間年休では必須。
      result[key] = source.leave_kind === "time_hourly"
        ? retirementFingerprintId_(value, key) : retirementFingerprintOptionalId_(value, key);
    } else if (/^(grant_date|valid_from|valid_to|use_date|grant_valid_to)$/.test(key)) {
      result[key] = retirementFingerprintDate_(value);
    } else if (key === "year") {
      result[key] = value == null || value === "" ? null : retirementFingerprintYear_(value);
    } else if (key === "is_expired" || key === "validity_needs_review") {
      if (typeof value !== "boolean") throw new Error("FIFO論理値が不正です: " + key);
      result[key] = value;
    } else if (/(_minutes|_days)(_total)?$/.test(key) || key === "days" ||
        key === "scheduled_minutes_per_day") {
      if (typeof value !== "number" || !Number.isFinite(value) ||
          ((key.endsWith("_minutes") || key === "scheduled_minutes_per_day") && !Number.isInteger(value))) {
        throw new Error("FIFO計算値が不正です: " + key);
      }
      result[key] = value === 0 ? 0 : value;
    } else if (value == null || value === "") {
      if (key === "calculation_mode" || key === "leave_kind" || key === "lot_type") {
        throw new Error("FIFO必須項目がありません: " + key);
      }
      result[key] = null;
    } else if (typeof value === "string") {
      result[key] = value.trim();
    } else {
      throw new Error("FIFO計算値の形式が不正です: " + key);
    }
  });
  return result;
}

function retirementFingerprintId_(value, key) {
  if (typeof value !== "string" && typeof value !== "number") throw new Error("FIFO必須IDがありません: " + key);
  const id = String(value).trim();
  if (!id) throw new Error("FIFO必須IDがありません: " + key);
  return id;
}

function retirementFingerprintOptionalId_(value, key) {
  return value == null || (typeof value === "string" && !value.trim())
    ? null : retirementFingerprintId_(value, key);
}

function retirementFingerprintYear_(value) {
  const year = String(value).trim();
  if (!/^\d{4}$/.test(year)) throw new Error("FIFO年が不正です");
  return year; // Spreadsheet の year 列は数値・文字列の両方があり得る。
}

function retirementFingerprintDate_(value) {
  if (value instanceof Date) {
    if (isNaN(value.getTime())) throw new Error("FIFO日付が不正です");
    // FIFO context の Date は parseLocalDate 済みの script-local 業務日。
    return [value.getFullYear(), String(value.getMonth() + 1).padStart(2, "0"),
      String(value.getDate()).padStart(2, "0")].join("-");
  }
  if (typeof value !== "string") throw new Error("FIFO日付が不正です");
  const normalized = value.trim().replace(/\//g, "-");
  if (!/^\d{4}-\d{2}-\d{2}$/.test(normalized)) throw new Error("FIFO日付が不正です");
  parseLocalDate(normalized); // 存在しない日付を拒否。返り値は timezone 変換に使わない。
  return normalized;
}

function getRetirementLeaveBalancePreview(employeeId, adminSessionToken) {
  requireAdminSession_(adminSessionToken);
  const employee = assertRetiredEmployeeForRetirementRecord_(employeeId);
  const leaveDate = parseLocalDate(employee.leave_date);
  const context = createSpreadsheetFifoBalanceContext_(leaveDate);

  // 退職日当日の承認済み取得は含める。既存FIFOは use_date > asOfDate のみ除外する。
  const fifoBalance = calculateFifoBalanceWithOpeningBalanceFromContext_(
    String(employee.employee_id || "").trim(),
    leaveDate,
    context
  );
  const existingRecord = getCompletedRetirementRecord_(employee.employee_id, leaveDate);

  return {
    ok: true,
    employee: {
      employee_id: String(employee.employee_id || "").trim(),
      employee_name: getDisplayName(employee) || String(employee.name || "").trim(),
      leave_date: formatDateValue(leaveDate)
    },
    balance: {
      current_remaining_days: Number(fifoBalance.current_remaining_days || 0),
      current_remaining_minutes: fifoBalance.current_remaining_minutes == null
        ? null : Number(fifoBalance.current_remaining_minutes),
      grant_details: (fifoBalance.grant_details || []).map(grant => ({
        grant_date: String(grant.grant_date || ""),
        valid_to: String(grant.valid_to || ""),
        total_days: Number(grant.total_days || 0),
        used_days: Number(grant.used_days || 0),
        active_remaining_days: Number(grant.active_remaining_days || 0),
        is_expired: grant.is_expired === true
      }))
    },
    existing_record: existingRecord ? toRetirementRecordView_(existingRecord) : null
  };
}

function getCompletedRetirementRecord_(employeeId, leaveDate) {
  const sheet = ensureLeaveRetirementRecordsSheet_();
  const headerInfo = requireHeaders(sheet, LEAVE_RETIREMENT_RECORD_HEADERS);
  const data = sheet.getDataRange().getValues();
  const targetId = String(employeeId || "").trim();
  const targetDate = toDateKey(leaveDate);

  return data.slice(1)
    .map(row => rowToObject(row, headerInfo.headers))
    .find(row =>
      String(row.employee_id || "").trim() === targetId &&
      row.leave_date &&
      toDateKey(row.leave_date) === targetDate &&
      String(row.record_status || "").trim().toLowerCase() === "completed"
    ) || null;
}

function toRetirementRecordView_(record) {
  return {
    retirement_record_id: String(record.retirement_record_id || ""),
    employee_id: String(record.employee_id || ""),
    leave_date: record.leave_date ? formatDateValue(record.leave_date) : "",
    balance_as_of_leave_date: Number(record.balance_as_of_leave_date || 0),
    adjustment_days: Number(record.adjustment_days || 0),
    balance_after_adjustment: Number(record.balance_after_adjustment || 0),
    remaining_minutes: record.remaining_minutes === "" || record.remaining_minutes == null ? null : Number(record.remaining_minutes),
    remaining_full_days: record.remaining_full_days === "" || record.remaining_full_days == null ? null : Number(record.remaining_full_days),
    remaining_hours: record.remaining_hours === "" || record.remaining_hours == null ? null : Number(record.remaining_hours),
    remaining_remainder_minutes: record.remaining_remainder_minutes === "" || record.remaining_remainder_minutes == null ? null : Number(record.remaining_remainder_minutes),
    adjustment_minutes: record.adjustment_minutes === "" || record.adjustment_minutes == null ? null : Number(record.adjustment_minutes),
    adjusted_remaining_minutes: record.adjusted_remaining_minutes === "" || record.adjusted_remaining_minutes == null ? null : Number(record.adjusted_remaining_minutes),
    adjustment_type: String(record.adjustment_type || ""),
    reason: String(record.reason || ""),
    notes: String(record.notes || ""),
    created_at: record.created_at ? formatDateValue(record.created_at) : ""
  };
}

function resolveRetirementAdjustmentMinutes_(mode, rawAdjustmentMinutes, balanceMinutes, rawAdjustmentDays, scheduledMinutesPerDay) {
  const adjustmentMode = String(mode || "").trim();
  const base = Number(balanceMinutes || 0);
  const perDay = Number(scheduledMinutesPerDay || 420);
  if (!Number.isInteger(base) || base < 0) throw new Error("退職時残高分数が不正です");
  if (!Number.isInteger(perDay) || perDay <= 0) throw new Error("所定勤務時間の分数が不正です");
  let adjustment;
  if (adjustmentMode === "record_only") adjustment = 0;
  else if (adjustmentMode === "zero_out") adjustment = -base;
  else if (adjustmentMode === "manual") {
    const hasExplicitMinutes = rawAdjustmentMinutes !== "" && rawAdjustmentMinutes != null;
    if (hasExplicitMinutes) {
      adjustment = Number(rawAdjustmentMinutes);
    } else {
      // 既存UIは adjustment_days（step=0.5）だけを送る。MAINでもこの入力を
      // 受け付け、監査上は整数分へ正規化して保存する。
      const days = Number(rawAdjustmentDays);
      if (!Number.isFinite(days) || !Number.isInteger(days * 2)) {
        throw new Error("MAINの手動調整日数は0.5日単位で入力してください");
      }
      adjustment = days * perDay;
    }
    if (!Number.isInteger(adjustment)) throw new Error("MAINの手動調整は整数分で入力してください");
  } else throw new Error("残有給の扱いが不正です");
  if (adjustment > 0 || base + adjustment < 0) {
    throw new Error("調整後管理残高を0分未満にはできません");
  }
  return {
    adjustment_minutes: adjustment,
    adjusted_remaining_minutes: base + adjustment,
    adjustment_type: adjustmentMode === "record_only" ? "none" :
      adjustmentMode === "zero_out" ? "retirement_zero_out" : "retirement_manual"
  };
}

function resolveRetirementAdjustment_(mode, rawAdjustmentDays, balance) {
  const adjustmentMode = String(mode || "").trim();
  const baseBalance = Number(balance || 0);
  let adjustmentDays;
  let adjustmentType;

  if (adjustmentMode === "record_only") {
    adjustmentDays = 0;
    adjustmentType = "none";
  } else if (adjustmentMode === "zero_out") {
    adjustmentDays = -baseBalance;
    adjustmentType = "retirement_zero_out";
  } else if (adjustmentMode === "manual") {
    adjustmentDays = Number(rawAdjustmentDays);
    adjustmentType = "retirement_manual";
    if (!isFinite(adjustmentDays)) throw new Error("調整日数は数値で入力してください");
  } else {
    throw new Error("残有給の扱いが不正です");
  }

  if (adjustmentDays > 0) throw new Error("Phase 1では正の調整日数は登録できません");
  const afterBalance = baseBalance + adjustmentDays;
  if (afterBalance < -0.000001) {
    throw new Error("調整後管理残高を0日未満にはできません");
  }

  return {
    adjustment_days: adjustmentDays,
    adjustment_type: adjustmentType,
    balance_after_adjustment: Math.max(0, afterBalance)
  };
}

function serializeRetirementFifoEvidence_(fifoBalance) {
  const compactGrantDetails = (fifoBalance.grant_details || []).map(lot => ({
    grant_id: lot.source_grant_id || lot.grant_id || "",
    lot_type: lot.lot_type || "",
    grant_date: lot.grant_date || "",
    valid_from: lot.valid_from || "",
    valid_to: lot.valid_to || "",
    total_days: Number(lot.total_days || 0),
    used_days: Number(lot.used_days || 0),
    remaining_days: Number(lot.remaining_days || 0),
    active_remaining_days: Number(lot.active_remaining_days || 0),
    expired_days: Number(lot.expired_days || 0),
    is_expired: lot.is_expired === true
  }));
  const compactAllocations = (fifoBalance.allocations || []).map(item => ({
    request_id: item.request_id || "", time_leave_id: item.time_leave_id || "",
    use_date: item.use_date || "", leave_kind: item.leave_kind || "",
    grant_id: item.grant_id || "", consumed_minutes: Number(item.consumed_minutes || 0),
    grant_valid_to: item.grant_valid_to || "", calculation_version: item.calculation_version || "",
    consumed_days: Number(item.consumed_days || 0)
  }));

  // Spreadsheetの1セル上限を避けるため、証跡は必要なフィールドに限定する。
  // 上限を超える場合もJSON文字列を途中で切らず、先頭の明細と件数だけを有効JSONで保存する。
  return {
    grant_details_json: serializeBoundedRetirementEvidence_(compactGrantDetails),
    allocations_json: serializeBoundedRetirementEvidence_(compactAllocations)
  };
}

function serializeBoundedRetirementEvidence_(items) {
  const source = Array.isArray(items) ? items : [];
  const kept = [];
  const maxLength = 45000;

  for (let i = 0; i < source.length; i++) {
    const candidate = kept.concat([source[i]]);
    const text = JSON.stringify({ truncated: false, total_count: source.length, items: candidate });
    if (text.length > maxLength) {
      return JSON.stringify({ truncated: true, total_count: source.length, items: kept });
    }
    kept.push(source[i]);
  }

  return JSON.stringify({ truncated: false, total_count: source.length, items: kept });
}

function createRetirementLeaveRecordForRetiredEmployee(payload, adminSessionToken) {
  const adminUser = requireAdminSession_(adminSessionToken);
  const data = payload || {};
  const employeeId = String(data.employee_id || "").trim();
  if (!employeeId) throw new Error("employee_id がありません");

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const employee = assertRetiredEmployeeForRetirementRecord_(employeeId);
    const leaveDate = parseLocalDate(employee.leave_date);
    const existing = getCompletedRetirementRecord_(employeeId, leaveDate);
    if (existing) throw new Error("この社員・退職日の退職時有給記録はすでに登録済みです");

    const context = createSpreadsheetFifoBalanceContext_(leaveDate);
    const fifoBalance = calculateFifoBalanceWithOpeningBalanceFromContext_(employeeId, leaveDate, context);
    const balance = Number(fifoBalance.current_remaining_days || 0);
    const isMainMinuteFifo = isMainTimeLeaveEmployeeForFifo_(employeeId, context);
    const minuteBalance = Number(fifoBalance.current_remaining_minutes || 0);
    // MAIN は日数への小数換算を調整値として使わない。従来日数列は新規分記録では空欄にし、
    // 分列を監査上の正とする。PARTNER は既存日数方式をそのまま利用する。
    const adjustment = isMainMinuteFifo
      ? { adjustment_days: "", balance_after_adjustment: "", adjustment_type: "" }
      : resolveRetirementAdjustment_(data.adjustment_mode, data.adjustment_days, balance);
    const minuteAdjustment = isMainMinuteFifo
      ? resolveRetirementAdjustmentMinutes_(
        data.adjustment_mode,
        data.adjustment_minutes,
        minuteBalance,
        data.adjustment_days,
        getCompanyLeavePolicy("MAIN").scheduledMinutesPerDay
      )
      : null;
    const reason = String(data.reason || "").trim();
    if (!["retirement_settlement", "company_adjustment", "data_correction", "other"].includes(reason)) {
      throw new Error("調整理由が不正です");
    }

    const evidence = serializeRetirementFifoEvidence_(fifoBalance);
    const now = new Date();
    const sheet = ensureLeaveRetirementRecordsSheet_();
    const headerInfo = requireHeaders(sheet, LEAVE_RETIREMENT_RECORD_HEADERS);
    const rowObj = createEmptyRowObject(headerInfo.headers);
    rowObj.retirement_record_id = Utilities.getUuid();
    rowObj.employee_id = employeeId;
    rowObj.leave_date = leaveDate;
    rowObj.balance_as_of_leave_date = balance;
    rowObj.adjustment_days = adjustment.adjustment_days;
    rowObj.balance_after_adjustment = adjustment.balance_after_adjustment;
    rowObj.adjustment_type = isMainMinuteFifo ? minuteAdjustment.adjustment_type : adjustment.adjustment_type;
    rowObj.reason = reason;
    rowObj.notes = String(data.notes || "").trim();
    rowObj.calculation_version = isMainMinuteFifo ? "fifo_minutes_v1" : "fifo_with_opening_balance_v1";
    rowObj.fifo_grant_details_json = evidence.grant_details_json;
    rowObj.fifo_allocations_json = evidence.allocations_json;
    rowObj.calculated_at = now;
    rowObj.operator_id = adminUser.admin_id;
    rowObj.operator_name = adminUser.admin_name;
    rowObj.record_status = "completed";
    rowObj.revision = 1;
    rowObj.supersedes_record_id = "";
    rowObj.created_at = now;
    rowObj.updated_at = now;
    if (isMainMinuteFifo) {
      rowObj.remaining_minutes = minuteBalance;
      rowObj.remaining_full_days = Number(fifoBalance.remaining_full_days || 0);
      rowObj.remaining_hours = Number(fifoBalance.remaining_hours || 0);
      rowObj.remaining_remainder_minutes = Number(fifoBalance.remaining_remainder_minutes || 0);
      rowObj.adjustment_minutes = minuteAdjustment.adjustment_minutes;
      rowObj.adjusted_remaining_minutes = minuteAdjustment.adjusted_remaining_minutes;
    }
    appendRowFast_(sheet, objectToRow(rowObj, headerInfo.headers));

    appendUsageLog({
      request_id: employeeId,
      action_type: "retirement_record_created",
      operator_id: rowObj.operator_id,
      operator_name: rowObj.operator_name,
      comment: "退職時有給記録を登録しました: 退職日=" + formatDateValue(leaveDate) +
        " / FIFO残高=" + (isMainMinuteFifo ? minuteBalance + "分" : balance + "日") + " / 調整=" +
        (isMainMinuteFifo ? minuteAdjustment.adjustment_minutes + "分" : adjustment.adjustment_days + "日") +
        " / 調整後=" + (isMainMinuteFifo
          ? minuteAdjustment.adjusted_remaining_minutes + "分"
          : adjustment.balance_after_adjustment + "日") + " / reason=" + reason
    });

    return {
      ok: true,
      record: toRetirementRecordView_(rowObj)
    };
  } finally {
    lock.releaseLock();
  }
}

/* 退職前プレビューに対応する確定API。既存の後追い記録APIとは独立させる。 */
function finalizeRetirementWithLeaveRecord(payload, token) {
  const adminUser = requireAdminSession_(token);
  const input = payload || {};
  const employeeId = String(input.employee_id || "").trim();
  const dateKey = String(input.planned_leave_date || "").trim();
  const previewFingerprint = String(input.preview_fingerprint || "").trim();
  if (!employeeId) throw new Error("RETIREMENT_INVALID_EMPLOYEE");
  if (!/^\d{4}-\d{2}-\d{2}$/.test(dateKey)) throw new Error("RETIREMENT_INVALID_DATE");
  let leaveDate;
  try { leaveDate = parseLocalDate(dateKey); } catch (err) { throw new Error("RETIREMENT_INVALID_DATE"); }
  if (!/^[a-f0-9]{64}$/.test(previewFingerprint)) throw new Error("RETIREMENT_PREVIEW_STALE");

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    // ロック後に社員・記録・FIFO入力をすべて再読取りする。ensure* は書込み得るため使用しない。
    const employeeSource = getRetirementFinalizeEmployeeRow_(employeeId);
    const recordSource = getRetirementFinalizeRecordRows_(employeeId);
    const state = classifyRetirementFinalizeState_(employeeSource.employee, recordSource.records, dateKey);
    const context = createSpreadsheetFifoBalanceContext_(leaveDate);
    const companyCode = String(context.company_code_by_employee[employeeId] || "").trim().toUpperCase();
    if (companyCode !== String(employeeSource.employee.company_code || "").trim().toUpperCase()) {
      throw new Error("RETIREMENT_STATE_CONFLICT");
    }
    let policy, balance, isMain, fingerprint;
    try {
      policy = getCompanyLeavePolicy(companyCode);
      balance = calculateFifoBalanceWithOpeningBalanceFromContext_(employeeId, leaveDate, context);
      isMain = isMainTimeLeaveEmployeeForFifo_(employeeId, context);
      fingerprint = createRetirementPreviewFingerprint_(employeeId, dateKey, companyCode, policy, balance, isMain);
    } catch (err) {
      throw new Error("RETIREMENT_FIFO_DATA_CONFLICT: " + err.message);
    }
    if (fingerprint !== previewFingerprint) throw new Error("RETIREMENT_PREVIEW_STALE");

    const record = state.kind === "already_completed"
      ? state.record
      : buildRetirementFinalizeRecord_(recordSource.headers, employeeId, leaveDate, adminUser, balance, isMain);
    if (state.kind === "already_completed") {
      assertRetirementFinalizeRecordOnly_(record, balance, isMain);
      // 再送は完全に読取り専用。付随処理の補完は通常の再送に混ぜない。
      return {
        ok: true, status: "already_completed", employee_id: employeeId, leave_date: dateKey,
        retirement_record_id: String(record.retirement_record_id),
        balance: {
          unit: isMain ? "minutes" : "days",
          remaining_days: Number(balance.current_remaining_days),
          remaining_minutes: isMain ? Number(balance.current_remaining_minutes) : null
        },
        warning: ""
      };
    }

    // 付随処理のヘッダーも先に検証し、中核処理後の失敗を減らす。
    requireHeaders(getSheet("usage_log"), ["log_id", "request_id", "action_type", "operator_id", "operator_name", "action_date", "comment"]);
    requireHeaders(employeeSource.sheet, ["display_employee_id", "name", "name_kana", "display_order"]);

    let completed;
    try {
      if (state.kind === "new" || state.kind === "partial_recovery") {
        const map = employeeSource.headers.map;
        if (state.kind === "new") {
          employeeSource.sheet.getRange(employeeSource.rowNumber, map.employment_status + 1).setValue("retired");
        }
        if (state.kind === "new" || state.stage === "status_only") {
          employeeSource.sheet.getRange(employeeSource.rowNumber, map.leave_date + 1).setValue(leaveDate);
        }
        employeeSource.sheet.getRange(employeeSource.rowNumber, map.leave_management_target + 1).setValue(false);
        employeeSource.sheet.getRange(employeeSource.rowNumber, map.updated_at + 1).setValue(new Date());
        SpreadsheetApp.flush();
      }
      if (state.kind !== "already_completed") {
        appendRowFast_(recordSource.sheet, objectToRow(record, recordSource.headers.headers));
        SpreadsheetApp.flush();
      }
      completed = verifyRetirementFinalizeCore_(employeeId, dateKey, record.retirement_record_id);
    } catch (err) {
      if (err.message === "RETIREMENT_INCOMPLETE" ||
          String(err.message).indexOf("RETIREMENT_INCOMPLETE: ") === 0) throw err;
      throw new Error("RETIREMENT_INCOMPLETE: " + err.message);
    }
    const warning = finishRetirementFinalizeSideEffects_(employeeId, dateKey, completed, adminUser);
    return {
      ok: true,
      status: state.kind === "new" ? "completed" : "recovered",
      employee_id: employeeId,
      leave_date: dateKey,
      retirement_record_id: String(completed.retirement_record_id),
      balance: {
        unit: isMain ? "minutes" : "days",
        remaining_days: Number(balance.current_remaining_days),
        remaining_minutes: isMain ? Number(balance.current_remaining_minutes) : null
      },
      warning: warning
    };
  } finally {
    lock.releaseLock();
  }
}

function getRetirementFinalizeEmployeeRow_(employeeId) {
  const sheet = getSheet("employees");
  const headers = requireHeaders(sheet, ["employee_id", "company_code", "employment_status",
    "leave_date", "leave_management_target", "updated_at"]);
  const matches = sheet.getDataRange().getValues().map((row, index) => ({ row: row, rowNumber: index + 1 }))
    .slice(1).filter(item => String(item.row[headers.map.employee_id] || "").trim() === employeeId);
  if (matches.length !== 1) throw new Error("RETIREMENT_STATE_CONFLICT");
  return { sheet: sheet, headers: headers, row: matches[0].row,
    rowNumber: matches[0].rowNumber, employee: rowToObject(matches[0].row, headers.headers) };
}

function getRetirementFinalizeRecordRows_(employeeId) {
  const sheet = getSheet(LEAVE_RETIREMENT_RECORDS_SHEET);
  const headers = requireHeaders(sheet,
    LEAVE_RETIREMENT_RECORD_HEADERS.concat(LEAVE_RETIREMENT_RECORD_MINUTE_HEADERS));
  const records = sheet.getDataRange().getValues().slice(1)
    .map(row => rowToObject(row, headers.headers))
    .filter(row => String(row.employee_id || "").trim() === employeeId);
  return { sheet: sheet, headers: headers, records: records };
}

function classifyRetirementFinalizeState_(employee, records, dateKey) {
  const status = String(employee.employment_status || "").trim().toLowerCase();
  const employeeDate = employee.leave_date ? toDateKey(employee.leave_date) : "";
  const completed = records.filter(row => String(row.record_status || "").trim().toLowerCase() === "completed");
  if (completed.length > 1) throw new Error("RETIREMENT_RECORD_CONFLICT");
  if (completed.length === 1 && (!completed[0].leave_date || toDateKey(completed[0].leave_date) !== dateKey)) {
    throw new Error("RETIREMENT_RECORD_CONFLICT");
  }
  if (records.some(row => String(row.record_status || "").trim().toLowerCase() !== "completed")) {
    throw new Error("RETIREMENT_RECOVERY_REQUIRED");
  }
  if (status === "active" || status === "在職") {
    if (completed.length || employeeDate) throw new Error("RETIREMENT_STATE_CONFLICT");
    return { kind: "new" };
  }
  if (status !== "retired" || (employeeDate && employeeDate !== dateKey)) {
    throw new Error("RETIREMENT_STATE_CONFLICT");
  }
  if (completed.length) {
    if (employeeDate !== dateKey || employee.leave_management_target !== false) {
      throw new Error("RETIREMENT_STATE_CONFLICT");
    }
    return { kind: "already_completed", record: completed[0] };
  }
  if (records.length !== 0) throw new Error("RETIREMENT_RECOVERY_REQUIRED");
  if (employeeDate === dateKey && employee.leave_management_target === false) return { kind: "recovery" };
  if (employee.leave_management_target === true) {
    if (!employeeDate) return { kind: "partial_recovery", stage: "status_only" };
    if (employeeDate === dateKey) return { kind: "partial_recovery", stage: "date_written" };
  }
  throw new Error("RETIREMENT_STATE_CONFLICT");
}

function buildRetirementFinalizeRecord_(headers, employeeId, leaveDate, adminUser, balance, isMain) {
  const days = Number(balance.current_remaining_days);
  const minutes = isMain ? Number(balance.current_remaining_minutes) : null;
  if (!Number.isFinite(days) || days < 0 || (isMain && (!Number.isInteger(minutes) || minutes < 0))) {
    throw new Error("RETIREMENT_FIFO_DATA_CONFLICT");
  }
  const evidence = serializeRetirementFifoEvidence_(balance);
  const now = new Date();
  const row = createEmptyRowObject(headers.headers);
  row.retirement_record_id = Utilities.getUuid();
  row.employee_id = employeeId;
  row.leave_date = leaveDate;
  row.balance_as_of_leave_date = days;
  row.adjustment_days = isMain ? "" : 0;
  row.balance_after_adjustment = isMain ? "" : days;
  row.adjustment_type = "none";
  row.reason = "retirement_settlement";
  row.notes = "";
  row.calculation_version = isMain ? "fifo_minutes_v1" : "fifo_with_opening_balance_v1";
  row.fifo_grant_details_json = evidence.grant_details_json;
  row.fifo_allocations_json = evidence.allocations_json;
  row.calculated_at = now;
  row.operator_id = adminUser.admin_id;
  row.operator_name = adminUser.admin_name;
  row.record_status = "completed";
  row.revision = 1;
  row.supersedes_record_id = "";
  row.created_at = now;
  row.updated_at = now;
  if (isMain) {
    row.remaining_minutes = minutes;
    row.remaining_full_days = Number(balance.remaining_full_days || 0);
    row.remaining_hours = Number(balance.remaining_hours || 0);
    row.remaining_remainder_minutes = Number(balance.remaining_remainder_minutes || 0);
    row.adjustment_minutes = 0;
    row.adjusted_remaining_minutes = minutes;
  }
  return row;
}

function assertRetirementFinalizeRecordOnly_(record, balance, isMain) {
  const days = Number(balance.current_remaining_days);
  const evidence = serializeRetirementFifoEvidence_(balance);
  const requiredNumber = value => value !== "" && value !== null && value !== undefined &&
    typeof value !== "boolean" && String(value).trim() !== "" && Number.isFinite(Number(value));
  if (!String(record.retirement_record_id || "").trim()) {
    throw new Error("RETIREMENT_RECORD_CONFLICT");
  }
  const hasNoAdjustment = String(record.adjustment_type || "").trim() === "none";
  const sameDays = requiredNumber(record.balance_as_of_leave_date) &&
    Number(record.balance_as_of_leave_date) === days;
  const sameUnit = isMain
    ? requiredNumber(record.remaining_minutes) && requiredNumber(record.adjustment_minutes) &&
      requiredNumber(record.adjusted_remaining_minutes) &&
      Number(record.remaining_minutes) === Number(balance.current_remaining_minutes) &&
      Number(record.adjustment_minutes) === 0 &&
      Number(record.adjusted_remaining_minutes) === Number(balance.current_remaining_minutes) &&
      record.adjustment_days === "" && record.balance_after_adjustment === ""
    : requiredNumber(record.adjustment_days) && requiredNumber(record.balance_after_adjustment) &&
      Number(record.adjustment_days) === 0 && Number(record.balance_after_adjustment) === days;
  if (!hasNoAdjustment || !sameDays || !sameUnit ||
      String(record.fifo_grant_details_json || "") !== evidence.grant_details_json ||
      String(record.fifo_allocations_json || "") !== evidence.allocations_json ||
      String(record.calculation_version || "").trim() !==
        (isMain ? "fifo_minutes_v1" : "fifo_with_opening_balance_v1")) {
    throw new Error("RETIREMENT_RECORD_CONFLICT");
  }
}

function verifyRetirementFinalizeCore_(employeeId, dateKey, recordId) {
  const employee = getRetirementFinalizeEmployeeRow_(employeeId).employee;
  const records = getRetirementFinalizeRecordRows_(employeeId).records
    .filter(row => String(row.record_status || "").trim().toLowerCase() === "completed" &&
      row.leave_date && toDateKey(row.leave_date) === dateKey);
  if (String(employee.employment_status || "").trim().toLowerCase() !== "retired" ||
      !employee.leave_date || toDateKey(employee.leave_date) !== dateKey ||
      employee.leave_management_target !== false || records.length !== 1 ||
      String(records[0].retirement_record_id || "").trim() !== String(recordId || "").trim()) {
    throw new Error("RETIREMENT_INCOMPLETE");
  }
  return records[0];
}

function finishRetirementFinalizeSideEffects_(employeeId, dateKey, record, adminUser) {
  const warnings = [];
  try { maintainEmployeeDisplayOrderOnly_(); } catch (err) { warnings.push("display_order: " + err.message); }
  // maintainEmployeeDisplayOrderOnly_ は内部で cache を消すが、失敗時にも個別に再試行する。
  try { clearAppCache(); } catch (err) { warnings.push("cache: " + err.message); }
  try {
    const logSheet = getSheet("usage_log");
    const headers = requireHeaders(logSheet, ["request_id", "action_type", "comment"]);
    const marker = "retirement_record_id=" + String(record.retirement_record_id);
    const logged = logSheet.getDataRange().getValues().slice(1).some(row => {
      const item = rowToObject(row, headers.headers);
      return String(item.request_id || "").trim() === employeeId &&
        String(item.action_type || "").trim() === "employee_retire_with_leave_record" &&
        String(item.comment || "").indexOf(marker) !== -1;
    });
    if (!logged) appendUsageLog({
      request_id: employeeId,
      action_type: "employee_retire_with_leave_record",
      operator_id: adminUser.admin_id,
      operator_name: adminUser.admin_name,
      comment: "退職・退職時有給記録を確定しました: 退職日=" + dateKey + " / " + marker
    });
  } catch (err) { warnings.push("audit_log: " + err.message); }
  return warnings.join("; ");
}

/* 読み取り・純粋計算のみの手動実行テスト。Spreadsheetデータは変更しない。 */
function testRetirementLeaveRecordValidationNoWrite_() {
  const recordOnly = resolveRetirementAdjustment_("record_only", "", 7);
  const zeroOut = resolveRetirementAdjustment_("zero_out", "", 7);
  const manual = resolveRetirementAdjustment_("manual", -2.5, 7);
  let positiveRejected = false;
  let negativeBalanceRejected = false;
  try { resolveRetirementAdjustment_("manual", 1, 7); } catch (e) { positiveRejected = true; }
  try { resolveRetirementAdjustment_("manual", -7.5, 7); } catch (e) { negativeBalanceRejected = true; }
  const minuteRecordOnly = resolveRetirementAdjustmentMinutes_("record_only", "", 2280);
  const minuteZeroOut = resolveRetirementAdjustmentMinutes_("zero_out", "", 2280);
  const minuteManual = resolveRetirementAdjustmentMinutes_("manual", -180, 2280);
  const minuteLegacyHalfDay = resolveRetirementAdjustmentMinutes_("manual", "", 2280, -0.5, 420);
  const minuteLegacyFullDay = resolveRetirementAdjustmentMinutes_("manual", "", 2280, -1, 420);
  let nonIntegerMinuteRejected = false;
  let negativeMinuteBalanceRejected = false;
  let invalidLegacyDayRejected = false;
  try { resolveRetirementAdjustmentMinutes_("manual", -0.5, 2280); } catch (e) { nonIntegerMinuteRejected = true; }
  try { resolveRetirementAdjustmentMinutes_("manual", -2400, 2280); } catch (e) { negativeMinuteBalanceRejected = true; }
  try { resolveRetirementAdjustmentMinutes_("manual", "", 2280, -0.25, 420); } catch (e) { invalidLegacyDayRejected = true; }

  return {
    ok: recordOnly.balance_after_adjustment === 7 &&
      zeroOut.balance_after_adjustment === 0 &&
      manual.balance_after_adjustment === 4.5 &&
      positiveRejected && negativeBalanceRejected &&
      minuteRecordOnly.adjusted_remaining_minutes === 2280 &&
      minuteZeroOut.adjusted_remaining_minutes === 0 &&
      minuteManual.adjusted_remaining_minutes === 2100 &&
      minuteLegacyHalfDay.adjustment_minutes === -210 &&
      minuteLegacyFullDay.adjustment_minutes === -420 &&
      nonIntegerMinuteRejected && negativeMinuteBalanceRejected && invalidLegacyDayRejected,
    record_only: recordOnly,
    zero_out: zeroOut,
    manual: manual,
    positive_adjustment_rejected: positiveRejected,
    negative_balance_rejected: negativeBalanceRejected,
    minute_record_only: minuteRecordOnly,
    minute_zero_out: minuteZeroOut,
    minute_manual: minuteManual,
    minute_legacy_half_day: minuteLegacyHalfDay,
    minute_legacy_full_day: minuteLegacyFullDay,
    non_integer_minute_rejected: nonIntegerMinuteRejected,
    negative_minute_balance_rejected: negativeMinuteBalanceRejected,
    invalid_legacy_day_rejected: invalidLegacyDayRejected
  };
}
