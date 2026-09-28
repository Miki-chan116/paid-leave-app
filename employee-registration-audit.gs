/*
 * 社員登録前の読取専用監査。
 * このファイルの本番入口は Spreadsheet/Cache/Script Properties の取得だけを行う。
 * 判定本体は渡された配列だけを読み、データを変更しない。
 */

function auditEmployeeRegistrationReadOnly(adminSessionToken) {
  const spreadsheet = SpreadsheetApp.openById(SS_ID);
  assertEmployeeAuditAdminSessionReadOnly_(adminSessionToken, spreadsheet);
  const names = [
    "employees", "leave_requests", "paid_leave_grants",
    "leave_retirement_records", "time_leave_segments", "usage_log"
  ];
  const sheets = {};
  names.forEach(name => {
    const sheet = spreadsheet.getSheetByName(name);
    const values = sheet ? sheet.getDataRange().getValues() : null;
    sheets[name] = values ? { headers: values[0] || [], rows: values.slice(1) } : null;
  });
  return auditEmployeeRegistrationSnapshot_(sheets, {
    dateToKey: value => Utilities.formatDate(
      value, spreadsheet.getSpreadsheetTimeZone(), "yyyy-MM-dd"
    )
  });
}

function getEmployeeRegistrationAuditFlagsReadOnly(adminSessionToken) {
  const spreadsheet = SpreadsheetApp.openById(SS_ID);
  assertEmployeeAuditAdminSessionReadOnly_(adminSessionToken, spreadsheet);
  const properties = PropertiesService.getScriptProperties();
  const value = properties.getProperty("USE_SUPABASE_READS");
  return {
    useSupabaseReads: String(value || "").trim().toLowerCase() === "true",
    configuredValue: value == null ? null : String(value)
  };
}

function getEmployeeRegistrationAuditSheetSizesReadOnly(adminSessionToken) {
  const spreadsheet = SpreadsheetApp.openById(SS_ID);
  assertEmployeeAuditAdminSessionReadOnly_(adminSessionToken, spreadsheet);
  const names = [
    "admin_users", "employees", "leave_requests", "paid_leave_grants",
    "leave_retirement_records", "time_leave_segments", "usage_log"
  ];
  return {
    sheets: names.map(name => {
      const sheet = spreadsheet.getSheetByName(name);
      return sheet ? {
        name, exists: true, lastRow: sheet.getLastRow(), lastColumn: sheet.getLastColumn()
      } : { name, exists: false, lastRow: null, lastColumn: null };
    })
  };
}

function assertEmployeeAuditAdminSessionReadOnly_(token, spreadsheet) {
  const value = String(token || "").trim();
  if (!value) throw new Error("ADMIN_SESSION_REQUIRED");
  const raw = CacheService.getScriptCache().get(getAdminSessionCacheKey_(value));
  if (!raw) throw new Error("ADMIN_SESSION_EXPIRED");
  let session;
  try { session = JSON.parse(raw); } catch (err) { throw new Error("ADMIN_SESSION_INVALID"); }
  const expiresAt = session && session.expires_at ? new Date(session.expires_at).getTime() : NaN;
  if (!session || !session.admin_id || !Number.isFinite(expiresAt) || expiresAt <= Date.now()) {
    throw new Error("ADMIN_SESSION_EXPIRED");
  }
  const sheet = spreadsheet.getSheetByName("admin_users");
  if (!sheet) throw new Error("ADMIN_SESSION_INVALID");
  const values = sheet.getDataRange().getValues();
  const headers = (values[0] || []).map(item => String(item || "").trim());
  const idColumn = headers.indexOf("admin_id");
  const nameColumn = headers.indexOf("admin_name");
  const pinColumn = headers.indexOf("pin");
  const activeColumn = headers.indexOf("is_active");
  if (idColumn < 0 || nameColumn < 0 || pinColumn < 0 || activeColumn < 0)
    throw new Error("ADMIN_SESSION_INVALID");
  const admin = values.slice(1).find(row =>
    String(row[idColumn] || "").trim() === String(session.admin_id).trim()
  );
  if (!admin || String(admin[activeColumn] || "").trim().toUpperCase() !== "TRUE")
    throw new Error("ADMIN_SESSION_INVALID");
}

const EMPLOYEE_REGISTRATION_AUDIT_DETAIL_LIMIT_ = 200;

function auditEmployeeRegistrationSnapshot_(sheets, options) {
  const source = sheets || {};
  const settings = options || {};
  const result = {
    summary: {
      employeeRows: 0,
      counts: { ERROR: 0, WARN: 0, REVIEW: 0 },
      returnedCounts: { ERROR: 0, WARN: 0, REVIEW: 0 },
      truncated: false,
      detailLimitPerSeverity: EMPLOYEE_REGISTRATION_AUDIT_DETAIL_LIMIT_,
      sheetRows: {}
    },
    errors: [], warnings: [], reviews: [],
    idStatistics: {
      strictDecimalMax: { EMP: null, W: null, P: null },
      currentAllocatorInterpretedMax: { EMP: null, W: null, P: null },
      recognizedCounts: { EMP: 0, W: 0, P: 0 },
      // 現在の行からの統計であり、削除済みIDの再利用可否を保証しない。
      scope: "current_rows_only"
    }
  };
  const add = (severity, rule, sheet, row, employeeId, message) => {
    const issue = { severity, rule, sheet, row: row == null ? null : row,
      employeeId: String(employeeId || "").trim().slice(0, 128), message };
    const target = severity === "ERROR" ? result.errors :
      severity === "WARN" ? result.warnings : result.reviews;
    result.summary.counts[severity]++;
    if (target.length < EMPLOYEE_REGISTRATION_AUDIT_DETAIL_LIMIT_) {
      target.push(issue);
      result.summary.returnedCounts[severity]++;
    } else result.summary.truncated = true;
  };
  const text = value => String(value == null ? "" : value).trim();
  const blank = value => value == null || String(value).trim() === "";
  const booleanValue = value => {
    if (value === true || text(value).toUpperCase() === "TRUE") return true;
    if (value === false || text(value).toUpperCase() === "FALSE") return false;
    return null;
  };
  const dateKey = value => {
    if (blank(value)) return { missing: true, valid: false, key: "" };
    if (Object.prototype.toString.call(value) === "[object Date]") {
      if (isNaN(value.getTime())) return { missing: false, valid: false, key: "" };
      const key = settings.dateToKey ? settings.dateToKey(value) :
        [value.getFullYear(), String(value.getMonth() + 1).padStart(2, "0"),
          String(value.getDate()).padStart(2, "0")].join("-");
      return { missing: false, valid: true, key };
    }
    const match = text(value).match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/);
    if (!match) return { missing: false, valid: false, key: "" };
    const year = Number(match[1]), month = Number(match[2]), day = Number(match[3]);
    const parsed = new Date(year, month - 1, day);
    if (parsed.getFullYear() !== year || parsed.getMonth() !== month - 1 ||
        parsed.getDate() !== day) return { missing: false, valid: false, key: "" };
    return { missing: false, valid: true,
      key: [match[1], String(month).padStart(2, "0"), String(day).padStart(2, "0")].join("-") };
  };
  const tables = {};
  const sheetNames = ["employees", "leave_requests", "paid_leave_grants",
    "leave_retirement_records", "time_leave_segments", "usage_log"];
  sheetNames.forEach(name => {
    const input = source[name];
    if (!input) {
      add(name === "employees" || name === "leave_requests" || name === "paid_leave_grants"
        ? "ERROR" : "REVIEW", "missing_sheet", name, null, "", "監査対象シートがありません");
      return;
    }
    const headers = (input.headers || []).map(text);
    const positions = {};
    headers.forEach((header, index) => {
      if (!header) return;
      if (Object.prototype.hasOwnProperty.call(positions, header)) {
        add("ERROR", "duplicate_header", name, 1, "",
          "ヘッダー列 " + (positions[header] + 1) + " と " + (index + 1) + " が重複しています");
      } else positions[header] = index;
    });
    const rows = (input.rows || []).map((cells, index) => ({
      row: index + 2, cells: cells || []
    }));
    tables[name] = { positions, rows };
    result.summary.sheetRows[name] = rows.length;
  });
  const has = (table, column) => !!table &&
    Object.prototype.hasOwnProperty.call(table.positions, column);
  const cell = (table, item, column) => has(table, column) ?
    item.cells[table.positions[column]] : "";
  const requireColumns = (name, columns) => {
    const table = tables[name];
    if (!table) return;
    columns.forEach(column => {
      if (!has(table, column)) add("ERROR", "missing_header", name, 1, "", "必要な列がありません: " + column);
    });
  };
  requireColumns("employees", ["employee_id", "display_employee_id", "name", "name_kana",
    "company_code", "employment_status", "hire_date", "leave_date", "work_days_per_week",
    "leave_management_target", "fiscal_start_month", "initial_grant_check_target"]);
  requireColumns("leave_requests", ["employee_id"]);
  requireColumns("paid_leave_grants", ["employee_id"]);
  if (tables.leave_retirement_records) requireColumns("leave_retirement_records",
    ["employee_id", "leave_date", "record_status"]);
  if (tables.time_leave_segments) requireColumns("time_leave_segments", ["employee_id"]);
  if (tables.usage_log) requireColumns("usage_log", ["request_id", "action_type"]);

  const employees = tables.employees;
  if (!employees) return result;
  result.summary.employeeRows = employees.rows.length;
  const employeeIds = new Map(), displayIds = new Map();
  const canonicalIds = new Map(), canonicalDisplayIds = new Map();
  const serials = { EMP: new Map(), W: new Map(), P: new Map() };
  const allocatorMax = { EMP: 0, W: 0, P: 0 };
  const allocatorMaxCanonical = { EMP: true, W: true, P: true };
  const maxSafe = "9007199254740991";
  const normalizeDigits = digits => digits.replace(/^0+/, "") || "0";
  const larger = (left, right) => !right || left.length > right.length ||
    (left.length === right.length && left > right);
  const inspectId = (rawValue, prefix, item, employeeId, label) => {
    const raw = String(rawValue == null ? "" : rawValue);
    const value = raw.trim();
    const map = label === "employee_id" ? employeeIds : displayIds;
    const canonicalMap = label === "employee_id" ? canonicalIds : canonicalDisplayIds;
    if (!value) {
      add("ERROR", label + "_missing", "employees", item.row, employeeId, label + " が空欄です");
      return;
    }
    if (raw !== value) add("WARN", label + "_outer_whitespace", "employees", item.row,
      employeeId, label + " に前後空白があります");
    if (map.has(value)) add("ERROR", label + "_duplicate", "employees", item.row,
      employeeId, label + " が行 " + map.get(value) + " と重複します");
    else map.set(value, item.row);
    const folded = value.toUpperCase();
    if (canonicalMap.has(folded) && canonicalMap.get(folded).value !== value) {
      add("WARN", label + "_case_collision", "employees", item.row, employeeId,
        label + " が行 " + canonicalMap.get(folded).row + " と大文字小文字だけ異なります");
    } else if (!canonicalMap.has(folded)) canonicalMap.set(folded, { value, row: item.row });
    // 現行 getNextIdNumber_ の Number(id.slice(prefix.length)) を読取専用で再現する。
    // 次番号の加算は行わず、危険な値も文字列で報告する。
    if (value.startsWith(prefix)) {
      const suffix = value.slice(prefix.length);
      const interpreted = Number(suffix);
      const canonical = new RegExp("^" + prefix + "\\d{4,}$").test(value);
      if (suffix && !Number.isNaN(interpreted) && interpreted > allocatorMax[prefix]) {
        allocatorMax[prefix] = interpreted;
        allocatorMaxCanonical[prefix] = canonical;
        result.idStatistics.currentAllocatorInterpretedMax[prefix] = {
          value: String(interpreted), safeInteger: Number.isSafeInteger(interpreted),
          canonicalSource: canonical, row: item.row
        };
      }
    }
    const match = value.match(prefix === "EMP" ? /^EMP(\d+)$/ : /^([WP])(\d+)$/);
    if (!match || (prefix === "EMP" ? false : match[1] !== prefix)) {
      add("WARN", label + "_invalid_format", "employees", item.row, employeeId,
        label + " が正規形式外です");
      const suffix = value.startsWith(prefix) ? value.slice(prefix.length) : "";
      if (suffix && !Number.isNaN(Number(suffix)))
        add("ERROR", label + "_number_interpretation_risk", "employees", item.row,
          employeeId, label + " の異常形式を現行採番の Number() が数値として解釈します");
      return;
    }
    const digits = prefix === "EMP" ? match[1] : match[2];
    if (digits.length < 4) add("WARN", label + "_invalid_format", "employees",
      item.row, employeeId, label + " の番号が4桁未満です");
    const numeric = normalizeDigits(digits);
    if (numeric === "0") add("WARN", label + "_zero_serial", "employees", item.row,
      employeeId, label + " の番号が0です");
    if (larger(numeric, maxSafe)) add("ERROR", label + "_unsafe_number", "employees",
      item.row, employeeId, label + " の番号がJavaScript安全整数を超えます");
    const previous = serials[prefix].get(numeric);
    if (previous && previous.value !== value) add("WARN", label + "_same_serial", "employees",
      item.row, employeeId, label + " が行 " + previous.row + " と同じ番号を表します");
    else if (!previous) serials[prefix].set(numeric, { value, row: item.row });
    if (digits.length >= 4 && larger(numeric, result.idStatistics.strictDecimalMax[prefix]))
      result.idStatistics.strictDecimalMax[prefix] = numeric;
    result.idStatistics.recognizedCounts[prefix]++;
  };

  employees.rows.forEach(item => {
    const employeeId = text(cell(employees, item, "employee_id"));
    if (has(employees, "employee_id"))
      inspectId(cell(employees, item, "employee_id"), "EMP", item, employeeId, "employee_id");
    if (has(employees, "display_employee_id")) {
      const display = text(cell(employees, item, "display_employee_id"));
      inspectId(cell(employees, item, "display_employee_id"),
        display.startsWith("P") ? "P" : "W", item, employeeId, "display_employee_id");
    }
    ["name", "name_kana", "company_code", "employment_status"].forEach(column => {
      if (has(employees, column) && blank(cell(employees, item, column)))
        add("ERROR", column + "_missing", "employees", item.row, employeeId,
          column + " が空欄です");
    });
    const companyRaw = text(cell(employees, item, "company_code"));
    const company = companyRaw.toUpperCase();
    if (has(employees, "company_code") && companyRaw &&
        company !== "MAIN" && company !== "PARTNER")
      add("ERROR", "company_code_invalid", "employees", item.row, employeeId,
        "company_code が許可値ではありません");
    else if (companyRaw && companyRaw !== company)
      add("WARN", "company_code_noncanonical", "employees", item.row, employeeId,
        "company_code の大文字表記を確認してください");
    const statusRaw = text(cell(employees, item, "employment_status"));
    const status = statusRaw.toLowerCase();
    const allowedStatus = ["active", "leave", "retired", "在職", "休職", "退職"];
    if (has(employees, "employment_status") && statusRaw &&
        allowedStatus.indexOf(status) < 0)
      add("ERROR", "employment_status_invalid", "employees", item.row, employeeId,
        "employment_status が許可値ではありません");
    if (["在職", "休職", "退職"].indexOf(status) >= 0)
      add("WARN", "employment_status_legacy", "employees", item.row, employeeId,
        "旧表記の employment_status は後続処理の正規値と一致しません");
    const active = status === "active" || status === "在職";
    const retired = status === "retired" || status === "退職";
    const hire = dateKey(cell(employees, item, "hire_date"));
    const leave = dateKey(cell(employees, item, "leave_date"));
    if (has(employees, "hire_date") && !hire.missing && !hire.valid)
      add("ERROR", "hire_date_invalid", "employees", item.row, employeeId, "hire_date が不正です");
    if (has(employees, "leave_date") && !leave.missing && !leave.valid)
      add("ERROR", "leave_date_invalid", "employees", item.row, employeeId, "leave_date が不正です");
    if (hire.valid && leave.valid && leave.key < hire.key)
      add("ERROR", "leave_before_hire", "employees", item.row, employeeId,
        "leave_date が hire_date より前です");
    if (active && !leave.missing)
      add("ERROR", "active_with_leave_date", "employees", item.row, employeeId,
        "active 社員に leave_date があります");
    if (retired && leave.missing)
      add("WARN", "retired_without_leave_date", "employees", item.row, employeeId,
        "retired 社員に leave_date がありません");
    const leaveTarget = booleanValue(cell(employees, item, "leave_management_target"));
    if (has(employees, "leave_management_target") && !blank(cell(employees, item, "leave_management_target")) &&
        leaveTarget === null)
      add("ERROR", "leave_management_target_invalid", "employees", item.row,
        employeeId, "leave_management_target が真偽値ではありません");
    if (has(employees, "leave_management_target") &&
        typeof cell(employees, item, "leave_management_target") === "string" &&
        cell(employees, item, "leave_management_target") !==
          text(cell(employees, item, "leave_management_target")))
      add("WARN", "leave_management_target_whitespace", "employees", item.row,
        employeeId, "leave_management_target に前後空白があります");
    if (active && leaveTarget === true) {
      if (hire.missing) add("ERROR", "paid_leave_hire_date_missing", "employees",
        item.row, employeeId, "有給対象の在職社員に hire_date がありません");
      // 不正日付は共通の日付検査で報告済み。
      const days = cell(employees, item, "work_days_per_week");
      if (has(employees, "work_days_per_week")) {
        if (blank(days)) add("ERROR", "paid_leave_work_days_missing", "employees",
          item.row, employeeId, "有給対象の在職社員に work_days_per_week がありません");
        else if (!(typeof days === "number" ? Number.isInteger(days) && days >= 1 && days <= 5 :
            /^[1-5]$/.test(String(days).trim())))
          add("ERROR", "paid_leave_work_days_invalid", "employees", item.row,
            employeeId, "work_days_per_week は1～5の整数が必要です");
      }
      const month = cell(employees, item, "fiscal_start_month");
      if (has(employees, "fiscal_start_month")) {
        if (blank(month) || !(typeof month === "number" ?
            Number.isInteger(month) && month >= 1 && month <= 12 :
            /^(?:[1-9]|1[0-2])$/.test(String(month).trim())))
          add("ERROR", "paid_leave_fiscal_month_invalid", "employees", item.row,
            employeeId, "fiscal_start_month が不正です");
        else if ((company === "MAIN" || company === "PARTNER") &&
            Number(month) !== (company === "MAIN" ? 4 : 6))
          add("WARN", "paid_leave_company_month_mismatch", "employees", item.row,
            employeeId, "会社コードと有給年度開始月が一致しません");
      }
    }
    if (has(employees, "initial_grant_check_target")) {
      const initial = cell(employees, item, "initial_grant_check_target");
      if (blank(initial)) add("REVIEW", "initial_grant_target_blank", "employees",
        item.row, employeeId, "初回付与チェック対象が空欄です（旧データの意図を確認）");
      else if (booleanValue(initial) === null)
        add("ERROR", "initial_grant_target_invalid", "employees", item.row,
          employeeId, "初回付与チェック対象が真偽値ではありません");
      else if (typeof initial === "string" && initial !== text(initial))
        add("WARN", "initial_grant_target_whitespace", "employees", item.row,
          employeeId, "初回付与チェック対象に前後空白があります");
    }
  });
  ["EMP", "W", "P"].forEach(prefix => {
    const current = result.idStatistics.currentAllocatorInterpretedMax[prefix];
    if (!current) return;
    if (!current.safeInteger)
      add("ERROR", "current_allocator_unsafe_max", "employees", current.row, "",
        prefix + " 系列の現行採番最大値は安全な整数ではありません");
    if (!allocatorMaxCanonical[prefix])
      add("WARN", "current_allocator_noncanonical_max", "employees", current.row, "",
        prefix + " 系列の正規形式外IDが現行採番の最大値を決めています");
  });

  const referenceSheets = ["leave_requests", "paid_leave_grants",
    "leave_retirement_records", "time_leave_segments"];
  referenceSheets.forEach(name => {
    const table = tables[name];
    if (!has(table, "employee_id")) return;
    table.rows.forEach(item => {
      const id = text(cell(table, item, "employee_id"));
      if (!id) add("ERROR", "reference_employee_id_missing", name, item.row, "",
        "参照先 employee_id が空欄です");
      else if (!employeeIds.has(id)) add("ERROR", "orphan_employee_reference", name,
        item.row, id, "employees にない employee_id を参照しています");
    });
  });
  const logs = tables.usage_log;
  if (has(logs, "request_id") && has(logs, "action_type")) {
    logs.rows.forEach(item => {
      const action = text(cell(logs, item, "action_type"));
      if (!/^(employee_add|employee_update|employee_retire|employee_retire_with_leave_record|six_month_grant|yearly_grant)$/.test(action)) return;
      const id = text(cell(logs, item, "request_id"));
      if (!id) add("REVIEW", "employee_log_target_missing", "usage_log", item.row,
        "", "社員関連ログの対象IDが空欄です");
      else if (!employeeIds.has(id)) add("WARN", "employee_log_orphan_target", "usage_log",
        item.row, id, "社員関連ログが存在しない社員IDを参照しています");
    });
  }
  const records = tables.leave_retirement_records;
  const completed = new Map(), allRetirementRecords = new Map();
  if (has(records, "employee_id") && has(records, "record_status")) {
    records.rows.forEach(item => {
      const id = text(cell(records, item, "employee_id"));
      const recordStatus = text(cell(records, item, "record_status")).toLowerCase();
      if (recordStatus !== "completed")
        add("WARN", "retirement_record_noncompleted", "leave_retirement_records", item.row,
          id, "未完了または異常な状態の退職記録があります");
      if (has(records, "leave_date")) {
        const recordDate = dateKey(cell(records, item, "leave_date"));
        if (recordDate.missing)
          add("ERROR", "retirement_record_leave_date_missing", "leave_retirement_records",
            item.row, id, "退職記録の leave_date が空欄です");
        else if (!recordDate.valid)
          add("ERROR", "retirement_record_leave_date_invalid", "leave_retirement_records",
            item.row, id, "退職記録の leave_date が不正です");
      }
      if (!id) return;
      if (!allRetirementRecords.has(id)) allRetirementRecords.set(id, []);
      allRetirementRecords.get(id).push(item);
      if (recordStatus !== "completed") return;
      if (!completed.has(id)) completed.set(id, []);
      completed.get(id).push(item);
    });
    completed.forEach((items, id) => {
      if (items.length > 1) add("ERROR", "multiple_completed_retirement_records",
        "leave_retirement_records", items[1].row, id, "完了済み退職記録が複数あります");
    });
    employees.rows.forEach(item => {
      const id = text(cell(employees, item, "employee_id"));
      const status = text(cell(employees, item, "employment_status")).toLowerCase();
      if (!id) return;
      const items = completed.get(id) || [];
      if ((status === "retired" || status === "退職") &&
          !allRetirementRecords.has(id))
        add("REVIEW", "retired_without_completed_record", "employees", item.row,
          id, "退職済み社員に完了済み退職記録がありません（過去登録の可能性を確認）");
      if (status !== "retired" && status !== "退職" && items.length > 0)
        add("ERROR", "completed_record_for_nonretired", "employees", item.row,
          id, "退職していない社員に完了済み退職記録があります");
      if (items.length > 0 && status === "retired" &&
          cell(employees, item, "leave_management_target") !== false)
        add("ERROR", "completed_retirement_leave_target_not_false", "employees", item.row,
          id, "完了済み退職記録がある退職社員の leave_management_target が false ではありません");
      if (items.length === 1 && has(records, "leave_date")) {
        const employeeDate = dateKey(cell(employees, item, "leave_date"));
        const recordDate = dateKey(cell(records, items[0], "leave_date"));
        if (recordDate.valid && (!employeeDate.valid || employeeDate.key !== recordDate.key))
          add("ERROR", "retirement_date_mismatch", "employees", item.row, id,
            "社員マスタと退職記録の退職日が一致しません");
      }
    });
  }
  return result;
}
