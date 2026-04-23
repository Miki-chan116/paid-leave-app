const SS_ID = "1o7KbHHsPMiL684YJq_Fpzg6gHjD6HebjeL0BhQImkt4";
const OUTPUT_SS_ID = "1SP7kD0wuxKQAwJ5YrMBGj3HAkMHW2U6Rdwzfhlu_Z_E";

/* =========================
   ステータス定義
========================= */
const STATUS = {
  PENDING: "pending",
  APPROVED: "approved",
  REJECTED: "rejected"
};

/* =========================
   カレンダー種別
========================= */
const CALENDAR_TYPE = {
  WORKDAY: "workday",
  HOLIDAY: "holiday",
  NO_LEAVE: "no_leave"
};

/* =========================
   出力シート名
========================= */
const OUTPUT_SHEET = {
  MONTHLY: "月間有給取得一覧",
  YEARLY: "年間有給取得一覧"
};

/* =========================
   キャッシュキー
========================= */
const CACHE_KEY = {
  EMPLOYEE_MAP: "employee_map_v2",
  COMPANY_CALENDAR: "company_calendar_v2",
  EMPLOYEES_FOR_REQUEST_PREFIX: "employees_for_request_v2_"
};

/* =========================
   実行中メモリキャッシュ
========================= */
let APP_SS_CACHE = null;
let OUTPUT_SS_CACHE = null;
let TZ_CACHE = null;

/* =========================
   スプレッドシート取得
========================= */
function getAppSpreadsheet() {
  if (APP_SS_CACHE) return APP_SS_CACHE;
  APP_SS_CACHE = SpreadsheetApp.openById(SS_ID);
  return APP_SS_CACHE;
}

function getOutputSpreadsheet() {
  if (OUTPUT_SS_CACHE) return OUTPUT_SS_CACHE;
  OUTPUT_SS_CACHE = SpreadsheetApp.openById(OUTPUT_SS_ID);
  return OUTPUT_SS_CACHE;
}

/* =========================
   アプリで使うタイムゾーン
========================= */
function getAppTimeZone() {
  if (TZ_CACHE) return TZ_CACHE;
  TZ_CACHE = getAppSpreadsheet().getSpreadsheetTimeZone();
  return TZ_CACHE;
}

/* =========================
   画面表示
========================= */
function doGet(e) {
  const page = e && e.parameter && e.parameter.page ? e.parameter.page : "index";

  return HtmlService.createTemplateFromFile(page).evaluate()
    .setTitle(page === "admin" ? "Paid Leave Admin" : "Paid Leave App");
}

/* =========================
   include
========================= */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/* =========================
   シート取得
========================= */
function getSheet(name) {
  const sheet = getAppSpreadsheet().getSheetByName(name);

  if (!sheet) {
    throw new Error(name + " シートが見つかりません");
  }

  return sheet;
}

function getOutputSheet(name) {
  const ss = getOutputSpreadsheet();
  let sheet = ss.getSheetByName(name);

  if (!sheet) {
    sheet = ss.insertSheet(name);
  }

  return sheet;
}

/* =========================
   キャッシュクリア
========================= */
function clearAppCache() {
  const cache = CacheService.getScriptCache();
  const currentFiscalYear = getCurrentFiscalYear();

  cache.remove(CACHE_KEY.EMPLOYEE_MAP);
  cache.remove(CACHE_KEY.COMPANY_CALENDAR);
  cache.remove(CACHE_KEY.EMPLOYEES_FOR_REQUEST_PREFIX + currentFiscalYear);
}

/* =========================
   文字正規化
========================= */
function norm(value) {
  return String(value == null ? "" : value)
    .replace(/\s/g, "")
    .toLowerCase();
}

/* =========================
   日付表示
========================= */
function formatDateValue(value) {
  if (!value) return "";

  const date = new Date(value);
  if (isNaN(date.getTime())) return String(value);

  return Utilities.formatDate(date, getAppTimeZone(), "yyyy/MM/dd");
}

/* =========================
   ローカル日付安全変換
========================= */
function parseLocalDate(value) {
  if (value instanceof Date) {
    const ymd = Utilities.formatDate(value, getAppTimeZone(), "yyyy-MM-dd");
    const parts = ymd.split("-");
    const year = Number(parts[0]);
    const month = Number(parts[1]);
    const day = Number(parts[2]);

    const d = new Date(year, month - 1, day);
    if (isNaN(d.getTime())) {
      throw new Error("日付が不正です");
    }
    return d;
  }

  const str = String(value || "").trim();
  if (!str) {
    throw new Error("日付が空です");
  }

  const normalized = str.replace(/\//g, "-");
  const parts = normalized.split("-");

  if (parts.length !== 3) {
    throw new Error("日付形式が不正です: " + str);
  }

  const year = Number(parts[0]);
  const month = Number(parts[1]);
  const day = Number(parts[2]);

  if (!year || !month || !day) {
    throw new Error("日付形式が不正です: " + str);
  }

  const date = new Date(year, month - 1, day);

  if (
    date.getFullYear() !== year ||
    date.getMonth() !== month - 1 ||
    date.getDate() !== day
  ) {
    throw new Error("存在しない日付です: " + str);
  }

  return date;
}

function toDateKey(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, getAppTimeZone(), "yyyy-MM-dd");
  }

  const date = parseLocalDate(value);
  return Utilities.formatDate(date, getAppTimeZone(), "yyyy-MM-dd");
}

/* =========================
   ヘッダー取得
========================= */
function getHeaderMap(sheet) {
  const lastColumn = sheet.getLastColumn();

  if (lastColumn === 0) {
    throw new Error(sheet.getName() + " シートにヘッダーがありません");
  }

  const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  const map = {};

  headers.forEach((header, index) => {
    const key = String(header || "").trim();
    if (key) {
      map[key] = index;
    }
  });

  return { headers, map };
}

/* =========================
   必須ヘッダーチェック
========================= */
function requireHeaders(sheet, requiredHeaders) {
  const headerInfo = getHeaderMap(sheet);
  const missing = requiredHeaders.filter(h => !(h in headerInfo.map));

  if (missing.length > 0) {
    throw new Error(sheet.getName() + " に不足ヘッダーがあります: " + missing.join(", "));
  }

  return headerInfo;
}

/* =========================
   行 → オブジェクト変換
========================= */
function rowToObject(row, headers) {
  const obj = {};

  headers.forEach((header, index) => {
    obj[String(header || "").trim()] = row[index];
  });

  return obj;
}

/* =========================
   空行オブジェクト
========================= */
function createEmptyRowObject(headers) {
  const obj = {};

  headers.forEach(header => {
    obj[String(header || "").trim()] = "";
  });

  return obj;
}

/* =========================
   オブジェクト → 行配列
========================= */
function objectToRow(obj, headers) {
  return headers.map(header => obj[String(header || "").trim()]);
}

/* =========================
   期間計算
========================= */
function getFiscalYearRange(fiscalYear) {
  const start = new Date(fiscalYear, 3, 1);
  const end = new Date(fiscalYear + 1, 2, 31);
  return { start, end };
}

function getFiscalYearFromDate(dateValue) {
  const date = parseLocalDate(dateValue);
  const year = date.getFullYear();
  const month = date.getMonth() + 1;
  return month >= 4 ? year : year - 1;
}

function getClosingMonthRange(targetYear, targetMonth) {
  const start = new Date(targetYear, targetMonth - 2, 26);
  const end = new Date(targetYear, targetMonth - 1, 25);
  return { start, end };
}

function isDateInRange(dateValue, start, end) {
  const date = parseLocalDate(dateValue);
  const target = new Date(date.getFullYear(), date.getMonth(), date.getDate());
  const from = new Date(start.getFullYear(), start.getMonth(), start.getDate());
  const to = new Date(end.getFullYear(), end.getMonth(), end.getDate());

  return target >= from && target <= to;
}

/* =========================
   company_calendar 取得
========================= */
function getCompanyCalendarMap() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get(CACHE_KEY.COMPANY_CALENDAR);
  if (cached) {
    return JSON.parse(cached);
  }

  const sheet = getSheet("company_calendar");
  const headerInfo = requireHeaders(sheet, ["date", "type"]);
  const data = sheet.getDataRange().getValues();
  const map = {};

  if (data.length > 1) {
    data.slice(1).forEach(row => {
      const rowObj = rowToObject(row, headerInfo.headers);
      const rawDate = rowObj.date;
      const rawType = norm(rowObj.type);

      if (!rawDate) return;

      const key = toDateKey(rawDate);
      map[key] = rawType;
    });
  }

  cache.put(CACHE_KEY.COMPANY_CALENDAR, JSON.stringify(map), 300);
  return map;
}

function getCalendarTypeForDate(dateValue, calendarMap) {
  const key = toDateKey(dateValue);

  if (calendarMap && key in calendarMap) {
    return calendarMap[key];
  }

  const date = parseLocalDate(dateValue);
  return date.getDay() === 0 ? CALENDAR_TYPE.HOLIDAY : CALENDAR_TYPE.WORKDAY;
}

function isLeaveAllowedDate(dateValue, calendarMap) {
  const type = getCalendarTypeForDate(dateValue, calendarMap);
  return type === CALENDAR_TYPE.WORKDAY;
}

function getCalendarLabel(type) {
  if (type === CALENDAR_TYPE.WORKDAY) return "営業日";
  if (type === CALENDAR_TYPE.HOLIDAY) return "休日";
  if (type === CALENDAR_TYPE.NO_LEAVE) return "有給NG";
  return type || "";
}

function validateLeaveRequestDates(startDateValue, endDateValue, halfDayValue) {
  const calendarMap = getCompanyCalendarMap();
  const start = parseLocalDate(startDateValue);
  const end = parseLocalDate(endDateValue);
  const normalizedHalfDay = norm(halfDayValue);

  if (normalizedHalfDay) {
    const type = getCalendarTypeForDate(start, calendarMap);
    if (type !== CALENDAR_TYPE.WORKDAY) {
      throw new Error(
        formatDateValue(start) + " は " + getCalendarLabel(type) + " のため有給申請できません"
      );
    }
    return;
  }

  let cursor = new Date(start);

  while (cursor <= end) {
    const type = getCalendarTypeForDate(cursor, calendarMap);
    if (type !== CALENDAR_TYPE.WORKDAY) {
      throw new Error(
        formatDateValue(cursor) + " は " + getCalendarLabel(type) + " のため有給申請できません"
      );
    }
    cursor.setDate(cursor.getDate() + 1);
  }
}

/* =========================
   日別展開
========================= */
function expandLeaveRequestToDailyRows(startDateValue, endDateValue, days, halfDayValue, calendarMap) {
  const result = [];
  const map = calendarMap || getCompanyCalendarMap();

  const start = parseLocalDate(startDateValue);
  const end = parseLocalDate(endDateValue);
  const normalizedHalfDay = norm(halfDayValue);

  if (normalizedHalfDay) {
    if (isLeaveAllowedDate(start, map)) {
      result.push({
        date: new Date(start),
        days: 0.5
      });
    }
    return result;
  }

  let cursor = new Date(start);

  while (cursor <= end) {
    if (isLeaveAllowedDate(cursor, map)) {
      result.push({
        date: new Date(cursor),
        days: 1
      });
    }
    cursor.setDate(cursor.getDate() + 1);
  }

  if (result.length === 0 && Number(days || 0) > 0 && isLeaveAllowedDate(start, map)) {
    result.push({
      date: new Date(start),
      days: Number(days || 0)
    });
  }

  return result;
}

/* =========================
   社員一覧取得
========================= */
function getEmployees() {
  const sheet = getSheet("employees");
  const headerInfo = requireHeaders(sheet, ["employee_id", "name"]);
  const data = sheet.getDataRange().getValues();

  if (data.length <= 1) return [];

  return data.slice(1)
    .map(row => {
      const rowObj = rowToObject(row, headerInfo.headers);
      return {
        id: String(rowObj.employee_id || "").trim(),
        name: String(rowObj.name || rowObj.employee_id || "").trim()
      };
    })
    .filter(emp => emp.id);
}

function getCurrentFiscalYear() {
  return getFiscalYearFromDate(new Date());
}

function getGrantMapByFiscalYear(fiscalYear) {
  const sheet = getSheet("paid_leave_grants");
  const headerInfo = requireHeaders(sheet, [
    "employee_id",
    "grant_days",
    "carry_over_days",
    "year"
  ]);

  const data = sheet.getDataRange().getValues();
  const result = {};

  if (data.length <= 1) return result;

  data.slice(1).forEach(row => {
    const rowObj = rowToObject(row, headerInfo.headers);
    const employeeId = String(rowObj.employee_id || "").trim();
    const rowYear = Number(rowObj.year);

    if (!employeeId) return;
    if (rowYear !== Number(fiscalYear)) return;

    if (!result[employeeId]) {
      result[employeeId] = {
        employee_id: employeeId,
        grant_days: 0,
        carry_over_days: 0
      };
    }

    result[employeeId].grant_days += Number(rowObj.grant_days || 0);
    result[employeeId].carry_over_days += Number(rowObj.carry_over_days || 0);
  });

  return result;
}

function getApprovedUsedDaysByFiscalYear(fiscalYear) {
  const sheet = getSheet("leave_requests");
  const headerInfo = requireHeaders(sheet, [
    "employee_id",
    "start_date",
    "end_date",
    "days",
    "half_day",
    "status"
  ]);

  const data = sheet.getDataRange().getValues();
  const result = {};
  const range = getFiscalYearRange(fiscalYear);
  const calendarMap = getCompanyCalendarMap();

  if (data.length <= 1) return result;

  data.slice(1).forEach(row => {
    const rowObj = rowToObject(row, headerInfo.headers);
    const employeeId = String(rowObj.employee_id || "").trim();
    const status = norm(rowObj.status);

    if (!employeeId) return;
    if (status !== STATUS.APPROVED) return;

    const dailyRows = expandLeaveRequestToDailyRows(
      rowObj.start_date,
      rowObj.end_date,
      rowObj.days,
      rowObj.half_day,
      calendarMap
    );

    dailyRows.forEach(item => {
      if (!isDateInRange(item.date, range.start, range.end)) return;

      if (!result[employeeId]) {
        result[employeeId] = 0;
      }
      result[employeeId] += Number(item.days || 0);
    });
  });

  return result;
}

/* =========================
   対象社員だけの承認済み使用日数集計
========================= */
function getApprovedUsedDaysByFiscalYearForEmployeeIds(fiscalYear, employeeIds) {
  const targetIds = new Set(
    (employeeIds || [])
      .map(id => String(id || "").trim())
      .filter(Boolean)
  );

  if (targetIds.size === 0) return {};

  const sheet = getSheet("leave_requests");
  const headerInfo = requireHeaders(sheet, [
    "employee_id",
    "start_date",
    "end_date",
    "days",
    "half_day",
    "status"
  ]);

  const data = sheet.getDataRange().getValues();
  const result = {};
  const range = getFiscalYearRange(fiscalYear);
  const calendarMap = getCompanyCalendarMap();

  if (data.length <= 1) return result;

  data.slice(1).forEach(row => {
    const rowObj = rowToObject(row, headerInfo.headers);
    const employeeId = String(rowObj.employee_id || "").trim();
    const status = norm(rowObj.status);

    if (!employeeId) return;
    if (!targetIds.has(employeeId)) return;
    if (status !== STATUS.APPROVED) return;

    const dailyRows = expandLeaveRequestToDailyRows(
      rowObj.start_date,
      rowObj.end_date,
      rowObj.days,
      rowObj.half_day,
      calendarMap
    );

    dailyRows.forEach(item => {
      if (!isDateInRange(item.date, range.start, range.end)) return;

      if (!result[employeeId]) {
        result[employeeId] = 0;
      }

      result[employeeId] += Number(item.days || 0);
    });
  });

  return result;
}

function getEmployeeBalanceMapForFiscalYear(fiscalYear) {
  const grantMap = getGrantMapByFiscalYear(fiscalYear);
  const usedMap = getApprovedUsedDaysByFiscalYear(fiscalYear);
  const employees = getEmployees();

  const result = {};

  employees.forEach(emp => {
    const employeeId = emp.id;
    const grantInfo = grantMap[employeeId] || {
      employee_id: employeeId,
      grant_days: 0,
      carry_over_days: 0
    };

    const previousDays = Number(grantInfo.carry_over_days || 0);
    const grantDays = Number(grantInfo.grant_days || 0);
    const usedDays = Number(usedMap[employeeId] || 0);

    const remainingFromPrevious = previousDays - usedDays;

    let nextCarryOverDays = 0;
    let expiredDays = 0;

    if (remainingFromPrevious >= 0) {
      expiredDays = remainingFromPrevious;
      nextCarryOverDays = grantDays;
    } else {
      expiredDays = 0;
      nextCarryOverDays = grantDays + remainingFromPrevious;
    }

    if (nextCarryOverDays < 0) {
      nextCarryOverDays = 0;
    }

    const currentRemainingDays = previousDays + grantDays - usedDays;

    result[employeeId] = {
      current_remaining_days: currentRemainingDays < 0 ? 0 : currentRemainingDays,
      carry_over_days: previousDays,
      grant_days: grantDays,
      used_days: usedDays,
      next_carry_over_days: nextCarryOverDays,
      expired_days: expiredDays
    };
  });

  return result;
}

/* =========================
   対象社員だけの残日数計算
========================= */
function getEmployeeBalanceMapForEmployeeIdsForFiscalYear(fiscalYear, employeeIds) {
  const ids = (employeeIds || [])
    .map(id => String(id || "").trim())
    .filter(Boolean);

  const grantMap = getGrantMapByFiscalYear(fiscalYear);
  const usedMap = getApprovedUsedDaysByFiscalYearForEmployeeIds(fiscalYear, ids);

  const result = {};

  ids.forEach(employeeId => {
    const grantInfo = grantMap[employeeId] || {
      employee_id: employeeId,
      grant_days: 0,
      carry_over_days: 0
    };

    const previousDays = Number(grantInfo.carry_over_days || 0);
    const grantDays = Number(grantInfo.grant_days || 0);
    const usedDays = Number(usedMap[employeeId] || 0);

    const remainingFromPrevious = previousDays - usedDays;

    let nextCarryOverDays = 0;
    let expiredDays = 0;

    if (remainingFromPrevious >= 0) {
      expiredDays = remainingFromPrevious;
      nextCarryOverDays = grantDays;
    } else {
      expiredDays = 0;
      nextCarryOverDays = grantDays + remainingFromPrevious;
    }

    if (nextCarryOverDays < 0) {
      nextCarryOverDays = 0;
    }

    const currentRemainingDays = previousDays + grantDays - usedDays;

    result[employeeId] = {
      current_remaining_days: currentRemainingDays < 0 ? 0 : currentRemainingDays,
      carry_over_days: previousDays,
      grant_days: grantDays,
      used_days: usedDays,
      next_carry_over_days: nextCarryOverDays,
      expired_days: expiredDays
    };
  });

  return result;
}

function calculateYearlyBalanceByEmployee(employeeId, fiscalYear) {
  const grantMap = getGrantMapByFiscalYear(fiscalYear);
  const usedMap = getApprovedUsedDaysByFiscalYear(fiscalYear);

  const grantInfo = grantMap[employeeId] || {
    employee_id: employeeId,
    grant_days: 0,
    carry_over_days: 0
  };

  const previousDays = Number(grantInfo.carry_over_days || 0);
  const grantDays = Number(grantInfo.grant_days || 0);
  const usedDays = Number(usedMap[employeeId] || 0);

  const remainingFromPrevious = previousDays - usedDays;

  let nextCarryOverDays = 0;
  let expiredDays = 0;

  if (remainingFromPrevious >= 0) {
    expiredDays = remainingFromPrevious;
    nextCarryOverDays = grantDays;
  } else {
    expiredDays = 0;
    nextCarryOverDays = grantDays + remainingFromPrevious;
  }

  if (nextCarryOverDays < 0) {
    nextCarryOverDays = 0;
  }

  const currentRemainingDays = previousDays + grantDays - usedDays;

  return {
    employee_id: employeeId,
    carry_over_days: previousDays,
    grant_days: grantDays,
    used_days: usedDays,
    next_carry_over_days: nextCarryOverDays,
    expired_days: expiredDays,
    current_remaining_days: currentRemainingDays < 0 ? 0 : currentRemainingDays
  };
}

/* =========================
   社員名マップ
========================= */
function getEmployeeMap() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get(CACHE_KEY.EMPLOYEE_MAP);
  if (cached) {
    return JSON.parse(cached);
  }

  const employees = getEmployees();
  const map = {};

  employees.forEach(emp => {
    map[emp.id] = emp.name;
  });

  cache.put(CACHE_KEY.EMPLOYEE_MAP, JSON.stringify(map), 300);
  return map;
}

/* =========================
   有給日数計算
========================= */
function calculateLeaveDays(startDate, endDate) {
  const calendarMap = getCompanyCalendarMap();
  const start = parseLocalDate(startDate);
  const end = parseLocalDate(endDate);

  let count = 0;

  while (start <= end) {
    if (isLeaveAllowedDate(start, calendarMap)) {
      count++;
    }
    start.setDate(start.getDate() + 1);
  }

  return count;
}

/* =========================
   使用ログ追加
========================= */
function appendUsageLog(logData) {
  const sheet = getSheet("usage_log");
  const headerInfo = requireHeaders(sheet, [
    "log_id",
    "request_id",
    "action_type",
    "operator_id",
    "operator_name",
    "action_date",
    "comment"
  ]);

  const rowObj = createEmptyRowObject(headerInfo.headers);
  rowObj.log_id = Utilities.getUuid();
  rowObj.request_id = logData.request_id || "";
  rowObj.action_type = logData.action_type || "";
  rowObj.operator_id = logData.operator_id || "";
  rowObj.operator_name = logData.operator_name || "";
  rowObj.action_date = new Date();
  rowObj.comment = logData.comment || "";

  sheet.appendRow(objectToRow(rowObj, headerInfo.headers));
}

/* =========================
   申請登録
========================= */
function submitLeaveRequest(data) {
  if (!data || typeof data !== "object") {
    throw new Error("submitLeaveRequest は画面からデータを受け取って実行してください");
  }

  const sheet = getSheet("leave_requests");
  const headerInfo = requireHeaders(sheet, [
    "request_id",
    "employee_id",
    "request_date",
    "start_date",
    "end_date",
    "days",
    "type",
    "half_day",
    "reason",
    "reason_detail",
    "status",
    "approver_id",
    "approver_name",
    "approved_at",
    "rejected_reason",
    "year",
    "created_at",
    "updated_at"
  ]);

  if (!data.employee_id) {
    throw new Error("employee_id がありません");
  }

  if (!data.start_date || !data.end_date) {
    throw new Error("start_date または end_date がありません");
  }

  const start = parseLocalDate(data.start_date);
  const end = parseLocalDate(data.end_date);

  validateLeaveRequestDates(start, end, data.half_type || (data.half_day ? "half" : ""));

  const isHalf = data.half_day === true;
  const days = isHalf ? 0.5 : calculateLeaveDays(start, end);
  const now = new Date();

  const rowObj = createEmptyRowObject(headerInfo.headers);
  rowObj.request_id = Utilities.getUuid();
  rowObj.employee_id = data.employee_id || "";
  rowObj.request_date = now;
  rowObj.start_date = start;
  rowObj.end_date = end;
  rowObj.days = days;
  rowObj.type = data.type || "paid_leave";
  rowObj.half_day = isHalf ? (data.half_type || "") : "";
  rowObj.reason = data.reason || "";
  rowObj.reason_detail = data.reason_detail || "";
  rowObj.status = STATUS.PENDING;
  rowObj.approver_id = "";
  rowObj.approver_name = "";
  rowObj.approved_at = "";
  rowObj.rejected_reason = "";
  rowObj.year = getFiscalYearFromDate(start);
  rowObj.created_at = now;
  rowObj.updated_at = now;

  sheet.appendRow(objectToRow(rowObj, headerInfo.headers));

  appendUsageLog({
    request_id: rowObj.request_id,
    action_type: "submit",
    operator_id: String(data.employee_id || ""),
    operator_name: "申請者",
    comment: "Leave request submitted"
  });

  clearAppCache();

  return {
    ok: true,
    request_id: rowObj.request_id
  };
}

/* =========================
   管理画面用：申請一覧取得
========================= */
function getRequestsByStatus(status) {
  const sheet = getSheet("leave_requests");
  const headerInfo = requireHeaders(sheet, [
    "request_id",
    "employee_id",
    "start_date",
    "end_date",
    "days",
    "half_day",
    "reason",
    "status"
  ]);

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const employeeMap = getEmployeeMap();
  const target = norm(status);

  const fiscalYears = [...new Set(
    data.slice(1)
      .map(row => {
        const rowObj = rowToObject(row, headerInfo.headers);
        if (!rowObj.start_date) return null;
        return getFiscalYearFromDate(rowObj.start_date);
      })
      .filter(v => v != null)
  )];

  const balanceMapByYear = {};
  fiscalYears.forEach(year => {
    balanceMapByYear[year] = getEmployeeBalanceMapForFiscalYear(year);
  });

  const result = data.slice(1)
    .map(row => {
      const rowObj = rowToObject(row, headerInfo.headers);
      const rowStatus = norm(rowObj.status);
      const employeeId = String(rowObj.employee_id || "").trim();

      if (!rowObj.start_date) return null;

      const fiscalYear = getFiscalYearFromDate(rowObj.start_date);
      const balanceMap = balanceMapByYear[fiscalYear] || {};
      const balance = balanceMap[employeeId] || {
        current_remaining_days: 0,
        grant_days: 0,
        carry_over_days: 0,
        used_days: 0
      };

      return {
        request_id: String(rowObj.request_id || ""),
        employee_id: employeeId,
        employee_name: String(employeeMap[employeeId] || employeeId || "Unknown"),
        start_date: formatDateValue(rowObj.start_date),
        end_date: formatDateValue(rowObj.end_date),
        date_label:
          formatDateValue(rowObj.start_date) +
          (
            formatDateValue(rowObj.start_date) !== formatDateValue(rowObj.end_date)
              ? " 〜 " + formatDateValue(rowObj.end_date)
              : ""
          ),
        days: rowObj.days || 0,
        half_day: String(rowObj.half_day || ""),
        reason: String(rowObj.reason || ""),
        status: rowStatus,
        current_remaining_days: balance.current_remaining_days,
        grant_days: balance.grant_days,
        carry_over_days: balance.carry_over_days,
        used_days: balance.used_days
      };
    })
    .filter(item => item && item.status === target);

  return result;
}

/* =========================
   承認
========================= */
function approveRequest(requestId) {
  if (!requestId) {
    throw new Error("requestId がありません");
  }

  const sheet = getSheet("leave_requests");
  const headerInfo = requireHeaders(sheet, [
    "request_id",
    "status",
    "approver_id",
    "approver_name",
    "approved_at",
    "updated_at"
  ]);

  const data = sheet.getDataRange().getValues();

  if (data.length <= 1) {
    throw new Error("申請データがありません");
  }

  const rowIndex = data.findIndex((row, index) => {
    if (index === 0) return false;
    const rowObj = rowToObject(row, headerInfo.headers);
    return String(rowObj.request_id) === String(requestId);
  });

  if (rowIndex === -1) {
    throw new Error("対象の申請が見つかりません");
  }

  const sheetRow = rowIndex + 1;
  const now = new Date();

  sheet.getRange(sheetRow, headerInfo.map.status + 1).setValue(STATUS.APPROVED);
  sheet.getRange(sheetRow, headerInfo.map.approver_id + 1).setValue("A001");
  sheet.getRange(sheetRow, headerInfo.map.approver_name + 1).setValue("管理者");
  sheet.getRange(sheetRow, headerInfo.map.approved_at + 1).setValue(now);
  sheet.getRange(sheetRow, headerInfo.map.updated_at + 1).setValue(now);

  appendUsageLog({
    request_id: requestId,
    action_type: "approve",
    operator_id: "A001",
    operator_name: "管理者",
    comment: "Approved"
  });

  clearAppCache();
  return { ok: true };
}

/* =========================
   否認
========================= */
function rejectRequest(requestId, reason) {
  if (!requestId) {
    throw new Error("requestId がありません");
  }

  const sheet = getSheet("leave_requests");
  const headerInfo = requireHeaders(sheet, [
    "request_id",
    "status",
    "rejected_reason",
    "updated_at"
  ]);

  const data = sheet.getDataRange().getValues();

  if (data.length <= 1) {
    throw new Error("申請データがありません");
  }

  const rowIndex = data.findIndex((row, index) => {
    if (index === 0) return false;
    const rowObj = rowToObject(row, headerInfo.headers);
    return String(rowObj.request_id) === String(requestId);
  });

  if (rowIndex === -1) {
    throw new Error("対象の申請が見つかりません");
  }

  const sheetRow = rowIndex + 1;
  const now = new Date();

  sheet.getRange(sheetRow, headerInfo.map.status + 1).setValue(STATUS.REJECTED);
  sheet.getRange(sheetRow, headerInfo.map.rejected_reason + 1).setValue(reason || "");
  sheet.getRange(sheetRow, headerInfo.map.updated_at + 1).setValue(now);

  appendUsageLog({
    request_id: requestId,
    action_type: "reject",
    operator_id: "A001",
    operator_name: "管理者",
    comment: reason || ""
  });

  clearAppCache();
  return { ok: true };
}

/* =========================
   ログ取得
========================= */
function getUsageLogs() {
  const sheet = getSheet("usage_log");
  const headerInfo = requireHeaders(sheet, [
    "log_id",
    "request_id",
    "action_type",
    "operator_id",
    "operator_name",
    "action_date",
    "comment"
  ]);

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  return data.slice(1)
    .map(row => {
      const rowObj = rowToObject(row, headerInfo.headers);
      return {
        log_id: rowObj.log_id,
        request_id: rowObj.request_id,
        type: rowObj.action_type,
        user_id: rowObj.operator_id,
        user_name: rowObj.operator_name,
        date: formatDateValue(rowObj.action_date),
        comment: rowObj.comment
      };
    })
    .sort((a, b) => new Date(b.date) - new Date(a.date));
}

/* =========================
   月間取得一覧出力
========================= */
function exportMonthlyPaidLeaveReport(targetYear, targetMonth) {
  if (!targetYear || !targetMonth) {
    const today = new Date();
    targetYear = today.getFullYear();
    targetMonth = today.getMonth() + 1;
  }

  const range = getClosingMonthRange(Number(targetYear), Number(targetMonth));
  const leaveSheet = getSheet("leave_requests");
  const leaveHeaderInfo = requireHeaders(leaveSheet, [
    "employee_id",
    "start_date",
    "end_date",
    "days",
    "half_day",
    "status"
  ]);

  const leaveData = leaveSheet.getDataRange().getValues();
  const employeeMap = getEmployeeMap();
  const calendarMap = getCompanyCalendarMap();

  const detailRows = [];
  const totalMap = {};

  if (leaveData.length > 1) {
    leaveData.slice(1).forEach(row => {
      const rowObj = rowToObject(row, leaveHeaderInfo.headers);
      const employeeId = String(rowObj.employee_id || "").trim();
      const status = norm(rowObj.status);

      if (!employeeId) return;
      if (status !== STATUS.APPROVED) return;

      const employeeName = employeeMap[employeeId] || employeeId;
      const dailyRows = expandLeaveRequestToDailyRows(
        rowObj.start_date,
        rowObj.end_date,
        rowObj.days,
        rowObj.half_day,
        calendarMap
      );

      dailyRows.forEach(item => {
        if (!isDateInRange(item.date, range.start, range.end)) return;

        detailRows.push([
          employeeId,
          employeeName,
          formatDateValue(item.date),
          Number(item.days || 0)
        ]);

        if (!totalMap[employeeId]) {
          totalMap[employeeId] = {
            employee_id: employeeId,
            name: employeeName,
            total_days: 0
          };
        }

        totalMap[employeeId].total_days += Number(item.days || 0);
      });
    });
  }

  detailRows.sort((a, b) => {
    if (a[0] !== b[0]) return a[0] > b[0] ? 1 : -1;
    return a[2] > b[2] ? 1 : -1;
  });

  const totalRows = Object.values(totalMap)
    .sort((a, b) => a.employee_id > b.employee_id ? 1 : -1)
    .map(item => [item.employee_id, item.name, item.total_days]);

  const outputSheet = getOutputSheet(OUTPUT_SHEET.MONTHLY);
  outputSheet.clearContents();

  const values = [];
  values.push(["月間有給取得一覧"]);
  values.push(["対象期間：" + formatDateValue(range.start) + " ～ " + formatDateValue(range.end)]);
  values.push([]);
  values.push(["社員ID", "氏名", "取得日", "取得日数"]);

  if (detailRows.length > 0) {
    detailRows.forEach(row => values.push(row));
  } else {
    values.push(["該当データなし", "", "", ""]);
  }

  values.push([]);
  values.push(["月間合計"]);
  values.push(["社員ID", "氏名", "月間合計取得日数"]);

  if (totalRows.length > 0) {
    totalRows.forEach(row => values.push(row));
  } else {
    values.push(["該当データなし", "", ""]);
  }

  const maxLen = Math.max(...values.map(r => r.length || 1));
  const normalizedValues = values.map(row => {
    const newRow = row.slice();
    while (newRow.length < maxLen) newRow.push("");
    return newRow;
  });

  outputSheet.getRange(1, 1, normalizedValues.length, maxLen).setValues(normalizedValues);

  return {
    ok: true,
    period_start: formatDateValue(range.start),
    period_end: formatDateValue(range.end),
    detail_count: detailRows.length,
    total_count: totalRows.length
  };
}

/* =========================
   年間取得一覧出力
========================= */
function exportYearlyPaidLeaveReport(fiscalYear) {
  if (!fiscalYear) {
    fiscalYear = getFiscalYearFromDate(new Date());
  }

  const yearRange = getFiscalYearRange(Number(fiscalYear));
  const employees = getEmployees();

  const reportRows = employees
    .map(emp => {
      const balance = calculateYearlyBalanceByEmployee(emp.id, Number(fiscalYear));
      return [
        emp.id,
        emp.name,
        balance.carry_over_days,
        balance.grant_days,
        balance.used_days,
        balance.next_carry_over_days,
        balance.expired_days
      ];
    })
    .sort((a, b) => a[0] > b[0] ? 1 : -1);

  const outputSheet = getOutputSheet(OUTPUT_SHEET.YEARLY);
  outputSheet.clearContents();

  const values = [];
  values.push(["年間有給取得一覧"]);
  values.push(["対象年度：" + formatDateValue(yearRange.start) + " ～ " + formatDateValue(yearRange.end)]);
  values.push([]);
  values.push([
    "社員ID",
    "氏名",
    "前年度残日数",
    "今年度付与日数",
    "今年度取得済み日数",
    "来年度繰越日数",
    "消滅日数"
  ]);

  if (reportRows.length > 0) {
    reportRows.forEach(row => values.push(row));
  } else {
    values.push(["該当データなし", "", "", "", "", "", ""]);
  }

  const normalizedValues = values.map(row => {
    const newRow = row.slice();
    while (newRow.length < 7) newRow.push("");
    return newRow;
  });

  outputSheet.getRange(1, 1, normalizedValues.length, 7).setValues(normalizedValues);

  return {
    ok: true,
    fiscal_year: Number(fiscalYear),
    period_start: formatDateValue(yearRange.start),
    period_end: formatDateValue(yearRange.end),
    row_count: reportRows.length
  };
}

/* =========================
   申請画面用社員一覧
   速度改善のためキャッシュ対応
========================= */
function getEmployeesForRequest() {
  const fiscalYear = getCurrentFiscalYear();
  const cache = CacheService.getScriptCache();
  const cacheKey = CACHE_KEY.EMPLOYEES_FOR_REQUEST_PREFIX + fiscalYear;
  const cached = cache.get(cacheKey);

  if (cached) {
    return JSON.parse(cached);
  }

  const sheet = getSheet("employees");
  const headerInfo = requireHeaders(sheet, ["employee_id", "name", "name_kana"]);
  const data = sheet.getDataRange().getValues();

  if (data.length <= 1) return [];

  const employeeRows = data.slice(1)
    .map(row => rowToObject(row, headerInfo.headers))
    .filter(rowObj => {
      const employeeId = String(rowObj.employee_id || "").trim();
      const name = String(rowObj.name || "").trim();
      return employeeId && name;
    });

  const employeeIds = employeeRows.map(rowObj => String(rowObj.employee_id || "").trim());
  const balanceMap = getEmployeeBalanceMapForEmployeeIdsForFiscalYear(fiscalYear, employeeIds);

  const result = employeeRows.map(rowObj => {
    const employeeId = String(rowObj.employee_id || "").trim();
    const balance = balanceMap[employeeId] || {
      current_remaining_days: 0,
      carry_over_days: 0,
      grant_days: 0,
      used_days: 0
    };

    const usedDays = Number(balance.used_days || 0);
    const fiveDayUsed = Math.min(usedDays, 5);
    const fiveDayRemaining = Math.max(0, 5 - usedDays);

    return {
      employee_id: employeeId,
      name: String(rowObj.name || "").trim(),
      name_kana: String(rowObj.name_kana || "").trim(),
      current_remaining_days: Number(balance.current_remaining_days || 0),
      carry_over_days: Number(balance.carry_over_days || 0),
      grant_days: Number(balance.grant_days || 0),
      used_days: usedDays,
      five_day_used: fiveDayUsed,
      five_day_remaining: fiveDayRemaining,
      five_day_completed: fiveDayRemaining === 0
    };
  });

  cache.put(cacheKey, JSON.stringify(result), 300);
  return result;
}

/* =========================
   フロント用返却
========================= */
function getCalendarRules() {
  return getCompanyCalendarMap();
}

function validateRequestDatesOnly(startDate, endDate, halfDay, halfType) {
  validateLeaveRequestDates(
    startDate,
    endDate,
    halfType || (halfDay ? "half" : "")
  );

  return { ok: true };
}