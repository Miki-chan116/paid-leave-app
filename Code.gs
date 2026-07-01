const SS_ID = "1o7KbHHsPMiL684YJq_Fpzg6gHjD6HebjeL0BhQImkt4";
const OUTPUT_SS_ID = "1SP7kD0wuxKQAwJ5YrMBGj3HAkMHW2U6Rdwzfhlu_Z_E";

/* =========================
   ステータス定義
========================= */
const STATUS = {
  PENDING: "pending",
  APPROVED: "approved",
  REJECTED: "rejected",
  CANCELED: "canceled"
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
  MONTHLY_MAIN: "月間有給取得一覧_MAIN",
  YEARLY_MAIN: "年間有給取得一覧_MAIN",

  MONTHLY_PARTNER: "月間有給取得一覧_PARTNER",
  YEARLY_PARTNER: "年間有給取得一覧_PARTNER"
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
  const p = e && e.parameter && e.parameter.p ? e.parameter.p : "";

  if (p === "manifest") {
    const manifest = {
      name: "有給申請システム",
      short_name: "有給申請",
      start_url: ".",
      display: "standalone",
      background_color: "#f3f9fb",
      theme_color: "#4f9fba",
      icons: [
        {
          src: "ここにicon192.pngの画像URL",
          sizes: "192x192",
          type: "image/png"
        },
        {
          src: "ここにicon512.pngの画像URL",
          sizes: "512x512",
          type: "image/png"
        }
      ]
    };

    return ContentService
      .createTextOutput(JSON.stringify(manifest))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const page = e && e.parameter && e.parameter.page
    ? e.parameter.page
    : "menu";

  const template = HtmlService.createTemplateFromFile(page);

  template.initialEmployeeId =
    e && e.parameter && e.parameter.employee_id
      ? String(e.parameter.employee_id).trim()
      : "";

  return template.evaluate()
    .setTitle("有給管理システム")
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

/* =========================
   include
========================= */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function includeTemplate(filename) {
  return HtmlService.createTemplateFromFile(filename).evaluate().getContent();
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

function getOutputSheetName(type, companyCode) {
  const code = String(companyCode || "MAIN")
    .trim()
    .toUpperCase();

  if (type === "monthly") {
    return code === "PARTNER"
      ? OUTPUT_SHEET.MONTHLY_PARTNER
      : OUTPUT_SHEET.MONTHLY_MAIN;
  }

  if (type === "yearly") {
    return code === "PARTNER"
      ? OUTPUT_SHEET.YEARLY_PARTNER
      : OUTPUT_SHEET.YEARLY_MAIN;
  }

  throw new Error("不正な出力タイプです");
}

/* =========================
   キャッシュクリア
========================= */
function clearAppCache() {
  const cache = CacheService.getScriptCache();
  const currentFiscalYear = getCurrentFiscalYear();

  cache.remove(CACHE_KEY.EMPLOYEE_MAP);
  cache.remove(CACHE_KEY.COMPANY_CALENDAR);

  [
    currentFiscalYear - 1,
    currentFiscalYear,
    currentFiscalYear + 1
  ].forEach(year => {
    cache.remove(CACHE_KEY.EMPLOYEES_FOR_REQUEST_PREFIX + year);
  });
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

function ensureSheetColumn_(sheet, headerName) {
  const name = String(headerName || "").trim();
  if (!name) throw new Error("追加するヘッダー名が空です");

  let headerInfo = getHeaderMap(sheet);
  if (name in headerInfo.map) return headerInfo;

  const nextColumn = headerInfo.headers.length + 1;
  sheet.getRange(1, nextColumn).setValue(name);
  return getHeaderMap(sheet);
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

function getDisplayName(employee) {
  if (!employee) return "";
  return String(employee.display_name || employee.name || "").trim();
}

function appendRowFast_(sheet, values) {
  const nextRow = sheet.getLastRow() + 1;
  sheet.getRange(nextRow, 1, 1, values.length).setValues([values]);
}

function updateSheetRowFast_(sheet, sheetRow, rowValues) {
  sheet.getRange(sheetRow, 1, 1, rowValues.length).setValues([rowValues]);
}

/* =========================
   期間計算
========================= */
function getFiscalYearRange(fiscalYear) {
  return getFiscalYearRangeWithStart(fiscalYear, 4);
}

function getFiscalYearRangeWithStart(fiscalYear, startMonth) {
  const fiscalStartMonth = Number(startMonth || 4);

  const start = new Date(Number(fiscalYear), fiscalStartMonth - 1, 1);
  const end = new Date(Number(fiscalYear) + 1, fiscalStartMonth - 1, 0);

  return { start, end };
}

function getFiscalYearFromDateWithStart(dateValue, startMonth) {
  const date = parseLocalDate(dateValue);
  const year = date.getFullYear();
  const month = date.getMonth() + 1;
  const fiscalStartMonth = Number(startMonth || 4);

  return month >= fiscalStartMonth ? year : year - 1;
}

function getFiscalYearFromDate(dateValue) {
  return getFiscalYearFromDateWithStart(dateValue, 4);
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
   admin初期表示用：前月＋当月の期間
========================= */
function getAdminRecentRange() {
  const today = new Date();
  const start = new Date(today.getFullYear(), today.getMonth() - 1, 1);
  const end = new Date(today.getFullYear(), today.getMonth() + 1, 0);

  return { start, end };
}

function getAdminPendingFocusRange() {
  const today = new Date();
  const start = new Date(today.getFullYear(), today.getMonth() - 1, 26);

  return { start };
}

function isRequestInDateRange(rowObj, start, end) {
  if (!rowObj.start_date || !rowObj.end_date) return false;

  const requestStart = parseLocalDate(rowObj.start_date);
  const requestEnd = parseLocalDate(rowObj.end_date);

  const from = new Date(start.getFullYear(), start.getMonth(), start.getDate());
  const to = new Date(end.getFullYear(), end.getMonth(), end.getDate());

  return requestStart <= to && requestEnd >= from;
}

function isRequestOnOrAfterDate(rowObj, start) {
  if (!rowObj.start_date || !rowObj.end_date) return false;

  const requestEnd = parseLocalDate(rowObj.end_date);
  const from = new Date(start.getFullYear(), start.getMonth(), start.getDate());

  return requestEnd >= from;
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
  const date = parseLocalDate(dateValue);

  if (date.getDay() === 0) {
    return CALENDAR_TYPE.HOLIDAY;
  }

  const key = toDateKey(date);

  if (calendarMap && key in calendarMap) {
    return calendarMap[key];
  }

  return CALENDAR_TYPE.WORKDAY;
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

  // 半休・1日申請は、その日が営業日でないとNG
  if (normalizedHalfDay || toDateKey(start) === toDateKey(end)) {
    const type = getCalendarTypeForDate(start, calendarMap);

    if (type !== CALENDAR_TYPE.WORKDAY) {
      throw new Error(
        formatDateValue(start) + " は " + getCalendarLabel(type) + " のため有給申請できません"
      );
    }

    return;
  }

  // 複数日申請は、日曜日・休日・有給NG日を飛ばしてOK
  // ただし、期間内に1日も申請可能日がない場合はNG
  let cursor = new Date(start);
  let allowedCount = 0;

  while (cursor <= end) {
    const type = getCalendarTypeForDate(cursor, calendarMap);

    if (type === CALENDAR_TYPE.WORKDAY) {
      allowedCount++;
    }

    cursor.setDate(cursor.getDate() + 1);
  }

  if (allowedCount === 0) {
    throw new Error("選択した期間に有給申請できる日がありません");
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
       name: String(rowObj.name || rowObj.employee_id || "").trim(),
       display_name: String(rowObj.display_name || "").trim(),

       company_code: String(rowObj.company_code || "").trim(),
       company_name: String(rowObj.company_name || "").trim(),

       fiscal_start_month: Number(rowObj.fiscal_start_month || 4),

       leave_management_target: String(rowObj.leave_management_target || "").toUpperCase() === "TRUE",

       employment_status: String(rowObj.employment_status || "").trim()
       };
    })
    .filter(emp => emp.id);
}

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

function getCurrentFiscalYear() {
  return getFiscalYearFromDate(new Date());
}

/* =========================
   付与情報
========================= */
function getGrantMapByFiscalYear(fiscalYear) {
  const sheet = getSheet("paid_leave_grants");
  const headerInfo = requireHeaders(sheet, [
    "employee_id",
    "grant_date",
    "grant_days",
    "carry_over_days"
  ]);

  const data = sheet.getDataRange().getValues();
  const result = {};
  const employeeDetailMap = getEmployeeDetailMap();

  if (data.length <= 1) return result;

  data.slice(1).forEach(row => {
    const rowObj = rowToObject(row, headerInfo.headers);
    const employeeId = String(rowObj.employee_id || "").trim();

    if (!employeeId) return;
    if (!rowObj.grant_date) return;

    const fiscalStartMonth = getFiscalStartMonthByEmployeeId(employeeId, employeeDetailMap);
    const rowYear = getFiscalYearFromDateWithStart(rowObj.grant_date, fiscalStartMonth);

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

/* =========================
   承認済み取得日数
========================= */
function getApprovedUsedDaysByFiscalYear(fiscalYear) {
  const employees = getEmployees();
  const employeeIds = employees.map(emp => emp.id);

  return getApprovedUsedDaysByFiscalYearForEmployeeIds(fiscalYear, employeeIds);
}
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
  const calendarMap = getCompanyCalendarMap();
  const employeeDetailMap = getEmployeeDetailMap();

  if (data.length <= 1) return result;

  data.slice(1).forEach(row => {
    const rowObj = rowToObject(row, headerInfo.headers);
    const employeeId = String(rowObj.employee_id || "").trim();
    const status = norm(rowObj.status);

    if (!employeeId) return;
    if (!targetIds.has(employeeId)) return;
    if (status !== STATUS.APPROVED) return;
    if (!rowObj.start_date || !rowObj.end_date) return;

    const fiscalStartMonth = getFiscalStartMonthByEmployeeId(employeeId, employeeDetailMap);
    const range = getFiscalYearRangeWithStart(fiscalYear, fiscalStartMonth);

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
   残日数計算
========================= */
function buildBalance(employeeId, grantInfo, usedDays) {
  const previousDays = Number(grantInfo.carry_over_days || 0);
  const grantDays = Number(grantInfo.grant_days || 0);
  const used = Number(usedDays || 0);

  const remainingFromPrevious = previousDays - used;

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

  const currentRemainingDays = previousDays + grantDays - used;

  return {
    employee_id: employeeId,
    current_remaining_days: currentRemainingDays < 0 ? 0 : currentRemainingDays,
    carry_over_days: previousDays,
    grant_days: grantDays,
    used_days: used,
    next_carry_over_days: nextCarryOverDays,
    expired_days: expiredDays
  };
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

    result[employeeId] = buildBalance(employeeId, grantInfo, usedMap[employeeId] || 0);
  });

  return result;
}

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

    result[employeeId] = buildBalance(employeeId, grantInfo, usedMap[employeeId] || 0);
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

  return buildBalance(employeeId, grantInfo, usedMap[employeeId] || 0);
}

/* =========================
   FIFO残日数計算（試験実装）
   既存表示には未接続
========================= */
function calculateFifoPaidLeaveBalance(employeeId, asOfDateValue) {
  const targetEmployeeId = String(employeeId || "").trim();
  if (!targetEmployeeId) throw new Error("employeeId がありません");

  const asOfDate = asOfDateValue ? parseLocalDate(asOfDateValue) : parseLocalDate(new Date());
  const grants = getFifoPaidLeaveGrantRows_(targetEmployeeId, asOfDate);
  const usedRows = getFifoApprovedLeaveUseRows_(targetEmployeeId, asOfDate);
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

  const totalGrantedDays = grants.reduce((sum, grant) => sum + grant.total_days, 0);
  const usedDays = usedRows.reduce((sum, row) => sum + Number(row.days || 0), 0);
  const allocatedUsedDays = allocations.reduce((sum, row) => sum + Number(row.consumed_days || 0), 0);
  const unallocatedUsedDays = usedRows.reduce((sum, row) => sum + Number(row.unallocated_days || 0), 0);
  const expiredDays = grants.reduce((sum, grant) => sum + grant.expired_days, 0);
  const currentRemainingDays = grants.reduce((sum, grant) => sum + grant.active_remaining_days, 0);

  return {
    employee_id: targetEmployeeId,
    as_of_date: formatDateValue(asOfDate),
    current_remaining_days: currentRemainingDays,
    total_granted_days: totalGrantedDays,
    used_days: usedDays,
    allocated_used_days: allocatedUsedDays,
    unallocated_used_days: unallocatedUsedDays,
    expired_days: expiredDays,
    grant_details: grants.map(grant => ({
      grant_id: grant.grant_id,
      grant_date: formatDateValue(grant.grant_date),
      valid_from: formatDateValue(grant.valid_from_date),
      valid_to: formatDateValue(grant.valid_to_date),
      grant_type: grant.grant_type,
      year: grant.year,
      grant_days: grant.grant_days,
      carry_over_days: grant.carry_over_days,
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

/* =========================
   初期導入残高を仮想付与ロットとして含めるFIFO
   通常表示・CSVには未接続。管理者試算・年跨ぎ候補で使用。
========================= */


function calculateFifoBalanceWithOpeningBalance_(employeeId, asOfDate) {
  const context = createFifoBalanceComparisonContext_(asOfDate);
  return calculateFifoBalanceWithOpeningBalanceFromContext_(
    employeeId,
    asOfDate,
    context
  );
}

function calculateFifoBalanceWithOpeningBalanceFromContext_(employeeId, asOfDate, context) {
  const grantData = getFifoGrantRowsWithOpeningBalanceFromContext_(
    employeeId,
    asOfDate,
    context
  );
  const grants = grantData.grants;
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
        lot_type: grant.lot_type,
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
    calculation_mode: "grant_days_plus_opening_balance_virtual_lots",
    current_remaining_days: grants.reduce((sum, grant) => sum + grant.active_remaining_days, 0),
    total_granted_days: grants.reduce((sum, grant) => sum + grant.total_days, 0),
    opening_balance_days_total: grantData.opening_balance_records.reduce(
      (sum, row) => sum + Number(row.carry_over_days || 0),
      0
    ),
    excluded_non_opening_carry_over_days_total: grantData.excluded_carry_over_records.reduce(
      (sum, row) => sum + Number(row.carry_over_days || 0),
      0
    ),
    expiry_unconfirmed_opening_balance_days_total: grantData.opening_balance_records
      .filter(row => row.validity_needs_review)
      .reduce((sum, row) => sum + Number(row.carry_over_days || 0), 0),
    used_days: usedRows.reduce((sum, row) => sum + Number(row.days || 0), 0),
    allocated_used_days: allocations.reduce((sum, row) => sum + Number(row.consumed_days || 0), 0),
    unallocated_used_days: usedRows.reduce((sum, row) => sum + Number(row.unallocated_days || 0), 0),
    expired_days: grants.reduce((sum, grant) => sum + grant.expired_days, 0),
    opening_balance_records: grantData.opening_balance_records,
    excluded_carry_over_records: grantData.excluded_carry_over_records,
    grant_details: grants.map(grant => ({
      grant_id: grant.grant_id,
      source_grant_id: grant.source_grant_id,
      lot_type: grant.lot_type,
      grant_date: formatDateValue(grant.grant_date),
      valid_from: formatDateValue(grant.valid_from_date),
      valid_to: formatDateValue(grant.valid_to_date),
      validity_basis: grant.validity_basis,
      validity_needs_review: grant.validity_needs_review,
      grant_type: grant.grant_type,
      year: grant.year,
      grant_days: grant.grant_days,
      opening_balance_days: grant.opening_balance_days,
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
    allocations: allocations,
    validity_warning: grantData.opening_balance_records.some(row => row.validity_needs_review)
      ? "初期導入残高に valid_from または valid_to が不足する行があります。試算では grant_date 起点の2年期限を仮定していますが、本番切替前に期限確認が必要です。"
      : ""
  };
}

function getFifoGrantRowsWithOpeningBalance_(employeeId, asOfDate) {
  const context = createFifoBalanceComparisonContext_(asOfDate);
  return getFifoGrantRowsWithOpeningBalanceFromContext_(
    employeeId,
    asOfDate,
    context
  );
}

function getFifoGrantRowsWithOpeningBalanceFromContext_(employeeId, asOfDate, context) {
  const grants = [];
  const openingBalanceRecords = [];
  const excludedCarryOverRecords = [];

  (context.grants_by_employee[employeeId] || []).forEach(rowObj => {
    const grantDate = rowObj.grant_date;
    const validFromDate = rowObj.valid_from_date;
    const validToDate = rowObj.valid_to_date;
    if (validFromDate > asOfDate) return;
    if (!rowObj.is_finalized) return;

    const grantId = rowObj.grant_id;
    const grantDays = Number(rowObj.grant_days || 0);
    const carryOverDays = Number(rowObj.carry_over_days || 0);
    const notes = String(rowObj.notes || "");
    const isOpeningBalance = isOpeningBalanceRecordForFifo_(notes);
    const validityNeedsReview =
      !rowObj.has_recorded_valid_from ||
      !rowObj.has_recorded_valid_to;
    const validityBasis = validityNeedsReview
      ? "推定: grant_date から2年後の前日"
      : "記録済み valid_from / valid_to";

    if (grantDays > 0) {
      grants.push({
        grant_id: grantId,
        source_grant_id: grantId,
        lot_type: "grant_days",
        grant_date: grantDate,
        valid_from_date: validFromDate,
        valid_to_date: validToDate,
        validity_basis: validityBasis,
        validity_needs_review: false,
        grant_type: rowObj.grant_type,
        year: rowObj.year || "",
        grant_days: grantDays,
        opening_balance_days: 0,
        total_days: grantDays,
        used_days: 0,
        remaining_days: grantDays,
        active_remaining_days: 0,
        expired_days: 0,
        is_expired: false
      });
    }

    if (carryOverDays <= 0) return;

    const carryOverRecord = {
      grant_id: grantId,
      grant_date: formatDateValue(grantDate),
      grant_type: rowObj.grant_type,
      year: rowObj.year || "",
      carry_over_days: carryOverDays,
      valid_from: formatDateValue(validFromDate),
      valid_to: formatDateValue(validToDate),
      validity_basis: validityBasis,
      validity_needs_review: validityNeedsReview,
      notes: notes
    };

    if (!isOpeningBalance) {
      excludedCarryOverRecords.push(carryOverRecord);
      return;
    }

    openingBalanceRecords.push(carryOverRecord);
    grants.push({
      grant_id: grantId + "#opening_balance",
      source_grant_id: grantId,
      lot_type: "opening_balance_virtual_lot",
      grant_date: grantDate,
      valid_from_date: validFromDate,
      valid_to_date: validToDate,
      validity_basis: validityBasis,
      validity_needs_review: validityNeedsReview,
      grant_type: "opening_balance_virtual_lot",
      year: rowObj.year || "",
      grant_days: 0,
      opening_balance_days: carryOverDays,
      total_days: carryOverDays,
      used_days: 0,
      remaining_days: carryOverDays,
      active_remaining_days: 0,
      expired_days: 0,
      is_expired: false
    });
  });

  grants.sort((a, b) => {
    if (a.grant_date.getTime() !== b.grant_date.getTime()) {
      return a.grant_date - b.grant_date;
    }
    return String(a.grant_id).localeCompare(String(b.grant_id));
  });

  return {
    grants: grants,
    opening_balance_records: openingBalanceRecords,
    excluded_carry_over_records: excludedCarryOverRecords
  };
}

function isOpeningBalanceRecordForFifo_(notes) {
  return String(notes || "").indexOf("初期導入残高") !== -1;
}

function getFifoOpeningBalanceDifferenceReason_(info) {
  const legacyDays = Number(info.legacy_remaining_days || 0);
  const withoutCarryOverDays = Number(info.without_carry_over_remaining_days || 0);
  const withOpeningBalanceDays = Number(info.with_opening_balance_remaining_days || 0);
  const openingBalanceDays = Number(info.opening_balance_days_total || 0);
  const expiryUnconfirmedDays = Number(info.expiry_unconfirmed_days_total || 0);

  if (expiryUnconfirmedDays > 0) {
    return "初期導入残高を仮想ロットとして含めていますが、期限未確認の日数があります。";
  }

  if (withOpeningBalanceDays === legacyDays && withoutCarryOverDays !== legacyDays) {
    return "初期導入残高を仮想ロットとして含めると旧計算と一致します。";
  }

  if (
    openingBalanceDays > 0 &&
    Math.abs(withOpeningBalanceDays - legacyDays) <
      Math.abs(withoutCarryOverDays - legacyDays)
  ) {
    return "初期導入残高を含めることで旧計算との差分が縮小します。残る差分は期限または使用割当の確認が必要です。";
  }

  if (openingBalanceDays === 0) {
    return "初期導入残高に該当する行はありません。";
  }

  return "初期導入残高以外にも、有効期限または年度集計方式による差分がある可能性があります。";
}

function getYearlyGrantFinalizedMap_(fiscalYear) {
  const sheet = getSheet("paid_leave_grants");
  const headerInfo = requireHeaders(sheet, [
    "employee_id",
    "grant_type",
    "year"
  ]);
  const data = sheet.getDataRange().getValues();
  const result = {};

  if (data.length <= 1) return result;

  data.slice(1).forEach(row => {
    const rowObj = rowToObject(row, headerInfo.headers);
    const employeeId = String(rowObj.employee_id || "").trim();

    if (!employeeId) return;
    if (String(rowObj.grant_type || "").trim() !== "yearly") return;
    if (Number(rowObj.year) !== Number(fiscalYear)) return;

    result[employeeId] = true;
  });

  return result;
}

function normalizePagingOptions_(options) {
  const opts = options || {};
  const rawLimit = Number(opts.limit || 20);
  const rawOffset = Number(opts.offset || 0);
  const limit = Math.max(1, Math.min(isFinite(rawLimit) ? rawLimit : 20, 20));
  const offset = Math.max(0, isFinite(rawOffset) ? rawOffset : 0);

  return {
    limit: limit,
    offset: offset
  };
}

function buildPagedResponse_(rows, options) {
  const allRows = Array.isArray(rows) ? rows : [];
  const page = normalizePagingOptions_(options);
  const pageRows = allRows.slice(page.offset, page.offset + page.limit);

  return {
    ok: true,
    total_count: allRows.length,
    row_count: pageRows.length,
    offset: page.offset,
    limit: page.limit,
    has_prev: page.offset > 0,
    has_next: page.offset + page.limit < allRows.length,
    rows: pageRows
  };
}

function createFifoBalanceComparisonContext_(asOfDate) {
  return {
    as_of_date: asOfDate,
    calendar_map: getCompanyCalendarMap(),
    grants_by_employee: getPaidLeaveGrantRowsByEmployeeForFifoCompare_(),
    requests_by_employee: getLeaveRequestRowsByEmployeeForFifoCompare_()
  };
}

function getPaidLeaveGrantRowsByEmployeeForFifoCompare_() {
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
  const result = {};

  if (data.length <= 1) return result;

  data.slice(1).forEach(row => {
    const rowObj = rowToObject(row, headerInfo.headers);
    const employeeId = String(rowObj.employee_id || "").trim();
    if (!employeeId || !rowObj.grant_date) return;

    const grantDate = parseLocalDate(rowObj.grant_date);
    const validFromDate = rowObj.valid_from ? parseLocalDate(rowObj.valid_from) : grantDate;
    const validToDate = rowObj.valid_to
      ? parseLocalDate(rowObj.valid_to)
      : addDaysLocal_(addYearsLocal_(grantDate, 2), -1);
    const grantDays = Number(rowObj.grant_days || 0);
    const carryOverDays = Number(rowObj.carry_over_days || 0);
    const finalizedValue = String(rowObj.is_finalized == null ? "" : rowObj.is_finalized)
      .trim()
      .toUpperCase();

    if (!result[employeeId]) result[employeeId] = [];

    result[employeeId].push({
      grant_id: String(rowObj.grant_id || ""),
      employee_id: employeeId,
      grant_date: grantDate,
      valid_from_date: validFromDate,
      valid_to_date: validToDate,
      grant_type: String(rowObj.grant_type || ""),
      year: rowObj.year || "",
      grant_days: grantDays,
      carry_over_days: carryOverDays,
      total_days: grantDays + carryOverDays,
      notes: String(rowObj.notes || ""),
      has_recorded_valid_from: !!rowObj.valid_from,
      has_recorded_valid_to: !!rowObj.valid_to,
      is_finalized: finalizedValue !== "FALSE"
    });
  });

  return result;
}

function getLeaveRequestRowsByEmployeeForFifoCompare_() {
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
  const result = {};

  if (data.length <= 1) return result;

  data.slice(1).forEach(row => {
    const rowObj = rowToObject(row, headerInfo.headers);
    const employeeId = String(rowObj.employee_id || "").trim();
    if (!employeeId) return;

    if (!result[employeeId]) result[employeeId] = [];
    result[employeeId].push(rowObj);
  });

  return result;
}


function calculateFifoBalanceFromContext_(employeeId, asOfDate, context) {
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
      carry_over_days: grant.carry_over_days,
      total_days: grant.total_days,
      used_days: 0,
      remaining_days: grant.total_days,
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
    current_remaining_days: grants.reduce((sum, grant) => sum + grant.active_remaining_days, 0),
    total_granted_days: grants.reduce((sum, grant) => sum + grant.total_days, 0),
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
      carry_over_days: grant.carry_over_days,
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

function getFifoApprovedLeaveUseRowsFromContext_(employeeId, asOfDate, context) {
  const requests = context.requests_by_employee[employeeId] || [];
  const result = [];

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
      if (useDate > asOfDate) return;

      result.push({
        request_id: String(rowObj.request_id || ""),
        use_date: useDate,
        days: Number(item.days || 0),
        unallocated_days: 0
      });
    });
  });

  return result.sort((a, b) => {
    if (a.use_date.getTime() !== b.use_date.getTime()) {
      return a.use_date - b.use_date;
    }
    return String(a.request_id).localeCompare(String(b.request_id));
  });
}

function isFifoBalanceCompareTargetEmployee_(emp) {
  const status = String(emp.employment_status || "").trim().toLowerCase();
  const isActive = status === "active" || status === "在職";
  return isActive && emp.leave_management_target === true;
}


function getFifoPaidLeaveGrantRows_(employeeId, asOfDate) {
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
    "year"
  ]);
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  return data.slice(1)
    .map(row => rowToObject(row, headerInfo.headers))
    .filter(rowObj => {
      if (String(rowObj.employee_id || "").trim() !== String(employeeId)) return false;
      if (!rowObj.grant_date) return false;

      const validFromDate = rowObj.valid_from
        ? parseLocalDate(rowObj.valid_from)
        : parseLocalDate(rowObj.grant_date);
      if (validFromDate > asOfDate) return false;

      const finalizedValue = String(rowObj.is_finalized == null ? "" : rowObj.is_finalized)
        .trim()
        .toUpperCase();

      return finalizedValue !== "FALSE";
    })
    .map(rowObj => {
      const grantDate = parseLocalDate(rowObj.grant_date);
      const validFromDate = rowObj.valid_from ? parseLocalDate(rowObj.valid_from) : grantDate;
      const validToDate = rowObj.valid_to
        ? parseLocalDate(rowObj.valid_to)
        : addDaysLocal_(addYearsLocal_(grantDate, 2), -1);
      const grantDays = Number(rowObj.grant_days || 0);
      const carryOverDays = Number(rowObj.carry_over_days || 0);
      const totalDays = grantDays + carryOverDays;

      return {
        grant_id: String(rowObj.grant_id || ""),
        grant_date: grantDate,
        valid_from_date: validFromDate,
        valid_to_date: validToDate,
        grant_type: String(rowObj.grant_type || ""),
        year: rowObj.year || "",
        grant_days: grantDays,
        carry_over_days: carryOverDays,
        total_days: totalDays,
        used_days: 0,
        remaining_days: totalDays,
        active_remaining_days: 0,
        expired_days: 0,
        is_expired: false
      };
    })
    .sort((a, b) => {
      if (a.grant_date.getTime() !== b.grant_date.getTime()) {
        return a.grant_date - b.grant_date;
      }
      return String(a.grant_id).localeCompare(String(b.grant_id));
    });
}

function getFifoApprovedLeaveUseRows_(employeeId, asOfDate) {
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
  const result = [];

  if (data.length <= 1) return result;

  data.slice(1).forEach(row => {
    const rowObj = rowToObject(row, headerInfo.headers);
    const targetEmployeeId = String(rowObj.employee_id || "").trim();
    const status = norm(rowObj.status);
    const requestType = String(rowObj.type || "paid_leave").trim();

    if (targetEmployeeId !== String(employeeId)) return;
    if (status !== STATUS.APPROVED) return;
    if (requestType && requestType !== "paid_leave") return;
    if (!rowObj.start_date || !rowObj.end_date) return;

    const dailyRows = expandLeaveRequestToDailyRows(
      rowObj.start_date,
      rowObj.end_date,
      rowObj.days,
      rowObj.half_day,
      calendarMap
    );

    dailyRows.forEach(item => {
      const useDate = parseLocalDate(item.date);
      if (useDate > asOfDate) return;

      result.push({
        request_id: String(rowObj.request_id || ""),
        use_date: useDate,
        days: Number(item.days || 0),
        unallocated_days: 0
      });
    });
  });

  return result.sort((a, b) => {
    if (a.use_date.getTime() !== b.use_date.getTime()) {
      return a.use_date - b.use_date;
    }
    return String(a.request_id).localeCompare(String(b.request_id));
  });
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

  appendRowFast_(
    sheet,
    objectToRow(rowObj, headerInfo.headers)
  );
}

function appendEmployeeMasterLog(actionType, employeeId, comment) {
  appendUsageLog({
    request_id: employeeId || "",
    action_type: actionType || "",
    operator_id: "admin",
    operator_name: "管理者",
    comment: comment || ""
  });
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

  if (!data.employee_id) throw new Error("employee_id がありません");
  if (!data.start_date || !data.end_date) throw new Error("start_date または end_date がありません");

  const start = parseLocalDate(data.start_date);
  const end = parseLocalDate(data.end_date);

  const isHalf =
    data.half_day === true ||
    String(data.half_day || "").toLowerCase() === "true";

  validateLeaveRequestDates(start, end, isHalf ? (data.half_type || "half") : "");

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

  const employeeMap = getEmployeeDetailMap();
  const employee = employeeMap[String(data.employee_id || "").trim()];
  const fiscalStartMonth = employee ? Number(employee.fiscal_start_month || 4) : 4;

  rowObj.year = getFiscalYearFromDateWithStart(start, fiscalStartMonth);
  rowObj.created_at = now;
  rowObj.updated_at = now;

appendRowFast_(
  sheet,
  objectToRow(rowObj, headerInfo.headers)
);

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
   個人ページ用：承認待ち申請の修正
========================= */
function updatePendingLeaveRequestForEmployee(requestId, employeeId, data) {
  const targetRequestId = String(requestId || "").trim();
  const targetEmployeeId = String(employeeId || "").trim();

  if (!targetRequestId) throw new Error("requestId がありません");
  if (!targetEmployeeId) throw new Error("employeeId がありません");
  if (!data || typeof data !== "object") {
    throw new Error("更新データがありません");
  }
  if (!data.start_date || !data.end_date) {
    throw new Error("start_date または end_date がありません");
  }

  const sheet = getSheet("leave_requests");
  const headerInfo = requireHeaders(sheet, [
    "request_id",
    "employee_id",
    "start_date",
    "end_date",
    "days",
    "type",
    "half_day",
    "reason",
    "reason_detail",
    "status",
    "year",
    "updated_at"
  ]);

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow <= 1) {
    throw new Error("申請データがありません");
  }

  const values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  const map = headerInfo.map;

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const rowRequestId = String(row[map.request_id] || "").trim();
    const rowEmployeeId = String(row[map.employee_id] || "").trim();

    if (rowRequestId !== targetRequestId || rowEmployeeId !== targetEmployeeId) {
      continue;
    }

    const rowObj = rowToObject(row, headerInfo.headers);
    const status = norm(rowObj.status || STATUS.PENDING);

    if (status !== STATUS.PENDING) {
      throw new Error("承認待ちの申請だけ修正できます");
    }

    const start = parseLocalDate(data.start_date);
    const end = parseLocalDate(data.end_date);
    const isHalf =
      data.half_day === true ||
      String(data.half_day || "").toLowerCase() === "true";
    const effectiveEnd = isHalf ? start : end;

    validateLeaveRequestDates(start, effectiveEnd, isHalf ? (data.half_type || "half") : "");

    const days = isHalf ? 0.5 : calculateLeaveDays(start, effectiveEnd);
    const employeeMap = getEmployeeDetailMap();
    const fiscalStartMonth = getFiscalStartMonthByEmployeeId(targetEmployeeId, employeeMap);

    rowObj.start_date = start;
    rowObj.end_date = effectiveEnd;
    rowObj.days = days;
    rowObj.type = data.type || "paid_leave";
    rowObj.half_day = isHalf ? (data.half_type || "") : "";
    rowObj.reason = data.reason || "";
    rowObj.reason_detail = data.reason_detail || "";
    rowObj.status = STATUS.PENDING;
    rowObj.year = getFiscalYearFromDateWithStart(start, fiscalStartMonth);
    rowObj.updated_at = new Date();

    updateSheetRowFast_(sheet, i + 1, objectToRow(rowObj, headerInfo.headers));

    appendUsageLog({
      request_id: targetRequestId,
      action_type: "request_update",
      operator_id: targetEmployeeId,
      operator_name: "申請者",
      comment: "Pending leave request updated"
    });

    clearAppCache();

    return {
      ok: true,
      request_id: targetRequestId
    };
  }

  throw new Error("対象の申請が見つかりません");
}

/* =========================
   個人ページ用：承認待ち申請の取消
========================= */
function cancelPendingLeaveRequestForEmployee(requestId, employeeId) {
  const targetRequestId = String(requestId || "").trim();
  const targetEmployeeId = String(employeeId || "").trim();

  if (!targetRequestId) throw new Error("requestId がありません");
  if (!targetEmployeeId) throw new Error("employeeId がありません");

  const sheet = getSheet("leave_requests");
  const headerInfo = requireHeaders(sheet, [
    "request_id",
    "employee_id",
    "status",
    "updated_at"
  ]);

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow <= 1) {
    throw new Error("申請データがありません");
  }

  const values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  const map = headerInfo.map;

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const rowRequestId = String(row[map.request_id] || "").trim();
    const rowEmployeeId = String(row[map.employee_id] || "").trim();

    if (rowRequestId !== targetRequestId || rowEmployeeId !== targetEmployeeId) {
      continue;
    }

    const rowObj = rowToObject(row, headerInfo.headers);
    const status = norm(rowObj.status || STATUS.PENDING);

    if (status !== STATUS.PENDING) {
      if (status === STATUS.APPROVED) {
        throw new Error("この申請はすでに承認済みのため、本人画面からは取消できません。管理者へ連絡してください。");
      }

      if (status === STATUS.REJECTED || status === STATUS.CANCELED) {
        throw new Error("すでに処理済みです。履歴を更新してください。");
      }

      throw new Error("承認待ちの申請だけ取消できます");
    }

    rowObj.status = STATUS.CANCELED;
    rowObj.updated_at = new Date();

    updateSheetRowFast_(sheet, i + 1, objectToRow(rowObj, headerInfo.headers));

    appendUsageLog({
      request_id: targetRequestId,
      action_type: "request_cancel",
      operator_id: targetEmployeeId,
      operator_name: "申請者",
      comment: "Pending leave request canceled"
    });

    clearAppCache();

    return {
      ok: true,
      request_id: targetRequestId
    };
  }

  throw new Error("対象の申請が見つかりません");
}

/* =========================
   社員詳細MAP
========================= */
function getEmployeeDetailMap() {
  const employees = getEmployees();
  const map = {};

  employees.forEach(emp => {
    map[emp.id] = emp;
  });

  return map;
}

/* =========================
   社員ごとの年度開始月取得
========================= */
function getFiscalStartMonthByEmployeeId(employeeId, employeeMap) {
  const map = employeeMap || getEmployeeDetailMap();
  const employee = map[String(employeeId || "").trim()];

  return employee ? Number(employee.fiscal_start_month || 4) : 4;
}

/* =========================
   管理画面用：初期表示
   前月＋当月のみ
========================= */
function getRequestsByStatus(status) {
  if (norm(status) === STATUS.PENDING) {
    const pendingRange = getAdminPendingFocusRange();

    return searchRequests({
      status: status,
      start_date: formatDateValue(pendingRange.start)
    });
  }

  const range = getAdminRecentRange();

  return searchRequests({
    status: status,
    start_date: formatDateValue(range.start),
    end_date: formatDateValue(range.end)
  });
}

/* =========================
   管理画面用：承認待ち軽量一覧
========================= */
function getPendingRequestsForAdminLight() {
  const range = getAdminPendingFocusRange();
  const requestSheet = getSheet("leave_requests");
  const requestHeaderInfo = requireHeaders(requestSheet, [
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
    "year",
    "created_at",
    "updated_at"
  ]);

  const employeeSheet = getSheet("employees");
  const employeeHeaderInfo = requireHeaders(employeeSheet, [
    "employee_id",
    "name"
  ]);

  const employeeColumns = [
    "employee_id",
    "name",
    "display_name",
    "display_employee_id",
    "company_name",
    "department"
  ].filter(key => key in employeeHeaderInfo.map);
  const employeeMinCol = Math.min(...employeeColumns.map(key => employeeHeaderInfo.map[key]));
  const employeeMaxCol = Math.max(...employeeColumns.map(key => employeeHeaderInfo.map[key]));
  const employeeLastRow = employeeSheet.getLastRow();
  const employeeValues = employeeLastRow > 1
    ? employeeSheet
      .getRange(2, employeeMinCol + 1, employeeLastRow - 1, employeeMaxCol - employeeMinCol + 1)
      .getValues()
    : [];
  const empIdx = key =>
    key in employeeHeaderInfo.map ? employeeHeaderInfo.map[key] - employeeMinCol : -1;
  const employeeMap = {};
  const empEmployeeIdIdx = empIdx("employee_id");
  const empNameIdx = empIdx("name");
  const empDisplayNameIdx = empIdx("display_name");
  const empDisplayIdIdx = empIdx("display_employee_id");
  const empCompanyNameIdx = empIdx("company_name");
  const empDepartmentIdx = empIdx("department");

  employeeValues.forEach(row => {
    const employeeId = String(row[empEmployeeIdIdx] || "").trim();
    if (!employeeId) return;

    employeeMap[employeeId] = {
      display_employee_id: empDisplayIdIdx >= 0 ? String(row[empDisplayIdIdx] || "").trim() : "",
      employee_name: String(
        (empDisplayNameIdx >= 0 ? row[empDisplayNameIdx] : "") ||
        row[empNameIdx] ||
        employeeId
      ).trim(),
      company_name: empCompanyNameIdx >= 0 ? String(row[empCompanyNameIdx] || "").trim() : "",
      department: empDepartmentIdx >= 0 ? String(row[empDepartmentIdx] || "").trim() : ""
    };
  });

  const requestColumns = [
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
    "year",
    "created_at",
    "updated_at"
  ];
  const requestMinCol = Math.min(...requestColumns.map(key => requestHeaderInfo.map[key]));
  const requestMaxCol = Math.max(...requestColumns.map(key => requestHeaderInfo.map[key]));
  const requestLastRow = requestSheet.getLastRow();
  const values = requestLastRow > 1
    ? requestSheet
      .getRange(2, requestMinCol + 1, requestLastRow - 1, requestMaxCol - requestMinCol + 1)
      .getValues()
    : [];

  if (values.length === 0) return [];

  const reqIdx = key => requestHeaderInfo.map[key] - requestMinCol;
  const requestIdIdx = reqIdx("request_id");
  const employeeIdIdx = reqIdx("employee_id");
  const requestDateIdx = reqIdx("request_date");
  const startDateIdx = reqIdx("start_date");
  const endDateIdx = reqIdx("end_date");
  const daysIdx = reqIdx("days");
  const typeIdx = reqIdx("type");
  const halfDayIdx = reqIdx("half_day");
  const reasonIdx = reqIdx("reason");
  const reasonDetailIdx = reqIdx("reason_detail");
  const statusIdx = reqIdx("status");
  const yearIdx = reqIdx("year");
  const createdAtIdx = reqIdx("created_at");
  const updatedAtIdx = reqIdx("updated_at");

  return values
    .map(row => {
      const rowStatus = norm(row[statusIdx] || STATUS.PENDING);
      const startDate = row[startDateIdx];
      const endDate = row[endDateIdx];

      if (rowStatus !== STATUS.PENDING) return null;
      if (!startDate || !endDate) return null;
      if (!isRequestOnOrAfterDate({ start_date: startDate, end_date: endDate }, range.start)) return null;

      const employeeId = String(row[employeeIdIdx] || "").trim();
      const employee = employeeMap[employeeId] || {};
      const startText = formatDateValue(startDate);
      const endText = formatDateValue(endDate);
      const halfDay = String(row[halfDayIdx] || "");

      return {
        request_id: String(row[requestIdIdx] || ""),
        employee_id: employeeId,
        display_employee_id: employee.display_employee_id || "",
        employee_name: employee.employee_name || employeeId || "Unknown",
        start_date: startText,
        end_date: endText,
        days: row[daysIdx] || 0,
        type: String(row[typeIdx] || "paid_leave"),
        half_day: halfDay,
        reason: String(row[reasonIdx] || ""),
        reason_detail: String(row[reasonDetailIdx] || ""),
        status: rowStatus,
        request_date: formatDateValue(row[requestDateIdx]),
        created_at: row[createdAtIdx] ? formatDateValue(row[createdAtIdx]) : "",
        updated_at: row[updatedAtIdx] ? formatDateValue(row[updatedAtIdx]) : "",
        year: row[yearIdx] || "",
        date_label: startText !== endText ? startText + " 〜 " + endText : startText,
        leave_type_label: getRequestHistoryLeaveTypeLabel_(halfDay, startDate, endDate),
        status_label: getRequestHistoryStatusLabel_(rowStatus),
        company_name: employee.company_name || "",
        department: employee.department || "",
        current_remaining_days: "",
        used_days: ""
      };
    })
    .filter(item => item)
    .sort((a, b) => {
      if (a.start_date !== b.start_date) return a.start_date < b.start_date ? 1 : -1;
      return a.employee_id > b.employee_id ? 1 : -1;
    });
}

/* =========================
   管理画面用：申請検索
========================= */
function searchRequests(filters) {
  filters = filters || {};

  const sheet = getSheet("leave_requests");
  const headerInfo = requireHeaders(sheet, [
    "request_id",
    "employee_id",
    "start_date",
    "end_date",
    "days",
    "half_day",
    "reason",
    "reason_detail",
    "status"
  ]);

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const employeeMap = getEmployeeDetailMap();
  const targetStatus = norm(filters.status || "");
  const keyword = norm(filters.employeeKeyword || "");

  const startFilter = filters.start_date ? parseLocalDate(filters.start_date) : null;
  const endFilter = filters.end_date ? parseLocalDate(filters.end_date) : null;

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

  return data.slice(1)
    .map(row => {
      const rowObj = rowToObject(row, headerInfo.headers);
      const rowStatus = norm(rowObj.status);
      const employeeId = String(rowObj.employee_id || "").trim();
      const employee = employeeMap[employeeId];
      const employeeName = getDisplayName(employee) || employeeId || "Unknown";

      if (!rowObj.start_date || !rowObj.end_date) return null;

      if (targetStatus && targetStatus !== "all" && rowStatus !== targetStatus) {
        return null;
      }

      if (keyword) {
        const targetText = norm(
          employeeId +
          employeeName +
          String(employee && employee.name ? employee.name : "")
        );
        if (!targetText.includes(keyword)) return null;
      }

      if (startFilter && endFilter) {
        if (!isRequestInDateRange(rowObj, startFilter, endFilter)) return null;
      } else if (startFilter) {
        const requestEnd = parseLocalDate(rowObj.end_date);
        if (requestEnd < startFilter) return null;
      } else if (endFilter) {
        const requestStart = parseLocalDate(rowObj.start_date);
        if (requestStart > endFilter) return null;
      }

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
        employee_name: employeeName,
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
        reason_detail: String(rowObj.reason_detail || ""),
        status: rowStatus,
        current_remaining_days: balance.current_remaining_days,
        grant_days: balance.grant_days,
        carry_over_days: balance.carry_over_days,
        used_days: balance.used_days
      };
    })
    .filter(item => item)
    .sort((a, b) => {
      if (a.start_date !== b.start_date) return a.start_date < b.start_date ? 1 : -1;
      return a.employee_id > b.employee_id ? 1 : -1;
    });
}

/* =========================
   個人ページ用：本人申請履歴
========================= */
function getEmployeeLeaveHistoryForRequest(employeeId, limit) {
  const targetEmployeeId = String(employeeId || "").trim();
  const maxRows = Math.max(1, Math.min(Number(limit || 50), 100));

  if (!targetEmployeeId) {
    return [];
  }

  const sheet = getSheet("leave_requests");
  const headerInfo = requireHeaders(sheet, [
    "request_id",
    "employee_id",
    "start_date",
    "end_date",
    "days",
    "half_day",
    "reason",
    "reason_detail",
    "status",
    "created_at"
  ]);

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow <= 1) {
    return [];
  }

  const values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  const map = headerInfo.map;
  const rows = [];

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const rowEmployeeId = String(row[map.employee_id] || "").trim();

    if (rowEmployeeId !== targetEmployeeId) continue;
    if (!row[map.start_date] || !row[map.end_date]) continue;

    const startDate = new Date(row[map.start_date]);
    const endDate = new Date(row[map.end_date]);
    const createdAt = row[map.created_at] ? new Date(row[map.created_at]) : startDate;

    rows.push({
      requestId: String(row[map.request_id] || ""),
      startDate: startDate,
      endDate: endDate,
      createdAt: createdAt,
      days: row[map.days] || 0,
      halfDay: String(row[map.half_day] || ""),
      reason: String(row[map.reason] || ""),
      reasonDetail: String(row[map.reason_detail] || ""),
      status: norm(row[map.status] || STATUS.PENDING)
    });
  }

  rows.sort((a, b) => {
    const startDiff = b.startDate.getTime() - a.startDate.getTime();
    if (startDiff !== 0) return startDiff;
    return b.createdAt.getTime() - a.createdAt.getTime();
  });

  return rows.slice(0, maxRows).map(row => {
    const startText = formatDateValue(row.startDate);
    const endText = formatDateValue(row.endDate);

    return {
      request_id: row.requestId,
      start_date: toDateKey(row.startDate),
      end_date: toDateKey(row.endDate),
      date_label: startText !== endText ? startText + " 〜 " + endText : startText,
      leave_type_label: getRequestHistoryLeaveTypeLabel_(row.halfDay, row.startDate, row.endDate),
      days: row.days,
      half_day: row.halfDay,
      reason: row.reason,
      reason_detail: row.reasonDetail,
      status: row.status,
      status_label: getRequestHistoryStatusLabel_(row.status),
      can_edit: row.status === STATUS.PENDING
    };
  });
}

function getRequestHistoryLeaveTypeLabel_(halfDay, startDate, endDate) {
  const value = norm(halfDay);

  if (value === "am") return "午前半休";
  if (value === "pm") return "午後半休";
  if (toDateKey(startDate) !== toDateKey(endDate)) return "複数日有給";

  return "1日有給";
}

function getRequestHistoryStatusLabel_(status) {
  const value = norm(status);

  if (value === STATUS.APPROVED) return "承認済み";
  if (value === STATUS.REJECTED) return "否認";
  if (value === STATUS.CANCELED) return "取消済み";

  return "承認待ち";
}

function approveRequestsBatch(requestIds, adminUser) {
  if (!Array.isArray(requestIds) || requestIds.length === 0) {
    throw new Error("承認対象が選択されていません");
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

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow <= 1) {
    throw new Error("申請データがありません");
  }

  const data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  const targetIdSet = new Set(requestIds.map(id => String(id)));

  const now = new Date();

  const operatorId = adminUser && adminUser.admin_id
    ? String(adminUser.admin_id).trim()
    : "admin";

  const operatorName = adminUser && adminUser.admin_name
    ? String(adminUser.admin_name).trim()
    : "管理者";

  let updatedCount = 0;

  const updatedRows = data.slice(1).map(row => {
    const requestId = String(row[headerInfo.map.request_id] || "");

    if (!targetIdSet.has(requestId)) {
      return row;
    }

    row[headerInfo.map.status] = STATUS.APPROVED;
    row[headerInfo.map.approver_id] = operatorId;
    row[headerInfo.map.approver_name] = operatorName;
    row[headerInfo.map.approved_at] = now;
    row[headerInfo.map.updated_at] = now;

    updatedCount++;

    return row;
  });

  if (updatedCount === 0) {
    throw new Error("承認対象の申請が見つかりません");
  }

  sheet.getRange(2, 1, updatedRows.length, lastCol).setValues(updatedRows);

  const logSheet = getSheet("usage_log");
  const logHeaderInfo = requireHeaders(logSheet, [
    "log_id",
    "request_id",
    "action_type",
    "operator_id",
    "operator_name",
    "action_date",
    "comment"
  ]);

  const logRows = requestIds.map(requestId => {
    const rowObj = createEmptyRowObject(logHeaderInfo.headers);

    rowObj.log_id = Utilities.getUuid();
    rowObj.request_id = requestId;
    rowObj.action_type = "approve";
    rowObj.operator_id = operatorId;
    rowObj.operator_name = operatorName;
    rowObj.action_date = now;
    rowObj.comment = "Batch approved by " + operatorName;

    return objectToRow(rowObj, logHeaderInfo.headers);
  });

  const logStartRow = logSheet.getLastRow() + 1;
  logSheet
    .getRange(logStartRow, 1, logRows.length, logRows[0].length)
    .setValues(logRows);

  clearAppCache();

  return {
    ok: true,
    count: updatedCount
  };
}

function approveRequest(requestId, adminUser) {
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

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow <= 1) {
    throw new Error("申請データがありません");
  }

  const data = sheet.getRange(1, 1, lastRow, lastCol).getValues();

  const rowIndex = data.findIndex((row, index) => {
    if (index === 0) return false;
    return String(row[headerInfo.map.request_id]) === String(requestId);
  });

  if (rowIndex === -1) {
    throw new Error("対象の申請が見つかりません");
  }

  const sheetRow = rowIndex + 1;
  const rowValues = data[rowIndex].slice();
  const now = new Date();

  const operatorId = adminUser && adminUser.admin_id
    ? String(adminUser.admin_id).trim()
    : "admin";

  const operatorName = adminUser && adminUser.admin_name
    ? String(adminUser.admin_name).trim()
    : "管理者";

  rowValues[headerInfo.map.status] = STATUS.APPROVED;
  rowValues[headerInfo.map.approver_id] = operatorId;
  rowValues[headerInfo.map.approver_name] = operatorName;
  rowValues[headerInfo.map.approved_at] = now;
  rowValues[headerInfo.map.updated_at] = now;

  updateSheetRowFast_(sheet, sheetRow, rowValues);

  appendUsageLog({
    request_id: requestId,
    action_type: "approve",
    operator_id: operatorId,
    operator_name: operatorName,
    comment: "Approved by " + operatorName
  });

  clearAppCache();

  return { ok: true };
}

/* =========================
   否認
========================= */
function rejectRequest(requestId, reason, adminUser) {
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
  const operatorId = adminUser && adminUser.admin_id
  ? String(adminUser.admin_id).trim()
  : "admin";
  const operatorName = adminUser && adminUser.admin_name
  ? String(adminUser.admin_name).trim()
  : "管理者";


  appendUsageLog({
  request_id: requestId,
  action_type: "reject",
  operator_id: operatorId,
  operator_name: operatorName,
  comment: reason || ""
});

  clearAppCache();

  return { ok: true };
}

/* =========================
   ログ取得
   初期表示は前月＋当月のみ
========================= */
function getUsageLogs() {
  const range = getAdminRecentRange();

  return searchUsageLogs({
    start_date: formatDateValue(range.start),
    end_date: formatDateValue(range.end)
  });
}

/* =========================
   ログ検索
========================= */
function searchUsageLogs(filters) {
  filters = filters || {};

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

  const keyword = norm(filters.keyword || "");
  const actionType = norm(filters.action_type || "");

  const startFilter = filters.start_date ? parseLocalDate(filters.start_date) : null;
  const endFilter = filters.end_date ? parseLocalDate(filters.end_date) : null;

  const employeeMap = getEmployeeDetailMap();

  return data.slice(1)
    .map(row => {
      const rowObj = rowToObject(row, headerInfo.headers);
      const actionDate = rowObj.action_date ? parseLocalDate(rowObj.action_date) : null;

      if (!actionDate) return null;

      if (startFilter && actionDate < startFilter) return null;
      if (endFilter && actionDate > endFilter) return null;

      const rowActionType = String(rowObj.action_type || "");
      if (actionType && norm(rowActionType) !== actionType) return null;

      const requestId = String(rowObj.request_id || "");
      const employee = employeeMap[requestId];
      const employeeName = getDisplayName(employee) || "";

      if (keyword) {
        const targetText = norm(
          requestId +
          employeeName +
          String(employee && employee.name ? employee.name : "") +
          String(rowObj.operator_id || "") +
          String(rowObj.operator_name || "") +
          String(rowObj.comment || "") +
          rowActionType +
          getLogActionLabel(rowActionType)
        );

        if (!targetText.includes(keyword)) return null;
      }

      return {
        log_id: rowObj.log_id,
        request_id: requestId,
        employee_name: employeeName,
        type: rowActionType,
        type_label: getLogActionLabel(rowActionType),
        type_class: getLogActionClass(rowActionType),
        user_id: rowObj.operator_id,
        user_name: rowObj.operator_name,
        date: formatDateValue(rowObj.action_date),
        comment: rowObj.comment
      };
    })
    .filter(item => item)
    .sort((a, b) => {
      const dateA = new Date(a.date);
      const dateB = new Date(b.date);
      return dateB - dateA;
    });
}

/* =========================
   月間取得一覧出力
========================= */
function exportMonthlyPaidLeaveReport(targetYear, targetMonth, companyCode) {
  if (!targetYear || !targetMonth) {
    const today = new Date();
    targetYear = today.getFullYear();
    targetMonth = today.getMonth() + 1;
  }

  const code = String(companyCode || "MAIN").trim().toUpperCase();

  const range = getClosingMonthRange(Number(targetYear), Number(targetMonth));
  const preview = getMonthlyPaidLeaveReportPreview({
    target_year: targetYear,
    target_month: targetMonth,
    company_code: code
  });

  const outputSheet = getOutputSheet(
    getOutputSheetName("monthly", code)
  );

  outputSheet.clearContents();

  const values = [];
  values.push(["表示用氏名", "取得日", "取得日数"]);

  if (preview.detail_rows.length > 0) {
    preview.detail_rows.forEach(row => {
      values.push([
        row.employee_name,
        row.date,
        row.days
      ]);
    });
  }

  outputSheet.getRange(1, 1, values.length, 3).setValues(values);

  return {
    ok: true,
    company_code: code,
    period_start: formatDateValue(range.start),
    period_end: formatDateValue(range.end),
    detail_count: preview.detail_count,
    total_count: preview.total_count
  };
}

/* =========================
   月間取得一覧プレビュー
   画面表示・CSV用
========================= */
function getMonthlyPaidLeaveReportPreview(filters) {
  filters = filters || {};

  const targetYear = Number(filters.target_year || new Date().getFullYear());
  const targetMonth = Number(filters.target_month || (new Date().getMonth() + 1));

  const companyCodeFilter = String(filters.company_code || "").trim().toUpperCase();
  const companyNameFilter = String(filters.company_name || "").trim();

  const range = getClosingMonthRange(targetYear, targetMonth);

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
  const employees = getEmployeesForAdmin();
  const employeeMap = {};
  const calendarMap = getCompanyCalendarMap();

  employees.forEach(emp => {
    employeeMap[emp.employee_id] = emp;
  });

  const detailRows = [];
  const totalMap = {};

  if (leaveData.length > 1) {
    leaveData.slice(1).forEach(row => {
      const rowObj = rowToObject(row, leaveHeaderInfo.headers);
      const employeeId = String(rowObj.employee_id || "").trim();
      const status = norm(rowObj.status);

      if (!employeeId) return;
      if (status !== STATUS.APPROVED) return;

      const emp = employeeMap[employeeId];
      if (!emp) return;

      const empCompanyCode = String(emp.company_code || "").trim().toUpperCase();
      const empCompanyName = String(emp.company_name || "").trim();

      if (companyCodeFilter && empCompanyCode !== companyCodeFilter) return;
      if (companyNameFilter && empCompanyName !== companyNameFilter) return;

      const dailyRows = expandLeaveRequestToDailyRows(
        rowObj.start_date,
        rowObj.end_date,
        rowObj.days,
        rowObj.half_day,
        calendarMap
      );

      dailyRows.forEach(item => {
        if (!isDateInRange(item.date, range.start, range.end)) return;

        const dateText = formatDateValue(item.date);
        const days = Number(item.days || 0);

        detailRows.push({
          employee_id: employeeId,
          display_employee_id: emp.display_employee_id || "",
          employee_name: getDisplayName(emp) || employeeId,
          company_code: empCompanyCode,
          company_name: empCompanyName,
          date: dateText,
          days: days
        });

        if (!totalMap[employeeId]) {
          totalMap[employeeId] = {
            employee_id: employeeId,
            display_employee_id: emp.display_employee_id || "",
            employee_name: getDisplayName(emp) || employeeId,
            company_code: empCompanyCode,
            company_name: empCompanyName,
            total_days: 0
          };
        }

        totalMap[employeeId].total_days += days;
      });
    });
  }

  detailRows.sort((a, b) => {
    if (a.employee_id !== b.employee_id) {
      return a.employee_id > b.employee_id ? 1 : -1;
    }
    return a.date > b.date ? 1 : -1;
  });

  const totalRows = Object.values(totalMap)
    .sort((a, b) => a.employee_id > b.employee_id ? 1 : -1);

  return {
    ok: true,
    target_year: targetYear,
    target_month: targetMonth,
    period_start: formatDateValue(range.start),
    period_end: formatDateValue(range.end),
    company_code: companyCodeFilter || "ALL",
    company_name: companyNameFilter || "",
    detail_rows: detailRows,
    total_rows: totalRows,
    detail_count: detailRows.length,
    total_count: totalRows.length
  };
}

/* =========================
   年間取得一覧出力
========================= */
function exportYearlyPaidLeaveReport(fiscalYear, companyCode) {
  const code = String(companyCode || "MAIN").trim().toUpperCase();

  if (!fiscalYear) {
    fiscalYear = getFiscalYearFromDate(new Date());
  }

  const employees = getEmployees().filter(emp => {
  return (
    String(emp.company_code || "").trim().toUpperCase() === code &&
    String(emp.employment_status || "").trim().toLowerCase() === "active" &&
    emp.leave_management_target === true
  );
});

  const fiscalStartMonth =
    employees.length > 0
      ? Number(employees[0].fiscal_start_month || 4)
      : code === "PARTNER" ? 6 : 4;

  const yearRange = getFiscalYearRangeWithStart(Number(fiscalYear), fiscalStartMonth);

  const grantMap = getGrantMapByFiscalYear(Number(fiscalYear));
  const usedMap = getApprovedUsedDaysByFiscalYear(Number(fiscalYear));

  const reportRows = employees
    .map(emp => {
      const grantInfo = grantMap[emp.id] || {
        employee_id: emp.id,
        grant_days: 0,
        carry_over_days: 0
      };

      const balance = buildBalance(
        emp.id,
        grantInfo,
        usedMap[emp.id] || 0
      );

      return [
        emp.id,
        getDisplayName(emp) || emp.id,
        balance.carry_over_days,
        balance.grant_days,
        balance.used_days,
        balance.next_carry_over_days,
        balance.expired_days
      ];
    })
    .sort((a, b) => a[0] > b[0] ? 1 : -1);

  const outputSheet = getOutputSheet(
    getOutputSheetName("yearly", code)
  );

  outputSheet.clearContents();

  const values = [];
  values.push(["年間有給取得一覧_" + code]);
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
    company_code: code,
    fiscal_year: Number(fiscalYear),
    period_start: formatDateValue(yearRange.start),
    period_end: formatDateValue(yearRange.end),
    row_count: reportRows.length
  };
}

/* =========================
   申請画面用社員一覧
   社員ごとの年度開始月対応版
========================= */
function getEmployeesForRequest() {
  const sheet = getSheet("employees");
  const headerInfo = requireHeaders(sheet, [
    "employee_id",
    "name",
    "name_kana",
    "employment_type",
    "employment_status",
    "leave_management_target",
    "fiscal_start_month",
    "display_order"
  ]);

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const employeeRows = data.slice(1)
    .map(row => rowToObject(row, headerInfo.headers))
    .filter(rowObj => {
      const employeeId = String(rowObj.employee_id || "").trim();
      const name = String(rowObj.name || "").trim();

      const employmentStatus = String(rowObj.employment_status || "")
        .trim()
        .toLowerCase();

      const leaveTargetRaw = String(rowObj.leave_management_target || "")
        .trim()
        .toUpperCase();

      const isActive =
        employmentStatus === "active" ||
        employmentStatus === "在職";

      const isLeaveTarget =
        rowObj.leave_management_target === true ||
        leaveTargetRaw === "TRUE" ||
        leaveTargetRaw === "1" ||
        leaveTargetRaw === "YES" ||
        leaveTargetRaw === "対象";

      return employeeId && name && isActive && isLeaveTarget;
    })
    .sort((a, b) => Number(a.display_order || 9999) - Number(b.display_order || 9999));

  const fiscalYearGroups = {};

  employeeRows.forEach(rowObj => {
    const employeeId = String(rowObj.employee_id || "").trim();
    const fiscalStartMonth = Number(rowObj.fiscal_start_month || 4);
    const fiscalYear = getFiscalYearFromDateWithStart(new Date(), fiscalStartMonth);

    if (!fiscalYearGroups[fiscalYear]) {
      fiscalYearGroups[fiscalYear] = [];
    }

    fiscalYearGroups[fiscalYear].push(employeeId);
  });

  const balanceMapByFiscalYear = {};

  Object.keys(fiscalYearGroups).forEach(fiscalYear => {
    balanceMapByFiscalYear[fiscalYear] =
      getEmployeeBalanceMapForEmployeeIdsForFiscalYear(
        Number(fiscalYear),
        fiscalYearGroups[fiscalYear]
      );
  });

  return employeeRows.map(rowObj => {
    const employeeId = String(rowObj.employee_id || "").trim();
    const fiscalStartMonth = Number(rowObj.fiscal_start_month || 4);
    const fiscalYear = getFiscalYearFromDateWithStart(new Date(), fiscalStartMonth);

    const balanceMap = balanceMapByFiscalYear[fiscalYear] || {};
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
      employment_type: String(rowObj.employment_type || "").trim(),

      fiscal_year: fiscalYear,
      fiscal_start_month: fiscalStartMonth,

      current_remaining_days: Number(balance.current_remaining_days || 0),
      carry_over_days: Number(balance.carry_over_days || 0),
      grant_days: Number(balance.grant_days || 0),
      used_days: usedDays,
      five_day_used: fiveDayUsed,
      five_day_remaining: fiveDayRemaining,
      five_day_completed: fiveDayRemaining === 0
    };
  });
}

/* =========================
   フロント用返却
========================= */
function getCalendarRules() {
  return getCompanyCalendarMap();
}

function validateRequestDatesOnly(startDate, endDate, halfDay, halfType) {
  const isHalf =
    halfDay === true ||
    String(halfDay || "").toLowerCase() === "true";

  validateLeaveRequestDates(
    startDate,
    endDate,
    isHalf ? (halfType || "half") : ""
  );

  return { ok: true };
}

/* =========================
   社員マスター整備
   互換用：表示順整理のみ実行
========================= */
function maintainEmployeeMaster() {
  return maintainEmployeeDisplayOrderOnly_();
}

/* =========================
   社員表示順整理
   ID系列の列は更新しない
========================= */
function maintainEmployeeDisplayOrderOnly_() {
  const sheet = getSheet("employees");
  const headerInfo = requireHeaders(sheet, [
    "employee_id",
    "display_employee_id",
    "company_code",
    "name",
    "name_kana",
    "employment_status",
    "display_order"
  ]);

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { ok: true, message: "社員データがありません", count: 0 };

  const companyCodeIndex = headerInfo.map.company_code;
  const kanaIndex = headerInfo.map.name_kana;
  const statusIndex = headerInfo.map.employment_status;
  const orderIndex = headerInfo.map.display_order;

  const rowItems = data.slice(1).map((row, index) => ({
    row: row,
    rowNumber: index + 2
  }));

  rowItems.sort((a, b) => {
    const statusA = getEmploymentStatusOrder_(a.row[statusIndex]);
    const statusB = getEmploymentStatusOrder_(b.row[statusIndex]);

    if (statusA !== statusB) return statusA - statusB;

    const companyA = getCompanyOrder_(a.row[companyCodeIndex]);
    const companyB = getCompanyOrder_(b.row[companyCodeIndex]);

    if (companyA !== companyB) return companyA - companyB;

    const kanaA = String(a.row[kanaIndex] || "");
    const kanaB = String(b.row[kanaIndex] || "");

    return kanaA.localeCompare(kanaB, "ja");
  });

  rowItems.forEach((item, index) => {
    sheet.getRange(item.rowNumber, orderIndex + 1).setValue(index + 1);
  });

  clearAppCache();

  return {
    ok: true,
    message: "社員の表示順を整理しました",
    count: rowItems.length
  };
}

/* =========================

   管理画面用：表示順整理

========================= */

function runMaintainEmployeeMasterFromAdmin() {

  const result = maintainEmployeeDisplayOrderOnly_();

  appendEmployeeMasterLog(

    "employee_maintain",

    "",

    "表示順整理を実行しました。対象件数: " + result.count

  );

  return result;

}

function getLogActionLabel(actionType) {
  const type = String(actionType || "");

  const labels = {
    submit: "申請",
    approve: "承認",
    reject: "否認",

    employee_add: "社員追加",
    employee_update: "社員編集",
    employee_retire: "退職処理",
    employee_maintain: "表示順整理",

    six_month_grant: "6か月有給付与",
    yearly_grant: "年次有給付与"
  };

  return labels[type] || type;
}

function getLogActionClass(actionType) {
  const type = String(actionType || "");

  if (type === "approve") return "log-approve";
  if (type === "reject") return "log-reject";
  if (type === "submit") return "log-submit";

  if (type === "employee_add") return "log-employee-add";
  if (type === "employee_update") return "log-employee-update";
  if (type === "employee_retire") return "log-employee-retire";
  if (type === "employee_maintain") return "log-employee-maintain";

  if (type === "six_month_grant") return "log-employee-update";
  if (type === "yearly_grant") return "log-employee-update";

  return "log-default";
}

/* =========================
   IDの次番号取得
========================= */
function getNextIdNumber_(usedIds, prefix) {
  let max = 0;

  usedIds.forEach(id => {
    const text = String(id || "").trim();
    if (!text.startsWith(prefix)) return;

    const numberPart = text.replace(prefix, "");
    const num = Number(numberPart);

    if (!isNaN(num) && num > max) {
      max = num;
    }
  });

  return max + 1;
}

function generateEmployeeIdsForNewEmployee_(sheet, headerInfo, companyCode) {
  const data = sheet.getDataRange().getValues();
  const employeeIdIndex = headerInfo.map.employee_id;
  const displayIdIndex = headerInfo.map.display_employee_id;
  const usedEmployeeIds = new Set();
  const usedDisplayIds = {
    W: new Set(),
    P: new Set()
  };

  data.slice(1).forEach(row => {
    const employeeId = String(row[employeeIdIndex] || "").trim();
    const displayId = String(row[displayIdIndex] || "").trim();

    if (employeeId) usedEmployeeIds.add(employeeId);
    if (displayId.startsWith("W")) usedDisplayIds.W.add(displayId);
    if (displayId.startsWith("P")) usedDisplayIds.P.add(displayId);
  });

  const normalizedCompanyCode = normalizeCompanyCode_(companyCode);
  const displayPrefix = normalizedCompanyCode === "PARTNER" ? "P" : "W";

  return {
    employee_id: "EMP" + String(getNextIdNumber_(usedEmployeeIds, "EMP")).padStart(4, "0"),
    display_employee_id:
      displayPrefix +
      String(getNextIdNumber_(usedDisplayIds[displayPrefix], displayPrefix)).padStart(4, "0")
  };
}

/* =========================
   company_code 正規化
========================= */
function normalizeCompanyCode_(companyCode) {
  const value = String(companyCode || "").trim().toUpperCase();

  if (value === "PARTONER") return "PARTNER";
  if (value === "PARTNER") return "PARTNER";

  return "MAIN";
}

/* =========================
   company_code 並び順
========================= */
function getCompanyOrder_(companyCode) {
  const value = normalizeCompanyCode_(companyCode);

  if (value === "MAIN") return 1;
  if (value === "PARTNER") return 2;

  return 9;
}

/* =========================
   在職状況の並び順
========================= */
function getEmploymentStatusOrder_(status) {
  const value = String(status || "").trim().toLowerCase();

  if (value === "active") return 1;
  if (value === "leave") return 2;
  if (value === "retired") return 3;

  return 9;
}

/* =========================
   社員追加
========================= */
function addEmployeeFromAdmin(data) {
  if (!data || typeof data !== "object") {
    throw new Error("社員データがありません");
  }

  const sheet = getSheet("employees");
  ensureEmployeeInitialGrantCheckTargetColumn_(sheet);
  const headerInfo = requireHeaders(sheet, [
    "employee_id",
    "display_employee_id",
    "name",
    "display_name",
    "name_kana",
    "company_code",
    "company_name",
    "department",
    "employment_type",
    "employment_status",
    "hire_date",
    "leave_date",
    "work_days_per_week",
    "fiscal_start_month",
    "leave_management_target",
    "initial_grant_check_target",
    "is_driver",
    "driver_type",
    "default_vehicle_id",
    "display_order",
    "notes",
    "created_at",
    "updated_at"
  ]);

  if (!data.name) throw new Error("氏名を入力してください");
  if (!data.name_kana) throw new Error("ふりがなを入力してください");
  if (!data.company_code) throw new Error("会社区分を選択してください");

  const now = new Date();
  const rowObj = createEmptyRowObject(headerInfo.headers);
  const newIds = generateEmployeeIdsForNewEmployee_(
    sheet,
    headerInfo,
    data.company_code
  );

  rowObj.employee_id = newIds.employee_id;
  rowObj.display_employee_id = newIds.display_employee_id;
  rowObj.name = String(data.name || "").trim();
  rowObj.display_name = String(data.display_name || "").trim();
  rowObj.name_kana = String(data.name_kana || "").trim();
  rowObj.company_code = String(data.company_code || "").trim().toUpperCase();
  rowObj.company_name = String(data.company_name || "").trim();
  rowObj.department = String(data.department || "").trim();
  rowObj.employment_type = String(data.employment_type || "").trim();
  rowObj.employment_status = String(data.employment_status || "active").trim();
  rowObj.hire_date = data.hire_date ? parseLocalDate(data.hire_date) : "";
  rowObj.leave_date = data.leave_date ? parseLocalDate(data.leave_date) : "";
  rowObj.work_days_per_week = data.work_days_per_week ? Number(data.work_days_per_week) : "";
  rowObj.fiscal_start_month = getFiscalStartMonthForCompanyCode_(
    rowObj.company_code,
    data.fiscal_start_month
  );
  rowObj.leave_management_target = String(data.leave_management_target || "").toUpperCase() === "TRUE";
  rowObj.initial_grant_check_target = true;
  rowObj.is_driver = String(data.is_driver || "").toUpperCase() === "TRUE";
  rowObj.driver_type = String(data.driver_type || "").trim();
  rowObj.default_vehicle_id = String(data.default_vehicle_id || "").trim();
  rowObj.display_order = "";
  rowObj.notes = String(data.notes || "").trim();
  rowObj.created_at = now;
  rowObj.updated_at = now;

  appendRowFast_(
  sheet,
  objectToRow(rowObj, headerInfo.headers)
);

  maintainEmployeeDisplayOrderOnly_();

  appendEmployeeMasterLog(
    "employee_add",
    "",
    "社員を追加しました: " + rowObj.name
  );

  return {
    ok: true,
    message: "社員を追加しました"
  };
}

function ensureEmployeeInitialGrantCheckTargetColumn_(sheet) {
  return ensureSheetColumn_(sheet || getSheet("employees"), "initial_grant_check_target");
}

function getFiscalStartMonthForCompanyCode_(companyCode, fallbackMonth) {
  const code = String(companyCode || "").trim().toUpperCase();
  if (code === "PARTNER") return 6;
  if (code === "MAIN") return 4;

  const month = Number(fallbackMonth || 4);
  return month >= 1 && month <= 12 ? month : 4;
}

/* =========================
   社員一覧取得（管理画面用）
========================= */
function getEmployeesForAdmin() {
  const sheet = getSheet("employees");
  const headerInfo = requireHeaders(sheet, [
    "employee_id",
    "display_employee_id",
    "name",
    "display_name",
    "name_kana",
    "company_code",
    "company_name",
    "department",
    "employment_type",
    "employment_status",
    "hire_date",
    "leave_date",
    "work_days_per_week",
    "fiscal_start_month",
    "leave_management_target",
    "is_driver",
    "driver_type",
    "default_vehicle_id",
    "display_order",
    "notes"
  ]);

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  return data.slice(1)
    .map(row => {
      const obj = rowToObject(row, headerInfo.headers);

      return {
        employee_id: String(obj.employee_id || "").trim(),
        display_employee_id: String(obj.display_employee_id || "").trim(),
        name: String(obj.name || "").trim(),
        display_name: String(obj.display_name || "").trim(),
        name_kana: String(obj.name_kana || "").trim(),
        company_code: String(obj.company_code || "").trim(),
        company_name: String(obj.company_name || "").trim(),
        department: String(obj.department || "").trim(),
        employment_type: String(obj.employment_type || "").trim(),
        employment_status: String(obj.employment_status || "").trim(),
        hire_date: formatDateValue(obj.hire_date),
        leave_date: formatDateValue(obj.leave_date),
        work_days_per_week: obj.work_days_per_week || "",
        fiscal_start_month: obj.fiscal_start_month || "",
        leave_management_target:
          String(obj.leave_management_target || "").toUpperCase() === "TRUE",
        initial_grant_check_target:
          String(obj.initial_grant_check_target || "").toUpperCase() === "TRUE",
        is_driver: String(obj.is_driver || "").toUpperCase() === "TRUE",
        driver_type: String(obj.driver_type || "").trim(),
        default_vehicle_id: String(obj.default_vehicle_id || "").trim(),
        display_order: obj.display_order || "",
        notes: String(obj.notes || "")
      };
    })
    .filter(emp => emp.employee_id)
    .sort((a, b) => Number(a.display_order || 9999) - Number(b.display_order || 9999));
}

function buildEmployeeUpdateDiffComment(beforeObj, afterData) {
  const fields = [
    { key: "name", label: "氏名" },
    { key: "display_name", label: "表示用氏名" },
    { key: "name_kana", label: "ふりがな" },
    { key: "company_code", label: "会社区分" },
    { key: "company_name", label: "会社名" },
    { key: "department", label: "部署" },
    { key: "employment_type", label: "雇用区分" },
    { key: "employment_status", label: "在職状況" },
    { key: "hire_date", label: "入社日", type: "date" },
    { key: "leave_date", label: "退職日", type: "date" },
    { key: "work_days_per_week", label: "週所定労働日数" },
    { key: "fiscal_start_month", label: "有給年度開始月" },
    { key: "leave_management_target", label: "有給管理対象", type: "boolean" },
    { key: "is_driver", label: "運転手区分", type: "driver_boolean" },
    { key: "driver_type", label: "運転手種別" },
    { key: "default_vehicle_id", label: "標準車両ID" },
    { key: "notes", label: "備考" }
  ];

  const diffs = [];

  fields.forEach(field => {
    const beforeValue = normalizeEmployeeLogValue(beforeObj[field.key], field.type);
    const afterValue = normalizeEmployeeLogValue(afterData[field.key], field.type);

    if (beforeValue !== afterValue) {
      diffs.push(
        field.label + "「" + beforeValue + "」→「" + afterValue + "」"
      );
    }
  });

  const name = String(afterData.name || beforeObj.name || "").trim();

  if (diffs.length === 0) {
    return "社員情報を更新しました: " + name + "（変更差分なし）";
  }

  return "社員情報を更新しました: " + name + " / 変更: " + diffs.join("、");
}

function normalizeEmployeeLogValue(value, type) {
  if (type === "date") {
    if (!value) return "";
    return formatDateValue(value).replace(/\//g, "-");
  }

  if (type === "boolean") {
    const text = String(value || "").trim().toUpperCase();
    return text === "TRUE" || value === true ? "対象" : "対象外";
  }

  if (type === "driver_boolean") {
    const text = String(value || "").trim().toUpperCase();
    return text === "TRUE" || value === true ? "運転手" : "運転手ではない";
  }

  return String(value == null ? "" : value).trim();
}

/* =========================
   社員情報更新
========================= */
function updateEmployeeFromAdmin(data) {
  if (!data || typeof data !== "object") {
    throw new Error("社員データがありません");
  }

  if (!data.employee_id) {
    throw new Error("employee_id がありません");
  }

  const sheet = getSheet("employees");
  const headerInfo = requireHeaders(sheet, [
    "employee_id",
    "name",
    "display_name",
    "name_kana",
    "company_code",
    "company_name",
    "department",
    "employment_type",
    "employment_status",
    "hire_date",
    "leave_date",
    "work_days_per_week",
    "fiscal_start_month",
    "leave_management_target",
    "is_driver",
    "driver_type",
    "default_vehicle_id",
    "notes",
    "updated_at"
  ]);

  const dataRange = sheet.getDataRange().getValues();

  const rowIndex = dataRange.findIndex((row, index) => {
    if (index === 0) return false;
    const rowObj = rowToObject(row, headerInfo.headers);
    return String(rowObj.employee_id || "").trim() === String(data.employee_id || "").trim();
  });

  if (rowIndex === -1) {
    throw new Error("対象社員が見つかりません");
  }

  const sheetRow = rowIndex + 1;
  const beforeObj = rowToObject(dataRange[rowIndex], headerInfo.headers);

  sheet.getRange(sheetRow, headerInfo.map.name + 1).setValue(String(data.name || "").trim());
  sheet.getRange(sheetRow, headerInfo.map.display_name + 1).setValue(String(data.display_name || "").trim());
  sheet.getRange(sheetRow, headerInfo.map.name_kana + 1).setValue(String(data.name_kana || "").trim());
  sheet.getRange(sheetRow, headerInfo.map.company_code + 1).setValue(String(data.company_code || "").trim().toUpperCase());
  sheet.getRange(sheetRow, headerInfo.map.company_name + 1).setValue(String(data.company_name || "").trim());
  sheet.getRange(sheetRow, headerInfo.map.department + 1).setValue(String(data.department || "").trim());
  sheet.getRange(sheetRow, headerInfo.map.employment_type + 1).setValue(String(data.employment_type || "").trim());
  sheet.getRange(sheetRow, headerInfo.map.employment_status + 1).setValue(String(data.employment_status || "").trim());

  sheet.getRange(sheetRow, headerInfo.map.hire_date + 1)
    .setValue(data.hire_date ? parseLocalDate(data.hire_date) : "");

  sheet.getRange(sheetRow, headerInfo.map.leave_date + 1)
    .setValue(data.leave_date ? parseLocalDate(data.leave_date) : "");

  sheet.getRange(sheetRow, headerInfo.map.work_days_per_week + 1)
    .setValue(data.work_days_per_week ? Number(data.work_days_per_week) : "");

  sheet.getRange(sheetRow, headerInfo.map.fiscal_start_month + 1)
    .setValue(getFiscalStartMonthForCompanyCode_(data.company_code, data.fiscal_start_month));

  sheet.getRange(sheetRow, headerInfo.map.leave_management_target + 1)
    .setValue(String(data.leave_management_target || "").toUpperCase() === "TRUE");

  sheet.getRange(sheetRow, headerInfo.map.is_driver + 1)
    .setValue(String(data.is_driver || "").toUpperCase() === "TRUE");

  sheet.getRange(sheetRow, headerInfo.map.driver_type + 1)
    .setValue(String(data.driver_type || "").trim());

  sheet.getRange(sheetRow, headerInfo.map.default_vehicle_id + 1)
    .setValue(String(data.default_vehicle_id || "").trim());

  sheet.getRange(sheetRow, headerInfo.map.notes + 1).setValue(String(data.notes || "").trim());
  sheet.getRange(sheetRow, headerInfo.map.updated_at + 1).setValue(new Date());

  maintainEmployeeDisplayOrderOnly_();
  clearAppCache();

  const diffComment = buildEmployeeUpdateDiffComment(beforeObj, data);

  appendEmployeeMasterLog(
    "employee_update",
    data.employee_id,
    diffComment
  );

  return {
    ok: true,
    message: "社員情報を更新しました"
  };
}

/* =========================
   退職処理
========================= */
function retireEmployeeFromAdmin(employeeId, leaveDate) {
  if (!employeeId) {
    throw new Error("employeeId がありません");
  }

  if (!leaveDate) {
    throw new Error("退職日を入力してください");
  }

  const sheet = getSheet("employees");
  const headerInfo = requireHeaders(sheet, [
    "employee_id",
    "employment_status",
    "leave_date",
    "leave_management_target",
    "updated_at"
  ]);

  const data = sheet.getDataRange().getValues();

  const rowIndex = data.findIndex((row, index) => {
    if (index === 0) return false;
    const rowObj = rowToObject(row, headerInfo.headers);
    return String(rowObj.employee_id || "").trim() === String(employeeId || "").trim();
  });

  if (rowIndex === -1) {
    throw new Error("対象社員が見つかりません");
  }

  const sheetRow = rowIndex + 1;

  sheet.getRange(sheetRow, headerInfo.map.employment_status + 1).setValue("retired");
  sheet.getRange(sheetRow, headerInfo.map.leave_date + 1).setValue(parseLocalDate(leaveDate));
  sheet.getRange(sheetRow, headerInfo.map.leave_management_target + 1).setValue(false);
  sheet.getRange(sheetRow, headerInfo.map.updated_at + 1).setValue(new Date());

  maintainEmployeeDisplayOrderOnly_();
  clearAppCache();

  appendEmployeeMasterLog(
    "employee_retire",
    employeeId,
    "退職処理を実行しました。退職日: " + leaveDate
  );

  return {
    ok: true,
    message: "退職処理を完了しました"
  };
}

function getCompanyCalendarMapForRequest() {
  const map = getCompanyCalendarMap();
  const result = {};

  Object.keys(map).forEach(dateKey => {
    result[dateKey] = {
      type: map[dateKey],
      notes: ""
    };
  });

  return result;
}

/* =========================
   company_calendar 管理
========================= */
function ensureCompanyCalendarNotesColumn_() {
  const sheet = getSheet("company_calendar");
  const headerInfo = requireHeaders(sheet, ["date", "type"]);

  if (!("notes" in headerInfo.map)) {
    sheet.getRange(1, headerInfo.headers.length + 1).setValue("notes");
  }

  requireHeaders(sheet, ["date", "type", "notes"]);
  return sheet;
}

function getCompanyCalendarDateRowMap_(sheet, headerInfo) {
  const data = sheet.getDataRange().getValues();
  const map = {};

  if (data.length <= 1) return map;

  data.slice(1).forEach((row, index) => {
    const rowObj = rowToObject(row, headerInfo.headers);
    if (!rowObj.date) return;
    map[toDateKey(rowObj.date)] = {
      rowNumber: index + 2,
      row,
      rowObj
    };
  });

  return map;
}

function getCompanyCalendarPeriod_(fiscalYear, fiscalStartMonth) {
  const year = Number(fiscalYear || 0);
  const startMonth = Number(fiscalStartMonth || 4);

  if (!year) throw new Error("年度を入力してください");
  if (startMonth < 1 || startMonth > 12) {
    throw new Error("年度開始月は1〜12で入力してください");
  }

  const range = getFiscalYearRangeWithStart(year, startMonth);
  return {
    fiscal_year: year,
    fiscal_start_month: startMonth,
    start: range.start,
    end: range.end,
    start_date: toDateKey(range.start),
    end_date: toDateKey(range.end)
  };
}

function getCompanyCalendarRowsForAdmin(fiscalYear, fiscalStartMonth) {
  const period = getCompanyCalendarPeriod_(fiscalYear, fiscalStartMonth);
  const sheet = ensureCompanyCalendarNotesColumn_();
  const headerInfo = requireHeaders(sheet, ["date", "type", "notes"]);
  const rowMap = getCompanyCalendarDateRowMap_(sheet, headerInfo);
  const rows = [];
  let cursor = new Date(period.start);

  while (cursor <= period.end) {
    const dateKey = toDateKey(cursor);
    const existing = rowMap[dateKey];
    const type = cursor.getDay() === 0
      ? CALENDAR_TYPE.HOLIDAY
      : (existing ? norm(existing.rowObj.type) : CALENDAR_TYPE.WORKDAY);

    rows.push({
      date: dateKey,
      day_of_week: ["日", "月", "火", "水", "木", "金", "土"][cursor.getDay()],
      type: type || CALENDAR_TYPE.WORKDAY,
      notes: existing ? String(existing.rowObj.notes || "") : "",
      registered: !!existing
    });

    cursor.setDate(cursor.getDate() + 1);
  }

  return {
    fiscal_year: period.fiscal_year,
    fiscal_start_month: period.fiscal_start_month,
    start_date: period.start_date,
    end_date: period.end_date,
    rows
  };
}

function generateCompanyCalendarFiscalYear(fiscalYear, fiscalStartMonth) {
  const period = getCompanyCalendarPeriod_(fiscalYear, fiscalStartMonth);
  const sheet = ensureCompanyCalendarNotesColumn_();
  const headerInfo = requireHeaders(sheet, ["date", "type", "notes"]);
  const rowMap = getCompanyCalendarDateRowMap_(sheet, headerInfo);
  const rowsToAppend = [];
  let skippedCount = 0;
  let cursor = new Date(period.start);

  while (cursor <= period.end) {
    const dateKey = toDateKey(cursor);

    if (rowMap[dateKey]) {
      skippedCount++;
      cursor.setDate(cursor.getDate() + 1);
      continue;
    }

    const rowObj = createEmptyRowObject(headerInfo.headers);
    rowObj.date = new Date(cursor);
    rowObj.type = cursor.getDay() === 0
      ? CALENDAR_TYPE.HOLIDAY
      : CALENDAR_TYPE.WORKDAY;
    rowObj.notes = "";
    rowsToAppend.push(objectToRow(rowObj, headerInfo.headers));

    cursor.setDate(cursor.getDate() + 1);
  }

  if (rowsToAppend.length > 0) {
    sheet
      .getRange(sheet.getLastRow() + 1, 1, rowsToAppend.length, headerInfo.headers.length)
      .setValues(rowsToAppend);
  }

  clearAppCache();

  return {
    ok: true,
    fiscal_year: period.fiscal_year,
    fiscal_start_month: period.fiscal_start_month,
    start_date: period.start_date,
    end_date: period.end_date,
    added_count: rowsToAppend.length,
    skipped_count: skippedCount
  };
}

function updateCompanyCalendarRowsForAdmin(rows) {
  if (!Array.isArray(rows)) {
    throw new Error("更新データが不正です");
  }

  const sheet = ensureCompanyCalendarNotesColumn_();
  const headerInfo = requireHeaders(sheet, ["date", "type", "notes"]);
  const rowMap = getCompanyCalendarDateRowMap_(sheet, headerInfo);
  const validTypes = [
    CALENDAR_TYPE.WORKDAY,
    CALENDAR_TYPE.HOLIDAY,
    CALENDAR_TYPE.NO_LEAVE
  ];
  let updatedCount = 0;
  let addedCount = 0;

  rows.forEach(item => {
    const dateKey = toDateKey(item.date);
    const type = norm(item.type);

    if (validTypes.indexOf(type) === -1) {
      throw new Error(dateKey + " の区分が不正です");
    }

    const notes = String(item.notes || "").trim();
    const existing = rowMap[dateKey];
    const rowObj = existing
      ? rowToObject(existing.row, headerInfo.headers)
      : createEmptyRowObject(headerInfo.headers);

    rowObj.date = parseLocalDate(dateKey);
    rowObj.type = type;
    rowObj.notes = notes;

    if (existing) {
      updateSheetRowFast_(sheet, existing.rowNumber, objectToRow(rowObj, headerInfo.headers));
      updatedCount++;
    } else {
      appendRowFast_(sheet, objectToRow(rowObj, headerInfo.headers));
      addedCount++;
    }
  });

  clearAppCache();

  return {
    ok: true,
    updated_count: updatedCount,
    added_count: addedCount
  };
}

function overwriteCompanyCalendarFiscalYear_(fiscalYear, fiscalStartMonth) {
  const data = getCompanyCalendarRowsForAdmin(fiscalYear, fiscalStartMonth);
  const rows = data.rows.map(row => ({
    date: row.date,
    type: row.day_of_week === "日" ? CALENDAR_TYPE.HOLIDAY : CALENDAR_TYPE.WORKDAY,
    notes: row.notes || ""
  }));

  return updateCompanyCalendarRowsForAdmin(rows);
}

/* =========================
   年間一覧CSV用データ取得
========================= */
function getYearlyPaidLeaveReportCsvData(fiscalYear, companyCode) {
  const code = String(companyCode || "MAIN").trim().toUpperCase();

  if (!fiscalYear) {
    fiscalYear = getFiscalYearFromDate(new Date());
  }

  const employees = getEmployees().filter(emp => {
    return (
      String(emp.company_code || "").trim().toUpperCase() === code &&
      String(emp.employment_status || "").trim().toLowerCase() === "active" &&
      emp.leave_management_target === true
    );
  });

  const fiscalStartMonth =
    employees.length > 0
      ? Number(employees[0].fiscal_start_month || 4)
      : code === "PARTNER" ? 6 : 4;

  const yearRange = getFiscalYearRangeWithStart(Number(fiscalYear), fiscalStartMonth);

  const grantMap = getGrantMapByFiscalYear(Number(fiscalYear));
  const usedMap = getApprovedUsedDaysByFiscalYear(Number(fiscalYear));

  const rows = employees
    .map(emp => {
      const grantInfo = grantMap[emp.id] || {
        employee_id: emp.id,
        grant_days: 0,
        carry_over_days: 0
      };

      const balance = buildBalance(
        emp.id,
        grantInfo,
        usedMap[emp.id] || 0
      );

      return [
        emp.id,
        getDisplayName(emp) || emp.id,
        balance.carry_over_days,
        balance.grant_days,
        balance.used_days,
        balance.next_carry_over_days,
        balance.expired_days
      ];
    })
    .sort((a, b) => a[0] > b[0] ? 1 : -1);

  return {
    ok: true,
    company_code: code,
    fiscal_year: Number(fiscalYear),
    period_start: formatDateValue(yearRange.start),
    period_end: formatDateValue(yearRange.end),
    rows: rows,
    row_count: rows.length
  };
}

function getYearlyPaidLeaveReportPreview(filters) {
  filters = filters || {};

  const fiscalYear = Number(filters.fiscal_year || getFiscalYearFromDate(new Date()));
  const companyCodeFilter = String(filters.company_code || "").trim().toUpperCase();
  const companyNameFilter = String(filters.company_name || "").trim();

  const employees = getEmployeesForAdmin().filter(emp => {
    if (String(emp.employment_status || "").trim().toLowerCase() !== "active") return false;
    if (emp.leave_management_target !== true) return false;

    const empCompanyCode = String(emp.company_code || "").trim().toUpperCase();
    const empCompanyName = String(emp.company_name || "").trim();

    if (companyCodeFilter && empCompanyCode !== companyCodeFilter) return false;
    if (companyNameFilter && empCompanyName !== companyNameFilter) return false;

    return true;
  });

  const fiscalStartMonth =
    employees.length > 0
      ? Number(employees[0].fiscal_start_month || 4)
      : companyCodeFilter === "PARTNER" ? 6 : 4;

  const yearRange = getFiscalYearRangeWithStart(fiscalYear, fiscalStartMonth);

  const grantMap = getGrantMapByFiscalYear(fiscalYear);
  const usedMap = getApprovedUsedDaysByFiscalYear(fiscalYear);

  const rows = employees.map(emp => {
    const grantInfo = grantMap[emp.employee_id] || {
      employee_id: emp.employee_id,
      grant_days: 0,
      carry_over_days: 0
    };

    const balance = buildBalance(
      emp.employee_id,
      grantInfo,
      usedMap[emp.employee_id] || 0
    );

    return {
      employee_id: emp.employee_id,
      display_employee_id: emp.display_employee_id || "",
      employee_name: getDisplayName(emp) || emp.employee_id,
      company_code: emp.company_code || "",
      company_name: emp.company_name || "",
      carry_over_days: balance.carry_over_days,
      grant_days: balance.grant_days,
      used_days: balance.used_days,
      next_carry_over_days: balance.next_carry_over_days,
      expired_days: balance.expired_days
    };
  }).sort((a, b) => {
    return String(a.employee_id).localeCompare(String(b.employee_id));
  });

  return {
    ok: true,
    fiscal_year: fiscalYear,
    period_start: formatDateValue(yearRange.start),
    period_end: formatDateValue(yearRange.end),
    company_code: companyCodeFilter || "ALL",
    company_name: companyNameFilter || "",
    rows: rows,
    row_count: rows.length
  };
}

/* =========================
   管理者ログイン：ユーザー一覧取得
========================= */
function getAdminUsersForLogin() {
  const sheet = getSheet("admin_users");
  const headerInfo = requireHeaders(sheet, [
    "admin_id",
    "admin_name",
    "pin",
    "is_active"
  ]);

  const data = sheet.getDataRange().getValues();

  if (data.length <= 1) return [];

  return data.slice(1)
    .map(row => rowToObject(row, headerInfo.headers))
    .filter(rowObj => {
      return String(rowObj.is_active || "").trim().toUpperCase() === "TRUE";
    })
    .map(rowObj => {
      return {
        admin_id: String(rowObj.admin_id || "").trim(),
        admin_name: String(rowObj.admin_name || "").trim()
      };
    })
    .filter(user => user.admin_id && user.admin_name);
}

/* =========================
   管理者ログイン：PIN確認
========================= */
function verifyAdminLogin(adminId, pin) {
  const sheet = getSheet("admin_users");

  const headerInfo = requireHeaders(sheet, [
    "admin_id",
    "admin_name",
    "pin",
    "is_active"
  ]);

  const data = sheet.getDataRange().getValues();

  const targetAdminId = String(adminId || "").trim();
  const targetPin = String(pin || "").trim();

  if (!targetAdminId) {
    throw new Error("管理者を選択してください");
  }

  if (!targetPin) {
    throw new Error("PINを入力してください");
  }

  const matched = data.slice(1)
    .map(row => rowToObject(row, headerInfo.headers))
    .find(rowObj => {
      return (
        String(rowObj.admin_id || "").trim() === targetAdminId &&
        String(rowObj.is_active || "").trim().toUpperCase() === "TRUE"
      );
    });

  if (!matched) {
    throw new Error("管理者が見つかりません");
  }

  if (String(matched.pin || "").trim() !== targetPin) {
    throw new Error("PINが違います");
  }

  return {
    ok: true,
    admin_id: String(matched.admin_id || "").trim(),
    admin_name: String(matched.admin_name || "").trim()
  };
}

/* =========================
   6か月到達者：初回有給付与候補取得
========================= */
function getSixMonthGrantCandidates(options) {
  const today = parseLocalDate(new Date());
  const employees = getEmployeesForAdmin();
  const grantedMap = getSixMonthGrantProcessedMap_();
  const opts = options || null;

  const rows = employees
    .filter(emp => {
      return isInitialPaidLeaveGrantCandidateEmployee_(emp, today, grantedMap);
    })
    .map(emp => {
      const grantInfo = getInitialPaidLeaveGrantInfo_(emp);
      const grantDays = getSixMonthGrantDays_(emp.work_days_per_week);
      const fiscalStartMonth = getInitialPaidLeaveFiscalStartMonth_(emp);

      return {
        employee_id: emp.employee_id,
        display_employee_id: emp.display_employee_id,
        name: getDisplayName(emp) || emp.name,
        hire_date: emp.hire_date,
        six_month_date: formatDateValue(grantInfo.six_month_date),
        company_basis_date: formatDateValue(grantInfo.company_basis_date),
        grant_date: formatDateValue(grantInfo.grant_date),
        grant_reason: grantInfo.grant_reason,
        grant_days: grantDays,
        work_days_per_week: emp.work_days_per_week || "",
        company_code: emp.company_code || "",
        company_name: emp.company_name || "",
        department: emp.department || "",
        fiscal_start_month: fiscalStartMonth,
        initial_grant_check_target: emp.initial_grant_check_target === true,
        preview_only: true
      };
    });

  if (opts) {
    const response = buildPagedResponse_(rows, opts);
    response.dry_run = true;
    response.data_changed = false;
    return response;
  }

  return rows;
}

/* =========================
   6か月到達者：1名付与
========================= */
function grantSixMonthPaidLeave(employeeId, adminUser, options) {
  if (!employeeId) throw new Error("employeeId がありません");

  const employees = getEmployeesForAdmin();
  const emp = employees.find(e => String(e.employee_id) === String(employeeId));

  validateInitialPaidLeaveGrantExecutionTarget_(emp, employeeId);

  const grantInfo = getInitialPaidLeaveGrantInfo_(emp);
  const grantDate = grantInfo.grant_date;
  const systemGrantDays = getSixMonthGrantDays_(emp.work_days_per_week);
  const grantDays = resolveGrantDaysOverride_(options, systemGrantDays);
  const now = new Date();

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

  const rowObj = createEmptyRowObject(headerInfo.headers);

  rowObj.grant_id = getNextGrantId_();
  rowObj.employee_id = employeeId;
  rowObj.grant_date = grantDate;
  rowObj.grant_days = grantDays;
  rowObj.carry_over_days = 0;
  rowObj.valid_from = grantDate;
  rowObj.valid_to = addDaysLocal_(addYearsLocal_(grantDate, 2), -1);
  rowObj.grant_type = "six_month";
  rowObj.year = getFiscalYearFromDateWithStart(
    grantDate,
    getInitialPaidLeaveFiscalStartMonth_(emp)
  );
  const baseNotes = grantInfo.grant_reason === "company_basis"
    ? "会社基準日による初回付与"
    : "入社6か月到達による初回付与";
  rowObj.notes = buildGrantDaysAdjustmentNotes_(baseNotes, systemGrantDays, grantDays);
  rowObj.created_at = now;
  rowObj.updated_at = now;

  appendRowFast_(
  sheet,
  objectToRow(rowObj, headerInfo.headers)
);

  const operatorId = adminUser && adminUser.admin_id ? adminUser.admin_id : "admin";
  const operatorName = adminUser && adminUser.admin_name ? adminUser.admin_name : "管理者";

  appendUsageLog({
    request_id: employeeId,
    action_type: "six_month_grant",
    operator_id: operatorId,
    operator_name: operatorName,
    comment: emp.name + " さんへ " + grantDays + "日を6か月到達付与しました"
  });

  clearAppCache();

  return {
    ok: true,
    employee_id: employeeId,
    name: emp.name,
    grant_date: formatDateValue(grantDate),
    grant_days: grantDays
  };
}

/* =========================
   6か月到達者：処理済みにする
========================= */
function markSixMonthGrantCandidateProcessed(employeeId, reason, adminUser) {
  if (!employeeId) throw new Error("employeeId がありません");

  const employees = getEmployeesForAdmin();
  const emp = employees.find(e => String(e.employee_id) === String(employeeId));

  validateInitialPaidLeaveGrantExecutionTarget_(emp, employeeId);

  const grantInfo = getInitialPaidLeaveGrantInfo_(emp);
  const grantDate = grantInfo.grant_date;
  const now = new Date();
  const note = String(reason || "").trim() ||
    "手動入力済みのため6か月付与チェックを処理済みにした";

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

  const rowObj = createEmptyRowObject(headerInfo.headers);

  rowObj.grant_id = getNextGrantId_();
  rowObj.employee_id = employeeId;
  rowObj.grant_date = grantDate;
  rowObj.grant_days = 0;
  rowObj.carry_over_days = 0;
  rowObj.valid_from = "";
  rowObj.valid_to = "";
  rowObj.grant_type = "six_month_processed";
  rowObj.year = getFiscalYearFromDateWithStart(
    grantDate,
    getInitialPaidLeaveFiscalStartMonth_(emp)
  );
  rowObj.notes = note;
  rowObj.created_at = now;
  rowObj.updated_at = now;

  appendRowFast_(
    sheet,
    objectToRow(rowObj, headerInfo.headers)
  );

  const operatorId = adminUser && adminUser.admin_id ? adminUser.admin_id : "admin";
  const operatorName = adminUser && adminUser.admin_name ? adminUser.admin_name : "管理者";

  appendUsageLog({
    request_id: employeeId,
    action_type: "six_month_processed",
    operator_id: operatorId,
    operator_name: operatorName,
    comment: emp.name + " さんの6か月付与チェックを処理済みにしました: " + note
  });

  clearAppCache();

  return {
    ok: true,
    employee_id: employeeId,
    name: emp.name,
    grant_date: formatDateValue(grantDate),
    grant_type: "six_month_processed"
  };
}

/* =========================
   6か月到達者：選択一括付与
========================= */
function grantSelectedSixMonthPaidLeave(employeeIds, adminUser) {
  return grantSelectedPaidLeave_(
    employeeIds,
    adminUser,
    grantSixMonthPaidLeave
  );
}

/* =========================
   初回付与予定日
========================= */
function getInitialPaidLeaveGrantInfo_(emp) {
  const hireDate = parseLocalDate(emp.hire_date);
  const fiscalStartMonth = getInitialPaidLeaveFiscalStartMonth_(emp);
  const sixMonthDate = addMonthsLocal_(hireDate, 6);
  let companyBasisDate = new Date(
    hireDate.getFullYear(),
    fiscalStartMonth - 1,
    1
  );

  if (companyBasisDate < hireDate) {
    companyBasisDate = new Date(
      hireDate.getFullYear() + 1,
      fiscalStartMonth - 1,
      1
    );
  }

  if (companyBasisDate < sixMonthDate) {
    return {
      grant_date: companyBasisDate,
      six_month_date: sixMonthDate,
      company_basis_date: companyBasisDate,
      grant_reason: "company_basis"
    };
  }

  return {
    grant_date: sixMonthDate,
    six_month_date: sixMonthDate,
    company_basis_date: companyBasisDate,
    grant_reason: "six_month"
  };
}

function getInitialPaidLeaveFiscalStartMonth_(emp) {
  const companyCode = String(emp && emp.company_code || "").trim().toUpperCase();
  if (companyCode === "PARTNER") return 6;
  return 4;
}

function isInitialPaidLeaveGrantCandidateEmployee_(emp, today, grantedMap) {
  if (!emp) return false;
  if (emp.initial_grant_check_target !== true) return false;

  const employeeId = String(emp.employee_id || "").trim();
  if (!employeeId) return false;
  if (grantedMap && grantedMap[employeeId]) return false;

  const status = String(emp.employment_status || "").trim().toLowerCase();
  const isActive = status === "active" || status === "在職";
  if (!isActive) return false;
  if (emp.leave_management_target !== true) return false;
  if (!emp.hire_date) return false;

  const targetDate = today ? parseLocalDate(today) : parseLocalDate(new Date());
  const oneYearDate = addYearsLocal_(parseLocalDate(emp.hire_date), 1);
  if (targetDate >= oneYearDate) return false;

  const grantInfo = getInitialPaidLeaveGrantInfo_(emp);
  return grantInfo.grant_date <= targetDate;
}

function validateInitialPaidLeaveGrantExecutionTarget_(emp, employeeId) {
  if (!emp) throw new Error("対象社員が見つかりません");
  if (String(emp.employee_id || "") !== String(employeeId || "")) {
    throw new Error("対象社員IDが一致しません");
  }

  const grantedMap = getSixMonthGrantProcessedMap_();
  if (grantedMap[employeeId]) {
    throw new Error("この社員の6か月付与チェックはすでに処理済みです");
  }

  if (emp.initial_grant_check_target !== true) {
    throw new Error("この社員は新規登録社員の初回付与チェック対象ではありません");
  }

  const today = parseLocalDate(new Date());
  if (!isInitialPaidLeaveGrantCandidateEmployee_(emp, today, grantedMap)) {
    throw new Error("この社員は初回付与の実行対象ではありません");
  }
}

/* =========================
   6か月付与済みチェック
========================= */
function getGrantedEmployeeMapByGrantType_(grantType) {
  const sheet = getSheet("paid_leave_grants");
  const headerInfo = requireHeaders(sheet, ["employee_id", "grant_type"]);
  const data = sheet.getDataRange().getValues();
  const result = {};
  const grantTypes = Array.isArray(grantType) ? grantType : [grantType];
  const targetTypes = grantTypes.map(type => String(type || "").trim());

  if (data.length <= 1) return result;

  data.slice(1).forEach(row => {
    const rowObj = rowToObject(row, headerInfo.headers);
    const employeeId = String(rowObj.employee_id || "").trim();
    const type = String(rowObj.grant_type || "").trim();

    if (employeeId && targetTypes.includes(type)) {
      result[employeeId] = true;
    }
  });

  return result;
}

function getSixMonthGrantProcessedMap_() {
  return getGrantedEmployeeMapByGrantType_([
    "six_month",
    "six_month_processed",
    "six_month_skipped"
  ]);
}

/* =========================
   6か月付与日数
   週5日以上は10日
   週4日以下は比例付与
========================= */
function getSixMonthGrantDays_(workDaysPerWeek) {
  const days = Number(workDaysPerWeek || 5);

  if (days >= 5) return 10;
  if (days === 4) return 7;
  if (days === 3) return 5;
  if (days === 2) return 3;
  if (days === 1) return 1;

  return 10;
}

/* =========================
   grant_id 自動採番
========================= */
function getNextGrantId_() {
  const sheet = getSheet("paid_leave_grants");
  const headerInfo = requireHeaders(sheet, ["grant_id"]);
  const data = sheet.getDataRange().getValues();

  let max = 0;

  if (data.length > 1) {
    data.slice(1).forEach(row => {
      const id = String(row[headerInfo.map.grant_id] || "").trim();
      const num = Number(id.replace("G", ""));
      if (!isNaN(num) && num > max) max = num;
    });
  }

  return "G" + String(max + 1).padStart(4, "0");
}

/* =========================
   日付加算ヘルパー
========================= */
function addMonthsLocal_(dateValue, months) {
  const date = parseLocalDate(dateValue);
  return new Date(date.getFullYear(), date.getMonth() + Number(months || 0), date.getDate());
}

function addYearsLocal_(dateValue, years) {
  const date = parseLocalDate(dateValue);
  return new Date(date.getFullYear() + Number(years || 0), date.getMonth(), date.getDate());
}

function addDaysLocal_(dateValue, days) {
  const date = parseLocalDate(dateValue);
  date.setDate(date.getDate() + Number(days || 0));
  return date;
}

function getAdminDashboardSummary() {
  const range = getAdminRecentRange();
  const pendingRange = getAdminPendingFocusRange();

  const sheet = getSheet("leave_requests");
  const headerInfo = requireHeaders(sheet, [
    "start_date",
    "end_date",
    "status"
  ]);

  const data = sheet.getDataRange().getValues();

  const result = {
    pending: 0,
    pending_out_of_range: 0,
    approved: 0,
    rejected: 0
  };

  if (data.length <= 1) return result;

  data.slice(1).forEach(row => {
    const rowObj = rowToObject(row, headerInfo.headers);
    const status = norm(rowObj.status);

    if (!rowObj.start_date || !rowObj.end_date) return;

    if (status === STATUS.PENDING) {
      if (isRequestOnOrAfterDate(rowObj, pendingRange.start)) {
        result.pending++;
      } else {
        result.pending_out_of_range++;
      }
      return;
    }

    if (!isRequestInDateRange(rowObj, range.start, range.end)) return;

    if (status === STATUS.APPROVED) result.approved++;
    if (status === STATUS.REJECTED) result.rejected++;
  });

  return result;
}

function getPaidLeaveDashboardData(filters) {
  const opts = filters || {};
  const fiscalYear = Number(opts.fiscal_year || getFiscalYearFromDate(new Date()));
  const companyCodeFilter = String(opts.company_code || "").trim().toUpperCase();
  const keyword = norm(opts.keyword || "");
  const fiveDayIncompleteOnly = opts.five_day_incomplete_only === true;
  const expiredOnly = opts.expired_only === true;
  const asOfDate = opts.as_of_date ? parseLocalDate(opts.as_of_date) : parseLocalDate(new Date());

  const employees = getEmployeesForAdmin()
    .filter(emp => isFifoBalanceCompareTargetEmployee_(emp))
    .filter(emp => {
      const empCompanyCode = String(emp.company_code || "").trim().toUpperCase();
      if (companyCodeFilter && companyCodeFilter !== "ALL" && empCompanyCode !== companyCodeFilter) {
        return false;
      }

      if (!keyword) return true;

      const targetText = norm(
        String(emp.employee_id || "") +
        String(emp.display_employee_id || "") +
        String(emp.name || "") +
        String(emp.display_name || "") +
        String(emp.name_kana || "") +
        String(emp.company_name || "") +
        String(emp.company_code || "")
      );

      return targetText.indexOf(keyword) !== -1;
    })
    .sort((a, b) => String(a.employee_id || "").localeCompare(String(b.employee_id || "")));

  const employeeIds = employees.map(emp => String(emp.employee_id || "").trim()).filter(Boolean);
  const fiscalBalanceMap = getEmployeeBalanceMapForEmployeeIdsForFiscalYear(
    fiscalYear,
    employeeIds
  );
  const context = createFifoBalanceComparisonContext_(asOfDate);

  const allRows = employees.map(emp => {
    const employeeId = String(emp.employee_id || "").trim();
    const fifoBalance = calculateFifoBalanceWithOpeningBalanceFromContext_(
      employeeId,
      asOfDate,
      context
    );
    const fiscalBalance = fiscalBalanceMap[employeeId] || {
      used_days: 0
    };
    const usedDays = Number(fiscalBalance.used_days || 0);
    const fiveDayRemaining = Math.max(0, 5 - usedDays);
    const expiryInfo = buildPaidLeaveDashboardExpiryInfo_(fifoBalance, asOfDate);

    return {
      employee_id: employeeId,
      display_employee_id: String(emp.display_employee_id || ""),
      name: String(emp.name || ""),
      display_name: String(emp.display_name || ""),
      employee_name: getDisplayName(emp) || employeeId,
      company_code: String(emp.company_code || ""),
      company_name: String(emp.company_name || ""),
      current_remaining_days: Number(fifoBalance.current_remaining_days || 0),
      fiscal_used_days: usedDays,
      five_day_used: Math.min(usedDays, 5),
      five_day_remaining: fiveDayRemaining,
      five_day_completed: fiveDayRemaining === 0,
      expired_days: Number(fifoBalance.expired_days || 0),
      nearest_expiry_date: expiryInfo.nearest_expiry_date,
      nearest_expiry_days: expiryInfo.nearest_expiry_days,
      expiry_status: expiryInfo.expiry_status,
      expiry_status_label: expiryInfo.expiry_status_label
    };
  });

  const summary = {
    target_employee_count: allRows.length,
    five_day_incomplete_count: allRows.filter(row => !row.five_day_completed).length,
    within_90_count: allRows.filter(row =>
      row.expiry_status === "within_90" ||
      row.expiry_status === "within_30"
    ).length,
    within_30_count: allRows.filter(row =>
      row.expiry_status === "within_30"
    ).length,
    expired_count: allRows.filter(row => Number(row.expired_days || 0) > 0).length
  };

  const rows = allRows.filter(row => {
    if (fiveDayIncompleteOnly && row.five_day_completed) return false;
    if (expiredOnly && Number(row.expired_days || 0) <= 0) return false;
    return true;
  });

  return {
    ok: true,
    fiscal_year: fiscalYear,
    as_of_date: formatDateValue(asOfDate),
    company_code: companyCodeFilter || "ALL",
    keyword: String(opts.keyword || ""),
    summary: summary,
    row_count: rows.length,
    rows: rows
  };
}


function getPaidLeaveBalanceSnapshotForAttendance(employeeIds, asOfDate) {
  const targetEmployeeIds = normalizeAttendanceSnapshotEmployeeIds_(employeeIds);
  const targetDate = asOfDate ? parseLocalDate(asOfDate) : parseLocalDate(new Date());
  const asOfDateKey = toDateKey(targetDate);

  if (targetEmployeeIds.length === 0) {
    return {
      ok: true,
      as_of_date: asOfDateKey,
      calculation_mode: "fifo_with_opening_balance",
      employees: []
    };
  }

  let employeesForAdmin = [];
  let context = null;
  let setupError = "";

  try {
    employeesForAdmin = getEmployeesForAdmin();
    context = createFifoBalanceComparisonContext_(targetDate);
  } catch (error) {
    setupError = error && error.message ? error.message : String(error || "");
  }

  const employeeMap = {};
  employeesForAdmin.forEach(emp => {
    const employeeId = String(emp.employee_id || "").trim();
    if (employeeId) employeeMap[employeeId] = emp;
  });

  const fiscalYearGroups = {};
  targetEmployeeIds.forEach(employeeId => {
    const emp = employeeMap[employeeId] || {};
    const fiscalStartMonth = Number(emp.fiscal_start_month || 4);
    const fiscalYear = getFiscalYearFromDateWithStart(targetDate, fiscalStartMonth);

    if (!fiscalYearGroups[fiscalYear]) fiscalYearGroups[fiscalYear] = [];
    fiscalYearGroups[fiscalYear].push(employeeId);
  });

  const fiscalBalanceByEmployee = {};
  Object.keys(fiscalYearGroups).forEach(fiscalYear => {
    try {
      const balanceMap = getEmployeeBalanceMapForEmployeeIdsForFiscalYear(
        Number(fiscalYear),
        fiscalYearGroups[fiscalYear]
      );

      fiscalYearGroups[fiscalYear].forEach(employeeId => {
        fiscalBalanceByEmployee[employeeId] = balanceMap[employeeId] || {
          employee_id: employeeId,
          current_remaining_days: 0,
          carry_over_days: 0,
          grant_days: 0,
          used_days: 0,
          next_carry_over_days: 0,
          expired_days: 0
        };
      });
    } catch (error) {
      fiscalYearGroups[fiscalYear].forEach(employeeId => {
        fiscalBalanceByEmployee[employeeId] = {
          employee_id: employeeId,
          current_remaining_days: null,
          carry_over_days: null,
          grant_days: null,
          used_days: null,
          next_carry_over_days: null,
          expired_days: null,
          error: error && error.message ? error.message : String(error || "")
        };
      });
    }
  });

  const rows = targetEmployeeIds.map(employeeId => {
    const emp = employeeMap[employeeId] || {};
    const fiscalStartMonth = Number(emp.fiscal_start_month || 4);
    const fiscalYear = getFiscalYearFromDateWithStart(targetDate, fiscalStartMonth);
    const fiscalBalance = fiscalBalanceByEmployee[employeeId] || {};

    if (setupError || !context) {
      return {
        employee_id: employeeId,
        employee_name: getDisplayName(emp) || employeeId,
        fiscal_year: String(fiscalYear),
        current_remaining_days: null,
        carry_over_days: fiscalBalance.carry_over_days == null ? null : Number(fiscalBalance.carry_over_days || 0),
        grant_days: fiscalBalance.grant_days == null ? null : Number(fiscalBalance.grant_days || 0),
        used_days: fiscalBalance.used_days == null ? null : Number(fiscalBalance.used_days || 0),
        expiring_soon_days: null,
        expired_days: null,
        expiry_lots: [],
        error: setupError || "有給残数計算の初期化に失敗しました"
      };
    }

    try {
      const fifoBalance = calculateFifoBalanceWithOpeningBalanceFromContext_(
        employeeId,
        targetDate,
        context
      );
      const expiryLots = buildAttendancePaidLeaveExpiryLots_(fifoBalance, targetDate);
      const expiringSoonDays = expiryLots
        .filter(lot => lot.days_until_expiry >= 0 && lot.days_until_expiry <= 90)
        .reduce((sum, lot) => sum + Number(lot.remaining_days || 0), 0);

      return {
        employee_id: employeeId,
        employee_name: getDisplayName(emp) || employeeId,
        fiscal_year: String(fiscalYear),
        current_remaining_days: Number(fifoBalance.current_remaining_days || 0),
        carry_over_days: Number(fiscalBalance.carry_over_days || 0),
        grant_days: Number(fiscalBalance.grant_days || 0),
        used_days: Number(fiscalBalance.used_days || 0),
        expiring_soon_days: expiringSoonDays,
        expired_days: Number(fifoBalance.expired_days || 0),
        expiry_lots: expiryLots.map(lot => ({
          expire_date: lot.expire_date,
          remaining_days: lot.remaining_days
        }))
      };
    } catch (error) {
      return {
        employee_id: employeeId,
        employee_name: getDisplayName(emp) || employeeId,
        fiscal_year: String(fiscalYear),
        current_remaining_days: null,
        carry_over_days: fiscalBalance.carry_over_days == null ? null : Number(fiscalBalance.carry_over_days || 0),
        grant_days: fiscalBalance.grant_days == null ? null : Number(fiscalBalance.grant_days || 0),
        used_days: fiscalBalance.used_days == null ? null : Number(fiscalBalance.used_days || 0),
        expiring_soon_days: null,
        expired_days: null,
        expiry_lots: [],
        error: error && error.message ? error.message : String(error || "")
      };
    }
  });

  return {
    ok: true,
    as_of_date: asOfDateKey,
    calculation_mode: "fifo_with_opening_balance",
    employees: rows
  };
}

function normalizeAttendanceSnapshotEmployeeIds_(employeeIds) {
  const source = Array.isArray(employeeIds)
    ? employeeIds
    : String(employeeIds || "").split(",");
  const seen = {};

  return source
    .map(id => String(id || "").trim())
    .filter(id => {
      if (!id || seen[id]) return false;
      seen[id] = true;
      return true;
    });
}

function buildAttendancePaidLeaveExpiryLots_(fifoBalance, asOfDate) {
  return (fifoBalance.grant_details || [])
    .map(lot => {
      const remainingDays = Number(lot.active_remaining_days || 0);
      if (remainingDays <= 0) return null;

      const validTo = parseLocalDate(lot.valid_to);
      const daysUntilExpiry = Math.round(
        (validTo.getTime() - asOfDate.getTime()) / (24 * 60 * 60 * 1000)
      );

      if (daysUntilExpiry < 0) return null;

      return {
        expire_date: toDateKey(validTo),
        remaining_days: remainingDays,
        days_until_expiry: daysUntilExpiry
      };
    })
    .filter(Boolean)
    .sort((a, b) => {
      if (a.days_until_expiry !== b.days_until_expiry) {
        return a.days_until_expiry - b.days_until_expiry;
      }
      return String(a.expire_date).localeCompare(String(b.expire_date));
    });
}


function getPaidLeaveDashboardEmployeeDetail(employeeId, filters) {
  const targetEmployeeId = String(employeeId || "").trim();
  if (!targetEmployeeId) throw new Error("社員IDがありません。");

  const opts = filters || {};
  const fiscalYear = Number(opts.fiscal_year || getFiscalYearFromDate(new Date()));
  const asOfDate = opts.as_of_date ? parseLocalDate(opts.as_of_date) : parseLocalDate(new Date());
  const emp = getEmployeesForAdmin().find(row =>
    String(row.employee_id || "").trim() === targetEmployeeId
  );

  if (!emp || !isFifoBalanceCompareTargetEmployee_(emp)) {
    throw new Error("対象社員が見つかりません。");
  }

  const context = createFifoBalanceComparisonContext_(asOfDate);
  const fifoBalance = calculateFifoBalanceWithOpeningBalanceFromContext_(
    targetEmployeeId,
    asOfDate,
    context
  );
  const fiscalBalanceMap = getEmployeeBalanceMapForEmployeeIdsForFiscalYear(
    fiscalYear,
    [targetEmployeeId]
  );
  const fiscalBalance = fiscalBalanceMap[targetEmployeeId] || {
    used_days: 0
  };
  const usedDays = Number(fiscalBalance.used_days || 0);
  const fiveDayRemaining = Math.max(0, 5 - usedDays);
  const expiryInfo = buildPaidLeaveDashboardExpiryInfo_(fifoBalance, asOfDate);
  const grantDetails = (fifoBalance.grant_details || [])
    .map(lot => buildPaidLeaveDashboardGrantDetail_(lot, asOfDate))
    .filter(Boolean)
    .sort((a, b) => {
      const priority = {
        expired: 1,
        within_30: 2,
        within_90: 3,
        normal: 4
      };
      const priorityDiff = (priority[a.status] || 9) - (priority[b.status] || 9);
      if (priorityDiff !== 0) return priorityDiff;
      if (a.days_until_expiry !== b.days_until_expiry) {
        return a.days_until_expiry - b.days_until_expiry;
      }
      if (a.grant_date === b.grant_date) return 0;
      return a.grant_date < b.grant_date ? -1 : 1;
    });

  return {
    ok: true,
    fiscal_year: fiscalYear,
    as_of_date: formatDateValue(asOfDate),
    employee: {
      employee_id: targetEmployeeId,
      display_employee_id: String(emp.display_employee_id || ""),
      name: String(emp.name || ""),
      display_name: String(emp.display_name || ""),
      employee_name: getDisplayName(emp) || targetEmployeeId,
      company_code: String(emp.company_code || ""),
      company_name: String(emp.company_name || "")
    },
    summary: {
      current_remaining_days: Number(fifoBalance.current_remaining_days || 0),
      fiscal_used_days: usedDays,
      five_day_used: Math.min(usedDays, 5),
      five_day_remaining: fiveDayRemaining,
      five_day_completed: fiveDayRemaining === 0,
      expired_days: Number(fifoBalance.expired_days || 0),
      nearest_expiry_date: expiryInfo.nearest_expiry_date,
      nearest_expiry_days: expiryInfo.nearest_expiry_days,
      expiry_status: expiryInfo.expiry_status,
      expiry_status_label: expiryInfo.expiry_status_label
    },
    grant_details: grantDetails
  };
}

function buildPaidLeaveDashboardGrantDetail_(lot, asOfDate) {
  const remainingDays = lot && lot.is_expired
    ? Number(lot.expired_days || 0)
    : Number(lot && lot.active_remaining_days || 0);
  if (remainingDays <= 0) return null;

  const validTo = parseLocalDate(lot.valid_to);
  const daysUntilExpiry = Math.round(
    (validTo.getTime() - asOfDate.getTime()) / (24 * 60 * 60 * 1000)
  );
  let status = "normal";
  let statusLabel = "通常";

  if (lot.is_expired || daysUntilExpiry < 0) {
    status = "expired";
    statusLabel = "期限切れ";
  } else if (daysUntilExpiry <= 30) {
    status = "within_30";
    statusLabel = "期限が近い（30日以内）";
  } else if (daysUntilExpiry <= 90) {
    status = "within_90";
    statusLabel = "期限が近い（90日以内）";
  }

  return {
    grant_date: String(lot.grant_date || ""),
    grant_days: Number(lot.total_days || 0),
    remaining_days: remainingDays,
    valid_to: formatDateValue(validTo),
    status: status,
    status_label: statusLabel,
    days_until_expiry: daysUntilExpiry
  };
}

function buildPaidLeaveDashboardExpiryInfo_(fifoBalance, asOfDate) {
  const activeLots = (fifoBalance.grant_details || [])
    .map(lot => {
      const remainingDays = Number(lot.active_remaining_days || 0);
      if (remainingDays <= 0) return null;

      const validTo = parseLocalDate(lot.valid_to);
      const daysUntilExpiry = Math.round(
        (validTo.getTime() - asOfDate.getTime()) / (24 * 60 * 60 * 1000)
      );

      return {
        valid_to: formatDateValue(validTo),
        days_until_expiry: daysUntilExpiry
      };
    })
    .filter(Boolean)
    .sort((a, b) => a.days_until_expiry - b.days_until_expiry);

  if (Number(fifoBalance.expired_days || 0) > 0) {
    return {
      nearest_expiry_date: activeLots.length > 0 ? activeLots[0].valid_to : "",
      nearest_expiry_days: activeLots.length > 0 ? activeLots[0].days_until_expiry : "",
      expiry_status: "expired",
      expiry_status_label: "期限切れ"
    };
  }

  if (activeLots.length === 0) {
    return {
      nearest_expiry_date: "",
      nearest_expiry_days: "",
      expiry_status: "normal",
      expiry_status_label: "正常"
    };
  }

  const nearest = activeLots[0];
  if (nearest.days_until_expiry <= 30) {
    return {
      nearest_expiry_date: nearest.valid_to,
      nearest_expiry_days: nearest.days_until_expiry,
      expiry_status: "within_30",
      expiry_status_label: "30日以内"
    };
  }

  if (nearest.days_until_expiry <= 90) {
    return {
      nearest_expiry_date: nearest.valid_to,
      nearest_expiry_days: nearest.days_until_expiry,
      expiry_status: "within_90",
      expiry_status_label: "90日以内"
    };
  }

  return {
    nearest_expiry_date: nearest.valid_to,
    nearest_expiry_days: nearest.days_until_expiry,
    expiry_status: "normal",
    expiry_status_label: "正常"
  };
}

/* =========================
   年次付与候補取得
========================= */
function getYearlyGrantCandidates(options) {
  const today = new Date();
  const employees = getEmployeesForAdmin();
  const opts = options || null;

  const rows = employees
    .filter(emp => {
      if (String(emp.employment_status || "").toLowerCase() !== "active") return false;
      if (emp.leave_management_target !== true) return false;
      if (!emp.hire_date) return false;

      const fiscalStartMonth = Number(emp.fiscal_start_month || 4);
      const fiscalYear = getFiscalYearFromDateWithStart(today, fiscalStartMonth);
      const basisDate = getFiscalYearRangeWithStart(fiscalYear, fiscalStartMonth).start;

      // 基準日前ならまだ表示しない
      if (today < basisDate) return false;

      const months = getMonthsWorked_(parseLocalDate(emp.hire_date), basisDate);

      // 年次付与は1年6か月以上から
      if (months < 18) return false;

      if (hasYearlyGrantForFiscalYear_(emp.employee_id, fiscalYear)) return false;

      return true;
    })
    .map(emp => {
      const fiscalStartMonth = Number(emp.fiscal_start_month || 4);
      const fiscalYear = getFiscalYearFromDateWithStart(today, fiscalStartMonth);
      const basisDate = getFiscalYearRangeWithStart(fiscalYear, fiscalStartMonth).start;
      const months = getMonthsWorked_(parseLocalDate(emp.hire_date), basisDate);
      const grantDays = getYearlyGrantDays_(months);

      return {
        employee_id: emp.employee_id,
        display_employee_id: emp.display_employee_id,
        name: getDisplayName(emp) || emp.name,
        hire_date: emp.hire_date,
        basis_date: formatDateValue(basisDate),
        months_worked: months,
        grant_days: grantDays,
        fiscal_year: fiscalYear,
        company_code: emp.company_code,
        company_name: emp.company_name,
        department: emp.department || "",
        fiscal_start_month: fiscalStartMonth
      };
    });

  return opts ? buildPagedResponse_(rows, opts) : rows;
}

/* =========================
   年次付与実行
========================= */
function grantYearlyPaidLeave(employeeId, adminUser, options) {
  if (!employeeId) throw new Error("employeeId がありません");

  const employees = getEmployeesForAdmin();
  const emp = employees.find(e => String(e.employee_id) === String(employeeId));

  if (!emp) throw new Error("対象社員が見つかりません");
  if (!emp.hire_date) throw new Error("入社日がありません");

  const today = new Date();
  const fiscalStartMonth = Number(emp.fiscal_start_month || 4);
  const fiscalYear = getFiscalYearFromDateWithStart(today, fiscalStartMonth);
  const basisDate = getFiscalYearRangeWithStart(fiscalYear, fiscalStartMonth).start;

  const months = getMonthsWorked_(parseLocalDate(emp.hire_date), basisDate);

  if (months < 18) {
    throw new Error("年次付与対象ではありません");
  }

  if (hasYearlyGrantForFiscalYear_(employeeId, fiscalYear)) {
    throw new Error("この年度はすでに年次付与済みです");
  }

  const systemGrantDays = getYearlyGrantDays_(months);
  const grantDays = resolveGrantDaysOverride_(options, systemGrantDays);
  const now = new Date();

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

  const rowObj = createEmptyRowObject(headerInfo.headers);

  rowObj.grant_id = getNextGrantId_();
  rowObj.employee_id = employeeId;
  rowObj.grant_date = basisDate;
  rowObj.grant_days = grantDays;
  rowObj.carry_over_days = 0;
  rowObj.valid_from = basisDate;
  rowObj.valid_to = addDaysLocal_(addYearsLocal_(basisDate, 2), -1);
  rowObj.grant_type = "yearly";
  rowObj.year = fiscalYear;
  rowObj.notes = buildGrantDaysAdjustmentNotes_("年次有給付与", systemGrantDays, grantDays);
  rowObj.created_at = now;
  rowObj.updated_at = now;

  appendRowFast_(
  sheet,
  objectToRow(rowObj, headerInfo.headers)
);

  const operatorId = adminUser && adminUser.admin_id ? adminUser.admin_id : "admin";
  const operatorName = adminUser && adminUser.admin_name ? adminUser.admin_name : "管理者";

  appendUsageLog({
    request_id: employeeId,
    action_type: "yearly_grant",
    operator_id: operatorId,
    operator_name: operatorName,
    comment: emp.name + " さんへ " + grantDays + "日を年次付与しました"
  });

  clearAppCache();

  return { ok: true };
}

/* =========================
   年次付与：選択一括付与
========================= */
function grantSelectedYearlyPaidLeave(employeeIds, adminUser) {
  return grantSelectedPaidLeave_(
    employeeIds,
    adminUser,
    grantYearlyPaidLeave
  );
}

function grantSelectedPaidLeave_(employeeIds, adminUser, grantFn) {
  const items = (employeeIds || [])
    .map(parseSelectedGrantItem_)
    .filter(item => item.employee_id);
  const result = {
    ok: true,
    total_count: items.length,
    success_count: 0,
    skipped_count: 0,
    error_count: 0,
    results: []
  };

  items.forEach(item => {
    const employeeId = item.employee_id;

    try {
      const res = grantFn(employeeId, adminUser, item.options);
      result.success_count++;
      result.results.push({
        employee_id: employeeId,
        status: "success",
        message: "付与しました",
        detail: res || null
      });
    } catch (e) {
      const message = e && e.message ? e.message : String(e);
      const isSkipped =
        message.indexOf("すでに") !== -1 ||
        message.indexOf("処理済み") !== -1 ||
        message.indexOf("付与済み") !== -1;

      if (isSkipped) {
        result.skipped_count++;
        result.results.push({
          employee_id: employeeId,
          status: "skipped",
          message: message
        });
      } else {
        result.error_count++;
        result.results.push({
          employee_id: employeeId,
          status: "error",
          message: message
        });
      }
    }
  });

  return result;
}

function parseSelectedGrantItem_(item) {
  if (item && typeof item === "object") {
    return {
      employee_id: String(item.employee_id || "").trim(),
      options: {
        grant_days_override: item.grant_days,
        original_grant_days: item.original_grant_days,
        manual_note: item.manual_note || "手入力調整"
      }
    };
  }

  return {
    employee_id: String(item || "").trim(),
    options: {}
  };
}

function resolveGrantDaysOverride_(options, systemGrantDays) {
  const opts = options || {};
  const rawValue = opts.grant_days_override;

  if (rawValue === "" || rawValue === null || rawValue === undefined) {
    return Number(systemGrantDays || 0);
  }

  const grantDays = Number(rawValue);

  if (!isFinite(grantDays)) {
    throw new Error("付与日数は数値で入力してください");
  }

  if (grantDays <= 0) {
    throw new Error("付与日数は0日より大きい値を入力してください");
  }

  if (grantDays > 20) {
    throw new Error("付与日数は20日以下で入力してください");
  }

  if (Math.abs(grantDays * 2 - Math.round(grantDays * 2)) > 0.000001) {
    throw new Error("付与日数は0.5日単位で入力してください");
  }

  return grantDays;
}

function buildGrantDaysAdjustmentNotes_(baseNotes, systemGrantDays, grantDays) {
  const systemDays = Number(systemGrantDays || 0);
  const actualDays = Number(grantDays || 0);

  if (Math.abs(systemDays - actualDays) < 0.000001) {
    return baseNotes;
  }

  return baseNotes +
    " / 手入力調整: システム計算 " +
    formatGrantDaysForNote_(systemDays) +
    "日 → 手入力 " +
    formatGrantDaysForNote_(actualDays) +
    "日";
}

function formatGrantDaysForNote_(value) {
  const num = Number(value || 0);
  return Number.isInteger(num) ? String(num) : String(num);
}

/* =========================
   勤続月数計算
========================= */
function getMonthsWorked_(startDate, endDate) {
  const start = parseLocalDate(startDate);
  const end = parseLocalDate(endDate);

  let months =
    (end.getFullYear() - start.getFullYear()) * 12 +
    (end.getMonth() - start.getMonth());

  if (end.getDate() < start.getDate()) {
    months--;
  }

  return months;
}

/* =========================
   年次付与日数
========================= */
function getYearlyGrantDays_(monthsWorked) {
  if (monthsWorked >= 78) return 20; // 6年6か月以上
  if (monthsWorked >= 66) return 18; // 5年6か月
  if (monthsWorked >= 54) return 16; // 4年6か月
  if (monthsWorked >= 42) return 14; // 3年6か月
  if (monthsWorked >= 30) return 12; // 2年6か月
  if (monthsWorked >= 18) return 11; // 1年6か月
  return 0;
}

/* =========================
   同年度付与済みチェック
========================= */
function hasYearlyGrantForFiscalYear_(
  employeeId,
  fiscalYear
) {
  const sheet = getSheet("paid_leave_grants");

  const headerInfo = requireHeaders(sheet, [
    "employee_id",
    "grant_type",
    "year"
  ]);

  const data = sheet.getDataRange().getValues();

  if (data.length <= 1) {
    return false;
  }

  return data.slice(1).some(row => {
    const rowObj = rowToObject(
      row,
      headerInfo.headers
    );

    return (
      String(rowObj.employee_id) ===
        String(employeeId) &&
      String(rowObj.grant_type) ===
        "yearly" &&
      Number(rowObj.year) ===
        Number(fiscalYear)
    );
  });
}
