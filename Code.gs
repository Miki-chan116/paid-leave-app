const SS_ID = "1o7KbHHsPMiL684YJq_Fpzg6gHjD6HebjeL0BhQImkt4";
const OUTPUT_SS_ID = "1SP7kD0wuxKQAwJ5YrMBGj3HAkMHW2U6Rdwzfhlu_Z_E";

/* =========================
   ステータス定義
========================= */
const STATUS = {
  PENDING: "pending",
  APPROVED: "approved",
  REJECTED: "rejected",
  CANCELED: "canceled",
  CANCELED_BY_ADMIN: "canceled_by_admin"
};

// leave_requests.status で許可する正式な値。Spreadsheet の入力規則と
// アプリケーションコードで同じ定義を使用する。
const LEAVE_REQUEST_STATUS_VALUES = [
  STATUS.PENDING,
  STATUS.APPROVED,
  STATUS.REJECTED,
  STATUS.CANCELED,
  STATUS.CANCELED_BY_ADMIN
];

/* =========================
   カレンダー種別
========================= */
const CALENDAR_TYPE = {
  WORKDAY: "workday",
  HOLIDAY: "holiday",
  NO_LEAVE: "no_leave"
};

/* =========================
   会社別有給制度設定

   現時点では静的設定。呼出側はこの関数だけを経由し、将来は
   company_leave_policies シート等の読取実装へ差し替える。
========================= */
function getCompanyLeavePolicy(companyCode) {
  const code = String(companyCode || "").trim().toUpperCase();
  const policies = {
    MAIN: {
      companyCode: "MAIN",
      fiscalStartMonth: 4,
      timeLeaveEnabled: true,
      scheduledMinutesPerDay: 420,
      timeLeaveUnitMinutes: 60,
      timeLeaveAnnualLimitMinutes: 2100,
      workStartMinute: 480,
      workEndMinute: 1020,
      breakPeriods: [
        { startMinute: 600, endMinute: 630 },
        { startMinute: 720, endMinute: 780 },
        { startMinute: 900, endMinute: 930 }
      ],
      policyVersion: "v1"
    },
    PARTNER: {
      companyCode: "PARTNER",
      fiscalStartMonth: 6,
      timeLeaveEnabled: false,
      scheduledMinutesPerDay: null,
      timeLeaveUnitMinutes: null,
      timeLeaveAnnualLimitMinutes: null,
      workStartMinute: null,
      workEndMinute: null,
      breakPeriods: [],
      policyVersion: "v1"
    }
  };

  if (!policies[code]) {
    throw new Error("有給制度設定がない会社コードです: " + code);
  }

  // 呼出側が配列・オブジェクトを変更しても静的定義を壊さないよう複製して返す。
  return Object.assign({}, policies[code], {
    breakPeriods: policies[code].breakPeriods.map(period => Object.assign({}, period))
  });
}

const EMPLOYEE_TIME_LEAVE_WORK_SCHEDULE_HEADERS = [
  "work_start_minute", "work_end_minute"
];

// 実処理。通常の申請・候補取得ではemployeesへ書き込まない。
function initializeEmployeeTimeLeaveWorkScheduleStructure_() {
  const sheet = getSheet("employees");
  const addedHeaders = [];
  EMPLOYEE_TIME_LEAVE_WORK_SCHEDULE_HEADERS.forEach(header => {
    const headerInfo = getHeaderMap(sheet);
    if (!(header in headerInfo.map)) {
      ensureSheetColumn_(sheet, header);
      addedHeaders.push(header);
    }
  });
  return {
    ok: true,
    employees_headers: getHeaderMap(sheet).headers,
    added_headers: addedHeaders,
    verified_headers: EMPLOYEE_TIME_LEAVE_WORK_SCHEDULE_HEADERS.slice()
  };
}

// Apps Scriptエディタから手動実行するための公開入口。
function initializeEmployeeTimeLeaveWorkScheduleStructure() {
  const result = initializeEmployeeTimeLeaveWorkScheduleStructure_();
  console.log(JSON.stringify(result, null, 2));
  return result;
}

function getOptionalEmployeeWorkMinute_(value, label) {
  if (value === "" || value === null || value === undefined) return null;
  const minute = Number(value);
  if (!Number.isInteger(minute) || minute < 0 || minute >= 24 * 60) {
    throw new Error((label || "社員別勤務時刻") + "は0から1439の整数分で指定してください");
  }
  return minute;
}

// 会社標準を基礎とし、employeesの開始・終了が両方設定されている場合だけ社員別勤務へ上書きする。
function resolveEmployeeTimeLeavePolicy_(employeeId, employeeRow) {
  const employee = employeeRow || getTimeLeaveEmployeeFromSpreadsheet_(employeeId);
  if (!employee) throw new Error("対象社員が見つかりません");
  const companyCode = String(employee.company_code || "").trim().toUpperCase();
  const basePolicy = getCompanyLeavePolicy(companyCode);
  const start = getOptionalEmployeeWorkMinute_(employee.work_start_minute, "work_start_minute");
  const end = getOptionalEmployeeWorkMinute_(employee.work_end_minute, "work_end_minute");
  if ((start == null) !== (end == null)) {
    throw new Error("社員別勤務時間は開始・終了を両方設定するか、両方空欄にしてください");
  }
  if (start != null && end <= start) {
    throw new Error("社員別勤務時間の終了は開始より後にしてください");
  }
  const policy = Object.assign({}, basePolicy, {
    workStartMinute: start == null ? basePolicy.workStartMinute : start,
    workEndMinute: end == null ? basePolicy.workEndMinute : end,
    breakPeriods: basePolicy.breakPeriods.map(period => Object.assign({}, period))
  });
  if (start != null && calculateTimeLeaveMinutes(policy.workStartMinute, policy.workEndMinute, policy) !==
      Number(policy.scheduledMinutesPerDay)) {
    throw new Error("社員別勤務時間と共通休憩から求めた実労働時間が所定420分と一致しません");
  }
  return policy;
}

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
  cache.remove(CACHE_KEY.EMPLOYEE_MAP + "_supabase");
  cache.remove(CACHE_KEY.COMPANY_CALENDAR + "_supabase");

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
   時間単位年休: 時刻・分の純粋計算

   時刻は Date を使用せず、HH:mm と 0:00 からの整数分だけで扱う。
========================= */
function parseTimeToMinute(value) {
  const text = String(value == null ? "" : value).trim();
  const match = text.match(/^([01][0-9]|2[0-3]):([0-5][0-9])$/);

  if (!match) {
    throw new Error("時刻は00:00から23:59のHH:mm形式で入力してください: " + text);
  }

  return Number(match[1]) * 60 + Number(match[2]);
}

// Spreadsheetの時刻セルはgetValues()でDateとして返ることがある。
// 業務計算ではDateをString化せず、SpreadsheetのタイムゾーンでHH:mmへ正規化する。
// 文字列は従来のparseTimeToMinuteと同じ厳格なHH:mm形式だけを許可する。
function normalizeTimeLeaveClockValue_(value, timezone) {
  if (value instanceof Date) {
    if (isNaN(value.getTime())) {
      throw new Error("時間年休明細の時刻が不正です: Invalid Date");
    }
    const clock = Utilities.formatDate(value, timezone || getAppTimeZone(), "HH:mm");
    parseTimeToMinute(clock);
    return clock;
  }

  const clock = String(value == null ? "" : value).trim();
  parseTimeToMinute(clock);
  return clock;
}

function assertMinuteRange_(startMinute, endMinute, label) {
  const start = Number(startMinute);
  const end = Number(endMinute);
  const prefix = label ? label + ": " : "";

  if (!Number.isInteger(start) || !Number.isInteger(end)) {
    throw new Error(prefix + "時刻は整数分で指定してください");
  }
  if (start < 0 || start >= 24 * 60 || end < 0 || end >= 24 * 60) {
    throw new Error(prefix + "時刻は0分から1439分の範囲で指定してください");
  }
  if (end <= start) {
    throw new Error(prefix + "終了時刻は開始時刻より後で指定してください");
  }

  return { startMinute: start, endMinute: end };
}

function calculateTimeLeaveMinutes(startMinute, endMinute, policy) {
  const rules = policy || {};
  const range = assertMinuteRange_(startMinute, endMinute, "時間有給");
  const workStart = Number(rules.workStartMinute);
  const workEnd = Number(rules.workEndMinute);

  if (!Number.isInteger(workStart) || !Number.isInteger(workEnd) || workEnd <= workStart) {
    throw new Error("勤務時間の制度設定が不正です");
  }
  if (range.startMinute < workStart || range.endMinute > workEnd) {
    throw new Error("時間有給は所定勤務時間内で指定してください");
  }

  const breakMinutes = (Array.isArray(rules.breakPeriods) ? rules.breakPeriods : [])
    .reduce((sum, period) => {
      const breakStart = Number(period && period.startMinute);
      const breakEnd = Number(period && period.endMinute);
      if (!Number.isInteger(breakStart) || !Number.isInteger(breakEnd) || breakEnd <= breakStart) {
        throw new Error("休憩時間の制度設定が不正です");
      }
      const overlap = Math.max(
        0,
        Math.min(range.endMinute, breakEnd) - Math.max(range.startMinute, breakStart)
      );
      return sum + overlap;
    }, 0);

  return range.endMinute - range.startMinute - breakMinutes;
}

function validateTimeLeaveUnitMinutes(minutes, policy) {
  const value = Number(minutes);
  const unit = Number(policy && policy.timeLeaveUnitMinutes);

  if (!Number.isInteger(value) || value <= 0) {
    throw new Error("時間有給の取得分は0より大きい整数分で指定してください");
  }
  if (!Number.isInteger(unit) || unit <= 0) {
    throw new Error("時間有給単位の制度設定が不正です");
  }
  if (value % unit !== 0) {
    throw new Error("時間有給の取得分は" + unit + "分単位で指定してください");
  }

  return { ok: true, minutes: value, unitMinutes: unit };
}

const TIME_LEAVE_CALCULATION_VERSION_V1 = "time_leave_v1";
const TIME_LEAVE_CALCULATION_VERSION_V2 = "time_leave_v2";
const MAX_STANDALONE_TIME_LEAVE_MINUTES = 180;

// 空欄は既存明細との互換性のため v1 として扱う。未知の版は黙って解釈しない。
function normalizeTimeLeaveCalculationVersion_(value) {
  const version = String(value || "").trim();
  if (!version || version === TIME_LEAVE_CALCULATION_VERSION_V1) return TIME_LEAVE_CALCULATION_VERSION_V1;
  if (version === TIME_LEAVE_CALCULATION_VERSION_V2) return TIME_LEAVE_CALCULATION_VERSION_V2;
  throw new Error("時間有給の計算方式が不正です: " + version);
}

// v1 は休憩控除後の実労働分、v2 は開始から終了までの時計上の経過分を消費分とする。
function calculateTimeLeaveRequestedMinutes_(startMinute, endMinute, policy, calculationVersion) {
  const version = normalizeTimeLeaveCalculationVersion_(calculationVersion);
  if (version === TIME_LEAVE_CALCULATION_VERSION_V1) {
    return calculateTimeLeaveMinutes(startMinute, endMinute, policy);
  }
  const range = assertMinuteRange_(startMinute, endMinute, "時間有給");
  const workStart = Number(policy && policy.workStartMinute);
  const workEnd = Number(policy && policy.workEndMinute);
  if (!Number.isInteger(workStart) || !Number.isInteger(workEnd) || workEnd <= workStart) {
    throw new Error("勤務時間の制度設定が不正です");
  }
  if (range.startMinute < workStart || range.endMinute > workEnd) {
    throw new Error("時間有給は所定勤務時間内で指定してください");
  }
  return range.endMinute - range.startMinute;
}

function hasTimeOverlap(startA, endA, startB, endB) {
  const a = assertMinuteRange_(startA, endA, "時間帯A");
  const b = assertMinuteRange_(startB, endB, "時間帯B");
  return a.startMinute < b.endMinute && b.startMinute < a.endMinute;
}

function validateDailyPaidLeaveMinutes(entries, limitMinutes) {
  const limit = Number(limitMinutes);
  if (!Number.isInteger(limit) || limit <= 0) {
    throw new Error("日次有給上限は0より大きい整数分で指定してください");
  }

  const totalMinutes = (Array.isArray(entries) ? entries : []).reduce((sum, entry) => {
    const minutes = Number(entry && entry.minutes);
    if (!Number.isInteger(minutes) || minutes < 0) {
      throw new Error("日次有給の各取得分は0以上の整数分で指定してください");
    }
    return sum + minutes;
  }, 0);

  if (totalMinutes > limit) {
    throw new Error("同日の有給消化は" + limit + "分を超えられません: " + totalMinutes + "分");
  }

  return {
    ok: true,
    totalMinutes: totalMinutes,
    limitMinutes: limit,
    remainingMinutes: limit - totalMinutes
  };
}

function validateAnnualTimeLeaveLimit(approvedMinutes, pendingMinutes, requestedMinutes, annualLimitMinutes) {
  const approved = Number(approvedMinutes);
  const pending = Number(pendingMinutes);
  const requested = Number(requestedMinutes);
  const limit = Number(annualLimitMinutes);
  const values = [approved, pending, requested, limit];

  if (!values.every(value => Number.isInteger(value)) || approved < 0 || pending < 0 || requested <= 0 || limit <= 0) {
    throw new Error("年間時間有給上限の分数は正しい整数分で指定してください");
  }

  const totalMinutes = approved + pending + requested;
  if (totalMinutes > limit) {
    throw new Error("年間時間有給上限" + limit + "分を超えます: " + totalMinutes + "分");
  }

  return {
    ok: true,
    approvedMinutes: approved,
    pendingMinutes: pending,
    requestedMinutes: requested,
    totalMinutes: totalMinutes,
    annualLimitMinutes: limit,
    remainingMinutes: limit - totalMinutes
  };
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
   leave_requests ステータス入力規則更新
========================= */
function updateLeaveRequestStatusValidation() {
  const sheet = getSheet("leave_requests");
  const headerInfo = requireHeaders(sheet, ["status"]);
  const statusColumn = headerInfo.map.status + 1;
  const rowCount = sheet.getMaxRows() - 1;

  if (rowCount <= 0) {
    Logger.log("leave_requests.status の入力規則は更新対象行がありません");
    return {
      ok: true,
      updated_rows: 0,
      status_column: statusColumn,
      allowed_values: LEAVE_REQUEST_STATUS_VALUES.slice()
    };
  }

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(LEAVE_REQUEST_STATUS_VALUES, true)
    .setAllowInvalid(false)
    .build();

  // ヘッダー行を除外し、既存行と今後 append される行に同じ規則を設定する。
  sheet.getRange(2, statusColumn, rowCount, 1).setDataValidation(rule);

  Logger.log(
    "leave_requests.status の入力規則を更新しました: column=" +
      statusColumn +
      ", rows=2-" +
      sheet.getMaxRows() +
      ", values=" +
      LEAVE_REQUEST_STATUS_VALUES.join(", ")
  );

  return {
    ok: true,
    updated_rows: rowCount,
    status_column: statusColumn,
    allowed_values: LEAVE_REQUEST_STATUS_VALUES.slice()
  };
}

/* =========================
   leave_requests ヘッダー診断
========================= */
function debugLeaveRequestHeaders() {
  const sheet = getSheet("leave_requests");
  const headerInfo = getHeaderMap(sheet);

  headerInfo.headers.forEach((header, index) => {
    Logger.log((index + 1) + " " + String(header || "").trim());
  });

  return headerInfo.headers.map((header, index) => ({
    column: index + 1,
    header: String(header || "").trim()
  }));
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
  const cacheKey = CACHE_KEY.COMPANY_CALENDAR + (shouldUseSupabaseReads_() ? "_supabase" : "");
  const cached = cache.get(cacheKey);
  if (cached) {
    return JSON.parse(cached);
  }

  if (shouldUseSupabaseReads_()) {
    const map = {};
    getCompanyCalendarFromSupabase_().forEach(rowObj => {
      if (!rowObj.date) return;
      map[toDateKey(rowObj.date)] = norm(rowObj.type);
    });
    cache.put(cacheKey, JSON.stringify(map), 300);
    return map;
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

  cache.put(cacheKey, JSON.stringify(map), 300);
  return map;
}

// 管理者向け予定API専用。キャッシュを一切変更しない読み取り経路。
function getCompanyCalendarMapReadOnly_() {
  if (shouldUseSupabaseReads_()) {
    return buildCompanyCalendarMapFromRows_(getCompanyCalendarFromSupabase_());
  }

  const sheet = getSheet("company_calendar");
  const headerInfo = requireHeaders(sheet, ["date", "type"]);
  const data = sheet.getDataRange().getValues();
  return buildCompanyCalendarMapFromRows_(
    data.slice(1).map(row => rowToObject(row, headerInfo.headers))
  );
}

// 入力配列だけからカレンダーMapを作る純粋関数。テストでも利用する。
function buildCompanyCalendarMapFromRows_(rows) {
  const map = {};
  (Array.isArray(rows) ? rows : []).forEach(rowObj => {
    if (!rowObj || !rowObj.date) return;
    map[toDateKey(rowObj.date)] = norm(rowObj.type);
  });
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
  if (shouldUseSupabaseReads_()) {
    return getEmployeesFromSupabase_()
      .map(rowObj => ({
       id: String(rowObj.employee_id || "").trim(),
       name: String(rowObj.name || rowObj.employee_id || "").trim(),
       display_name: String(rowObj.display_name || "").trim(),

       company_code: String(rowObj.company_code || "").trim(),
       company_name: String(rowObj.company_name || "").trim(),

       fiscal_start_month: Number(rowObj.fiscal_start_month || 4),

       leave_management_target: rowObj.leave_management_target === true,

       employment_status: String(rowObj.employment_status || "").trim()
       }))
      .filter(emp => emp.id);
  }

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
  const cacheKey = CACHE_KEY.EMPLOYEE_MAP + (shouldUseSupabaseReads_() ? "_supabase" : "");
  const cached = cache.get(cacheKey);
  if (cached) {
    return JSON.parse(cached);
  }

  const employees = getEmployees();
  const map = {};

  employees.forEach(emp => {
    map[emp.id] = emp.name;
  });

  cache.put(cacheKey, JSON.stringify(map), 300);
  return map;
}

function getCurrentFiscalYear() {
  return getFiscalYearFromDate(new Date());
}

/* =========================
   付与情報
========================= */
function getGrantMapByFiscalYear(fiscalYear) {
  if (shouldUseSupabaseReads_()) {
    const result = {};
    const employeeDetailMap = getEmployeeDetailMap();

    getPaidLeaveGrantsFromSupabase_().forEach(rowObj => {
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

  if (shouldUseSupabaseReads_()) {
    const result = {};
    const calendarMap = getCompanyCalendarMap();
    const employeeDetailMap = getEmployeeDetailMap();

    getLeaveRequestsFromSupabase_().forEach(rowObj => {
      const employeeId = String(rowObj.employee_id || "").trim();
      const status = norm(rowObj.status);

      if (!employeeId) return;
      if (!targetIds.has(employeeId)) return;
      if (status !== STATUS.APPROVED) return;
      // 時間有給の残高消化は分単位FIFOを導入するPhase 3まで保留する。
      if (isTimeLeaveRequestRow_(rowObj)) return;
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
    // 時間有給の残高消化は分単位FIFOを導入するPhase 3まで保留する。
    if (isTimeLeaveRequestRow_(rowObj)) return;
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
   年5日取得義務: 通常残高用の取得日数とは別集計

   時間単位年休（request_kind=time_hourly）は通常有給残高の消化対象でも、
   年5日取得義務には一切算入しない。既存行で request_kind が空欄の場合は
   従来どおり1日・半日・複数日の申請として扱う。
========================= */
function getFiveDayObligationDaysByFiscalYear(fiscalYear) {
  const employees = getEmployees();
  return getFiveDayObligationDaysByFiscalYearForEmployeeIds(
    fiscalYear,
    employees.map(emp => emp.id)
  );
}

function getFiveDayObligationContribution_(requestRow, dailyDays) {
  const requestKind = norm(requestRow && requestRow.request_kind);
  if (requestKind === "time_hourly") return 0;
  return Number(dailyDays || 0);
}

function getFiveDayObligationDaysByFiscalYearForEmployeeIds(fiscalYear, employeeIds) {
  const targetIds = new Set(
    (employeeIds || [])
      .map(id => String(id || "").trim())
      .filter(Boolean)
  );
  if (targetIds.size === 0) return {};

  const result = {};
  const calendarMap = getCompanyCalendarMap();
  const employeeDetailMap = getEmployeeDetailMap();
  const rows = shouldUseSupabaseReads_()
    ? getLeaveRequestsFromSupabase_()
    : (function() {
      const sheet = getSheet("leave_requests");
      const headerInfo = requireHeaders(sheet, [
        "employee_id", "start_date", "end_date", "days", "half_day", "status"
      ]);
      const data = sheet.getDataRange().getValues();
      return data.length <= 1
        ? []
        : data.slice(1).map(row => rowToObject(row, headerInfo.headers));
    })();

  rows.forEach(rowObj => {
    const employeeId = String(rowObj.employee_id || "").trim();
    if (!employeeId || !targetIds.has(employeeId)) return;
    if (norm(rowObj.status) !== STATUS.APPROVED) return;
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
      const contribution = getFiveDayObligationContribution_(rowObj, item.days);
      if (contribution === 0) return;
      result[employeeId] = Number(result[employeeId] || 0) + contribution;
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

// 申請画面等で使う現在時点の正式残高。MAIN のみ分単位 FIFO を正とし、
// PARTNER は既存の日数年度集計を呼出側で引き続き使う。
function getCurrentMinuteFifoBalanceMapForEmployeeIds_(employeeIds) {
  const ids = (employeeIds || []).map(id => String(id || "").trim()).filter(Boolean);
  if (ids.length === 0) return {};
  const asOfDate = parseLocalDate(new Date());
  const context = createFifoBalanceComparisonContext_(asOfDate);
  const result = {};
  ids.forEach(employeeId => {
    if (!isMainTimeLeaveEmployeeForFifo_(employeeId, context)) return;
    result[employeeId] = calculateFifoBalanceMinutesFromContext_(
      employeeId, asOfDate, context
    );
  });
  return result;
}

// 申請画面の「申請中」表示用。利用日が未来のpendingも含め、現在未確定の
// 有給予約総額を返す。候補・登録時の検証は利用日順に判定する既存関数を使う。
function getCurrentMainPendingPaidLeaveReservationMapForEmployeeIds_(employeeIds) {
  const ids = (employeeIds || []).map(id => String(id || "").trim()).filter(Boolean);
  if (ids.length === 0) return {};
  const context = createFifoBalanceComparisonContext_(parseLocalDate(new Date()));
  const result = {};
  ids.forEach(employeeId => {
    if (!isMainTimeLeaveEmployeeForFifo_(employeeId, context)) return;
    result[employeeId] = getAllPendingPaidLeaveReservationMinutes_(employeeId, context);
  });
  return result;
}

/* =========================
   FIFO残日数計算（試験実装）
   既存表示には未接続
========================= */
function calculateFifoPaidLeaveBalance(employeeId, asOfDateValue) {
  const targetEmployeeId = String(employeeId || "").trim();
  if (!targetEmployeeId) throw new Error("employeeId がありません");

  const asOfDate = asOfDateValue ? parseLocalDate(asOfDateValue) : parseLocalDate(new Date());
  const context = createFifoBalanceComparisonContext_(asOfDate);
  if (isMainTimeLeaveEmployeeForFifo_(targetEmployeeId, context)) {
    return calculateFifoBalanceMinutesFromContext_(targetEmployeeId, asOfDate, context);
  }
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
  if (isMainTimeLeaveEmployeeForFifo_(employeeId, context)) {
    return calculateFifoBalanceMinutesFromContext_(employeeId, asOfDate, context);
  }

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

/* =========================
   MAIN 時間有給対応: 分単位 FIFO
   既存の物理日数は変更せず、計算境界だけで整数分に正規化する。
========================= */
function isMainTimeLeaveEmployeeForFifo_(employeeId, context) {
  const code = String(
    context && context.company_code_by_employee && context.company_code_by_employee[employeeId] || ""
  ).trim().toUpperCase();
  return code === "MAIN" && getCompanyLeavePolicy(code).timeLeaveEnabled === true;
}

function getMinuteBalanceDisplay_(remainingMinutes, scheduledMinutesPerDay) {
  const minutes = Math.max(0, Number(remainingMinutes || 0));
  const perDay = Number(scheduledMinutesPerDay || 420);
  const fullDays = Math.floor(minutes / perDay);
  const remainderMinutes = minutes % perDay;
  return {
    remaining_minutes: minutes,
    remaining_days: minutes / perDay,
    remaining_full_days: fullDays,
    remaining_hours: Math.floor(remainderMinutes / 60),
    remaining_remainder_minutes: remainderMinutes
  };
}

// 旧データは carry_over_days に小数日（例: 1.5日）を保持している。
// 新方式では carry_over_days は整数日、端数は carry_over_minutes にだけ保持する。
// 小数日と分列が混在した不整合行では、既存の小数日を正として分列を加算しない。
// これにより旧形式を壊さず、移行途中の二重計上も防ぐ。
function getGrantCarryOverMinutes_(grant, scheduledMinutesPerDay) {
  const perDay = Number(scheduledMinutesPerDay || 420);
  const carryDays = Number(grant && grant.carry_over_days || 0);
  const rawCarryMinutes = grant && grant.carry_over_minutes;
  const carryMinutes = rawCarryMinutes === "" || rawCarryMinutes == null ? 0 : Number(rawCarryMinutes);
  if (!Number.isFinite(carryDays) || !Number.isInteger(carryMinutes) || carryMinutes < 0) {
    throw new Error("繰越有給の値が不正です");
  }
  if (!Number.isInteger(carryDays)) return carryDays * perDay;
  return carryDays * perDay + carryMinutes;
}

function calculateCarryOverMinutes_(remainingMinutes, scheduledMinutesPerDay) {
  const remaining = Math.max(0, Number(remainingMinutes || 0));
  const perDay = Number(scheduledMinutesPerDay || 420);
  const candidateMinutes = Math.min(remaining, 20 * perDay);
  return {
    carry_over_candidate_minutes: candidateMinutes,
    carry_over_days: Math.floor(candidateMinutes / perDay),
    carry_over_minutes: candidateMinutes % perDay,
    carry_over_limit_expired_minutes: Math.max(0, remaining - candidateMinutes)
  };
}

function getFifoApprovedLeaveUseRowsMinutesFromContext_(employeeId, asOfDate, context, policy) {
  const requests = context.requests_by_employee[employeeId] || [];
  const segmentsByRequest = context.time_leave_segments_by_request || {};
  const result = [];

  requests.forEach(rowObj => {
    const status = norm(rowObj.status);
    const requestType = String(rowObj.type || "paid_leave").trim();
    if (status !== STATUS.APPROVED || (requestType && requestType !== "paid_leave")) return;

    const requestId = String(rowObj.request_id || "").trim();
    if (isTimeLeaveSegmentRequestRow_(rowObj)) {
      if (isCombinedHalfDayTimeLeaveRequestRow_(rowObj)) {
        const halfDay = norm(rowObj.half_day);
        if (halfDay !== "am" && halfDay !== "pm") {
          throw new Error("複合申請の半休区分が不正です: " + requestId);
        }
        result.push({
          request_id: requestId,
          time_leave_id: "",
          use_date: parseLocalDate(rowObj.start_date),
          leave_kind: "half_day",
          consumed_minutes: policy.scheduledMinutesPerDay / 2,
          unallocated_minutes: 0
        });
      }
      (segmentsByRequest[requestId] || []).forEach(segment => {
        const useDate = parseLocalDate(segment.leave_date);
        if (useDate > asOfDate) return;
        const consumedMinutes = Number(segment.requested_minutes || 0);
        if (!Number.isInteger(consumedMinutes) || consumedMinutes <= 0) {
          throw new Error("時間有給明細の取得分が不正です: " + requestId);
        }
        result.push({
          request_id: requestId,
          time_leave_id: String(segment.time_leave_id || "").trim(),
          use_date: useDate,
          leave_kind: "time_hourly",
          consumed_minutes: consumedMinutes,
          unallocated_minutes: 0
        });
      });
      return;
    }

    if (!rowObj.start_date || !rowObj.end_date) return;
    const dailyRows = expandLeaveRequestToDailyRows(
      rowObj.start_date, rowObj.end_date, rowObj.days, rowObj.half_day, context.calendar_map
    );
    dailyRows.forEach(item => {
      const useDate = parseLocalDate(item.date);
      if (useDate > asOfDate) return;
      const isHalfDay = !!norm(rowObj.half_day);
      result.push({
        request_id: requestId,
        time_leave_id: "",
        use_date: useDate,
        leave_kind: isHalfDay ? "half_day" : "full_day",
        consumed_minutes: isHalfDay
          ? policy.scheduledMinutesPerDay / 2
          : policy.scheduledMinutesPerDay,
        unallocated_minutes: 0
      });
    });
  });

  return result.sort((a, b) => {
    if (a.use_date.getTime() !== b.use_date.getTime()) return a.use_date - b.use_date;
    if (a.request_id !== b.request_id) return a.request_id.localeCompare(b.request_id);
    return a.time_leave_id.localeCompare(b.time_leave_id);
  });
}

function calculateFifoBalanceMinutesFromContext_(employeeId, asOfDate, context) {
  const policy = getCompanyLeavePolicy("MAIN");
  const grants = (context.grants_by_employee[employeeId] || [])
    .filter(grant => grant.is_finalized && grant.valid_from_date <= asOfDate)
    .map(grant => {
      const grantMinutes = Number(grant.grant_days || 0) * policy.scheduledMinutesPerDay;
      const carryOverMinutes = getGrantCarryOverMinutes_(grant, policy.scheduledMinutesPerDay);
      const totalMinutes = grantMinutes + carryOverMinutes;
      return {
        grant_id: grant.grant_id,
        grant_date: grant.grant_date,
        valid_from_date: grant.valid_from_date,
        valid_to_date: grant.valid_to_date,
        grant_type: grant.grant_type,
        year: grant.year,
        grant_days: Number(grant.grant_days || 0),
        carry_over_days: Number(grant.carry_over_days || 0),
        carry_over_minutes: Number(grant.carry_over_minutes || 0),
        total_minutes: totalMinutes,
        used_minutes: 0,
        remaining_minutes: totalMinutes,
        active_remaining_minutes: 0,
        expired_minutes: 0,
        is_expired: false
      };
    })
    .sort((a, b) => a.grant_date.getTime() !== b.grant_date.getTime()
      ? a.grant_date - b.grant_date
      : String(a.grant_id).localeCompare(String(b.grant_id)));
  const usedRows = getFifoApprovedLeaveUseRowsMinutesFromContext_(employeeId, asOfDate, context, policy);
  const allocations = [];

  usedRows.forEach(useRow => {
    let remainingUseMinutes = useRow.consumed_minutes;
    grants.forEach(grant => {
      if (remainingUseMinutes <= 0 || grant.remaining_minutes <= 0) return;
      if (useRow.use_date < grant.valid_from_date || useRow.use_date > grant.valid_to_date) return;
      const consumedMinutes = Math.min(grant.remaining_minutes, remainingUseMinutes);
      grant.remaining_minutes -= consumedMinutes;
      grant.used_minutes += consumedMinutes;
      remainingUseMinutes -= consumedMinutes;
      allocations.push({
        request_id: useRow.request_id,
        time_leave_id: useRow.time_leave_id,
        use_date: formatDateValue(useRow.use_date),
        leave_kind: useRow.leave_kind,
        grant_id: grant.grant_id,
        consumed_minutes: consumedMinutes,
        grant_valid_to: formatDateValue(grant.valid_to_date),
        calculation_version: "fifo_minutes_v1"
      });
    });
    useRow.unallocated_minutes = Math.max(0, remainingUseMinutes);
  });

  grants.forEach(grant => {
    grant.is_expired = grant.valid_to_date < asOfDate;
    grant.expired_minutes = grant.is_expired ? grant.remaining_minutes : 0;
    grant.active_remaining_minutes = grant.is_expired ? 0 : grant.remaining_minutes;
  });
  const currentRemainingMinutes = grants.reduce((sum, grant) => sum + grant.active_remaining_minutes, 0);
  const display = getMinuteBalanceDisplay_(currentRemainingMinutes, policy.scheduledMinutesPerDay);
  return Object.assign({
    employee_id: employeeId,
    as_of_date: formatDateValue(asOfDate),
    calculation_mode: "fifo_minutes_v1",
    scheduled_minutes_per_day: policy.scheduledMinutesPerDay,
    current_remaining_minutes: currentRemainingMinutes,
    total_granted_minutes: grants.reduce((sum, grant) => sum + grant.total_minutes, 0),
    used_minutes: usedRows.reduce((sum, row) => sum + row.consumed_minutes, 0),
    allocated_used_minutes: allocations.reduce((sum, row) => sum + row.consumed_minutes, 0),
    unallocated_used_minutes: usedRows.reduce((sum, row) => sum + row.unallocated_minutes, 0),
    expired_minutes: grants.reduce((sum, grant) => sum + grant.expired_minutes, 0),
    grant_details: grants.map(grant => Object.assign({}, grant, {
      grant_date: formatDateValue(grant.grant_date),
      valid_from: formatDateValue(grant.valid_from_date),
      valid_to: formatDateValue(grant.valid_to_date),
      total_days: grant.total_minutes / policy.scheduledMinutesPerDay,
      used_days: grant.used_minutes / policy.scheduledMinutesPerDay,
      remaining_days: grant.remaining_minutes / policy.scheduledMinutesPerDay,
      active_remaining_days: grant.active_remaining_minutes / policy.scheduledMinutesPerDay,
      expired_days: grant.expired_minutes / policy.scheduledMinutesPerDay
    })),
    used_details: usedRows.map(row => Object.assign({}, row, {
      use_date: formatDateValue(row.use_date),
      days: row.consumed_minutes / policy.scheduledMinutesPerDay,
      unallocated_days: row.unallocated_minutes / policy.scheduledMinutesPerDay
    })),
    allocations: allocations
  }, display, {
    total_granted_days: grants.reduce((sum, grant) => sum + grant.total_minutes, 0) / policy.scheduledMinutesPerDay,
    used_days: usedRows.reduce((sum, row) => sum + row.consumed_minutes, 0) / policy.scheduledMinutesPerDay,
    allocated_used_days: allocations.reduce((sum, row) => sum + row.consumed_minutes, 0) / policy.scheduledMinutesPerDay,
    unallocated_used_days: usedRows.reduce((sum, row) => sum + row.unallocated_minutes, 0) / policy.scheduledMinutesPerDay,
    expired_days: grants.reduce((sum, grant) => sum + grant.expired_minutes, 0) / policy.scheduledMinutesPerDay,
    current_remaining_days: currentRemainingMinutes / policy.scheduledMinutesPerDay
  });
}

function getPendingTimeLeaveReservationMinutes_(employeeId, asOfDate, context, excludedRequestId) {
  const targetId = String(employeeId || "").trim();
  const excludedId = String(excludedRequestId || "").trim();
  const targetDate = parseLocalDate(asOfDate);
  const segmentsByRequest = context.time_leave_segments_by_request || {};
  let total = 0;
  (context.requests_by_employee[targetId] || []).forEach(parent => {
    const requestId = String(parent.request_id || "").trim();
    if (requestId === excludedId || !isTimeLeaveSegmentRequestRow_(parent)) return;
    if (norm(parent.status) !== STATUS.PENDING) return;
    (segmentsByRequest[requestId] || []).forEach(segment => {
      if (parseLocalDate(segment.leave_date) > targetDate) return;
      const minutes = Number(segment.requested_minutes || 0);
      if (!Number.isInteger(minutes) || minutes < 0) {
        throw new Error("時間有給明細の取得分が不正です: " + requestId);
      }
      total += minutes;
    });
  });
  return total;
}

function validateTimeLeaveFifoBalanceAvailability_(candidate, excludedRequestId) {
  const context = createFifoBalanceComparisonContext_(candidate.leave_date);
  const balance = calculateFifoBalanceWithOpeningBalanceFromContext_(
    candidate.employee_id,
    candidate.leave_date,
    context
  );
  const approvedRemainingMinutes = Number(balance.current_remaining_minutes || 0);
  const pendingReservedMinutes = getPendingPaidLeaveReservationMinutes_(
    candidate.employee_id,
    candidate.leave_date,
    context,
    excludedRequestId
  );
  return validateTimeLeaveFifoReservation_(
    approvedRemainingMinutes,
    pendingReservedMinutes,
    candidate.requested_minutes
  );
}

function validateTimeLeaveFifoReservation_(approvedRemainingMinutes, pendingReservedMinutes, requestedMinutes) {
  const approved = Number(approvedRemainingMinutes || 0);
  const pending = Number(pendingReservedMinutes || 0);
  const requested = Number(requestedMinutes || 0);
  if (![approved, pending, requested].every(Number.isInteger) || approved < 0 || pending < 0 || requested <= 0) {
    throw new Error("時間有給残高の分数が不正です");
  }
  const availableMinutes = approved - pending;
  if (availableMinutes < requested) {
    throw new Error("通常有給残高が不足しています（承認済み残高から他の時間有給申請を仮押さえしています）");
  }
  return {
    approved_remaining_minutes: approvedRemainingMinutes,
    pending_reserved_minutes: pendingReservedMinutes,
    available_minutes: availableMinutes
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
  if (shouldUseSupabaseReads_()) {
    const result = {};

    getPaidLeaveGrantsFromSupabase_().forEach(rowObj => {
      const employeeId = String(rowObj.employee_id || "").trim();

      if (!employeeId) return;
      if (String(rowObj.grant_type || "").trim() !== "yearly") return;
      if (Number(rowObj.year) !== Number(fiscalYear)) return;

      result[employeeId] = true;
    });

    return result;
  }

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

function createFifoBalanceComparisonContext_(asOfDate, options) {
  const opts = options || {};
  const companyCodeByEmployee = {};
  getEmployeesForAdmin().forEach(employee => {
    const employeeId = String(employee.employee_id || "").trim();
    if (employeeId) companyCodeByEmployee[employeeId] = String(employee.company_code || "").trim().toUpperCase();
  });
  return {
    as_of_date: asOfDate,
    calendar_map: opts.read_only === true
      ? getCompanyCalendarMapReadOnly_()
      : getCompanyCalendarMap(),
    grants_by_employee: getPaidLeaveGrantRowsByEmployeeForFifoCompare_(),
    requests_by_employee: getLeaveRequestRowsByEmployeeForFifoCompare_(),
    time_leave_segments_by_request: getTimeLeaveSegmentsByRequestForFifoCompare_(),
    company_code_by_employee: companyCodeByEmployee
  };
}

function getTimeLeaveSegmentsByRequestForFifoCompare_() {
  // 時間有給の子明細は、USE_SUPABASE_READS の設定に関係なくSpreadsheetを正とする。
  // 現行Supabase schemaには time_leave_segments を追加しないため、ここから読みに行かない。
  const sheet = getAppSpreadsheet().getSheetByName(TIME_LEAVE_SEGMENTS_SHEET);
  if (!sheet) return {};
  const headerInfo = requireHeaders(sheet, [
    "time_leave_id", "request_id", "leave_date", "requested_minutes"
  ]);
  const data = sheet.getDataRange().getValues();
  const result = {};
  data.slice(1).forEach(row => {
    const rowObj = rowToObject(row, headerInfo.headers);
    const requestId = String(rowObj.request_id || "").trim();
    if (!requestId) return;
    if (!result[requestId]) result[requestId] = [];
    result[requestId].push(rowObj);
  });
  return result;
}

function getPaidLeaveGrantRowsByEmployeeForFifoCompare_() {
  if (shouldUseSupabaseReads_()) {
    const result = {};

    getPaidLeaveGrantsFromSupabase_().forEach(rowObj => {
      const employeeId = String(rowObj.employee_id || "").trim();
      if (!employeeId || !rowObj.grant_date) return;

      const grantDate = parseLocalDate(rowObj.grant_date);
      const validFromDate = rowObj.valid_from ? parseLocalDate(rowObj.valid_from) : grantDate;
      const validToDate = rowObj.valid_to
        ? parseLocalDate(rowObj.valid_to)
        : addDaysLocal_(addYearsLocal_(grantDate, 2), -1);
      const grantDays = Number(rowObj.grant_days || 0);
      const carryOverDays = Number(rowObj.carry_over_days || 0);
      const carryOverMinutes = Number(rowObj.carry_over_minutes || 0);

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
        carry_over_minutes: carryOverMinutes,
        notes: String(rowObj.notes || ""),
        has_recorded_valid_from: !!rowObj.valid_from,
        has_recorded_valid_to: !!rowObj.valid_to,
        is_finalized: rowObj.is_finalized !== false
      });
    });

    // MAINの分FIFOには、Supabase未適用のcarry_over_minutesと新年度付与を含める必要がある。
    // そのためMAINのロットだけSpreadsheet正本で置換し、PARTNERの既存Supabase readは維持する。
    const mainEmployeeIds = {};
    getEmployeesForAdmin().forEach(employee => {
      if (String(employee.company_code || "").trim().toUpperCase() !== "MAIN") return;
      mainEmployeeIds[String(employee.employee_id || "").trim()] = true;
    });
    const spreadsheetMainGrants = getSpreadsheetPaidLeaveGrantRowsForFifoCompare_();
    Object.keys(mainEmployeeIds).forEach(employeeId => {
      delete result[employeeId];
      if (spreadsheetMainGrants[employeeId]) result[employeeId] = spreadsheetMainGrants[employeeId];
    });

    return result;
  }

  return getSpreadsheetPaidLeaveGrantRowsForFifoCompare_();
}

function getSpreadsheetPaidLeaveGrantRowsForFifoCompare_() {
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
    const carryOverMinutes = Number(rowObj.carry_over_minutes || 0);
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
      carry_over_minutes: carryOverMinutes,
      notes: String(rowObj.notes || ""),
      has_recorded_valid_from: !!rowObj.valid_from,
      has_recorded_valid_to: !!rowObj.valid_to,
      is_finalized: finalizedValue !== "FALSE"
    });
  });

  return result;
}

function getLeaveRequestRowsByEmployeeForFifoCompare_() {
  if (shouldUseSupabaseReads_()) {
    // 通常有給は既存どおりSupabaseを優先する。一方、時間有給親は子明細と同じ
    // Spreadsheet正本を優先し、同一IDの旧Supabase親（request_kindなし）で
    // combined/time_hourlyを通常有給として解釈しないようにする。
    const sheet = getSheet("leave_requests");
    const headerInfo = requireHeaders(sheet, [
      "request_id", "employee_id", "start_date", "end_date", "days", "half_day", "status"
    ]);
    const spreadsheetTimeLeaveParents = sheet.getDataRange().getValues().slice(1)
      .map(row => rowToObject(row, headerInfo.headers))
      .filter(rowObj => isTimeLeaveSegmentRequestRow_(rowObj));

    return buildFifoLeaveRequestRowsByEmployee_(
      getLeaveRequestsFromSupabase_(),
      spreadsheetTimeLeaveParents
    );
  }

  return getSpreadsheetLeaveRequestRowsByEmployeeForFifoCompare_();
}

// Supabase親を基礎にしつつ、Spreadsheetで識別できる時間有給親だけは同一request_idで置換する。
// 子明細もSpreadsheetを正としているため、親子の読取元を揃えるためのFIFO専用処理。
function buildFifoLeaveRequestRowsByEmployee_(supabaseRows, spreadsheetTimeLeaveParents) {
  const rowsByRequestId = {};
  const rowsWithoutRequestId = [];

  (supabaseRows || []).forEach(rowObj => {
    const requestId = String(rowObj && rowObj.request_id || "").trim();
    if (requestId) rowsByRequestId[requestId] = rowObj;
    else rowsWithoutRequestId.push(rowObj);
  });

  (spreadsheetTimeLeaveParents || []).forEach(rowObj => {
    const requestId = String(rowObj && rowObj.request_id || "").trim();
    if (!requestId || !isTimeLeaveSegmentRequestRow_(rowObj)) return;
    rowsByRequestId[requestId] = rowObj;
  });

  const result = {};
  Object.keys(rowsByRequestId).forEach(requestId => {
    const rowObj = rowsByRequestId[requestId];
    const employeeId = String(rowObj && rowObj.employee_id || "").trim();
    if (!employeeId) return;
    if (!result[employeeId]) result[employeeId] = [];
    result[employeeId].push(rowObj);
  });
  rowsWithoutRequestId.forEach(rowObj => {
    const employeeId = String(rowObj && rowObj.employee_id || "").trim();
    if (!employeeId) return;
    if (!result[employeeId]) result[employeeId] = [];
    result[employeeId].push(rowObj);
  });
  return result;
}

function getSpreadsheetLeaveRequestRowsByEmployeeForFifoCompare_() {
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
    if (isTimeLeaveRequestRow_(rowObj)) return;
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
  if (shouldUseSupabaseReads_()) {
    return getPaidLeaveGrantsFromSupabase_()
      .filter(rowObj => {
        if (String(rowObj.employee_id || "").trim() !== String(employeeId)) return false;
        if (!rowObj.grant_date) return false;

        const validFromDate = rowObj.valid_from
          ? parseLocalDate(rowObj.valid_from)
          : parseLocalDate(rowObj.grant_date);
        if (validFromDate > asOfDate) return false;

        return rowObj.is_finalized !== false;
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
  if (shouldUseSupabaseReads_()) {
    const calendarMap = getCompanyCalendarMap();
    const result = [];

    getLeaveRequestsFromSupabase_().forEach(rowObj => {
      const targetEmployeeId = String(rowObj.employee_id || "").trim();
      const status = norm(rowObj.status);
      const requestType = String(rowObj.type || "paid_leave").trim();

      if (targetEmployeeId !== String(employeeId)) return;
      if (status !== STATUS.APPROVED) return;
      if (requestType && requestType !== "paid_leave") return;
      if (isTimeLeaveRequestRow_(rowObj)) return;
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
    if (isTimeLeaveRequestRow_(rowObj)) return;
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
   時間単位年休: Spreadsheet 構造

   この初期化関数は時間休の登録処理からのみ呼ぶ。デプロイ時や通常の
   既存申請処理では実行しないため、既存シートの列・既存データを変更しない。
========================= */
const TIME_LEAVE_SEGMENTS_SHEET = "time_leave_segments";
const TIME_LEAVE_SEGMENTS_HEADERS = [
  "time_leave_id",
  "request_id",
  "employee_id",
  "company_code",
  "leave_date",
  "start_time",
  "end_time",
  "start_minute",
  "end_minute",
  "requested_minutes",
  "scheduled_minutes_per_day",
  "time_leave_unit_minutes",
  "work_start_minute",
  "work_end_minute",
  "break_periods_json",
  "calculation_version",
  "created_at",
  "updated_at"
];
const TIME_LEAVE_REQUEST_HEADERS = [
  "request_kind",
  "company_code_snapshot",
  "policy_version"
];

function ensureTimeLeaveInfrastructure_() {
  const requestSheet = getSheet("leave_requests");
  TIME_LEAVE_REQUEST_HEADERS.forEach(header => ensureSheetColumn_(requestSheet, header));

  const ss = getAppSpreadsheet();
  let segmentSheet = ss.getSheetByName(TIME_LEAVE_SEGMENTS_SHEET);
  if (!segmentSheet) {
    segmentSheet = ss.insertSheet(TIME_LEAVE_SEGMENTS_SHEET);
    segmentSheet.getRange(1, 1, 1, TIME_LEAVE_SEGMENTS_HEADERS.length)
      .setValues([TIME_LEAVE_SEGMENTS_HEADERS]);
    segmentSheet.setFrozenRows(1);
  }
  // 既存シートではヘッダーを末尾追加するだけに留める。列順・既存値・書式を変更しない。
  TIME_LEAVE_SEGMENTS_HEADERS.forEach(header => ensureSheetColumn_(segmentSheet, header));
  requireHeaders(segmentSheet, TIME_LEAVE_SEGMENTS_HEADERS);
  return { requestSheet: requestSheet, segmentSheet: segmentSheet };
}

// 年度切替の確定処理だけが呼び出す。通常の読取・残高計算ではシート構造を変更しない。
function ensurePaidLeaveMinutesInfrastructure_() {
  const grantSheet = getSheet("paid_leave_grants");
  ensureSheetColumn_(grantSheet, "carry_over_minutes");
  return grantSheet;
}

// 本番実行は呼出側の明示的操作に限る初期化入口。既存列やデータには触れず、足りない
// ヘッダーだけを末尾追加する。Phase 5ではこの関数を実行しない。
function initializeTimeLeaveProductionStructure_() {
  const timeLeave = ensureTimeLeaveInfrastructure_();
  const grantSheet = ensurePaidLeaveMinutesInfrastructure_();
  const retirementSheet = ensureLeaveRetirementRecordsSheet_();
  return {
    ok: true,
    leave_requests_headers: getHeaderMap(timeLeave.requestSheet).headers,
    time_leave_segments_headers: getHeaderMap(timeLeave.segmentSheet).headers,
    paid_leave_grants_headers: getHeaderMap(grantSheet).headers,
    leave_retirement_records_headers: getHeaderMap(retirementSheet).headers
  };
}

// Apps Scriptエディタの関数一覧から、本番Spreadsheet構造の初期化を明示実行する入口。
// 実処理は内部関数へ完全委譲し、この関数自身はSpreadsheet操作を持たない。
function initializeTimeLeaveProductionStructure() {
  const result = initializeTimeLeaveProductionStructure_();
  console.log(JSON.stringify(result, null, 2));
  return result;
}

function formatMinuteAsTime_(minute) {
  const value = Number(minute);
  if (!Number.isInteger(value) || value < 0 || value >= 24 * 60) {
    throw new Error("時刻分数が不正です");
  }
  return String(Math.floor(value / 60)).padStart(2, "0") + ":" +
    String(value % 60).padStart(2, "0");
}

function getLeaveRequestRecordById_(requestId) {
  const targetId = String(requestId || "").trim();
  const sheet = getSheet("leave_requests");
  const headerInfo = requireHeaders(sheet, ["request_id"]);
  const data = sheet.getDataRange().getValues();
  const index = data.findIndex((row, rowIndex) =>
    rowIndex > 0 && String(row[headerInfo.map.request_id] || "").trim() === targetId
  );
  if (index < 0) return null;
  return {
    sheet: sheet,
    headerInfo: headerInfo,
    sheetRow: index + 1,
    rowObj: rowToObject(data[index], headerInfo.headers)
  };
}

function deleteLeaveRequestById_(requestId) {
  const record = getLeaveRequestRecordById_(requestId);
  if (!record) return false;
  record.sheet.deleteRow(record.sheetRow);
  return true;
}

function isTimeLeaveRequestRow_(rowObj) {
  return norm(rowObj && rowObj.request_kind) === "time_hourly";
}

function isCombinedHalfDayTimeLeaveRequestRow_(rowObj) {
  return norm(rowObj && rowObj.request_kind) === "half_day_time_hourly";
}

// time_leave_segments を持つ親申請を判定する。年5日義務の扱いとは分け、
// 単独時間休と半休＋時間休の双方をここで扱う。
function isTimeLeaveSegmentRequestRow_(rowObj) {
  return isTimeLeaveRequestRow_(rowObj) || isCombinedHalfDayTimeLeaveRequestRow_(rowObj);
}

function isTimeLeaveRequestById_(requestId) {
  const record = getLeaveRequestRecordById_(requestId);
  return !!(record && isTimeLeaveSegmentRequestRow_(record.rowObj));
}

function getTimeLeaveSegmentRecordByRequestId_(requestId) {
  const targetId = String(requestId || "").trim();
  const sheet = getSheet(TIME_LEAVE_SEGMENTS_SHEET);
  const headerInfo = requireHeaders(sheet, TIME_LEAVE_SEGMENTS_HEADERS);
  const data = sheet.getDataRange().getValues();
  const matched = data
    .map((row, index) => ({ row: row, index: index }))
    .filter(item => item.index > 0 &&
      String(item.row[headerInfo.map.request_id] || "").trim() === targetId);

  if (matched.length === 0) return null;
  if (matched.length > 1) {
    throw new Error("時間有給明細が複数あります: request_id=" + targetId);
  }
  return {
    sheet: sheet,
    headerInfo: headerInfo,
    sheetRow: matched[0].index + 1,
    rowObj: rowToObject(matched[0].row, headerInfo.headers)
  };
}

function getTimeLeaveSegmentsForEmployeeDate_(employeeId, leaveDate) {
  const targetEmployeeId = String(employeeId || "").trim();
  const targetDate = toDateKey(leaveDate);
  const sheet = getSheet(TIME_LEAVE_SEGMENTS_SHEET);
  const headerInfo = requireHeaders(sheet, TIME_LEAVE_SEGMENTS_HEADERS);
  const data = sheet.getDataRange().getValues();

  return data.slice(1)
    .map((row, index) => ({ rowObj: rowToObject(row, headerInfo.headers), sheetRow: index + 2 }))
    .filter(item =>
      String(item.rowObj.employee_id || "").trim() === targetEmployeeId &&
      item.rowObj.leave_date && toDateKey(item.rowObj.leave_date) === targetDate
    );
}

function getTimeLeaveParentRows_() {
  const sheet = getSheet("leave_requests");
  const headerInfo = requireHeaders(sheet, [
    "request_id", "employee_id", "start_date", "end_date", "days", "half_day", "status"
  ]);
  const data = sheet.getDataRange().getValues();
  const result = {};
  data.slice(1).forEach(row => {
    const rowObj = rowToObject(row, headerInfo.headers);
    const requestId = String(rowObj.request_id || "").trim();
    if (requestId) result[requestId] = rowObj;
  });
  return result;
}

function isActiveTimeLeaveStatus_(status) {
  const value = norm(status);
  return value === STATUS.PENDING || value === STATUS.APPROVED;
}

function validateTimeLeaveEmployee_(employeeId) {
  const targetEmployeeId = String(employeeId || "").trim();
  const employee = getTimeLeaveEmployeeFromSpreadsheet_(targetEmployeeId);
  if (!employee) throw new Error("対象社員が見つかりません");

  const status = norm(employee.employment_status);
  if (status !== "active" && status !== "在職") {
    throw new Error("在職中の社員だけ時間有給を申請できます");
  }
  if (String(employee.leave_management_target || "").toUpperCase() !== "TRUE") {
    throw new Error("この社員は有給管理対象ではありません");
  }

  const companyCode = String(employee.company_code || "").trim().toUpperCase();
  const policy = resolveEmployeeTimeLeavePolicy_(targetEmployeeId, employee);
  if (!policy.timeLeaveEnabled) {
    throw new Error("この会社では時間有給を利用できません");
  }
  return { employee: employee, companyCode: companyCode, policy: policy };
}

function normalizeTimeLeavePayload_(payload, employeeInfo, calculationVersion) {
  const data = payload || {};
  if (!data.leave_date) throw new Error("leave_date がありません");
  const leaveDate = parseLocalDate(data.leave_date);
  const startMinute = parseTimeToMinute(data.start_time);
  const endMinute = parseTimeToMinute(data.end_time);
  const policy = employeeInfo.policy;
  const version = normalizeTimeLeaveCalculationVersion_(
    calculationVersion == null ? data.calculation_version : calculationVersion
  );

  validateLeaveRequestDates(leaveDate, leaveDate, "");
  const requestedMinutes = calculateTimeLeaveRequestedMinutes_(startMinute, endMinute, policy, version);
  validateTimeLeaveUnitMinutes(requestedMinutes, policy);
  if (data.requested_minutes !== undefined && data.requested_minutes !== null && data.requested_minutes !== "") {
    const suppliedMinutes = Number(data.requested_minutes);
    if (!Number.isInteger(suppliedMinutes) || suppliedMinutes !== requestedMinutes) {
      throw new Error("時間有給の取得分が開始・終了時刻の計算結果と一致しません");
    }
  }

  return {
    employee_id: String(data.employee_id || "").trim(),
    leave_date: leaveDate,
    leave_date_key: toDateKey(leaveDate),
    start_minute: startMinute,
    end_minute: endMinute,
    start_time: formatMinuteAsTime_(startMinute),
    end_time: formatMinuteAsTime_(endMinute),
    requested_minutes: requestedMinutes,
    calculation_version: version,
    reason: String(data.reason || "").trim(),
    reason_detail: String(data.reason_detail || "").trim(),
    company_code: employeeInfo.companyCode,
    policy: policy
  };
}

// 新規の単独時間年休だけに適用する運用上限。既存明細の承認・監査時再検証や
// 半休＋時間年休には適用しないため、過去データの有効性を遡って変更しない。
function validateNewStandaloneTimeLeaveMaximum_(candidate) {
  const requestedMinutes = Number(candidate && candidate.requested_minutes);
  if (!Number.isInteger(requestedMinutes) || requestedMinutes <= 0) {
    throw new Error("時間年休の取得分が不正です");
  }
  if (requestedMinutes > MAX_STANDALONE_TIME_LEAVE_MINUTES) {
    throw new Error("単独の時間年休は3時間までです。4時間以上は半休＋時間休または1日有給を利用してください");
  }
  return { ok: true, requested_minutes: requestedMinutes };
}

function getDailyPaidLeaveEntries_(employeeId, leaveDate, excludedRequestId) {
  const targetEmployeeId = String(employeeId || "").trim();
  const targetDateKey = toDateKey(leaveDate);
  const excludedId = String(excludedRequestId || "").trim();
  const parentRows = getTimeLeaveParentRows_();
  const calendarMap = getCompanyCalendarMap();
  const policy = validateTimeLeaveEmployee_(targetEmployeeId).policy;
  const entries = [];

  Object.keys(parentRows).forEach(requestId => {
    const rowObj = parentRows[requestId];
    if (requestId === excludedId) return;
    if (String(rowObj.employee_id || "").trim() !== targetEmployeeId) return;
    if (!isActiveTimeLeaveStatus_(rowObj.status)) return;
    if (isTimeLeaveRequestRow_(rowObj)) return;

    const dailyRows = expandLeaveRequestToDailyRows(
      rowObj.start_date, rowObj.end_date, rowObj.days, rowObj.half_day, calendarMap
    );
    dailyRows.forEach(item => {
      if (toDateKey(item.date) !== targetDateKey) return;
      entries.push({
        request_id: requestId,
        kind: norm(rowObj.half_day) ? "half_day" : "full_day",
        half_day: String(rowObj.half_day || ""),
        minutes: norm(rowObj.half_day)
          ? policy.scheduledMinutesPerDay / 2
          : policy.scheduledMinutesPerDay
      });
    });
  });

  getTimeLeaveSegmentsForEmployeeDate_(targetEmployeeId, targetDateKey).forEach(item => {
    const parent = parentRows[String(item.rowObj.request_id || "").trim()];
    if (!parent || !isActiveTimeLeaveStatus_(parent.status)) return;
    if (String(item.rowObj.request_id || "").trim() === excludedId) return;
    entries.push({
      request_id: String(item.rowObj.request_id || "").trim(),
      time_leave_id: String(item.rowObj.time_leave_id || "").trim(),
      kind: "time_hourly",
      minutes: Number(item.rowObj.requested_minutes || 0),
      start_minute: Number(item.rowObj.start_minute),
      end_minute: Number(item.rowObj.end_minute)
    });
  });

  return entries;
}

function getTimeLeaveMinutesForFiscalYear_(employeeId, leaveDate, policy, excludedRequestId) {
  const date = parseLocalDate(leaveDate);
  const fiscalYear = getFiscalYearFromDateWithStart(date, policy.fiscalStartMonth);
  const range = getFiscalYearRangeWithStart(fiscalYear, policy.fiscalStartMonth);
  const excludedId = String(excludedRequestId || "").trim();
  const parentRows = getTimeLeaveParentRows_();
  const result = { approvedMinutes: 0, pendingMinutes: 0, fiscalYear: fiscalYear };
  const sheet = getSheet(TIME_LEAVE_SEGMENTS_SHEET);
  const headerInfo = requireHeaders(sheet, TIME_LEAVE_SEGMENTS_HEADERS);
  const data = sheet.getDataRange().getValues();

  data.slice(1).forEach(row => {
    const segment = rowToObject(row, headerInfo.headers);
    if (String(segment.employee_id || "").trim() !== String(employeeId || "").trim()) return;
    const requestId = String(segment.request_id || "").trim();
    if (requestId === excludedId) return;
    const parent = parentRows[requestId];
    if (!parent || !isTimeLeaveSegmentRequestRow_(parent) || !segment.leave_date || !isDateInRange(segment.leave_date, range.start, range.end)) return;
    const minutes = Number(segment.requested_minutes || 0);
    if (!Number.isInteger(minutes) || minutes < 0) {
      throw new Error("時間有給明細の取得分が不正です: " + requestId);
    }
    if (norm(parent.status) === STATUS.APPROVED) result.approvedMinutes += minutes;
    if (norm(parent.status) === STATUS.PENDING) result.pendingMinutes += minutes;
  });
  return result;
}

function getHalfDayOccupiedRange_(halfDay, policy) {
  const value = norm(halfDay);
  const halfDayMinutes = Number(policy && policy.scheduledMinutesPerDay) / 2;
  if (!Number.isInteger(halfDayMinutes) || halfDayMinutes <= 0) {
    throw new Error("半日有給の所定労働時間設定が不正です");
  }
  if (value === "am") {
    const endMinute = calculateWorkingEndMinute_(policy.workStartMinute, halfDayMinutes, policy);
    if (endMinute == null) throw new Error("AM半休の時間帯を解決できません");
    return { startMinute: policy.workStartMinute, endMinute: endMinute };
  }
  if (value === "pm") {
    const startMinute = calculateWorkingStartMinute_(policy.workEndMinute, halfDayMinutes, policy);
    if (startMinute == null) throw new Error("PM半休の時間帯を解決できません");
    return { startMinute: startMinute, endMinute: policy.workEndMinute };
  }
  throw new Error("半日有給区分が不正です");
}

function normalizeCombinedHalfDayTimeLeavePayload_(payload, employeeInfo, calculationVersion) {
  const candidate = normalizeTimeLeavePayload_(payload, employeeInfo, calculationVersion);
  const halfDay = norm(payload && payload.half_day);
  if (halfDay !== "am" && halfDay !== "pm") {
    throw new Error("half_day は am または pm で指定してください");
  }
  return Object.assign({}, candidate, {
    half_day: halfDay,
    half_day_minutes: employeeInfo.policy.scheduledMinutesPerDay / 2
  });
}

function parseTimeLeaveBreakPeriodsSnapshot_(value) {
  let periods;
  try {
    periods = typeof value === "string" ? JSON.parse(value || "[]") : value;
  } catch (error) {
    throw new Error("時間有給明細の休憩時間スナップショットが不正です");
  }
  if (!Array.isArray(periods)) throw new Error("時間有給明細の休憩時間スナップショットが不正です");
  return periods.map(period => {
    const startMinute = Number(period && period.startMinute);
    const endMinute = Number(period && period.endMinute);
    if (!Number.isInteger(startMinute) || !Number.isInteger(endMinute) || endMinute <= startMinute) {
      throw new Error("時間有給明細の休憩時間スナップショットが不正です");
    }
    return { startMinute: startMinute, endMinute: endMinute };
  });
}

// 承認・編集時はemployeesの現在値ではなく、明細に保存した申請時policyを復元する。
function resolveTimeLeavePolicyFromSegmentSnapshot_(employeeInfo, segmentRow) {
  const segment = segmentRow || {};
  const basePolicy = getCompanyLeavePolicy(employeeInfo.companyCode);
  const scheduledMinutesPerDay = Number(segment.scheduled_minutes_per_day);
  const timeLeaveUnitMinutes = Number(segment.time_leave_unit_minutes);
  const workStartMinute = Number(segment.work_start_minute);
  const workEndMinute = Number(segment.work_end_minute);
  if (![scheduledMinutesPerDay, timeLeaveUnitMinutes, workStartMinute, workEndMinute].every(Number.isInteger) ||
      scheduledMinutesPerDay <= 0 || timeLeaveUnitMinutes <= 0 || workStartMinute < 0 ||
      workEndMinute <= workStartMinute || workEndMinute > 24 * 60) {
    throw new Error("時間有給明細の勤務時間スナップショットが不正です");
  }
  return Object.assign({}, basePolicy, {
    scheduledMinutesPerDay: scheduledMinutesPerDay,
    timeLeaveUnitMinutes: timeLeaveUnitMinutes,
    workStartMinute: workStartMinute,
    workEndMinute: workEndMinute,
    breakPeriods: parseTimeLeaveBreakPeriodsSnapshot_(segment.break_periods_json)
  });
}

// 保存済み明細は必ず作成時のcalculation_versionで再検証する。
function normalizeStoredTimeLeaveSegmentForValidation_(segmentRow, parentRow, employeeInfo) {
  const segment = segmentRow || {};
  const parent = parentRow || {};
  const snapshotEmployeeInfo = Object.assign({}, employeeInfo, {
    policy: resolveTimeLeavePolicyFromSegmentSnapshot_(employeeInfo, segment)
  });
  const payload = {
    employee_id: String(parent.employee_id || "").trim(),
    leave_date: segment.leave_date,
    start_time: normalizeTimeLeaveClockValue_(segment.start_time),
    end_time: normalizeTimeLeaveClockValue_(segment.end_time),
    requested_minutes: segment.requested_minutes,
    calculation_version: segment.calculation_version,
    reason: parent.reason,
    reason_detail: parent.reason_detail
  };
  if (isCombinedHalfDayTimeLeaveRequestRow_(parent)) {
    payload.half_day = parent.half_day;
    return normalizeCombinedHalfDayTimeLeavePayload_(payload, snapshotEmployeeInfo, payload.calculation_version);
  }
  return normalizeTimeLeavePayload_(payload, snapshotEmployeeInfo, payload.calculation_version);
}

function validateTimeLeaveEntriesAgainstCandidate_(candidate, entries) {
  const policy = candidate.policy;

  (entries || []).forEach(entry => {
    if (entry.kind === "full_day") {
      throw new Error("同日に1日有給の申請があります");
    }
    if (entry.kind === "half_day") {
      const halfRange = getHalfDayOccupiedRange_(entry.half_day, policy);
      if (hasTimeOverlap(candidate.start_minute, candidate.end_minute, halfRange.startMinute, halfRange.endMinute)) {
        throw new Error("同日に重複する半日有給の申請があります");
      }
      return;
    }
    if (entry.kind === "time_hourly" && hasTimeOverlap(
      candidate.start_minute, candidate.end_minute, entry.start_minute, entry.end_minute
    )) {
      throw new Error("同日に重複する時間有給の申請があります");
    }
  });

  validateDailyPaidLeaveMinutes(
    (entries || []).concat([{ kind: "time_hourly", minutes: candidate.requested_minutes }]),
    policy.scheduledMinutesPerDay
  );
  return { ok: true };
}

function validateTimeLeaveRequestConflicts_(candidate, excludedRequestId) {
  const policy = candidate.policy;
  const entries = getDailyPaidLeaveEntries_(
    candidate.employee_id,
    candidate.leave_date,
    excludedRequestId
  );
  validateTimeLeaveEntriesAgainstCandidate_(candidate, entries);
  const annual = getTimeLeaveMinutesForFiscalYear_(
    candidate.employee_id,
    candidate.leave_date,
    policy,
    excludedRequestId
  );
  validateAnnualTimeLeaveLimit(
    annual.approvedMinutes,
    annual.pendingMinutes,
    candidate.requested_minutes,
    policy.timeLeaveAnnualLimitMinutes
  );
  const balanceReservation = validateTimeLeaveFifoBalanceAvailability_(candidate, excludedRequestId);
  return { dailyEntries: entries, annual: annual, balance_reservation: balanceReservation };
}

function getPendingPaidLeaveReservationMinutes_(employeeId, asOfDate, context, excludedRequestId) {
  const targetEmployeeId = String(employeeId || "").trim();
  const excludedId = String(excludedRequestId || "").trim();
  const targetDate = parseLocalDate(asOfDate);
  let total = 0;

  (context.requests_by_employee[targetEmployeeId] || []).forEach(rowObj => {
    const requestId = String(rowObj.request_id || "").trim();
    if (requestId === excludedId || norm(rowObj.status) !== STATUS.PENDING) return;
    const useDate = getApprovalRequestEndDate_(rowObj, context);
    if (useDate > targetDate) return;
    total += getMainApprovalRequiredMinutes_(rowObj, context);
  });
  return total;
}

// 表示専用のpending総額。未来日の申請も含めるが、取消・否認・承認済みは含めない。
function getAllPendingPaidLeaveReservationMinutes_(employeeId, context) {
  const targetEmployeeId = String(employeeId || "").trim();
  return (context.requests_by_employee[targetEmployeeId] || []).reduce((total, rowObj) => {
    if (norm(rowObj.status) !== STATUS.PENDING) return total;
    return total + getMainApprovalRequiredMinutes_(rowObj, context);
  }, 0);
}

// 通常の1日・半日・複数日申請を、時間年休／combinedと同じFIFO予約規則で検証する。
// 実残高を書き換えず、既存pendingを加味した申請時点の仮押さえだけを行う。
function validateMainPaidLeaveRequestBalanceReservation_(employeeId, startDate, endDate, days, halfDay) {
  const targetEmployeeId = String(employeeId || "").trim();
  const context = createFifoBalanceComparisonContext_(endDate);
  return validateMainPaidLeaveRequestBalanceReservationFromContext_(
    targetEmployeeId, startDate, endDate, days, halfDay, context
  );
}

// 固定データテストでも同じ予約ロジックを検証できるよう、取得済みcontextを受け取る本体を分離する。
function validateMainPaidLeaveRequestBalanceReservationFromContext_(employeeId, startDate, endDate, days, halfDay, context) {
  const targetEmployeeId = String(employeeId || "").trim();
  const fifoBalance = calculateFifoBalanceWithOpeningBalanceFromContext_(
    targetEmployeeId, endDate, context
  );
  const request = {
    request_id: "",
    employee_id: targetEmployeeId,
    type: "paid_leave",
    start_date: startDate,
    end_date: endDate,
    days: Number(days || 0),
    half_day: halfDay || ""
  };
  const requestedMinutes = getMainApprovalRequiredMinutes_(request, context);
  const pendingReservedMinutes = getPendingPaidLeaveReservationMinutes_(
    targetEmployeeId, endDate, context, ""
  );
  return validateTimeLeaveFifoReservation_(
    Number(fifoBalance.current_remaining_minutes || 0),
    pendingReservedMinutes,
    requestedMinutes
  );
}

function validateCombinedHalfDayTimeLeaveRequestConflicts_(candidate, excludedRequestId) {
  const policy = candidate.policy;
  const entries = getDailyPaidLeaveEntries_(candidate.employee_id, candidate.leave_date, excludedRequestId);
  const entriesWithPlannedHalfDay = entries.concat([{
    kind: "half_day",
    half_day: candidate.half_day,
    minutes: candidate.half_day_minutes
  }]);

  validateTimeLeaveEntriesAgainstCandidate_(candidate, entriesWithPlannedHalfDay);

  const annual = getTimeLeaveMinutesForFiscalYear_(
    candidate.employee_id,
    candidate.leave_date,
    policy,
    excludedRequestId
  );
  // 半休210分は含めず、時間年休子明細の分だけを年間上限へ算入する。
  validateAnnualTimeLeaveLimit(
    annual.approvedMinutes,
    annual.pendingMinutes,
    candidate.requested_minutes,
    policy.timeLeaveAnnualLimitMinutes
  );

  const context = createFifoBalanceComparisonContext_(candidate.leave_date);
  const fifoBalance = calculateFifoBalanceWithOpeningBalanceFromContext_(
    candidate.employee_id,
    candidate.leave_date,
    context
  );
  const requiredMinutes = candidate.half_day_minutes + candidate.requested_minutes;
  const pendingReservedMinutes = getPendingPaidLeaveReservationMinutes_(
    candidate.employee_id,
    candidate.leave_date,
    context,
    excludedRequestId
  );
  const balanceReservation = validateTimeLeaveFifoReservation_(
    Number(fifoBalance.current_remaining_minutes || 0),
    pendingReservedMinutes,
    requiredMinutes
  );

  return {
    daily_entries: entries,
    annual: annual,
    required_minutes: requiredMinutes,
    balance_reservation: balanceReservation
  };
}

// 通常の半休を時間年休の後に申請する逆順でも、同日の重複・420分超過を防ぐ。
// PARTNERの既存日数処理には適用しない。
function validateMainHalfDayRequestAgainstTimeLeave_(employeeId, leaveDate, halfDay) {
  const employeeInfo = validateTimeLeaveEmployee_(employeeId);
  const policy = employeeInfo.policy;
  const normalizedHalfDay = norm(halfDay);
  if (normalizedHalfDay !== "am" && normalizedHalfDay !== "pm") {
    throw new Error("半休区分が不正です");
  }

  const halfRange = getHalfDayOccupiedRange_(normalizedHalfDay, policy);
  const entries = getDailyPaidLeaveEntries_(employeeId, leaveDate, "");
  entries.forEach(entry => {
    if (entry.kind === "time_hourly" && hasTimeOverlap(
      halfRange.startMinute,
      halfRange.endMinute,
      entry.start_minute,
      entry.end_minute
    )) {
      throw new Error("同日に重複する時間有給の申請があります");
    }
  });
  validateDailyPaidLeaveMinutes(entries.concat([{
    kind: "half_day",
    half_day: normalizedHalfDay,
    minutes: policy.scheduledMinutesPerDay / 2
  }]), policy.scheduledMinutesPerDay);
  return { ok: true };
}

function isMinuteInBreakPeriod_(minute, policy) {
  return (policy.breakPeriods || []).some(period =>
    minute >= Number(period.startMinute) && minute < Number(period.endMinute)
  );
}

// 休憩を除いた実労働分を、開始から前方向に進める。半休範囲の解決にも使う。
function calculateWorkingEndMinute_(startMinute, workingMinutes, policy) {
  const start = Number(startMinute);
  const required = Number(workingMinutes);
  const workStart = Number(policy.workStartMinute);
  const workEnd = Number(policy.workEndMinute);
  if (!Number.isInteger(start) || !Number.isInteger(required) || required <= 0) return null;
  if (start < workStart || start >= workEnd || isMinuteInBreakPeriod_(start, policy)) return null;
  let cursor = start;
  let remaining = required;
  const breaks = (policy.breakPeriods || []).slice().sort((a, b) => a.startMinute - b.startMinute);
  while (remaining > 0 && cursor < workEnd) {
    const currentBreak = breaks.find(period =>
      cursor >= Number(period.startMinute) && cursor < Number(period.endMinute)
    );
    if (currentBreak) {
      cursor = Number(currentBreak.endMinute);
      continue;
    }
    const nextBreak = breaks.find(period => Number(period.startMinute) > cursor);
    const boundary = Math.min(workEnd, nextBreak ? Number(nextBreak.startMinute) : workEnd);
    const consumed = Math.min(boundary - cursor, remaining);
    if (consumed <= 0) return null;
    cursor += consumed;
    remaining -= consumed;
  }
  return remaining === 0 && cursor <= workEnd ? cursor : null;
}

// 休憩を除いた実労働分を、終了から後ろ向きに戻す。PM半休の開始を求める。
function calculateWorkingStartMinute_(endMinute, workingMinutes, policy) {
  const end = Number(endMinute);
  const required = Number(workingMinutes);
  const workStart = Number(policy.workStartMinute);
  const workEnd = Number(policy.workEndMinute);
  if (!Number.isInteger(end) || !Number.isInteger(required) || required <= 0) return null;
  if (end <= workStart || end > workEnd || isMinuteInBreakPeriod_(end - 1, policy)) return null;
  let cursor = end;
  let remaining = required;
  const breaks = (policy.breakPeriods || []).slice().sort((a, b) => b.startMinute - a.startMinute);
  while (remaining > 0 && cursor > workStart) {
    const currentBreak = breaks.find(period =>
      cursor > Number(period.startMinute) && cursor <= Number(period.endMinute)
    );
    if (currentBreak) {
      cursor = Number(currentBreak.startMinute);
      continue;
    }
    const previousBreak = breaks.find(period => Number(period.endMinute) < cursor);
    const boundary = Math.max(workStart, previousBreak ? Number(previousBreak.endMinute) : workStart);
    const consumed = Math.min(cursor - boundary, remaining);
    if (consumed <= 0) return null;
    cursor -= consumed;
    remaining -= consumed;
  }
  return remaining === 0 && cursor >= workStart ? cursor : null;
}

// v1は実労働分を休憩越しに消化し、v2は時計時間をそのまま加算して終了時刻を求める。
// いずれも勤務終了を越える候補は null。
function calculateTimeLeaveEndMinute_(startMinute, requestedMinutes, policy, calculationVersion) {
  const start = Number(startMinute);
  const requested = Number(requestedMinutes);
  const workStart = Number(policy.workStartMinute);
  const workEnd = Number(policy.workEndMinute);
  const version = normalizeTimeLeaveCalculationVersion_(calculationVersion);
  if (!Number.isInteger(start) || !Number.isInteger(requested) || requested <= 0) return null;
  if (start < workStart || start >= workEnd || isMinuteInBreakPeriod_(start, policy)) return null;

  if (version === TIME_LEAVE_CALCULATION_VERSION_V2) {
    const end = start + requested;
    return end <= workEnd ? end : null;
  }

  let cursor = start;
  let remaining = requested;
  const breaks = (policy.breakPeriods || []).slice().sort((a, b) => a.startMinute - b.startMinute);
  while (remaining > 0 && cursor < workEnd) {
    const currentBreak = breaks.find(period =>
      cursor >= Number(period.startMinute) && cursor < Number(period.endMinute)
    );
    if (currentBreak) {
      cursor = Number(currentBreak.endMinute);
      continue;
    }

    const nextBreak = breaks.find(period => Number(period.startMinute) > cursor);
    const nextBoundary = Math.min(workEnd, nextBreak ? Number(nextBreak.startMinute) : workEnd);
    const available = nextBoundary - cursor;
    if (available <= 0) {
      cursor = nextBreak ? Number(nextBreak.endMinute) : workEnd;
      continue;
    }
    const consumed = Math.min(available, remaining);
    cursor += consumed;
    remaining -= consumed;
  }
  if (remaining !== 0 || cursor > workEnd) return null;

  // 15:00 は午後休憩の開始点であるため、候補UIでは休憩の直前で終了する
  // 表示を避け、休憩後の 15:30 を終了時刻として返す。保存時も 13:00-15:30
  // の休憩控除後分数は 120 分となり、制度上の取得分は変わらない。
  const afternoonBreak = breaks.find(period =>
    Number(period.startMinute) === 900 && Number(period.endMinute) === 930
  );
  if (afternoonBreak && cursor === Number(afternoonBreak.startMinute)) {
    cursor = Number(afternoonBreak.endMinute);
  }
  return cursor <= workEnd ? cursor : null;
}

// Spreadsheet を読まず、与えられた最新スナップショットだけから有効候補を作る純粋関数。
function buildTimeLeaveCandidateRows_(options) {
  const opts = options || {};
  const policy = opts.policy;
  const mode = String(opts.request_mode || "").trim();
  const halfDay = norm(opts.half_day);
  const entries = Array.isArray(opts.entries) ? opts.entries : [];
  const annual = opts.annual || { approvedMinutes: 0, pendingMinutes: 0 };
  const approvedRemainingMinutes = Number(opts.approved_remaining_minutes || 0);
  const pendingReservedMinutes = Number(opts.pending_reserved_minutes || 0);
  const calculationVersion = normalizeTimeLeaveCalculationVersion_(
    opts.calculation_version || TIME_LEAVE_CALCULATION_VERSION_V2
  );
  const isCombined = mode === "half_day_time_hourly";
  const plannedHalfDayMinutes = isCombined ? Number(policy.scheduledMinutesPerDay) / 2 : 0;

  if (mode !== "time_hourly" && !isCombined) {
    throw new Error("request_mode が不正です");
  }
  if (isCombined && halfDay !== "am" && halfDay !== "pm") {
    throw new Error("half_day は am または pm で指定してください");
  }

  const entriesForCandidate = isCombined
    ? entries.concat([{ kind: "half_day", half_day: halfDay, minutes: plannedHalfDayMinutes }])
    : entries.slice();
  // 単独時間年休は運用上3時間まで。複合申請は半休210分との合計420分という
  // 既存ルールで判定するため、ここで単独の上限を流用しない。
  const maxRequestedMinutes = isCombined
    ? Number(policy.scheduledMinutesPerDay)
    : Math.min(Number(policy.scheduledMinutesPerDay), MAX_STANDALONE_TIME_LEAVE_MINUTES);
  const candidates = [];

  // 開始時刻は毎正時のみ。休憩中・既存申請との競合などは以下の既存検証で除外する。
  for (let startMinute = Number(policy.workStartMinute); startMinute < Number(policy.workEndMinute); startMinute += 60) {
    if (isMinuteInBreakPeriod_(startMinute, policy)) continue;

    for (let requestedMinutes = Number(policy.timeLeaveUnitMinutes);
      requestedMinutes <= maxRequestedMinutes;
      requestedMinutes += Number(policy.timeLeaveUnitMinutes)) {
      const endMinute = calculateTimeLeaveEndMinute_(startMinute, requestedMinutes, policy, calculationVersion);
      if (endMinute == null) continue;

      const candidate = {
        policy: policy,
        start_minute: startMinute,
        end_minute: endMinute,
        requested_minutes: requestedMinutes
      };
      try {
        validateTimeLeaveEntriesAgainstCandidate_(candidate, entriesForCandidate);
        validateAnnualTimeLeaveLimit(
          Number(annual.approvedMinutes || 0),
          Number(annual.pendingMinutes || 0),
          requestedMinutes,
          Number(policy.timeLeaveAnnualLimitMinutes)
        );
        validateTimeLeaveFifoReservation_(
          approvedRemainingMinutes,
          pendingReservedMinutes,
          plannedHalfDayMinutes + requestedMinutes
        );
      } catch (error) {
        continue;
      }

      candidates.push({
        start_time: formatMinuteAsTime_(startMinute),
        requested_minutes: requestedMinutes,
        requested_hours: requestedMinutes / 60,
        end_time: formatMinuteAsTime_(endMinute)
      });
    }
  }
  return candidates;
}

function groupTimeLeaveCandidatesByStart_(candidates) {
  const map = {};
  (candidates || []).forEach(candidate => {
    if (!map[candidate.start_time]) {
      map[candidate.start_time] = { start_time: candidate.start_time, duration_options: [] };
    }
    map[candidate.start_time].duration_options.push({
      requested_minutes: candidate.requested_minutes,
      requested_hours: candidate.requested_hours,
      end_time: candidate.end_time
    });
  });
  return Object.keys(map).sort().map(key => map[key]);
}

// 候補API専用のSpreadsheet正本読取り。USE_SUPABASE_READS の設定に関係なく
// 時間年休候補がSupabaseへ通信しないよう、必要な最小列だけを直接参照する。
function getTimeLeaveEmployeeFromSpreadsheet_(employeeId) {
  const targetEmployeeId = String(employeeId || "").trim();
  const sheet = getSheet("employees");
  const headerInfo = requireHeaders(sheet, [
    "employee_id", "company_code", "employment_status", "leave_management_target"
  ]);
  const row = sheet.getDataRange().getValues().slice(1).find(values =>
    String(rowToObject(values, headerInfo.headers).employee_id || "").trim() === targetEmployeeId
  );
  return row ? rowToObject(row, headerInfo.headers) : null;
}

function validateTimeLeaveSpreadsheetEmployee_(employeeId) {
  return validateTimeLeaveEmployee_(employeeId);
}

function getSpreadsheetCompanyCalendarMapForTimeLeave_() {
  const sheet = getSheet("company_calendar");
  const headerInfo = requireHeaders(sheet, ["date", "type"]);
  return buildCompanyCalendarMapFromRows_(
    sheet.getDataRange().getValues().slice(1).map(row => rowToObject(row, headerInfo.headers))
  );
}

function validateTimeLeaveSpreadsheetDate_(leaveDate) {
  const date = parseLocalDate(leaveDate);
  const type = getCalendarTypeForDate(date, getSpreadsheetCompanyCalendarMapForTimeLeave_());
  if (type !== CALENDAR_TYPE.WORKDAY) {
    throw new Error(formatDateValue(date) + " は " + getCalendarLabel(type) + " のため有給申請できません");
  }
  return date;
}

function createTimeLeaveCandidateFifoContext_(employeeId, leaveDate, companyCode) {
  const targetEmployeeId = String(employeeId || "").trim();
  return {
    as_of_date: leaveDate,
    calendar_map: getSpreadsheetCompanyCalendarMapForTimeLeave_(),
    grants_by_employee: getSpreadsheetPaidLeaveGrantRowsForFifoCompare_(),
    requests_by_employee: getSpreadsheetLeaveRequestRowsByEmployeeForFifoCompare_(),
    time_leave_segments_by_request: getTimeLeaveSegmentsByRequestForFifoCompare_(),
    company_code_by_employee: { [targetEmployeeId]: String(companyCode || "").trim().toUpperCase() }
  };
}

function getDailyPaidLeaveEntriesFromSpreadsheetForTimeLeaveCandidates_(employeeId, leaveDate, policy) {
  const targetEmployeeId = String(employeeId || "").trim();
  const targetDateKey = toDateKey(leaveDate);
  const parentRows = getTimeLeaveParentRows_();
  const calendarMap = getSpreadsheetCompanyCalendarMapForTimeLeave_();
  const entries = [];

  Object.keys(parentRows).forEach(requestId => {
    const rowObj = parentRows[requestId];
    if (String(rowObj.employee_id || "").trim() !== targetEmployeeId) return;
    if (!isActiveTimeLeaveStatus_(rowObj.status)) return;
    if (isTimeLeaveRequestRow_(rowObj)) return;

    expandLeaveRequestToDailyRows(
      rowObj.start_date, rowObj.end_date, rowObj.days, rowObj.half_day, calendarMap
    ).forEach(item => {
      if (toDateKey(item.date) !== targetDateKey) return;
      entries.push({
        request_id: requestId,
        kind: norm(rowObj.half_day) ? "half_day" : "full_day",
        half_day: String(rowObj.half_day || ""),
        minutes: norm(rowObj.half_day)
          ? policy.scheduledMinutesPerDay / 2
          : policy.scheduledMinutesPerDay
      });
    });
  });

  getTimeLeaveSegmentsForEmployeeDate_(targetEmployeeId, targetDateKey).forEach(item => {
    const parent = parentRows[String(item.rowObj.request_id || "").trim()];
    if (!parent || !isActiveTimeLeaveStatus_(parent.status)) return;
    entries.push({
      request_id: String(item.rowObj.request_id || "").trim(),
      time_leave_id: String(item.rowObj.time_leave_id || "").trim(),
      kind: "time_hourly",
      minutes: Number(item.rowObj.requested_minutes || 0),
      start_minute: Number(item.rowObj.start_minute),
      end_minute: Number(item.rowObj.end_minute)
    });
  });
  return entries;
}

// UI候補表示専用。保存APIはこの結果を信用せず、ロック下で同じ検証を再実行する。
function getTimeLeaveCandidates(payload) {
  const data = payload || {};
  const employeeId = String(data.employee_id || "").trim();
  if (!employeeId) throw new Error("employee_id がありません");
  if (!data.leave_date) throw new Error("leave_date がありません");

  const mode = String(data.request_mode || "").trim();
  const halfDay = norm(data.half_day);
  const employeeInfo = validateTimeLeaveSpreadsheetEmployee_(employeeId);
  const policy = employeeInfo.policy;
  const leaveDate = validateTimeLeaveSpreadsheetDate_(data.leave_date);
  if (mode !== "time_hourly" && mode !== "half_day_time_hourly") {
    throw new Error("request_mode は time_hourly または half_day_time_hourly で指定してください");
  }
  if (mode === "half_day_time_hourly" && halfDay !== "am" && halfDay !== "pm") {
    throw new Error("half_day は am または pm で指定してください");
  }
  if (mode === "time_hourly" && halfDay) {
    throw new Error("単独時間年休では half_day を指定できません");
  }

  const entries = getDailyPaidLeaveEntriesFromSpreadsheetForTimeLeaveCandidates_(employeeId, leaveDate, policy);
  const annual = getTimeLeaveMinutesForFiscalYear_(employeeId, leaveDate, policy, "");
  const context = createTimeLeaveCandidateFifoContext_(employeeId, leaveDate, employeeInfo.companyCode);
  const fifoBalance = calculateFifoBalanceWithOpeningBalanceFromContext_(employeeId, leaveDate, context);
  const pendingReservedMinutes = getPendingPaidLeaveReservationMinutes_(employeeId, leaveDate, context, "");
  const plannedHalfDayMinutes = mode === "half_day_time_hourly"
    ? policy.scheduledMinutesPerDay / 2
    : 0;
  const existingMinutes = entries.reduce((sum, entry) => sum + Number(entry.minutes || 0), 0);
  const approvedRemainingMinutes = Number(fifoBalance.current_remaining_minutes || 0);
  const candidates = buildTimeLeaveCandidateRows_({
    policy: policy,
    request_mode: mode,
    half_day: halfDay,
    entries: entries,
    annual: annual,
    approved_remaining_minutes: approvedRemainingMinutes,
    pending_reserved_minutes: pendingReservedMinutes,
    calculation_version: TIME_LEAVE_CALCULATION_VERSION_V2
  });

  return {
    ok: true,
    request_mode: mode,
    half_day: halfDay,
    policy: {
      work_start_time: formatMinuteAsTime_(policy.workStartMinute),
      work_end_time: formatMinuteAsTime_(policy.workEndMinute),
      scheduled_minutes_per_day: policy.scheduledMinutesPerDay,
      unit_minutes: policy.timeLeaveUnitMinutes
    },
    daily_summary: {
      existing_minutes: existingMinutes,
      planned_half_day_minutes: plannedHalfDayMinutes,
      available_minutes: Math.max(0, policy.scheduledMinutesPerDay - existingMinutes - plannedHalfDayMinutes)
    },
    annual_time_leave: {
      approved_minutes: annual.approvedMinutes,
      pending_minutes: annual.pendingMinutes,
      remaining_minutes: Math.max(0, policy.timeLeaveAnnualLimitMinutes - annual.approvedMinutes - annual.pendingMinutes)
    },
    fifo: {
      approved_remaining_minutes: approvedRemainingMinutes,
      pending_reserved_minutes: pendingReservedMinutes,
      available_minutes: Math.max(0, approvedRemainingMinutes - pendingReservedMinutes)
    },
    candidates: candidates,
    start_options: groupTimeLeaveCandidatesByStart_(candidates)
  };
}

// 利用者画面上部の表示専用。候補APIと同じSpreadsheet正本・年間集計を用い、
// 指定日なしで当年度の時間年休利用状況だけを返す。
function getEmployeeTimeLeaveSummary(employeeId) {
  const targetEmployeeId = String(employeeId || "").trim();
  if (!targetEmployeeId) throw new Error("employee_id がありません");
  const employeeInfo = validateTimeLeaveSpreadsheetEmployee_(targetEmployeeId);
  const policy = employeeInfo.policy;
  const annual = getTimeLeaveMinutesForFiscalYear_(targetEmployeeId, new Date(), policy, "");
  const approvedMinutes = Number(annual.approvedMinutes || 0);
  const pendingMinutes = Number(annual.pendingMinutes || 0);
  const annualLimitMinutes = Number(policy.timeLeaveAnnualLimitMinutes || 0);
  return {
    ok: true,
    fiscal_year: annual.fiscalYear,
    annual_time_leave: {
      approved_minutes: approvedMinutes,
      pending_minutes: pendingMinutes,
      remaining_minutes: Math.max(0, annualLimitMinutes - approvedMinutes - pendingMinutes),
      annual_limit_minutes: annualLimitMinutes
    }
  };
}

function buildTimeLeaveSegmentRow_(headerInfo, requestId, candidate, now) {
  const policy = candidate.policy;
  const rowObj = createEmptyRowObject(headerInfo.headers);
  rowObj.time_leave_id = Utilities.getUuid();
  rowObj.request_id = requestId;
  rowObj.employee_id = candidate.employee_id;
  rowObj.company_code = candidate.company_code;
  rowObj.leave_date = candidate.leave_date;
  rowObj.start_time = candidate.start_time;
  rowObj.end_time = candidate.end_time;
  rowObj.start_minute = candidate.start_minute;
  rowObj.end_minute = candidate.end_minute;
  rowObj.requested_minutes = candidate.requested_minutes;
  rowObj.scheduled_minutes_per_day = policy.scheduledMinutesPerDay;
  rowObj.time_leave_unit_minutes = policy.timeLeaveUnitMinutes;
  rowObj.work_start_minute = policy.workStartMinute;
  rowObj.work_end_minute = policy.workEndMinute;
  rowObj.break_periods_json = JSON.stringify(policy.breakPeriods);
  rowObj.calculation_version = candidate.calculation_version;
  rowObj.created_at = now;
  rowObj.updated_at = now;
  return rowObj;
}

function submitTimeLeaveRequest(payload) {
  const data = payload || {};
  const employeeId = String(data.employee_id || "").trim();
  if (!employeeId) throw new Error("employee_id がありません");

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  let parentSheet;
  try {
    const infrastructure = ensureTimeLeaveInfrastructure_();
    parentSheet = infrastructure.requestSheet;
    const parentHeaderInfo = requireHeaders(parentSheet, [
      "request_id", "employee_id", "request_date", "start_date", "end_date", "days", "type",
      "half_day", "reason", "reason_detail", "status", "approver_id", "approver_name",
      "approved_at", "rejected_reason", "year", "created_at", "updated_at"
    ]);
    const segmentHeaderInfo = requireHeaders(infrastructure.segmentSheet, TIME_LEAVE_SEGMENTS_HEADERS);
    const employeeInfo = validateTimeLeaveEmployee_(employeeId);
    const candidate = normalizeTimeLeavePayload_(data, employeeInfo, TIME_LEAVE_CALCULATION_VERSION_V2);
    validateNewStandaloneTimeLeaveMaximum_(candidate);
    validateTimeLeaveRequestConflicts_(candidate, "");

    const now = new Date();
    const requestId = Utilities.getUuid();
    const parent = createEmptyRowObject(parentHeaderInfo.headers);
    parent.request_id = requestId;
    parent.employee_id = employeeId;
    parent.request_date = now;
    parent.start_date = candidate.leave_date;
    parent.end_date = candidate.leave_date;
    parent.days = 0;
    parent.type = "paid_leave";
    parent.half_day = "";
    parent.reason = candidate.reason;
    parent.reason_detail = candidate.reason_detail;
    parent.status = STATUS.PENDING;
    parent.approver_id = "";
    parent.approver_name = "";
    parent.approved_at = "";
    parent.rejected_reason = "";
    parent.year = getFiscalYearFromDateWithStart(candidate.leave_date, candidate.policy.fiscalStartMonth);
    parent.created_at = now;
    parent.updated_at = now;
    parent.request_kind = "time_hourly";
    parent.company_code_snapshot = candidate.company_code;
    parent.policy_version = candidate.policy.policyVersion;

    appendRowFast_(parentSheet, objectToRow(parent, parentHeaderInfo.headers));
    const segment = buildTimeLeaveSegmentRow_(segmentHeaderInfo, requestId, candidate, now);
    try {
      appendRowFast_(infrastructure.segmentSheet, objectToRow(segment, segmentHeaderInfo.headers));
    } catch (segmentError) {
      try {
        if (!deleteLeaveRequestById_(requestId)) {
          throw new Error("保存直後の親申請が見つかりません");
        }
      } catch (rollbackError) {
        throw new Error("時間有給明細の保存に失敗し、親申請のロールバックにも失敗しました: " +
          segmentError.message + " / " + rollbackError.message);
      }
      throw segmentError;
    }

    appendUsageLog({
      request_id: requestId,
      action_type: "time_leave_submit",
      operator_id: employeeId,
      operator_name: "申請者",
      comment: "Time leave request submitted: " + candidate.leave_date_key + " " +
        candidate.start_time + "-" + candidate.end_time + " / " + candidate.requested_minutes + "分"
    });
    clearAppCache();
    return { ok: true, request_id: requestId, time_leave_id: segment.time_leave_id, requested_minutes: candidate.requested_minutes };
  } finally {
    lock.releaseLock();
  }
}

// 半休（210分）と時間年休を、1件の親申請と1件の時間明細として原子的に登録する。
// 親 request_id が複合申請全体の識別子となるため、request_group_id は使用しない。
function submitCombinedHalfDayAndTimeLeaveRequest(payload) {
  const data = payload || {};
  const employeeId = String(data.employee_id || "").trim();
  if (!employeeId) throw new Error("employee_id がありません");

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const infrastructure = ensureTimeLeaveInfrastructure_();
    const parentSheet = infrastructure.requestSheet;
    const parentHeaderInfo = requireHeaders(parentSheet, [
      "request_id", "employee_id", "request_date", "start_date", "end_date", "days", "type",
      "half_day", "reason", "reason_detail", "status", "approver_id", "approver_name",
      "approved_at", "rejected_reason", "year", "created_at", "updated_at"
    ]);
    const segmentHeaderInfo = requireHeaders(infrastructure.segmentSheet, TIME_LEAVE_SEGMENTS_HEADERS);
    const employeeInfo = validateTimeLeaveEmployee_(employeeId);
    const candidate = normalizeCombinedHalfDayTimeLeavePayload_(data, employeeInfo, TIME_LEAVE_CALCULATION_VERSION_V2);
    const validation = validateCombinedHalfDayTimeLeaveRequestConflicts_(candidate, "");

    const now = new Date();
    const requestId = Utilities.getUuid();
    const parent = createEmptyRowObject(parentHeaderInfo.headers);
    parent.request_id = requestId;
    parent.employee_id = employeeId;
    parent.request_date = now;
    parent.start_date = candidate.leave_date;
    parent.end_date = candidate.leave_date;
    parent.days = 0.5;
    parent.type = "paid_leave";
    parent.half_day = candidate.half_day;
    parent.reason = candidate.reason;
    parent.reason_detail = candidate.reason_detail;
    parent.status = STATUS.PENDING;
    parent.approver_id = "";
    parent.approver_name = "";
    parent.approved_at = "";
    parent.rejected_reason = "";
    parent.year = getFiscalYearFromDateWithStart(candidate.leave_date, candidate.policy.fiscalStartMonth);
    parent.created_at = now;
    parent.updated_at = now;
    parent.request_kind = "half_day_time_hourly";
    parent.company_code_snapshot = candidate.company_code;
    parent.policy_version = candidate.policy.policyVersion;

    appendRowFast_(parentSheet, objectToRow(parent, parentHeaderInfo.headers));
    const segment = buildTimeLeaveSegmentRow_(segmentHeaderInfo, requestId, candidate, now);
    try {
      appendRowFast_(infrastructure.segmentSheet, objectToRow(segment, segmentHeaderInfo.headers));
    } catch (segmentError) {
      try {
        if (!deleteLeaveRequestById_(requestId)) {
          throw new Error("保存直後の複合親申請が見つかりません");
        }
      } catch (rollbackError) {
        throw new Error("複合申請の時間明細保存に失敗し、親申請のロールバックにも失敗しました: " +
          segmentError.message + " / " + rollbackError.message);
      }
      throw segmentError;
    }

    appendUsageLog({
      request_id: requestId,
      action_type: "combined_half_day_time_leave_submit",
      operator_id: employeeId,
      operator_name: "申請者",
      comment: "Combined half-day and time leave submitted: " + candidate.leave_date_key +
        " / " + candidate.half_day + " half-day / " + candidate.start_time + "-" +
        candidate.end_time + " / 合計" + validation.required_minutes + "分"
    });
    clearAppCache();
    return {
      ok: true,
      request_id: requestId,
      time_leave_id: segment.time_leave_id,
      half_day_minutes: candidate.half_day_minutes,
      requested_minutes: candidate.requested_minutes,
      total_requested_minutes: validation.required_minutes
    };
  } finally {
    lock.releaseLock();
  }
}

function updatePendingTimeLeaveRequest(requestId, employeeId, payload) {
  const targetRequestId = String(requestId || "").trim();
  const targetEmployeeId = String(employeeId || "").trim();
  if (!targetRequestId) throw new Error("requestId がありません");
  if (!targetEmployeeId) throw new Error("employeeId がありません");

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const parentRecord = getLeaveRequestRecordById_(targetRequestId);
    if (!parentRecord || !isTimeLeaveRequestRow_(parentRecord.rowObj)) {
      throw new Error("対象の時間有給申請が見つかりません");
    }
    if (String(parentRecord.rowObj.employee_id || "").trim() !== targetEmployeeId) {
      throw new Error("申請者が一致しません");
    }
    if (norm(parentRecord.rowObj.status) !== STATUS.PENDING) {
      throw new Error("承認待ちの時間有給申請だけ修正できます");
    }

    const segmentRecord = getTimeLeaveSegmentRecordByRequestId_(targetRequestId);
    if (!segmentRecord) throw new Error("時間有給明細が見つかりません");
    const currentEmployeeInfo = validateTimeLeaveEmployee_(targetEmployeeId);
    const employeeInfo = Object.assign({}, currentEmployeeInfo, {
      policy: resolveTimeLeavePolicyFromSegmentSnapshot_(currentEmployeeInfo, segmentRecord.rowObj)
    });
    const input = Object.assign({}, payload || {}, { employee_id: targetEmployeeId });
    const calculationVersion = normalizeTimeLeaveCalculationVersion_(segmentRecord.rowObj.calculation_version);
    const candidate = normalizeTimeLeavePayload_(input, employeeInfo, calculationVersion);
    validateTimeLeaveRequestConflicts_(candidate, targetRequestId);

    const now = new Date();
    const previousParent = objectToRow(parentRecord.rowObj, parentRecord.headerInfo.headers);
    const previousSegment = objectToRow(segmentRecord.rowObj, segmentRecord.headerInfo.headers);
    const parent = Object.assign({}, parentRecord.rowObj, {
      start_date: candidate.leave_date,
      end_date: candidate.leave_date,
      days: 0,
      half_day: "",
      reason: candidate.reason,
      reason_detail: candidate.reason_detail,
      year: getFiscalYearFromDateWithStart(candidate.leave_date, candidate.policy.fiscalStartMonth),
      company_code_snapshot: candidate.company_code,
      policy_version: candidate.policy.policyVersion,
      updated_at: now
    });
    const segment = buildTimeLeaveSegmentRow_(
      segmentRecord.headerInfo,
      targetRequestId,
      candidate,
      now
    );
    segment.time_leave_id = segmentRecord.rowObj.time_leave_id;
    segment.created_at = segmentRecord.rowObj.created_at;

    updateSheetRowFast_(
      parentRecord.sheet,
      parentRecord.sheetRow,
      objectToRow(parent, parentRecord.headerInfo.headers)
    );
    try {
      updateSheetRowFast_(
        segmentRecord.sheet,
        segmentRecord.sheetRow,
        objectToRow(segment, segmentRecord.headerInfo.headers)
      );
    } catch (segmentError) {
      try {
        updateSheetRowFast_(parentRecord.sheet, parentRecord.sheetRow, previousParent);
        updateSheetRowFast_(segmentRecord.sheet, segmentRecord.sheetRow, previousSegment);
      } catch (rollbackError) {
        throw new Error("時間有給編集に失敗し、ロールバックにも失敗しました: " +
          segmentError.message + " / " + rollbackError.message);
      }
      throw segmentError;
    }

    appendUsageLog({
      request_id: targetRequestId,
      action_type: "time_leave_update",
      operator_id: targetEmployeeId,
      operator_name: "申請者",
      comment: "Pending time leave request updated: " + candidate.leave_date_key + " " +
        candidate.start_time + "-" + candidate.end_time + " / " + candidate.requested_minutes + "分"
    });
    clearAppCache();
    return { ok: true, request_id: targetRequestId, requested_minutes: candidate.requested_minutes };
  } finally {
    lock.releaseLock();
  }
}

function validatePendingTimeLeaveRequestForApproval_(requestId, stageCallback) {
  const notifyStage = (stage, fields) => {
    if (typeof stageCallback === "function") stageCallback(stage, fields || {});
  };
  notifyStage("TIME_LEAVE_APPROVAL_PARENT_LOADING");
  const parentRecord = getLeaveRequestRecordById_(requestId);
  if (!parentRecord || !isTimeLeaveSegmentRequestRow_(parentRecord.rowObj)) {
    throw new Error("対象の時間有給申請が見つかりません");
  }
  if (norm(parentRecord.rowObj.status) !== STATUS.PENDING) {
    throw new Error("承認待ちの時間有給申請だけ承認できます");
  }
  notifyStage("TIME_LEAVE_APPROVAL_PARENT_LOADED", {
    employee_id: String(parentRecord.rowObj.employee_id || "").trim(),
    request_kind: String(parentRecord.rowObj.request_kind || "").trim(),
    current_status: String(parentRecord.rowObj.status || "").trim(),
    half_day: String(parentRecord.rowObj.half_day || "").trim()
  });
  notifyStage("TIME_LEAVE_APPROVAL_SEGMENT_LOADING");
  const segmentRecord = getTimeLeaveSegmentRecordByRequestId_(requestId);
  if (!segmentRecord) throw new Error("時間有給明細が見つかりません");
  notifyStage("TIME_LEAVE_APPROVAL_SEGMENT_LOADED", {
    employee_id: String(parentRecord.rowObj.employee_id || "").trim(),
    start_time_source_type: segmentRecord.rowObj.start_time instanceof Date ? "Date" : typeof segmentRecord.rowObj.start_time,
    end_time_source_type: segmentRecord.rowObj.end_time instanceof Date ? "Date" : typeof segmentRecord.rowObj.end_time,
    requested_minutes: Number(segmentRecord.rowObj.requested_minutes || 0),
    calculation_version: String(segmentRecord.rowObj.calculation_version || "").trim(),
    scheduled_minutes_per_day: Number(segmentRecord.rowObj.scheduled_minutes_per_day || 0),
    work_start_minute: Number(segmentRecord.rowObj.work_start_minute || 0),
    work_end_minute: Number(segmentRecord.rowObj.work_end_minute || 0)
  });
  const employeeId = String(parentRecord.rowObj.employee_id || "").trim();
  const employeeInfo = validateTimeLeaveEmployee_(employeeId);
  const candidate = normalizeStoredTimeLeaveSegmentForValidation_(
    segmentRecord.rowObj, parentRecord.rowObj, employeeInfo
  );
  let validationResult;
  notifyStage("TIME_LEAVE_APPROVAL_REQUEST_VALIDATING", {
    employee_id: employeeId,
    request_kind: String(parentRecord.rowObj.request_kind || "").trim(),
    half_day: String(parentRecord.rowObj.half_day || "").trim()
  });
  if (isCombinedHalfDayTimeLeaveRequestRow_(parentRecord.rowObj)) {
    validationResult = validateCombinedHalfDayTimeLeaveRequestConflicts_(candidate, requestId);
  } else {
    validationResult = validateTimeLeaveRequestConflicts_(candidate, requestId);
  }
  notifyStage("TIME_LEAVE_APPROVAL_REQUEST_VALIDATED", {
    employee_id: employeeId,
    request_kind: String(parentRecord.rowObj.request_kind || "").trim(),
    half_day: String(parentRecord.rowObj.half_day || "").trim(),
    requested_minutes: candidate.requested_minutes,
    calculation_version: candidate.calculation_version,
    scheduled_minutes_per_day: candidate.policy.scheduledMinutesPerDay,
    work_start_minute: candidate.policy.workStartMinute,
    work_end_minute: candidate.policy.workEndMinute,
    required_minutes: candidate.half_day_minutes
      ? candidate.half_day_minutes + candidate.requested_minutes
      : candidate.requested_minutes,
    annual_time_leave_minutes_before_current: validationResult && validationResult.annual
      ? Number(validationResult.annual.approvedMinutes || 0) + Number(validationResult.annual.pendingMinutes || 0)
      : null,
    annual_time_leave_minutes: validationResult && validationResult.annual
      ? Number(validationResult.annual.approvedMinutes || 0) + Number(validationResult.annual.pendingMinutes || 0) + candidate.requested_minutes
      : null
  });
  return {
    parentRecord: parentRecord,
    segmentRecord: segmentRecord,
    candidate: candidate,
    validation_result: validationResult
  };
}

function logTimeLeaveApprovalDiagnostic_(eventName, requestId, fields) {
  const payload = Object.assign({
    event: eventName,
    request_id: String(requestId || "").trim()
  }, fields || {});
  console.log(JSON.stringify(payload));
}

function approveTimeLeaveRequest_(requestId, adminUser) {
  const lock = LockService.getScriptLock();
  let lockAcquired = false;
  let stage = "TIME_LEAVE_APPROVAL_START";
  let statusUpdated = false;
  try {
    logTimeLeaveApprovalDiagnostic_(stage, requestId, { status_updated: statusUpdated });
    stage = "TIME_LEAVE_APPROVAL_LOCK_WAIT";
    lock.waitLock(30000);
    lockAcquired = true;

    stage = "TIME_LEAVE_APPROVAL_BALANCE_VALIDATING";
    const balanceResults = validateMainApprovalBalancesForRequests_([requestId]);
    const balanceResult = balanceResults[0] || {};
    stage = "TIME_LEAVE_APPROVAL_BALANCE_VALIDATED";
    logTimeLeaveApprovalDiagnostic_(stage, requestId, {
      required_minutes: balanceResult.required_minutes || null,
      status_updated: statusUpdated
    });

    const validation = validatePendingTimeLeaveRequestForApproval_(requestId, (eventName, fields) => {
      stage = eventName;
      logTimeLeaveApprovalDiagnostic_(stage, requestId, Object.assign({ status_updated: statusUpdated }, fields));
    });
    const headerInfo = validation.parentRecord.headerInfo;
    const rowObj = validation.parentRecord.rowObj;
    const isCombined = isCombinedHalfDayTimeLeaveRequestRow_(rowObj);
    const now = new Date();
    const operatorId = adminUser && adminUser.admin_id ? String(adminUser.admin_id).trim() : "admin";
    const operatorName = adminUser && adminUser.admin_name ? String(adminUser.admin_name).trim() : "管理者";
    rowObj.status = STATUS.APPROVED;
    rowObj.approver_id = operatorId;
    rowObj.approver_name = operatorName;
    rowObj.approved_at = now;
    rowObj.updated_at = now;
    stage = "TIME_LEAVE_APPROVAL_STATUS_UPDATING";
    updateSheetRowFast_(
      validation.parentRecord.sheet,
      validation.parentRecord.sheetRow,
      objectToRow(rowObj, headerInfo.headers)
    );
    statusUpdated = true;
    stage = "TIME_LEAVE_APPROVAL_STATUS_UPDATED";
    logTimeLeaveApprovalDiagnostic_(stage, requestId, {
      employee_id: String(rowObj.employee_id || "").trim(),
      request_kind: String(rowObj.request_kind || "").trim(),
      current_status: STATUS.APPROVED,
      status_updated: statusUpdated
    });

    stage = "TIME_LEAVE_APPROVAL_USAGE_LOG_CREATING";
    appendUsageLog({
      request_id: requestId,
      action_type: isCombined ? "combined_half_day_time_leave_approve" : "time_leave_approve",
      operator_id: operatorId,
      operator_name: operatorName,
      comment: (isCombined ? "Combined half-day and time leave" : "Time leave") + " approved by " + operatorName
    });
    stage = "TIME_LEAVE_APPROVAL_USAGE_LOG_CREATED";
    logTimeLeaveApprovalDiagnostic_(stage, requestId, { status_updated: statusUpdated });

    clearAppCache();
    stage = "TIME_LEAVE_APPROVAL_COMPLETE";
    logTimeLeaveApprovalDiagnostic_(stage, requestId, { status_updated: statusUpdated });
    return { ok: true };
  } catch (error) {
    const message = error && error.message ? error.message : String(error || "不明なエラー");
    console.error(JSON.stringify({
      event: "TIME_LEAVE_APPROVAL_ERROR",
      request_id: String(requestId || "").trim(),
      stage: stage,
      status_updated: statusUpdated,
      error_message: message
    }), error);
    throw error;
  } finally {
    if (lockAcquired) lock.releaseLock();
  }
}

function approveRequestsBatchWithFifoValidation_(requestIds, adminUser) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    validateMainApprovalBalancesForRequests_(requestIds);
    const timeRequestIds = requestIds.filter(requestId => isTimeLeaveRequestById_(requestId));
    // 状態を書き換える前に全時間有給を検証する。相互のpending申請も検証対象に残る。
    timeRequestIds.forEach(requestId => validatePendingTimeLeaveRequestForApproval_(requestId));

    const sheet = getSheet("leave_requests");
    const headerInfo = requireHeaders(sheet, [
      "request_id", "status", "approver_id", "approver_name", "approved_at", "updated_at"
    ]);
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow <= 1) throw new Error("申請データがありません");

    const data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
    const targetIdSet = new Set(requestIds.map(id => String(id)));
    const now = new Date();
    const operatorId = adminUser && adminUser.admin_id ? String(adminUser.admin_id).trim() : "admin";
    const operatorName = adminUser && adminUser.admin_name ? String(adminUser.admin_name).trim() : "管理者";
    let updatedCount = 0;
    const updatedRows = data.slice(1).map(row => {
      if (!targetIdSet.has(String(row[headerInfo.map.request_id] || ""))) return row;
      row[headerInfo.map.status] = STATUS.APPROVED;
      row[headerInfo.map.approver_id] = operatorId;
      row[headerInfo.map.approver_name] = operatorName;
      row[headerInfo.map.approved_at] = now;
      row[headerInfo.map.updated_at] = now;
      updatedCount++;
      return row;
    });
    if (updatedCount === 0) throw new Error("承認対象の申請が見つかりません");
    sheet.getRange(2, 1, updatedRows.length, lastCol).setValues(updatedRows);

    requestIds.forEach(requestId => {
      const requestRecord = getLeaveRequestRecordById_(requestId);
      const isCombined = !!(requestRecord && isCombinedHalfDayTimeLeaveRequestRow_(requestRecord.rowObj));
      appendUsageLog({
      request_id: requestId,
      action_type: isCombined
        ? "combined_half_day_time_leave_approve"
        : (timeRequestIds.includes(requestId) ? "time_leave_approve" : "approve"),
      operator_id: operatorId,
      operator_name: operatorName,
      comment: "Batch approved by " + operatorName
      });
    });
    clearAppCache();
    return { ok: true, count: updatedCount };
  } finally {
    lock.releaseLock();
  }
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
  const isMainEmployee = employee && String(employee.company_code || "").trim().toUpperCase() === "MAIN";

  if (isMainEmployee) {
    validateMainPaidLeaveRequestBalanceReservation_(
      String(data.employee_id || "").trim(),
      start,
      isHalf ? start : end,
      days,
      isHalf ? (data.half_type || "") : ""
    );
  }

  if (isHalf && isMainEmployee) {
    validateMainHalfDayRequestAgainstTimeLeave_(
      String(data.employee_id || "").trim(),
      start,
      data.half_type || ""
    );
  }

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

  tryInsertLeaveRequestToSupabase_(rowObj);

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
  if (shouldUseSupabaseReads_()) {
    const employeeMap = {};

    getEmployeesFromSupabase_().forEach(rowObj => {
      const employeeId = String(rowObj.employee_id || "").trim();
      if (!employeeId) return;

      employeeMap[employeeId] = {
        display_employee_id: String(rowObj.display_employee_id || "").trim(),
        employee_name: String(rowObj.display_name || rowObj.name || employeeId).trim(),
        company_name: String(rowObj.company_name || "").trim(),
        department: String(rowObj.department || "").trim()
      };
    });

    const pendingRequests = getLeaveRequestsFromSupabase_()
      .map(rowObj => {
        const rowStatus = norm(rowObj.status || STATUS.PENDING);
        const startDate = rowObj.start_date;
        const endDate = rowObj.end_date;

        if (rowStatus !== STATUS.PENDING) return null;
        if (!startDate || !endDate) return null;
        if (!isRequestOnOrAfterDate({ start_date: startDate, end_date: endDate }, range.start)) return null;

        const employeeId = String(rowObj.employee_id || "").trim();
        const employee = employeeMap[employeeId] || {};
        const startText = formatDateValue(startDate);
        const endText = formatDateValue(endDate);
        const halfDay = String(rowObj.half_day || "");

        return {
          request_id: String(rowObj.request_id || ""),
          employee_id: employeeId,
          display_employee_id: employee.display_employee_id || "",
          employee_name: employee.employee_name || employeeId || "Unknown",
          start_date: startText,
          end_date: endText,
          days: rowObj.days || 0,
          type: String(rowObj.type || "paid_leave"),
          half_day: halfDay,
          request_kind: String(rowObj.request_kind || ""),
          reason: String(rowObj.reason || ""),
          reason_detail: String(rowObj.reason_detail || ""),
          status: rowStatus,
          request_date: formatDateValue(rowObj.request_date),
          created_at: rowObj.created_at ? formatDateValue(rowObj.created_at) : "",
          updated_at: rowObj.updated_at ? formatDateValue(rowObj.updated_at) : "",
          year: rowObj.year || "",
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
    return enrichPendingRequestsForAdminTimeLeavePresentation_(pendingRequests);
  }

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
  if ("request_kind" in requestHeaderInfo.map) requestColumns.push("request_kind");
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
  const requestKindIdx = "request_kind" in requestHeaderInfo.map ? reqIdx("request_kind") : -1;

  const pendingRequests = values
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
        request_kind: requestKindIdx >= 0 ? String(row[requestKindIdx] || "") : "",
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
  return enrichPendingRequestsForAdminTimeLeavePresentation_(pendingRequests);
}

// 承認待ち軽量一覧専用。対象request_id群の子明細だけを一度に読み、表示情報を付与する。
// Supabase読取り中でも時間年休明細は既存方針どおりSpreadsheetを正とする。
function enrichPendingRequestsForAdminTimeLeavePresentation_(requests) {
  const rows = Array.isArray(requests) ? requests : [];
  const requestIds = rows.map(row => String(row.request_id || "").trim()).filter(Boolean);
  const segmentsByRequest = getTimeLeaveSegmentPresentationByRequestIds_(requestIds);
  return rows.map(row => buildPendingAdminTimeLeavePresentationRow_(row, segmentsByRequest[String(row.request_id || "").trim()]));
}

function buildPendingAdminTimeLeavePresentationRow_(request, segment) {
  const row = Object.assign({}, request || {});
  const originalKind = String(row.request_kind || "").trim();
  let effectiveKind = originalKind;
  const normalizedKind = norm(originalKind);

  // 現行Supabase schemaにrequest_kindがない場合だけ、既存Spreadsheet子明細とhalf_dayから
  // 表示用に復元する。保存データ・Supabaseデータは変更しない。
  if (!normalizedKind && segment) {
    const halfDay = norm(row.half_day);
    effectiveKind = halfDay === "am" || halfDay === "pm" ? "half_day_time_hourly" : "time_hourly";
  }

  const isTimeLeave = norm(effectiveKind) === "time_hourly" ||
    norm(effectiveKind) === "half_day_time_hourly";
  row.request_kind = effectiveKind;
  row.time_leave_presentation = isTimeLeave
    ? buildTimeLeaveHistoryPresentation_(segment, row.half_day)
    : null;
  return row;
}

function getTimeLeaveSegmentPresentationByRequestIds_(requestIds) {
  const targetIds = new Set((requestIds || []).map(id => String(id || "").trim()).filter(Boolean));
  if (targetIds.size === 0) return {};
  const sheet = getAppSpreadsheet().getSheetByName(TIME_LEAVE_SEGMENTS_SHEET);
  if (!sheet || sheet.getLastRow() <= 1) return {};
  const headerInfo = requireHeaders(sheet, [
    "request_id", "start_time", "end_time", "requested_minutes", "calculation_version"
  ]);
  const minimumColumn = Math.min(
    headerInfo.map.request_id,
    headerInfo.map.start_time,
    headerInfo.map.end_time,
    headerInfo.map.requested_minutes,
    headerInfo.map.calculation_version
  );
  const maximumColumn = Math.max(
    headerInfo.map.request_id,
    headerInfo.map.start_time,
    headerInfo.map.end_time,
    headerInfo.map.requested_minutes,
    headerInfo.map.calculation_version
  );
  const values = sheet.getRange(2, minimumColumn + 1, sheet.getLastRow() - 1, maximumColumn - minimumColumn + 1).getValues();
  const result = {};
  values.forEach(row => {
    const getValue = header => row[headerInfo.map[header] - minimumColumn];
    const requestId = String(getValue("request_id") || "").trim();
    if (!targetIds.has(requestId)) return;
    result[requestId] = {
      request_id: requestId,
      start_time: getValue("start_time"),
      end_time: getValue("end_time"),
      requested_minutes: getValue("requested_minutes"),
      calculation_version: getValue("calculation_version")
    };
  });
  return result;
}

/* =========================
   管理画面用：申請検索
========================= */
function searchRequests(filters) {
  filters = filters || {};
  const timeLeaveSegments = getTimeLeaveSegmentPresentationByRequest_();

  if (shouldUseSupabaseReads_()) {
    const rows = getLeaveRequestsFromSupabase_();
    if (rows.length === 0) return [];

    const employeeMap = getEmployeeDetailMap();
    const targetStatus = norm(filters.status || "");
    const keyword = norm(filters.employeeKeyword || "");

    const startFilter = filters.start_date ? parseLocalDate(filters.start_date) : null;
    const endFilter = filters.end_date ? parseLocalDate(filters.end_date) : null;

    const fiscalYears = [...new Set(
      rows
        .map(rowObj => {
          if (!rowObj.start_date) return null;
          return getFiscalYearFromDate(rowObj.start_date);
        })
        .filter(v => v != null)
    )];

    const balanceMapByYear = {};
    fiscalYears.forEach(year => {
      balanceMapByYear[year] = getEmployeeBalanceMapForFiscalYear(year);
    });

    return rows
      .map(rowObj => {
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

        const timeLeavePresentation = buildTimeLeaveHistoryPresentation_(
          timeLeaveSegments[String(rowObj.request_id || "").trim()], rowObj.half_day
        );
        const timeLeaveDetail = timeLeavePresentation
          ? formatTimeLeavePresentation_(timeLeaveSegments[String(rowObj.request_id || "").trim()], rowObj.half_day)
          : "";
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
          time_leave_detail: timeLeaveDetail,
          time_leave_presentation: timeLeavePresentation,
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

      const timeLeavePresentation = buildTimeLeaveHistoryPresentation_(
        timeLeaveSegments[String(rowObj.request_id || "").trim()], rowObj.half_day
      );
      const timeLeaveDetail = timeLeavePresentation
        ? formatTimeLeavePresentation_(timeLeaveSegments[String(rowObj.request_id || "").trim()], rowObj.half_day)
        : "";
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
        time_leave_detail: timeLeaveDetail,
        time_leave_presentation: timeLeavePresentation,
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
  const timeLeaveSegments = getTimeLeaveSegmentPresentationByRequest_();

  if (!targetEmployeeId) {
    return [];
  }

  if (shouldUseSupabaseReads_()) {
    return getLeaveRequestsFromSupabase_()
      .filter(rowObj => {
        return (
          String(rowObj.employee_id || "").trim() === targetEmployeeId &&
          rowObj.start_date &&
          rowObj.end_date
        );
      })
      .map(rowObj => ({
        requestId: String(rowObj.request_id || ""),
        startDate: parseLocalDate(rowObj.start_date),
        endDate: parseLocalDate(rowObj.end_date),
        createdAt: rowObj.created_at ? parseLocalDate(rowObj.created_at) : parseLocalDate(rowObj.start_date),
        days: rowObj.days || 0,
        halfDay: String(rowObj.half_day || ""),
        reason: String(rowObj.reason || ""),
        reasonDetail: String(rowObj.reason_detail || ""),
        status: norm(rowObj.status || STATUS.PENDING),
        timeLeaveSegment: timeLeaveSegments[String(rowObj.request_id || "").trim()]
      }))
      .sort((a, b) => {
        const startDiff = b.startDate.getTime() - a.startDate.getTime();
        if (startDiff !== 0) return startDiff;
        return b.createdAt.getTime() - a.createdAt.getTime();
      })
      .slice(0, maxRows)
      .map(row => {
        const startText = formatDateValue(row.startDate);
        const endText = formatDateValue(row.endDate);
        const timeLeaveHistory = buildTimeLeaveHistoryPresentation_(row.timeLeaveSegment, row.halfDay);

        return {
          request_id: row.requestId,
          start_date: toDateKey(row.startDate),
          end_date: toDateKey(row.endDate),
          date_label: startText !== endText ? startText + " 〜 " + endText : startText,
          leave_type_label: timeLeaveHistory
            ? timeLeaveHistory.leave_type_label
            : getRequestHistoryLeaveTypeLabel_(row.halfDay, row.startDate, row.endDate),
          time_leave_history: timeLeaveHistory,
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
      status: norm(row[map.status] || STATUS.PENDING),
      timeLeaveSegment: timeLeaveSegments[String(row[map.request_id] || "").trim()]
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
    const timeLeaveHistory = buildTimeLeaveHistoryPresentation_(row.timeLeaveSegment, row.halfDay);

    return {
      request_id: row.requestId,
      start_date: toDateKey(row.startDate),
      end_date: toDateKey(row.endDate),
      date_label: startText !== endText ? startText + " 〜 " + endText : startText,
      leave_type_label: timeLeaveHistory
        ? timeLeaveHistory.leave_type_label
        : getRequestHistoryLeaveTypeLabel_(row.halfDay, row.startDate, row.endDate),
      time_leave_history: timeLeaveHistory,
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
  if (value === STATUS.CANCELED_BY_ADMIN) return "管理者取消済み";
  if (value === STATUS.CANCELED) return "取消済み";

  return "承認待ち";
}

// 表示専用。時間年休の明細はSpreadsheetを正とし、Supabaseの読取り設定には従わない。
function getTimeLeaveSegmentPresentationByRequest_() {
  const sheet = getAppSpreadsheet().getSheetByName(TIME_LEAVE_SEGMENTS_SHEET);
  if (!sheet || sheet.getLastRow() <= 1) return {};
  const headerInfo = requireHeaders(sheet, [
    "request_id", "start_time", "end_time", "requested_minutes", "calculation_version"
  ]);
  const result = {};
  sheet.getDataRange().getValues().slice(1).forEach(row => {
    const item = rowToObject(row, headerInfo.headers);
    const requestId = String(item.request_id || "").trim();
    if (!requestId) return;
    result[requestId] = item;
  });
  return result;
}

function formatTimeLeavePresentation_(segment, halfDay) {
  const presentation = buildTimeLeaveHistoryPresentation_(segment, halfDay);
  if (!presentation) return "";
  return presentation.leave_type_label + " " + presentation.start_time + "〜" +
    presentation.end_time + "（" + formatTimeLeaveMinutesForPresentation_(presentation.requested_minutes) + "）";
}

// Spreadsheet の時刻セルは Date として返る。表示用APIでは日付・タイムゾーンを含めず必ず HH:mm にする。
function formatTimeLeaveClockForPresentation_(value, timezone) {
  if (value instanceof Date) {
    try {
      return normalizeTimeLeaveClockValue_(value, timezone);
    } catch (error) {
      return "";
    }
  }

  // 表示は従来どおり9:00も09:00として見せる。一方、業務計算は上記関数で厳格なHH:mmだけを許可する。
  const match = String(value == null ? "" : value).trim().match(/^(\d{1,2}):(\d{2})(?::\d{2})?$/);
  if (!match) return "";
  const hour = Number(match[1]);
  const minute = Number(match[2]);
  if (hour < 0 || hour > 23 || minute < 0 || minute > 59) return "";
  return (hour < 10 ? "0" : "") + hour + ":" + (minute < 10 ? "0" : "") + minute;
}

function formatTimeLeaveMinutesForPresentation_(minutes) {
  const value = Number(minutes || 0);
  const hours = Math.floor(value / 60);
  const remainder = value % 60;
  return remainder ? hours + "時間" + remainder + "分" : hours + "時間";
}

function buildTimeLeaveHistoryPresentation_(segment, halfDay) {
  if (!segment) return null;
  const requestedMinutes = Number(segment.requested_minutes || 0);
  if (!Number.isInteger(requestedMinutes) || requestedMinutes <= 0) return null;
  const startTime = formatTimeLeaveClockForPresentation_(segment.start_time);
  const endTime = formatTimeLeaveClockForPresentation_(segment.end_time);
  if (!startTime || !endTime) return null;

  const normalizedHalfDay = norm(halfDay);
  const isCombined = normalizedHalfDay === "am" || normalizedHalfDay === "pm";
  const halfDayLabel = normalizedHalfDay === "am" ? "AM半休＋" :
    (normalizedHalfDay === "pm" ? "PM半休＋" : "");
  return {
    leave_type_label: halfDayLabel + "時間年休",
    start_time: startTime,
    end_time: endTime,
    requested_minutes: requestedMinutes,
    is_combined: isCombined,
    total_paid_leave_minutes: isCombined ? 210 + requestedMinutes : requestedMinutes
  };
}

function getMainApprovalRequiredMinutes_(requestRow, context) {
  const policy = getCompanyLeavePolicy("MAIN");
  const requestId = String(requestRow.request_id || "").trim();
  if (isTimeLeaveSegmentRequestRow_(requestRow)) {
    const segments = (context.time_leave_segments_by_request || {})[requestId] || [];
    const minutes = segments.reduce((sum, segment) => sum + Number(segment.requested_minutes || 0), 0);
    if (!Number.isInteger(minutes) || minutes <= 0) {
      throw new Error("時間有給明細の取得分が不正です: " + requestId);
    }
    if (isCombinedHalfDayTimeLeaveRequestRow_(requestRow)) {
      const halfDay = norm(requestRow.half_day);
      if (halfDay !== "am" && halfDay !== "pm") {
        throw new Error("複合申請の半休区分が不正です: " + requestId);
      }
      return policy.scheduledMinutesPerDay / 2 + minutes;
    }
    return minutes;
  }
  const dailyRows = expandLeaveRequestToDailyRows(
    requestRow.start_date, requestRow.end_date, requestRow.days, requestRow.half_day, context.calendar_map
  );
  return dailyRows.length * policy.scheduledMinutesPerDay -
    (norm(requestRow.half_day) ? dailyRows.length * policy.scheduledMinutesPerDay / 2 : 0);
}

function getApprovalRequestEndDate_(requestRow, context) {
  if (!isTimeLeaveSegmentRequestRow_(requestRow)) return parseLocalDate(requestRow.end_date);
  const requestId = String(requestRow.request_id || "").trim();
  const segments = (context.time_leave_segments_by_request || {})[requestId] || [];
  if (!segments.length) throw new Error("時間有給明細が見つかりません: " + requestId);
  return segments.reduce((latest, segment) => {
    const date = parseLocalDate(segment.leave_date);
    return date > latest ? date : latest;
  }, parseLocalDate(segments[0].leave_date));
}

function formatPaidLeaveMinutesForError_(minutes) {
  const display = getMinuteBalanceDisplay_(minutes, getCompanyLeavePolicy("MAIN").scheduledMinutesPerDay);
  return display.remaining_full_days + "日" + display.remaining_hours + "時間" + display.remaining_remainder_minutes + "分";
}

function validateMainApprovalRemainingMinutes_(remainingMinutes, requiredMinutes) {
  const remaining = Number(remainingMinutes || 0);
  const required = Number(requiredMinutes || 0);
  if (!Number.isInteger(remaining) || !Number.isInteger(required) || remaining < 0 || required <= 0) {
    throw new Error("承認時有給残高の分数が不正です");
  }
  if (remaining < required) {
    throw new Error(
      "有給残高が不足しているため承認できません。現在残高：" +
      formatPaidLeaveMinutesForError_(remaining) +
      " / 申請必要量：" + formatPaidLeaveMinutesForError_(required)
    );
  }
  return { ok: true, remaining_minutes: remaining, required_minutes: required };
}

// 承認処理専用。入力順ではなく利用開始日、同日なら request_id の順で仮承認する。
// 既存一括承認と同じ all-or-nothing を維持し、書込み前に全件を検証する。
function validateMainApprovalBalancesForRequests_(requestIds) {
  const targetIds = new Set((requestIds || []).map(id => String(id || "").trim()).filter(Boolean));
  if (!targetIds.size) return [];
  const context = createFifoBalanceComparisonContext_(parseLocalDate(new Date()));
  const targets = [];
  Object.keys(context.requests_by_employee || {}).forEach(employeeId => {
    (context.requests_by_employee[employeeId] || []).forEach(rowObj => {
      if (!targetIds.has(String(rowObj.request_id || "").trim())) return;
      if (!isMainTimeLeaveEmployeeForFifo_(employeeId, context)) return;
      if (norm(rowObj.status) !== STATUS.PENDING) {
        throw new Error("承認待ちの申請だけ承認できます: " + rowObj.request_id);
      }
      targets.push(rowObj);
    });
  });
  if (targets.length !== [...targetIds].filter(id => {
    return Object.keys(context.requests_by_employee || {}).some(employeeId =>
      (context.requests_by_employee[employeeId] || []).some(row => String(row.request_id || "").trim() === id)
    );
  }).length) {
    throw new Error("承認対象の申請が見つかりません");
  }
  targets.sort((a, b) => {
    const aDate = getApprovalRequestEndDate_(a, context);
    const bDate = getApprovalRequestEndDate_(b, context);
    if (aDate.getTime() !== bDate.getTime()) return aDate - bDate;
    return String(a.request_id).localeCompare(String(b.request_id));
  });

  const result = [];
  targets.forEach(rowObj => {
    const employeeId = String(rowObj.employee_id || "").trim();
    const useDate = getApprovalRequestEndDate_(rowObj, context);
    const requiredMinutes = getMainApprovalRequiredMinutes_(rowObj, context);
    const before = calculateFifoBalanceMinutesFromContext_(employeeId, useDate, context);
    validateMainApprovalRemainingMinutes_(before.current_remaining_minutes, requiredMinutes);
    rowObj.status = STATUS.APPROVED;
    const after = calculateFifoBalanceMinutesFromContext_(employeeId, useDate, context);
    if (Number(after.unallocated_used_minutes || 0) > Number(before.unallocated_used_minutes || 0)) {
      throw new Error("有給残高が不足しているため承認できません: " + rowObj.request_id);
    }
    result.push({ request_id: String(rowObj.request_id || ""), required_minutes: requiredMinutes });
  });
  return result;
}

function approveRequestsBatch(requestIds, adminUser) {
  if (!Array.isArray(requestIds) || requestIds.length === 0) {
    throw new Error("承認対象が選択されていません");
  }

  return approveRequestsBatchWithFifoValidation_(requestIds, adminUser);
}

// 旧日数のみの一括承認実装。Phase 4以降は上記の共通FIFO検証経由でのみ呼び出す。
function approveRequestsBatchLegacy_(requestIds, adminUser) {
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

  if (isTimeLeaveRequestById_(requestId)) {
    return approveTimeLeaveRequest_(requestId, adminUser);
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    validateMainApprovalBalancesForRequests_([requestId]);
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
  } finally {
    lock.releaseLock();
  }
}

/* =========================
   管理画面用：承認後取消
========================= */
function cancelApprovedRequestByAdmin(requestId, reason, adminUser) {
  const targetRequestId = String(requestId || "").trim();

  if (!targetRequestId) {
    throw new Error("requestId がありません");
  }

  const sheet = getSheet("leave_requests");
  const headerInfo = requireHeaders(sheet, [
    "request_id",
    "status",
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
    return String(row[headerInfo.map.request_id] || "").trim() === targetRequestId;
  });

  if (rowIndex === -1) {
    throw new Error("対象の申請が見つかりません");
  }

  const rowValues = data[rowIndex].slice();
  const currentStatus = norm(rowValues[headerInfo.map.status]);

  if (currentStatus !== STATUS.APPROVED) {
    throw new Error("承認済みの申請だけ管理者取消できます");
  }

  const now = new Date();
  const operatorId = adminUser && adminUser.admin_id
    ? String(adminUser.admin_id).trim()
    : "admin";
  const operatorName = adminUser && adminUser.admin_name
    ? String(adminUser.admin_name).trim()
    : "管理者";
  const cancelReason = String(reason || "").trim();
  const map = headerInfo.map;

  rowValues[map.status] = STATUS.CANCELED_BY_ADMIN;
  rowValues[map.updated_at] = now;

  if ("updated_by" in map) rowValues[map.updated_by] = operatorName;
  if ("cancel_reason" in map) rowValues[map.cancel_reason] = cancelReason;
  if ("canceled_reason" in map) rowValues[map.canceled_reason] = cancelReason;
  if ("cancelled_reason" in map) rowValues[map.cancelled_reason] = cancelReason;
  if ("canceled_at" in map) rowValues[map.canceled_at] = now;
  if ("cancelled_at" in map) rowValues[map.cancelled_at] = now;
  if ("canceled_by" in map) rowValues[map.canceled_by] = operatorName;
  if ("cancelled_by" in map) rowValues[map.cancelled_by] = operatorName;

  updateSheetRowFast_(sheet, rowIndex + 1, rowValues);

  appendUsageLog({
    request_id: targetRequestId,
    action_type: "admin_cancel_approved",
    operator_id: operatorId,
    operator_name: operatorName,
    comment: cancelReason || "Approved leave request canceled by admin"
  });

  clearAppCache();

  return {
    ok: true,
    request_id: targetRequestId
  };
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

  if (shouldUseSupabaseReads_()) {
    const rows = getUsageLogsFromSupabase_();
    if (rows.length === 0) return [];

    const keyword = norm(filters.keyword || "");
    const actionType = norm(filters.action_type || "");

    const startFilter = filters.start_date ? parseLocalDate(filters.start_date) : null;
    const endFilter = filters.end_date ? parseLocalDate(filters.end_date) : null;

    const employeeMap = getEmployeeDetailMap();
    const requestEmployeeMap = {};

    getLeaveRequestsFromSupabase_().forEach(request => {
      const requestId = String(request.request_id || "").trim();
      const employeeId = String(request.employee_id || "").trim();
      if (requestId && employeeId) requestEmployeeMap[requestId] = employeeId;
    });

    return rows
      .map(rowObj => {
        const actionDate = rowObj.action_date ? parseLocalDate(rowObj.action_date) : null;

        if (!actionDate) return null;

        if (startFilter && actionDate < startFilter) return null;
        if (endFilter && actionDate > endFilter) return null;

        const rowActionType = String(rowObj.action_type || "");
        if (actionType && norm(rowActionType) !== actionType) return null;

        const requestId = String(rowObj.request_id || "");
        const linkedRequestId = String(rowObj.leave_request_id || "").trim();
        const resolvedEmployeeId = String(rowObj.employee_id || "").trim() ||
          requestEmployeeMap[linkedRequestId] ||
          requestEmployeeMap[requestId] ||
          requestId;
        const employee = employeeMap[resolvedEmployeeId];
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
  if (shouldUseSupabaseReads_()) {
    const employeeRows = getEmployeesFromSupabase_()
      .filter(rowObj => {
        const employeeId = String(rowObj.employee_id || "").trim();
        const name = String(rowObj.name || "").trim();

        const employmentStatus = String(rowObj.employment_status || "")
          .trim()
          .toLowerCase();

        const isActive =
          employmentStatus === "active" ||
          employmentStatus === "在職";

        return employeeId && name && isActive && rowObj.leave_management_target === true;
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
    const fiveDayMapByFiscalYear = {};

    Object.keys(fiscalYearGroups).forEach(fiscalYear => {
      balanceMapByFiscalYear[fiscalYear] =
        getEmployeeBalanceMapForEmployeeIdsForFiscalYear(
          Number(fiscalYear),
          fiscalYearGroups[fiscalYear]
        );
      fiveDayMapByFiscalYear[fiscalYear] =
        getFiveDayObligationDaysByFiscalYearForEmployeeIds(
          Number(fiscalYear),
          fiscalYearGroups[fiscalYear]
        );
    });
    const currentMinuteFifoBalanceMap = getCurrentMinuteFifoBalanceMapForEmployeeIds_(
      employeeRows.map(row => String(row.employee_id || "").trim())
    );
    const currentPendingReservationMap = getCurrentMainPendingPaidLeaveReservationMapForEmployeeIds_(
      employeeRows.map(row => String(row.employee_id || "").trim())
    );

    return employeeRows.map(rowObj => {
      const employeeId = String(rowObj.employee_id || "").trim();
      const fiscalStartMonth = Number(rowObj.fiscal_start_month || 4);
      const fiscalYear = getFiscalYearFromDateWithStart(new Date(), fiscalStartMonth);

      const balanceMap = balanceMapByFiscalYear[fiscalYear] || {};
      const fiveDayMap = fiveDayMapByFiscalYear[fiscalYear] || {};
      const balance = currentMinuteFifoBalanceMap[employeeId] || balanceMap[employeeId] || {
        current_remaining_days: 0,
        carry_over_days: 0,
        grant_days: 0,
        used_days: 0
      };

      const usedDays = Number(balance.used_days || 0);
      const fiveDayUsed = Math.min(Number(fiveDayMap[employeeId] || 0), 5);
      const fiveDayRemaining = Math.max(0, 5 - fiveDayUsed);
      const confirmedMinutes = balance.current_remaining_minutes == null ? null : Number(balance.current_remaining_minutes);
      const pendingReservedMinutes = confirmedMinutes == null ? null : Number(currentPendingReservationMap[employeeId] || 0);
      const availableMinutes = confirmedMinutes == null ? null : Math.max(0, confirmedMinutes - pendingReservedMinutes);

      return {
        employee_id: employeeId,
        name: String(rowObj.name || "").trim(),
        name_kana: String(rowObj.name_kana || "").trim(),
        employment_type: String(rowObj.employment_type || "").trim(),

        fiscal_year: fiscalYear,
        fiscal_start_month: fiscalStartMonth,

        current_remaining_days: Number(balance.current_remaining_days || 0),
        current_remaining_minutes: confirmedMinutes,
        confirmed_remaining_minutes: confirmedMinutes,
        pending_reserved_minutes: pendingReservedMinutes,
        available_remaining_minutes: availableMinutes,
        remaining_full_days: balance.remaining_full_days == null ? null : Number(balance.remaining_full_days),
        remaining_hours: balance.remaining_hours == null ? null : Number(balance.remaining_hours),
        remaining_remainder_minutes: balance.remaining_remainder_minutes == null ? null : Number(balance.remaining_remainder_minutes),
        carry_over_days: Number(balance.carry_over_days || 0),
        grant_days: Number(balance.grant_days || 0),
        used_days: usedDays,
        five_day_used: fiveDayUsed,
        five_day_remaining: fiveDayRemaining,
        five_day_completed: fiveDayRemaining === 0
      };
    });
  }

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
  const fiveDayMapByFiscalYear = {};

    Object.keys(fiscalYearGroups).forEach(fiscalYear => {
    balanceMapByFiscalYear[fiscalYear] =
      getEmployeeBalanceMapForEmployeeIdsForFiscalYear(
        Number(fiscalYear),
        fiscalYearGroups[fiscalYear]
      );
    fiveDayMapByFiscalYear[fiscalYear] =
      getFiveDayObligationDaysByFiscalYearForEmployeeIds(
        Number(fiscalYear),
        fiscalYearGroups[fiscalYear]
        );
    });
  const currentMinuteFifoBalanceMap = getCurrentMinuteFifoBalanceMapForEmployeeIds_(
    employeeRows.map(row => String(row.employee_id || "").trim())
  );
  const currentPendingReservationMap = getCurrentMainPendingPaidLeaveReservationMapForEmployeeIds_(
    employeeRows.map(row => String(row.employee_id || "").trim())
  );

  return employeeRows.map(rowObj => {
    const employeeId = String(rowObj.employee_id || "").trim();
    const fiscalStartMonth = Number(rowObj.fiscal_start_month || 4);
    const fiscalYear = getFiscalYearFromDateWithStart(new Date(), fiscalStartMonth);

    const balanceMap = balanceMapByFiscalYear[fiscalYear] || {};
    const fiveDayMap = fiveDayMapByFiscalYear[fiscalYear] || {};
    const balance = currentMinuteFifoBalanceMap[employeeId] || balanceMap[employeeId] || {
      current_remaining_days: 0,
      carry_over_days: 0,
      grant_days: 0,
      used_days: 0
    };

    const usedDays = Number(balance.used_days || 0);
    const fiveDayUsed = Math.min(Number(fiveDayMap[employeeId] || 0), 5);
    const fiveDayRemaining = Math.max(0, 5 - fiveDayUsed);
    const confirmedMinutes = balance.current_remaining_minutes == null ? null : Number(balance.current_remaining_minutes);
    const pendingReservedMinutes = confirmedMinutes == null ? null : Number(currentPendingReservationMap[employeeId] || 0);
    const availableMinutes = confirmedMinutes == null ? null : Math.max(0, confirmedMinutes - pendingReservedMinutes);

    return {
      employee_id: employeeId,
      name: String(rowObj.name || "").trim(),
      name_kana: String(rowObj.name_kana || "").trim(),
      employment_type: String(rowObj.employment_type || "").trim(),

      fiscal_year: fiscalYear,
      fiscal_start_month: fiscalStartMonth,

      current_remaining_days: Number(balance.current_remaining_days || 0),
      current_remaining_minutes: confirmedMinutes,
      confirmed_remaining_minutes: confirmedMinutes,
      pending_reserved_minutes: pendingReservedMinutes,
      available_remaining_minutes: availableMinutes,
      remaining_full_days: balance.remaining_full_days == null ? null : Number(balance.remaining_full_days),
      remaining_hours: balance.remaining_hours == null ? null : Number(balance.remaining_hours),
      remaining_remainder_minutes: balance.remaining_remainder_minutes == null ? null : Number(balance.remaining_remainder_minutes),
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
    admin_cancel_approved: "承認後取消",

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
  if (type === "admin_cancel_approved") return "log-reject";

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
  EMPLOYEE_TIME_LEAVE_WORK_SCHEDULE_HEADERS.forEach(header => ensureSheetColumn_(sheet, header));
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
    "work_start_minute",
    "work_end_minute",
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
  rowObj.work_start_minute = data.work_start_minute === "" || data.work_start_minute == null
    ? "" : getOptionalEmployeeWorkMinute_(data.work_start_minute, "work_start_minute");
  rowObj.work_end_minute = data.work_end_minute === "" || data.work_end_minute == null
    ? "" : getOptionalEmployeeWorkMinute_(data.work_end_minute, "work_end_minute");
  // 入力途中の片側だけは登録させず、空欄時は会社標準へfallbackする。
  resolveEmployeeTimeLeavePolicy_(rowObj.employee_id, rowObj);
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
function getEmployeeTimeLeaveWorkScheduleMapFromSpreadsheet_() {
  const sheet = getSheet("employees");
  const headerInfo = requireHeaders(sheet, ["employee_id"]);
  const hasStart = "work_start_minute" in headerInfo.map;
  const hasEnd = "work_end_minute" in headerInfo.map;
  const result = {};
  if (!hasStart && !hasEnd) return result;
  sheet.getDataRange().getValues().slice(1).forEach(row => {
    const employee = rowToObject(row, headerInfo.headers);
    const employeeId = String(employee.employee_id || "").trim();
    if (!employeeId) return;
    result[employeeId] = {
      work_start_minute: hasStart ? employee.work_start_minute : "",
      work_end_minute: hasEnd ? employee.work_end_minute : ""
    };
  });
  return result;
}

function getEmployeesForAdmin() {
  if (shouldUseSupabaseReads_()) {
    const timeLeaveWorkSchedules = getEmployeeTimeLeaveWorkScheduleMapFromSpreadsheet_();
    return getEmployeesFromSupabase_()
      .map(obj => {
        const schedule = timeLeaveWorkSchedules[String(obj.employee_id || "").trim()] || {};
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
        work_start_minute: schedule.work_start_minute == null ? "" : schedule.work_start_minute,
        work_end_minute: schedule.work_end_minute == null ? "" : schedule.work_end_minute,
        fiscal_start_month: obj.fiscal_start_month || "",
        leave_management_target: obj.leave_management_target === true,
        initial_grant_check_target: obj.initial_grant_check_target === true,
        is_driver: obj.is_driver === true,
        driver_type: String(obj.driver_type || "").trim(),
        default_vehicle_id: String(obj.default_vehicle_id || "").trim(),
        display_order: obj.display_order || "",
        notes: String(obj.notes || "")
        };
      })
      .filter(emp => emp.employee_id)
      .sort((a, b) => Number(a.display_order || 9999) - Number(b.display_order || 9999));
  }

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
        work_start_minute: obj.work_start_minute == null ? "" : obj.work_start_minute,
        work_end_minute: obj.work_end_minute == null ? "" : obj.work_end_minute,
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
    { key: "work_start_minute", label: "勤務開始時刻" },
    { key: "work_end_minute", label: "勤務終了時刻" },
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
  EMPLOYEE_TIME_LEAVE_WORK_SCHEDULE_HEADERS.forEach(header => ensureSheetColumn_(sheet, header));
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
    "work_start_minute",
    "work_end_minute",
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

  const workStartMinute = data.work_start_minute === "" || data.work_start_minute == null
    ? "" : getOptionalEmployeeWorkMinute_(data.work_start_minute, "work_start_minute");
  const workEndMinute = data.work_end_minute === "" || data.work_end_minute == null
    ? "" : getOptionalEmployeeWorkMinute_(data.work_end_minute, "work_end_minute");
  resolveEmployeeTimeLeavePolicy_(data.employee_id, Object.assign({}, beforeObj, {
    company_code: data.company_code, work_start_minute: workStartMinute, work_end_minute: workEndMinute
  }));
  sheet.getRange(sheetRow, headerInfo.map.work_start_minute + 1).setValue(workStartMinute);
  sheet.getRange(sheetRow, headerInfo.map.work_end_minute + 1).setValue(workEndMinute);

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
  if (shouldUseSupabaseReads_()) {
    const rowMap = {};
    getCompanyCalendarFromSupabase_().forEach(rowObj => {
      if (!rowObj.date) return;
      rowMap[toDateKey(rowObj.date)] = rowObj;
    });

    const rows = [];
    let cursor = new Date(period.start);

    while (cursor <= period.end) {
      const dateKey = toDateKey(cursor);
      const existing = rowMap[dateKey];
      const type = cursor.getDay() === 0
        ? CALENDAR_TYPE.HOLIDAY
        : (existing ? norm(existing.type) : CALENDAR_TYPE.WORKDAY);

      rows.push({
        date: dateKey,
        day_of_week: ["日", "月", "火", "水", "木", "金", "土"][cursor.getDay()],
        type: type || CALENDAR_TYPE.WORKDAY,
        notes: existing ? String(existing.notes || "") : "",
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
  const grantRows = getInitialPaidLeaveGrantHistoryRows_();
  const opts = options || null;

  const rows = employees
    .map(emp => {
      const eligibility = calculateInitialPaidLeaveGrantEligibility_(
        emp,
        grantRows,
        today
      );
      return { emp: emp, eligibility: eligibility };
    })
    .filter(item => {
      return isInitialPaidLeaveGrantExecutionCandidate_(
        item.emp,
        item.eligibility
      );
    })
    .map(item => {
      const emp = item.emp;
      const eligibility = item.eligibility;
      const fiscalStartMonth = getInitialPaidLeaveFiscalStartMonth_(emp);

      return {
        employee_id: emp.employee_id,
        display_employee_id: emp.display_employee_id,
        name: getDisplayName(emp) || emp.name,
        hire_date: emp.hire_date,
        six_month_date: eligibility.six_month_date || "",
        company_basis_date: eligibility.company_basis_date || "",
        grant_date: eligibility.next_grant_date,
        grant_reason: eligibility.grant_reason,
        grant_days: eligibility.expected_grant_days,
        eligibility_status: eligibility.status,
        is_provisional: eligibility.is_provisional,
        warning_codes: eligibility.warning_codes,
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

  return runInitialPaidLeaveGrantWithLock_(function() {
    const employees = getEmployeesForAdmin();
    const emp = employees.find(e => String(e.employee_id) === String(employeeId));
    validateInitialPaidLeaveGrantEmployee_(emp, employeeId);

    const eligibility = calculateInitialPaidLeaveGrantEligibility_(
      emp,
      getInitialPaidLeaveGrantHistoryRows_(),
      parseLocalDate(new Date())
    );
    assertInitialPaidLeaveGrantCanExecute_(eligibility);

    const grantDate = parseLocalDate(eligibility.next_grant_date);
    const grantDays = Number(eligibility.expected_grant_days);
    const now = new Date();
    const sheet = getSheet("paid_leave_grants");
    const headerInfo = requireHeaders(sheet, INITIAL_PAID_LEAVE_GRANT_HEADERS_);
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
    rowObj.notes = buildInitialPaidLeaveGrantNotes_(eligibility);
    rowObj.created_at = now;
    rowObj.updated_at = now;

    appendRowFast_(sheet, objectToRow(rowObj, headerInfo.headers));

    const operatorId = adminUser && adminUser.admin_id ? adminUser.admin_id : "admin";
    const operatorName = adminUser && adminUser.admin_name ? adminUser.admin_name : "管理者";
    appendUsageLog({
      request_id: employeeId,
      action_type: "six_month_grant",
      operator_id: operatorId,
      operator_name: operatorName,
      comment: emp.name + " さんへ " + grantDays + "日を初回有給付与しました（" + eligibility.grant_reason + "）"
    });

    clearAppCache();
    return {
      ok: true,
      employee_id: employeeId,
      name: emp.name,
      grant_date: eligibility.next_grant_date,
      grant_days: grantDays,
      grant_reason: eligibility.grant_reason
    };
  });
}

/* =========================
   6か月到達者：処理済みにする
========================= */
function markSixMonthGrantCandidateProcessed(employeeId, reason, adminUser) {
  if (!employeeId) throw new Error("employeeId がありません");

  return runInitialPaidLeaveGrantWithLock_(function() {
    const employees = getEmployeesForAdmin();
    const emp = employees.find(e => String(e.employee_id) === String(employeeId));
    validateInitialPaidLeaveGrantEmployee_(emp, employeeId);

    const eligibility = calculateInitialPaidLeaveGrantEligibility_(
      emp,
      getInitialPaidLeaveGrantHistoryRows_(),
      parseLocalDate(new Date())
    );
    assertInitialPaidLeaveGrantCanExecute_(eligibility);

    const grantDate = parseLocalDate(eligibility.next_grant_date);
    const now = new Date();
    const note = String(reason || "").trim() ||
      "手動入力済みのため6か月付与チェックを処理済みにした";
    const sheet = getSheet("paid_leave_grants");
    const headerInfo = requireHeaders(sheet, INITIAL_PAID_LEAVE_GRANT_HEADERS_);
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
    appendRowFast_(sheet, objectToRow(rowObj, headerInfo.headers));

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
      grant_date: eligibility.next_grant_date,
      grant_type: "six_month_processed"
    };
  });
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

/* =========================
   初回有給付与の正式規則（読み取り専用の純粋計算）
   実付与処理はこの段階では変更しない。
========================= */
function calculateInitialPaidLeaveGrantEligibility_(emp, grantRows, asOfDateValue) {
  const employee = emp || {};
  const employeeId = String(employee.employee_id || "").trim();
  const companyCode = String(employee.company_code || "").trim().toUpperCase();
  const warningCodes = [];
  const asOfDate = asOfDateValue
    ? parseLocalDate(asOfDateValue)
    : parseLocalDate(new Date());

  if (!employeeId) {
    return buildUnjudgeableInitialGrantEligibility_(
      asOfDate,
      ["EMPLOYEE_ID_MISSING"]
    );
  }

  if (companyCode !== "MAIN" && companyCode !== "PARTNER") {
    return buildUnjudgeableInitialGrantEligibility_(
      asOfDate,
      ["COMPANY_CODE_UNSUPPORTED"]
    );
  }

  let hireDate;
  try {
    hireDate = parseLocalDate(employee.hire_date);
  } catch (e) {
    return buildUnjudgeableInitialGrantEligibility_(
      asOfDate,
      ["HIRE_DATE_INVALID"]
    );
  }

  const processedGrant = findInitialGrantProcessedRecord_(employeeId, grantRows, asOfDate);
  const initialGrant = calculateInitialPaidLeaveGrantPlan_(employee, hireDate, companyCode);

  if (processedGrant) {
    if (processedGrant.is_future) {
      return buildUnjudgeableInitialGrantEligibility_(
        asOfDate,
        ["FUTURE_INITIAL_GRANT_HISTORY"]
      );
    }
    if (!processedGrant.grant_date) {
      warningCodes.push(
        processedGrant.history_date_invalid
          ? "INITIAL_GRANT_HISTORY_DATE_INVALID"
          : "INITIAL_GRANT_HISTORY_DATE_MISSING"
      );
    }

    return {
      grant_stage: "INITIAL",
      status: "PROCESSED",
      as_of_date: formatInitialGrantDateKey_(asOfDate),
      next_grant_date: initialGrant.next_grant_date,
      expected_grant_days: initialGrant.expected_grant_days,
      grant_reason: initialGrant.grant_reason,
      is_provisional: initialGrant.is_provisional,
      existing_grant_found: true,
      processed_grant_type: processedGrant.grant_type,
      processed_grant_date: processedGrant.grant_date,
      processed_grant_days: processedGrant.grant_days,
      processed_grant_notes: processedGrant.notes,
      warning_codes: initialGrant.warning_codes.concat(warningCodes)
    };
  }

  if (initialGrant.expected_grant_days === null) {
    return Object.assign({}, initialGrant, {
      status: "UNJUDGEABLE",
      as_of_date: formatInitialGrantDateKey_(asOfDate),
      existing_grant_found: false,
      processed_grant_type: "",
      processed_grant_date: "",
      processed_grant_days: "",
      processed_grant_notes: ""
    });
  }

  const grantDate = parseLocalDate(initialGrant.next_grant_date);
  let status = "UPCOMING";
  if (grantDate.getTime() === asOfDate.getTime()) {
    status = "DUE_TODAY";
  } else if (grantDate < asOfDate) {
    status = "OVERDUE";
  }

  return Object.assign({}, initialGrant, {
    status: status,
    as_of_date: formatInitialGrantDateKey_(asOfDate),
    existing_grant_found: false,
    processed_grant_type: "",
    processed_grant_date: "",
    processed_grant_days: "",
    processed_grant_notes: ""
  });
}

function calculateInitialPaidLeaveGrantPlan_(employee, hireDate, companyCode) {
  const fiscalStartMonth = companyCode === "PARTNER" ? 6 : 4;
  const sixMonthDate = addMonthsClampedLocal_(hireDate, 6);
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

  const isCompanyBasisAdvance = companyBasisDate < sixMonthDate;
  const isHireDateCompanyBasisDate =
    companyBasisDate.getTime() === hireDate.getTime();
  const warningCodes = isHireDateCompanyBasisDate
    ? ["HIRE_DATE_EQUALS_COMPANY_BASIS_DATE"]
    : [];

  const workDays = normalizeWorkDaysPerWeek_(employee.work_days_per_week);
  const workDaysWarning = !isCompanyBasisAdvance && !workDays.is_valid
    ? [workDays.warning_code]
    : [];
  return {
    grant_stage: "INITIAL",
    six_month_date: formatInitialGrantDateKey_(sixMonthDate),
    company_basis_date: formatInitialGrantDateKey_(companyBasisDate),
    next_grant_date: formatInitialGrantDateKey_(
      isCompanyBasisAdvance ? companyBasisDate : sixMonthDate
    ),
    expected_grant_days: isCompanyBasisAdvance ? 10 :
      (workDays.is_valid ? getSixMonthGrantDays_(workDays.value) : null),
    grant_reason: isCompanyBasisAdvance
      ? "INITIAL_COMPANY_BASIS"
      : "INITIAL_SIX_MONTHS",
    is_provisional: isHireDateCompanyBasisDate || workDaysWarning.length > 0,
    warning_codes: warningCodes.concat(workDaysWarning)
  };
}

const INITIAL_PAID_LEAVE_GRANT_HEADERS_ = [
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
];

function getInitialPaidLeaveGrantHistoryRows_() {
  const sheet = getSheet("paid_leave_grants");
  const headerInfo = requireHeaders(sheet, [
    "employee_id",
    "grant_type",
    "grant_date",
    "grant_days",
    "notes"
  ]);
  const data = sheet.getDataRange().getValues();

  if (data.length <= 1) return [];
  return data.slice(1).map(row => rowToObject(row, headerInfo.headers));
}

function isInitialPaidLeaveGrantExecutionCandidate_(emp, eligibility) {
  if (!isInitialPaidLeaveGrantEmployeeEligible_(emp)) return false;
  return getInitialPaidLeaveGrantExecutionDecision_(eligibility).can_execute;
}

function isInitialPaidLeaveGrantEmployeeEligible_(emp) {
  if (!emp || emp.initial_grant_check_target !== true) return false;

  const status = String(emp.employment_status || "").trim().toLowerCase();
  const isActive = status === "active" || status === "在職";
  return isActive && emp.leave_management_target === true;
}

function validateInitialPaidLeaveGrantEmployee_(emp, employeeId) {
  if (!emp) throw new Error("対象社員が見つかりません");
  if (String(emp.employee_id || "") !== String(employeeId || "")) {
    throw new Error("対象社員IDが一致しません");
  }
  if (emp.initial_grant_check_target !== true) {
    throw new Error("この社員は新規登録社員の初回付与チェック対象ではありません");
  }

  const status = String(emp.employment_status || "").trim().toLowerCase();
  if (status !== "active" && status !== "在職") {
    throw new Error("この社員は在職中ではないため初回付与を実行できません");
  }
  if (emp.leave_management_target !== true) {
    throw new Error("この社員は有給管理の対象ではありません");
  }
}

function getInitialPaidLeaveGrantExecutionDecision_(eligibility) {
  const status = String(eligibility && eligibility.status || "");
  if (status === "DUE_TODAY" || status === "OVERDUE") {
    return { can_execute: true, error_message: "" };
  }
  if (status === "UPCOMING") {
    return { can_execute: false, error_message: "この社員はまだ付与日を迎えていません。" };
  }
  if (status === "PROCESSED") {
    return { can_execute: false, error_message: "この社員の初回付与はすでに処理されています。" };
  }
  if (status === "UNJUDGEABLE") {
    return { can_execute: false, error_message: "会社情報または入社日の不整合により付与判定ができません。" };
  }
  return { can_execute: false, error_message: "この社員は初回付与の実行対象ではありません。" };
}

function assertInitialPaidLeaveGrantCanExecute_(eligibility) {
  const decision = getInitialPaidLeaveGrantExecutionDecision_(eligibility);
  if (!decision.can_execute) throw new Error(decision.error_message);
}

function buildInitialPaidLeaveGrantNotes_(eligibility) {
  const isCompanyBasis = eligibility.grant_reason === "INITIAL_COMPANY_BASIS";
  const baseNotes = isCompanyBasis
    ? "初回付与（INITIAL_COMPANY_BASIS）: 会社基準日による初回付与"
    : "初回付与（INITIAL_SIX_MONTHS）: 入社6か月到達による初回付与";
  return eligibility.is_provisional
    ? baseNotes + " / 要確認: 入社日と会社基準日が同日"
    : baseNotes;
}

function runInitialPaidLeaveGrantWithLock_(callback) {
  const lock = LockService.getScriptLock();
  let locked = false;
  try {
    locked = lock.tryLock(30000);
    if (!locked) {
      throw new Error("初回付与処理が混み合っています。しばらくしてから再操作してください。");
    }
    return callback();
  } finally {
    if (locked) lock.releaseLock();
  }
}

/* =========================
   年次有給付与予定（読み取り専用の純粋計算）
========================= */
function calculateYearlyPaidLeaveGrantEligibility_(emp, grantRows, asOfDateValue) {
  const employee = emp || {};
  const employeeId = String(employee.employee_id || "").trim();
  const asOfDate = parseLocalDate(asOfDateValue);
  const warningCodes = [];

  if (!employeeId) return buildUnjudgeableYearlyGrantEligibility_(asOfDate, ["EMPLOYEE_ID_MISSING"]);
  const employmentStatus = String(employee.employment_status || "").trim().toLowerCase();
  if (employmentStatus !== "active" && employmentStatus !== "在職") {
    return buildNotEligibleYearlyGrantEligibility_(asOfDate, ["EMPLOYMENT_STATUS_NOT_ACTIVE"]);
  }
  if (employee.leave_management_target !== true) {
    return buildNotEligibleYearlyGrantEligibility_(asOfDate, ["LEAVE_MANAGEMENT_TARGET_DISABLED"]);
  }
  const companyCode = String(employee.company_code || "").trim().toUpperCase();
  if (companyCode !== "MAIN" && companyCode !== "PARTNER") {
    return buildUnjudgeableYearlyGrantEligibility_(asOfDate, ["COMPANY_CODE_UNSUPPORTED"]);
  }

  let hireDate;
  try {
    hireDate = parseLocalDate(employee.hire_date);
  } catch (e) {
    return buildUnjudgeableYearlyGrantEligibility_(asOfDate, ["HIRE_DATE_INVALID"]);
  }
  if (hireDate > asOfDate) {
    return buildNotEligibleYearlyGrantEligibility_(asOfDate, ["HIRE_DATE_IN_FUTURE"]);
  }

  const initialHistory = findInitialGrantProcessedRecord_(employeeId, grantRows, asOfDate);
  if (initialHistory && initialHistory.is_future) {
    return buildUnjudgeableYearlyGrantEligibility_(asOfDate, ["FUTURE_INITIAL_GRANT_HISTORY"]);
  }
  if (!initialHistory) {
    return buildNotEligibleYearlyGrantEligibility_(asOfDate, ["INITIAL_GRANT_NOT_PROCESSED"]);
  }

  const expectedFiscalStartMonth = companyCode === "PARTNER" ? 6 : 4;
  const configuredFiscalStartMonth = Number(employee.fiscal_start_month || expectedFiscalStartMonth);
  if (configuredFiscalStartMonth !== expectedFiscalStartMonth) {
    warningCodes.push("COMPANY_AND_FISCAL_MONTH_MISMATCH");
  }

  const yearlyPlan = calculateYearlyPaidLeaveGrantPlan_(
    hireDate,
    expectedFiscalStartMonth,
    asOfDate
  );
  if (!yearlyPlan) {
    return buildNotEligibleYearlyGrantEligibility_(asOfDate, warningCodes.concat(["YEARLY_GRANT_NOT_YET_APPLICABLE"]));
  }

  const yearlyHistory = findYearlyGrantHistoryForFiscalYear_(
    employeeId,
    grantRows,
    yearlyPlan.fiscal_year,
    expectedFiscalStartMonth,
    asOfDate
  );
  warningCodes.push.apply(warningCodes, yearlyHistory.warning_codes);
  const workDays = normalizeWorkDaysPerWeek_(employee.work_days_per_week);
  const isShortTime = workDays.is_valid && workDays.value < 5;
  if (isShortTime) warningCodes.push("YEARLY_PROPORTIONAL_GRANT_RULE_NOT_IMPLEMENTED");
  if (!workDays.is_valid) warningCodes.push(workDays.warning_code);

  const base = {
    grant_stage: "YEARLY",
    yearly_grant_stage: yearlyPlan.grant_stage,
    as_of_date: formatInitialGrantDateKey_(asOfDate),
    next_grant_date: yearlyPlan.next_grant_date,
    fiscal_year: yearlyPlan.fiscal_year,
    months_worked: yearlyPlan.months_worked,
    grant_reason: "YEARLY_COMPANY_BASIS",
    expected_grant_days: workDays.is_valid && !isShortTime ? yearlyPlan.normal_grant_days : null,
    reference_grant_days: isShortTime ? yearlyPlan.normal_grant_days : null,
    is_provisional: isShortTime || !workDays.is_valid || warningCodes.indexOf("COMPANY_AND_FISCAL_MONTH_MISMATCH") !== -1,
    attendance_status: "ATTENDANCE_UNCONFIRMED",
    requires_manual_confirmation: true,
    existing_grant_found: yearlyHistory.rows.length > 0,
    warning_codes: warningCodes
  };

  if (yearlyHistory.has_data_issue) {
    return Object.assign(base, {
      status: "UNJUDGEABLE",
      processed_grant_type: "",
      processed_grant_date: "",
      processed_grant_days: "",
      processed_grant_notes: ""
    });
  }
  if (yearlyHistory.rows.length > 0) {
    return Object.assign(base, {
      status: "PROCESSED",
      processed_grant_type: "yearly",
      processed_grant_date: yearlyHistory.rows[0].grant_date || "",
      processed_grant_days: yearlyHistory.rows[0].grant_days,
      processed_grant_notes: yearlyHistory.rows[0].notes || ""
    });
  }

  const nextGrantDate = parseLocalDate(yearlyPlan.next_grant_date);
  return Object.assign(base, {
    status: nextGrantDate > asOfDate ? "UPCOMING" :
      (nextGrantDate.getTime() === asOfDate.getTime() ? "DUE_TODAY" : "OVERDUE"),
    processed_grant_type: "",
    processed_grant_date: "",
    processed_grant_days: "",
    processed_grant_notes: ""
  });
}

function calculateYearlyPaidLeaveGrantPlan_(hireDate, fiscalStartMonth, asOfDate) {
  let fiscalYear = getFiscalYearFromDateWithStart(asOfDate, fiscalStartMonth);
  let basisDate = getFiscalYearRangeWithStart(fiscalYear, fiscalStartMonth).start;
  let monthsWorked = getMonthsWorked_(hireDate, basisDate);

  if (monthsWorked < 18) {
    fiscalYear++;
    basisDate = getFiscalYearRangeWithStart(fiscalYear, fiscalStartMonth).start;
    monthsWorked = getMonthsWorked_(hireDate, basisDate);
  }
  if (monthsWorked < 18) return null;

  return {
    fiscal_year: fiscalYear,
    next_grant_date: formatInitialGrantDateKey_(basisDate),
    months_worked: monthsWorked,
    normal_grant_days: getYearlyGrantDays_(monthsWorked),
    grant_stage: getYearlyGrantStage_(monthsWorked)
  };
}

function getYearlyGrantStage_(monthsWorked) {
  if (monthsWorked >= 78) return "YEARLY_6_5_PLUS_YEARS";
  if (monthsWorked >= 66) return "YEARLY_5_5_YEARS";
  if (monthsWorked >= 54) return "YEARLY_4_5_YEARS";
  if (monthsWorked >= 42) return "YEARLY_3_5_YEARS";
  if (monthsWorked >= 30) return "YEARLY_2_5_YEARS";
  return "YEARLY_1_5_YEARS";
}

function findYearlyGrantHistoryForFiscalYear_(employeeId, grantRows, fiscalYear, fiscalStartMonth, asOfDateValue) {
  const rows = [];
  const warningCodes = [];
  let hasDataIssue = false;
  const asOfDate = parseLocalDate(asOfDateValue);
  const knownGrantTypes = {
    six_month: true,
    six_month_processed: true,
    six_month_skipped: true,
    initial: true,
    yearly: true,
    opening_balance_virtual_lot: true
  };

  (Array.isArray(grantRows) ? grantRows : []).forEach(row => {
    if (String(row && row.employee_id || "").trim() !== String(employeeId)) return;
    const grantType = String(row && row.grant_type || "").trim().toLowerCase();
    if (!knownGrantTypes[grantType]) warningCodes.push("UNKNOWN_GRANT_TYPE");
    if (grantType !== "yearly") return;

    const rowYear = Number(row.year);
    let dateFiscalYear = null;
    let grantDateKey = "";
    try {
      if (row && row.grant_date) {
        const parsedGrantDate = parseLocalDate(row.grant_date);
        dateFiscalYear = getFiscalYearFromDateWithStart(parsedGrantDate, fiscalStartMonth);
        grantDateKey = formatInitialGrantDateKey_(parsedGrantDate);
      }
    } catch (e) {
      warningCodes.push("YEAR_AND_GRANT_DATE_MISMATCH");
      hasDataIssue = true;
    }
    const isRelevant = rowYear === Number(fiscalYear) || dateFiscalYear === Number(fiscalYear);
    if (!isRelevant) return;
    if (!rowYear) {
      warningCodes.push("YEARLY_GRANT_YEAR_MISSING");
      hasDataIssue = true;
      return;
    }
    if (dateFiscalYear === null || dateFiscalYear !== rowYear) {
      warningCodes.push("YEAR_AND_GRANT_DATE_MISMATCH");
      hasDataIssue = true;
      return;
    }
    const grantDays = Number(row.grant_days || 0);
    if (grantDays <= 0) {
      warningCodes.push("ZERO_DAY_YEARLY_GRANT_HISTORY");
      hasDataIssue = true;
      return;
    }
    if (parseLocalDate(grantDateKey) > asOfDate) {
      warningCodes.push("FUTURE_YEARLY_GRANT_HISTORY");
      hasDataIssue = true;
      return;
    }
    rows.push({ grant_date: grantDateKey, grant_days: grantDays, notes: String(row.notes || "") });
  });
  if (rows.length > 1) warningCodes.push("DUPLICATE_YEARLY_GRANT_HISTORY");
  return { rows: rows, has_data_issue: hasDataIssue, warning_codes: uniqueWarningCodes_(warningCodes) };
}

function buildUnjudgeableYearlyGrantEligibility_(asOfDate, warningCodes) {
  return {
    grant_stage: "YEARLY", status: "UNJUDGEABLE",
    as_of_date: formatInitialGrantDateKey_(asOfDate), next_grant_date: "",
    expected_grant_days: "", reference_grant_days: null, grant_reason: "",
    attendance_status: "ATTENDANCE_UNCONFIRMED", requires_manual_confirmation: true,
    existing_grant_found: false, is_provisional: true,
    warning_codes: uniqueWarningCodes_(warningCodes || [])
  };
}

function buildNotEligibleYearlyGrantEligibility_(asOfDate, warningCodes) {
  return Object.assign(buildUnjudgeableYearlyGrantEligibility_(asOfDate, warningCodes), {
    status: "NOT_ELIGIBLE",
    is_provisional: false
  });
}

function calculateNextPaidLeaveGrantSchedule_(emp, grantRows, asOfDateValue) {
  const employee = emp || {};
  const asOfDate = parseLocalDate(asOfDateValue);
  try {
    if (employee.hire_date && parseLocalDate(employee.hire_date) > asOfDate) {
      return buildUnjudgeableInitialGrantEligibility_(asOfDate, ["HIRE_DATE_IN_FUTURE"]);
    }
  } catch (e) {
    return buildUnjudgeableInitialGrantEligibility_(asOfDate, ["HIRE_DATE_INVALID"]);
  }
  const employmentStatus = String(employee.employment_status || "").trim().toLowerCase();
  if (employmentStatus !== "active" && employmentStatus !== "在職") {
    return buildNotEligiblePaidLeaveGrantSchedule_(asOfDate, "INITIAL", ["EMPLOYMENT_STATUS_NOT_ACTIVE"]);
  }
  if (employee.leave_management_target !== true) {
    return buildNotEligiblePaidLeaveGrantSchedule_(asOfDate, "INITIAL", ["LEAVE_MANAGEMENT_TARGET_DISABLED"]);
  }
  const initial = calculateInitialPaidLeaveGrantEligibility_(emp, grantRows, asOfDateValue);
  if (initial.status !== "PROCESSED") return initial;
  return calculateYearlyPaidLeaveGrantEligibility_(emp, grantRows, asOfDateValue);
}

function buildNotEligiblePaidLeaveGrantSchedule_(asOfDate, grantStage, warningCodes) {
  return {
    grant_stage: grantStage,
    status: "NOT_ELIGIBLE",
    as_of_date: formatInitialGrantDateKey_(asOfDate),
    next_grant_date: "",
    expected_grant_days: "",
    reference_grant_days: null,
    grant_reason: "",
    attendance_status: "ATTENDANCE_UNCONFIRMED",
    requires_manual_confirmation: false,
    existing_grant_found: false,
    is_provisional: false,
    warning_codes: uniqueWarningCodes_(warningCodes || [])
  };
}

function uniqueWarningCodes_(codes) {
  return (codes || []).filter((code, index, all) => code && all.indexOf(code) === index);
}

function findInitialGrantProcessedRecord_(employeeId, grantRows, asOfDateValue) {
  const processedTypes = {
    six_month: true,
    six_month_processed: true,
    six_month_skipped: true,
    initial: true
  };

  return (Array.isArray(grantRows) ? grantRows : [])
    .filter(row => String(row && row.employee_id || "").trim() === String(employeeId))
    .map(row => {
      let grantDate = "";
      let historyDateInvalid = false;

      if (row && row.grant_date) {
        try {
          grantDate = formatInitialGrantDateKey_(parseLocalDate(row.grant_date));
        } catch (e) {
          historyDateInvalid = true;
        }
      }

      const isFuture = grantDate && asOfDateValue && parseLocalDate(grantDate) > parseLocalDate(asOfDateValue);
      return {
        grant_type: String(row && row.grant_type || "").trim(),
        grant_date: grantDate,
        grant_days: row && row.grant_days != null ? Number(row.grant_days) : "",
        notes: String(row && row.notes || ""),
        history_date_invalid: historyDateInvalid,
        is_future: isFuture
      };
    })
    .find(row => processedTypes[row.grant_type]) || null;
}

function buildUnjudgeableInitialGrantEligibility_(asOfDate, warningCodes) {
  return {
    grant_stage: "INITIAL",
    status: "UNJUDGEABLE",
    as_of_date: formatInitialGrantDateKey_(asOfDate),
    next_grant_date: "",
    expected_grant_days: "",
    grant_reason: "",
    is_provisional: false,
    existing_grant_found: false,
    processed_grant_type: "",
    processed_grant_date: "",
    processed_grant_days: "",
    processed_grant_notes: "",
    warning_codes: warningCodes.slice()
  };
}

function formatInitialGrantDateKey_(dateValue) {
  const date = dateValue instanceof Date
    ? dateValue
    : parseLocalDate(dateValue);
  return [
    String(date.getFullYear()).padStart(4, "0"),
    String(date.getMonth() + 1).padStart(2, "0"),
    String(date.getDate()).padStart(2, "0")
  ].join("-");
}

function addMonthsClampedLocal_(dateValue, months) {
  const date = dateValue instanceof Date
    ? dateValue
    : parseLocalDate(dateValue);
  const targetMonthStart = new Date(date.getFullYear(), date.getMonth() + Number(months || 0), 1);
  const lastDay = new Date(targetMonthStart.getFullYear(), targetMonthStart.getMonth() + 1, 0).getDate();
  return new Date(targetMonthStart.getFullYear(), targetMonthStart.getMonth(), Math.min(date.getDate(), lastDay));
}

function addMonthsForInitialGrant_(dateValue, months) {
  return addMonthsClampedLocal_(dateValue, months);
}

function normalizeWorkDaysPerWeek_(value) {
  if (value === null || value === undefined || String(value).trim() === "") {
    return { is_valid: false, value: null, warning_code: "WORK_DAYS_PER_WEEK_MISSING" };
  }
  const normalized = typeof value === "number" ? value :
    (/^[1-5]$/.test(String(value).trim()) ? Number(String(value).trim()) : NaN);
  if (!Number.isInteger(normalized) || normalized < 1 || normalized > 5) {
    return { is_valid: false, value: null, warning_code: "INVALID_WORK_DAYS_PER_WEEK" };
  }
  return { is_valid: true, value: normalized, warning_code: null };
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
  const fiveDayMap = getFiveDayObligationDaysByFiscalYearForEmployeeIds(
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
    const fiveDayUsed = Math.min(Number(fiveDayMap[employeeId] || 0), 5);
    const fiveDayRemaining = Math.max(0, 5 - fiveDayUsed);
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
      current_remaining_minutes: fifoBalance.current_remaining_minutes == null ? null : Number(fifoBalance.current_remaining_minutes),
      remaining_full_days: fifoBalance.remaining_full_days == null ? null : Number(fifoBalance.remaining_full_days),
      remaining_hours: fifoBalance.remaining_hours == null ? null : Number(fifoBalance.remaining_hours),
      remaining_remainder_minutes: fifoBalance.remaining_remainder_minutes == null ? null : Number(fifoBalance.remaining_remainder_minutes),
      fiscal_used_days: usedDays,
      five_day_used: fiveDayUsed,
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

/* =========================
   管理者向け付与予定・要確認（完全読み取り専用）
========================= */
function getPaidLeaveGrantScheduleForAdmin(params) {
  const opts = params || {};
  const asOfDate = opts.as_of_date ? parseLocalDate(opts.as_of_date) : parseLocalDate(new Date());
  const daysAhead = Math.max(0, Math.min(Number(opts.days_ahead || 30), 366));
  const includeProcessedDays = Math.max(0, Math.min(Number(opts.include_processed_days || 31), 366));
  const companyCodeFilter = String(opts.company_code || "ALL").trim().toUpperCase();
  const employees = getEmployeesForAdmin();
  const grantRows = getInitialPaidLeaveGrantHistoryRows_();
  const fifoContext = createFifoBalanceComparisonContext_(asOfDate, { read_only: true });
  const horizon = addDaysLocal_(asOfDate, daysAhead);
  const processedSince = addDaysLocal_(asOfDate, -includeProcessedDays);

  const rows = employees
    .filter(emp => isPaidLeaveGrantScheduleCompanyMatch_(emp, companyCodeFilter))
    .map(emp => buildPaidLeaveGrantScheduleAdminRow_(emp, grantRows, asOfDate, fifoContext))
    .filter(row => shouldIncludePaidLeaveGrantScheduleRow_(row, horizon, processedSince))
    .sort(comparePaidLeaveGrantScheduleRows_);

  const counts = {
    upcoming: 0, due_today: 0, overdue: 0, needs_review: 0,
    attendance_confirmation: 0, data_issue: 0, processed: 0, unjudgeable: 0
  };
  rows.forEach(row => {
    const status = String(row.eligibility_status || "").toLowerCase();
    if (Object.prototype.hasOwnProperty.call(counts, status)) counts[status]++;
    const dataIssue = row.data_issue === true;
    const attendanceConfirmation = row.attendance_confirmation_required === true;
    if (attendanceConfirmation) counts.attendance_confirmation++;
    if (dataIssue) counts.data_issue++;
    if (attendanceConfirmation || dataIssue) counts.needs_review++;
  });

  return {
    success: true,
    generated_at: new Date().toISOString(),
    as_of_date: formatInitialGrantDateKey_(asOfDate),
    days_ahead: daysAhead,
    company_code: companyCodeFilter,
    counts: counts,
    rows: rows
  };
}

/*
 * P0004補正後の付与予定APIを、シート等へ書き込まずに確認するための診断。
 * 氏名・notes本文はログおよび戻り値に含めない。
 */
function debugPaidLeaveGrantScheduleApiAfterP0004Repair() {
  const asOfDate = parseLocalDate("2026-07-25");
  const options = {
    as_of_date: formatInitialGrantDateKey_(asOfDate),
    company_code: "PARTNER"
  };
  Logger.log("[PAID_LEAVE_GRANT_SCHEDULE_API] API実行開始");

  try {
    const employees = getEmployeesForAdmin();
    const grantRows = getInitialPaidLeaveGrantHistoryRows_();
    const p0004Employee = employees.find(emp =>
      String(emp.employee_id || "").trim() === "EMP0062" &&
      String(emp.display_employee_id || "").trim() === "P0004"
    );
    const p0004GrantRows = grantRows.filter(row =>
      String(row.employee_id || "").trim() === "EMP0062" &&
      (String(row.grant_id || "").trim() === "G0058" || String(row.grant_id || "").trim() === "G0061")
    );
    Logger.log("[PAID_LEAVE_GRANT_SCHEDULE_API] 社員件数: %s", employees.length);
    Logger.log("[PAID_LEAVE_GRANT_SCHEDULE_API] 対象行件数: %s", p0004GrantRows.length);
    if (!p0004Employee) throw new Error("P0004 / EMP0062 が社員マスターに見つかりません。");

    Logger.log("[PAID_LEAVE_GRANT_SCHEDULE_API] P0004の計算開始");
    const response = getPaidLeaveGrantScheduleForAdmin(options);
    const fifoContext = createFifoBalanceComparisonContext_(asOfDate, { read_only: true });
    const fifoBalance = calculateFifoBalanceWithOpeningBalanceFromContext_(
      "EMP0062",
      asOfDate,
      fifoContext
    );
    const p0004Schedule = buildPaidLeaveGrantScheduleAdminRow_(
      p0004Employee,
      grantRows,
      asOfDate,
      fifoContext
    );
    const diagnostic = buildPaidLeaveGrantScheduleApiAfterP0004RepairDiagnostic_(
      response,
      p0004Schedule,
      fifoBalance
    );
    Logger.log("[PAID_LEAVE_GRANT_SCHEDULE_API] P0004の計算完了: %s", JSON.stringify(diagnostic.p0004));
    Logger.log("[PAID_LEAVE_GRANT_SCHEDULE_API] APIレスポンス生成完了");
    return diagnostic;
  } catch (error) {
    Logger.log("[PAID_LEAVE_GRANT_SCHEDULE_API] API実行失敗: %s", String(error && error.message || error));
    throw error;
  }
}

function buildPaidLeaveGrantScheduleApiAfterP0004RepairDiagnostic_(response, p0004Schedule, fifoBalance) {
  const apiResponse = response || {};
  const schedule = p0004Schedule || {};
  const fifo = fifoBalance || {};
  const activeLots = (fifo.grant_details || [])
    .filter(lot => Number(lot.active_remaining_days || 0) > 0)
    .slice()
    .sort((a, b) => String(a.valid_to || "").localeCompare(String(b.valid_to || "")));
  const nearestExpiryLot = activeLots[0] || null;
  const nearestExpiryDate = nearestExpiryLot ? String(nearestExpiryLot.valid_to || "") : "";
  const nearestExpiryDateKey = nearestExpiryDate
    ? formatInitialGrantDateKey_(nearestExpiryDate)
    : "";

  return {
    ok: apiResponse.success === true,
    read_only: true,
    api_employee_count: Array.isArray(apiResponse.rows) ? apiResponse.rows.length : 0,
    p0004: {
      employee_id: "EMP0062",
      display_employee_id: "P0004",
      current_remaining_days: Number(fifo.current_remaining_days || 0),
      expired_days: Number(fifo.expired_days || 0),
      nearest_expiry_date: nearestExpiryDate,
      grant_schedule_status: String(schedule.eligibility_status || ""),
      warning_codes: Array.isArray(schedule.warning_codes) ? schedule.warning_codes.slice() : [],
      fifo_lot_count: Array.isArray(fifo.grant_details) ? fifo.grant_details.length : 0
    },
    expected_values_check: {
      current_remaining_days_9_5: Number(fifo.current_remaining_days || 0) === 9.5,
      expired_days_0: Number(fifo.expired_days || 0) === 0,
      nearest_expiry_date_2028_05_31: nearestExpiryDateKey === "2028-05-31"
    }
  };
}

function isPaidLeaveGrantScheduleCompanyMatch_(emp, companyCodeFilter) {
  const filter = String(companyCodeFilter || "ALL").trim().toUpperCase();
  return filter === "ALL" || String(emp && emp.company_code || "").trim().toUpperCase() === filter;
}

function buildPaidLeaveGrantScheduleAdminRow_(emp, grantRows, asOfDate, fifoContext) {
  const employeeId = String(emp.employee_id || "").trim();
  const schedule = calculateNextPaidLeaveGrantSchedule_(emp, grantRows, asOfDate);
  const fifoBalance = calculateFifoBalanceWithOpeningBalanceFromContext_(
    employeeId,
    asOfDate,
    fifoContext
  );
  const fifo = buildPaidLeaveGrantScheduleFifoView_(fifoBalance, asOfDate);
  const warningCodes = uniqueWarningCodes_(
    (schedule.warning_codes || []).concat(fifo.warning_codes || [])
  );
  const dataIssue = hasPaidLeaveGrantScheduleDataIssue_(warningCodes);
  const attendanceConfirmation = schedule.status !== "PROCESSED" &&
    schedule.grant_stage === "YEARLY" &&
    schedule.attendance_status === "ATTENDANCE_UNCONFIRMED";
  const requiresManualConfirmation = attendanceConfirmation || dataIssue;

  return {
    employee_id: employeeId,
    display_employee_id: String(emp.display_employee_id || ""),
    employee_name: getDisplayName(emp) || emp.name || employeeId,
    company_code: String(emp.company_code || ""),
    company_name: String(emp.company_name || ""),
    hire_date: emp.hire_date || "",
    grant_stage: schedule.grant_stage || "",
    yearly_grant_stage: schedule.yearly_grant_stage || "",
    next_grant_date: schedule.next_grant_date || "",
    expected_grant_days: schedule.expected_grant_days,
    reference_grant_days: schedule.reference_grant_days,
    grant_reason: schedule.grant_reason || "",
    eligibility_status: schedule.status || "UNJUDGEABLE",
    attendance_status: schedule.attendance_status || "ATTENDANCE_UNCONFIRMED",
    attendance_confirmation_required: attendanceConfirmation,
    data_issue: dataIssue,
    requires_manual_confirmation: requiresManualConfirmation,
    existing_grant_found: schedule.existing_grant_found === true,
    is_provisional: schedule.is_provisional === true,
    exclusion_reasons: schedule.status === "NOT_ELIGIBLE" ? warningCodes : [],
    warning_codes: warningCodes,
    fifo_balance: fifo
  };
}

function buildPaidLeaveGrantScheduleFifoView_(fifoBalance, asOfDate) {
  const targetDate = parseLocalDate(asOfDate);
  const warningCodes = [];
  const details = (fifoBalance.grant_details || []).map((lot, index) => {
    const remainingDays = Number(lot.active_remaining_days != null
      ? lot.active_remaining_days
      : lot.remaining_days || 0);
    const validTo = lot.valid_to || "";
    let status = "ACTIVE";
    let daysUntilExpiry = null;
    if (validTo) {
      const validToDate = parseLocalDate(validTo);
      daysUntilExpiry = Math.floor((validToDate - targetDate) / 86400000);
      if (validToDate < targetDate || lot.is_expired) status = "EXPIRED";
      else if (parseLocalDate(lot.valid_from) > targetDate) status = "FUTURE";
      else if (remainingDays <= 0) status = "FULLY_USED";
      else if (daysUntilExpiry <= 30) status = "EXPIRING_SOON";
    } else {
      warningCodes.push("GRANT_VALIDITY_MISSING");
    }
    if (status === "EXPIRING_SOON") warningCodes.push("FIFO_LOT_EXPIRING_WITHIN_30_DAYS");
    if (remainingDays < 0) warningCodes.push("NEGATIVE_FIFO_REMAINING");
    if (status === "FUTURE") warningCodes.push("FUTURE_GRANT_INCLUDED");
    if (lot.validity_needs_review) warningCodes.push("GRANT_VALIDITY_MISSING");
    return {
      grant_id: lot.grant_id,
      grant_type: lot.grant_type,
      grant_date: lot.grant_date,
      valid_from: lot.valid_from,
      valid_to: validTo,
      original_days: Number(lot.total_days || 0),
      used_days: Number(lot.used_days || 0),
      remaining_days: remainingDays,
      consumption_priority: status === "ACTIVE" || status === "EXPIRING_SOON" ? index + 1 : null,
      status: status
    };
  });
  const activeLots = details.filter(lot => lot.status === "ACTIVE" || lot.status === "EXPIRING_SOON");
  activeLots.forEach((lot, index) => { lot.consumption_priority = index + 1; });
  const duplicateIds = details.map(lot => lot.grant_id).filter((id, index, all) => id && all.indexOf(id) !== index);
  if (duplicateIds.length > 0) warningCodes.push("DUPLICATE_GRANT_LOT");
  const lotTotal = activeLots.reduce((sum, lot) => sum + Number(lot.remaining_days || 0), 0);
  if (Math.abs(lotTotal - Number(fifoBalance.current_remaining_days || 0)) > 0.000001) {
    warningCodes.push("FIFO_TOTAL_MISMATCH");
  }
  return {
    total_remaining_days: Number(fifoBalance.current_remaining_days || 0),
    lots: details,
    warning_codes: uniqueWarningCodes_(warningCodes)
  };
}

function shouldIncludePaidLeaveGrantScheduleRow_(row, horizon, processedSince) {
  const status = String(row.eligibility_status || "");
  if (status === "NOT_ELIGIBLE") return false;
  if (status === "OVERDUE" || status === "DUE_TODAY" || status === "UNJUDGEABLE") return true;
  if (row.requires_manual_confirmation || (row.warning_codes || []).length > 0) return true;
  if (status === "PROCESSED") {
    if (!row.next_grant_date) return true;
    return parseLocalDate(row.next_grant_date) >= processedSince;
  }
  if (status !== "UPCOMING" || !row.next_grant_date) return false;
  return parseLocalDate(row.next_grant_date) <= horizon;
}

function hasPaidLeaveGrantScheduleDataIssue_(warningCodes) {
  const attendanceOnly = {
    YEARLY_PROPORTIONAL_GRANT_RULE_NOT_IMPLEMENTED: true
  };
  return (warningCodes || []).some(code => !attendanceOnly[code]);
}

function comparePaidLeaveGrantScheduleRows_(a, b) {
  const priority = { OVERDUE: 1, DUE_TODAY: 2, UNJUDGEABLE: 3, NOT_ELIGIBLE: 4, UPCOMING: 5, PROCESSED: 6 };
  const aPriority = priority[a.eligibility_status] || 99;
  const bPriority = priority[b.eligibility_status] || 99;
  if (aPriority !== bPriority) return aPriority - bPriority;
  const aWarning = a.requires_manual_confirmation || (a.warning_codes || []).length > 0;
  const bWarning = b.requires_manual_confirmation || (b.warning_codes || []).length > 0;
  if (aWarning !== bWarning) return aWarning ? -1 : 1;
  const aDate = a.next_grant_date || "9999-12-31";
  const bDate = b.next_grant_date || "9999-12-31";
  if (aDate !== bDate) return aDate < bDate ? -1 : 1;
  return String(a.employee_id || "").localeCompare(String(b.employee_id || ""));
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
        current_remaining_minutes: fifoBalance.current_remaining_minutes == null ? null : Number(fifoBalance.current_remaining_minutes),
        remaining_full_days: fifoBalance.remaining_full_days == null ? null : Number(fifoBalance.remaining_full_days),
        remaining_hours: fifoBalance.remaining_hours == null ? null : Number(fifoBalance.remaining_hours),
        remaining_remainder_minutes: fifoBalance.remaining_remainder_minutes == null ? null : Number(fifoBalance.remaining_remainder_minutes),
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
  const fiveDayUsed = Math.min(
    Number(getFiveDayObligationDaysByFiscalYearForEmployeeIds(
      fiscalYear,
      [targetEmployeeId]
    )[targetEmployeeId] || 0),
    5
  );
  const fiveDayRemaining = Math.max(0, 5 - fiveDayUsed);
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
      current_remaining_minutes: fifoBalance.current_remaining_minutes == null ? null : Number(fifoBalance.current_remaining_minutes),
      remaining_full_days: fifoBalance.remaining_full_days == null ? null : Number(fifoBalance.remaining_full_days),
      remaining_hours: fifoBalance.remaining_hours == null ? null : Number(fifoBalance.remaining_hours),
      remaining_remainder_minutes: fifoBalance.remaining_remainder_minutes == null ? null : Number(fifoBalance.remaining_remainder_minutes),
      fiscal_used_days: usedDays,
      five_day_used: fiveDayUsed,
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

/* =========================
   Supabase移行前DB監査（読み取り専用）
========================= */
function auditLeaveDbForSupabaseMigration() {
  const sheetConfigs = [
    {
      name: "employees",
      idColumn: "employee_id",
      dateColumns: ["hire_date", "leave_date", "created_at", "updated_at"],
      booleanColumns: ["leave_management_target", "initial_grant_check_target", "is_driver"],
      enumColumns: {
        employment_status: ["active", "leave", "retired", "在職", "休職", "退職"],
        company_code: ["MAIN", "PARTNER"]
      },
      numberColumns: {
        work_days_per_week: { max: 7 },
        fiscal_start_month: { min: 1, max: 12 },
        display_order: { min: 0, max: 100000 }
      }
    },
    {
      name: "leave_requests",
      idColumn: "request_id",
      dateColumns: ["request_date", "start_date", "end_date", "approved_at", "created_at", "updated_at"],
      booleanColumns: [],
      enumColumns: {
        status: ["pending", "approved", "rejected", "canceled", "canceled_by_admin"],
        type: ["paid_leave"],
        half_day: ["", "am", "pm"]
      },
      numberColumns: {
        days: { min: 0, max: 30 },
        year: { min: 2000, max: 2100 }
      }
    },
    {
      name: "paid_leave_grants",
      idColumn: "grant_id",
      dateColumns: ["grant_date", "valid_from", "valid_to", "created_at", "updated_at", "finalized_at"],
      booleanColumns: ["is_finalized"],
      enumColumns: {
        grant_type: ["six_month", "six_month_processed", "six_month_skipped", "yearly"]
      },
      numberColumns: {
        grant_days: { min: 0, max: 40 },
        carry_over_days: { min: 0, max: 80 },
        year: { min: 2000, max: 2100 }
      }
    },
    {
      name: "company_calendar",
      idColumn: "date",
      dateColumns: ["date"],
      booleanColumns: [],
      enumColumns: {
        type: ["workday", "holiday", "no_leave"]
      },
      numberColumns: {}
    },
    {
      name: "usage_log",
      idColumn: "log_id",
      dateColumns: ["action_date"],
      booleanColumns: [],
      enumColumns: {},
      numberColumns: {}
    },
    {
      name: "admin_users",
      idColumn: "admin_id",
      dateColumns: [],
      booleanColumns: ["is_active"],
      enumColumns: {},
      numberColumns: {}
    }
  ];
  const expectedSheets = sheetConfigs.map(config => config.name);
  const ss = getAppSpreadsheet();
  const report = {
    generated_at: new Date().toISOString(),
    mode: "read_only",
    summary: {
      expected_sheets: expectedSheets,
      sheets: {},
      representative_counts: {},
      status_distribution: {},
      top_request_employees: [],
      top_grant_employees: []
    },
    warnings: [],
    errors: [],
    details: {
      headers: {},
      duplicate_ids: {},
      referential_integrity: {},
      invalid_dates: {},
      boolean_variants: {},
      enum_variants: {},
      number_issues: {}
    }
  };
  const sheetData = {};
  const maxSamples = 20;

  function addWarning(code, message, context) {
    report.warnings.push({
      code: code,
      message: message,
      context: context || {}
    });
  }

  function addError(code, message, context) {
    report.errors.push({
      code: code,
      message: message,
      context: context || {}
    });
  }

  function isBlank(value) {
    return value === "" || value == null;
  }

  function normalizeText(value) {
    return String(value == null ? "" : value).trim();
  }

  function isValidDateValue(value) {
    if (isBlank(value)) return true;
    if (Object.prototype.toString.call(value) === "[object Date]") {
      return !isNaN(value.getTime());
    }
    const text = normalizeText(value);
    if (!text) return true;

    const ymd = text.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/);
    if (ymd) {
      const y = Number(ymd[1]);
      const m = Number(ymd[2]);
      const d = Number(ymd[3]);
      const date = new Date(y, m - 1, d);
      return (
        date.getFullYear() === y &&
        date.getMonth() === m - 1 &&
        date.getDate() === d
      );
    }

    const parsed = new Date(text);
    return !isNaN(parsed.getTime());
  }

  function toAuditDateKey(value) {
    if (isBlank(value)) return "";
    if (!isValidDateValue(value)) return normalizeText(value);
    if (Object.prototype.toString.call(value) === "[object Date]") {
      return Utilities.formatDate(value, getAppTimeZone(), "yyyy-MM-dd");
    }

    const text = normalizeText(value);
    const ymd = text.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/);
    if (ymd) {
      return [
        ymd[1],
        String(Number(ymd[2])).padStart(2, "0"),
        String(Number(ymd[3])).padStart(2, "0")
      ].join("-");
    }

    const parsed = new Date(text);
    return Utilities.formatDate(parsed, getAppTimeZone(), "yyyy-MM-dd");
  }

  function addSample(target, item) {
    if (target.length < maxSamples) target.push(item);
  }

  function getCell(rowObj, column) {
    return column in rowObj ? rowObj[column] : "";
  }

  function countBy(rows, column) {
    const result = {};
    rows.forEach(item => {
      const value = normalizeText(getCell(item.rowObj, column)) || "(blank)";
      result[value] = (result[value] || 0) + 1;
    });
    return result;
  }

  function topByEmployee(rows, column) {
    const counts = countBy(rows, column);
    return Object.keys(counts)
      .filter(key => key !== "(blank)")
      .map(key => ({ employee_id: key, count: counts[key] }))
      .sort((a, b) => b.count - a.count)
      .slice(0, 10);
  }

  function collectSheet(config) {
    const sheet = ss.getSheetByName(config.name);
    if (!sheet) {
      addError("missing_sheet", config.name + " シートが見つかりません", { sheet: config.name });
      report.summary.sheets[config.name] = {
        exists: false,
        row_count: 0,
        header_count: 0
      };
      return;
    }

    const values = sheet.getDataRange().getValues();
    const headers = values.length > 0
      ? values[0].map(header => normalizeText(header))
      : [];
    const rows = values.slice(1).map((row, index) => ({
      row_number: index + 2,
      values: row,
      rowObj: rowToObject(row, headers)
    }));
    const headerSet = new Set(headers.filter(Boolean));

    report.summary.sheets[config.name] = {
      exists: true,
      row_count: rows.length,
      header_count: headers.filter(Boolean).length
    };
    report.details.headers[config.name] = headers;
    sheetData[config.name] = {
      config: config,
      headers: headers,
      headerSet: headerSet,
      rows: rows
    };

    if (config.idColumn && !headerSet.has(config.idColumn)) {
      addError("missing_id_column", config.name + "." + config.idColumn + " がありません", {
        sheet: config.name,
        column: config.idColumn
      });
    }

    config.dateColumns.forEach(column => {
      if (!headerSet.has(column)) return;
      const invalid = [];
      rows.forEach(item => {
        const value = getCell(item.rowObj, column);
        if (!isBlank(value) && !isValidDateValue(value)) {
          addSample(invalid, {
            row: item.row_number,
            value: normalizeText(value)
          });
        }
      });
      if (invalid.length > 0) {
        report.details.invalid_dates[config.name + "." + column] = invalid;
        addWarning("invalid_date", config.name + "." + column + " に日付として扱えない値があります", {
          sheet: config.name,
          column: column,
          sample_count: invalid.length
        });
      }
    });

    config.booleanColumns.forEach(column => {
      if (!headerSet.has(column)) return;
      const distribution = {};
      const unexpected = [];
      rows.forEach(item => {
        const value = getCell(item.rowObj, column);
        const text = normalizeText(value);
        const key = value === true ? "TRUE(boolean)" :
          value === false ? "FALSE(boolean)" :
          text || "(blank)";
        distribution[key] = (distribution[key] || 0) + 1;

        if (
          !isBlank(value) &&
          value !== true &&
          value !== false &&
          text.toUpperCase() !== "TRUE" &&
          text.toUpperCase() !== "FALSE"
        ) {
          addSample(unexpected, {
            row: item.row_number,
            value: text
          });
        }
      });
      report.details.boolean_variants[config.name + "." + column] = {
        distribution: distribution,
        unexpected_samples: unexpected
      };
      if (unexpected.length > 0) {
        addWarning("boolean_variant", config.name + "." + column + " にTRUE/FALSE以外の値があります", {
          sheet: config.name,
          column: column,
          sample_count: unexpected.length
        });
      }
    });

    Object.keys(config.enumColumns).forEach(column => {
      if (!headerSet.has(column)) return;
      const allowed = config.enumColumns[column];
      const allowedSet = new Set(allowed.map(value => String(value).toLowerCase()));
      const distribution = countBy(rows, column);
      const unexpected = [];
      rows.forEach(item => {
        const raw = normalizeText(getCell(item.rowObj, column));
        const key = raw.toLowerCase();
        if (raw && !allowedSet.has(key)) {
          addSample(unexpected, {
            row: item.row_number,
            value: raw
          });
        }
      });
      report.details.enum_variants[config.name + "." + column] = {
        allowed: allowed,
        distribution: distribution,
        unexpected_samples: unexpected
      };
      if (unexpected.length > 0) {
        addWarning("enum_variant", config.name + "." + column + " に想定外の値があります", {
          sheet: config.name,
          column: column,
          sample_count: unexpected.length
        });
      }
    });

    Object.keys(config.numberColumns).forEach(column => {
      if (!headerSet.has(column)) return;
      const rule = config.numberColumns[column] || {};
      const issues = [];
      rows.forEach(item => {
        const value = getCell(item.rowObj, column);
        if (isBlank(value)) return;
        const num = Number(value);
        let reason = "";
        if (!isFinite(num)) reason = "not_numeric";
        else if (rule.min != null && num < rule.min) reason = "less_than_min";
        else if (rule.max != null && num > rule.max) reason = "greater_than_max";

        if (reason) {
          addSample(issues, {
            row: item.row_number,
            value: normalizeText(value),
            reason: reason
          });
        }
      });
      if (issues.length > 0) {
        report.details.number_issues[config.name + "." + column] = issues;
        addWarning("number_issue", config.name + "." + column + " に数値化不可または範囲外の値があります", {
          sheet: config.name,
          column: column,
          sample_count: issues.length
        });
      }
    });
  }

  function checkDuplicates(config) {
    const data = sheetData[config.name];
    if (!data || !config.idColumn || !data.headerSet.has(config.idColumn)) return;

    const seen = {};
    const duplicates = [];
    data.rows.forEach(item => {
      const raw = getCell(item.rowObj, config.idColumn);
      const id = config.name === "company_calendar"
        ? toAuditDateKey(raw)
        : normalizeText(raw);
      if (!id) return;

      if (seen[id]) {
        addSample(duplicates, {
          id: id,
          first_row: seen[id],
          duplicate_row: item.row_number
        });
      } else {
        seen[id] = item.row_number;
      }
    });

    report.details.duplicate_ids[config.name + "." + config.idColumn] = duplicates;
    if (duplicates.length > 0) {
      addError("duplicate_id", config.name + "." + config.idColumn + " に重複があります", {
        sheet: config.name,
        column: config.idColumn,
        sample_count: duplicates.length
      });
    }
  }

  function buildIdSet(sheetName, columnName, normalizeFn) {
    const data = sheetData[sheetName];
    const result = new Set();
    if (!data || !data.headerSet.has(columnName)) return result;

    data.rows.forEach(item => {
      const raw = getCell(item.rowObj, columnName);
      const id = normalizeFn ? normalizeFn(raw) : normalizeText(raw);
      if (id) result.add(id);
    });
    return result;
  }

  function checkEmployeeReferences(sheetName, columnName, employeeIds) {
    const data = sheetData[sheetName];
    if (!data || !data.headerSet.has(columnName)) return;

    const missing = [];
    data.rows.forEach(item => {
      const employeeId = normalizeText(getCell(item.rowObj, columnName));
      if (employeeId && !employeeIds.has(employeeId)) {
        addSample(missing, {
          row: item.row_number,
          employee_id: employeeId
        });
      }
    });
    report.details.referential_integrity[sheetName + "." + columnName + "_missing_employees"] = missing;
    if (missing.length > 0) {
      addError("missing_employee_reference", sheetName + "." + columnName + " にemployees未存在の社員IDがあります", {
        sheet: sheetName,
        column: columnName,
        sample_count: missing.length
      });
    }
  }

  function checkUsageLogRequestReferences(employeeIds, requestIds) {
    const data = sheetData.usage_log;
    if (!data || !data.headerSet.has("request_id")) return;

    const missing = [];
    const employeeIdLike = [];
    data.rows.forEach(item => {
      const requestId = normalizeText(getCell(item.rowObj, "request_id"));
      if (!requestId) return;
      if (requestIds.has(requestId)) return;

      if (employeeIds.has(requestId)) {
        addSample(employeeIdLike, {
          row: item.row_number,
          request_id: requestId,
          classification: "employee_id_in_request_id_column"
        });
      } else {
        addSample(missing, {
          row: item.row_number,
          request_id: requestId
        });
      }
    });

    report.details.referential_integrity.usage_log_request_id_employee_id_like = employeeIdLike;
    report.details.referential_integrity.usage_log_request_id_missing = missing;

    if (employeeIdLike.length > 0) {
      addWarning("usage_log_request_id_employee_id", "usage_log.request_id に社員IDらしき値があります（既存仕様上、要確認）", {
        sample_count: employeeIdLike.length
      });
    }
    if (missing.length > 0) {
      addWarning("usage_log_request_id_missing", "usage_log.request_id に申請ID/社員IDのどちらにも一致しない値があります", {
        sample_count: missing.length
      });
    }
  }

  sheetConfigs.forEach(collectSheet);
  sheetConfigs.forEach(checkDuplicates);

  const employeeIds = buildIdSet("employees", "employee_id");
  const requestIds = buildIdSet("leave_requests", "request_id");

  checkEmployeeReferences("leave_requests", "employee_id", employeeIds);
  checkEmployeeReferences("paid_leave_grants", "employee_id", employeeIds);
  checkUsageLogRequestReferences(employeeIds, requestIds);

  const employees = sheetData.employees ? sheetData.employees.rows : [];
  const leaveRequests = sheetData.leave_requests ? sheetData.leave_requests.rows : [];
  const paidLeaveGrants = sheetData.paid_leave_grants ? sheetData.paid_leave_grants.rows : [];

  report.summary.representative_counts = {
    employees: employees.length,
    active_employees: employees.filter(item => {
      const status = normalizeText(getCell(item.rowObj, "employment_status")).toLowerCase();
      return status === "active" || status === "在職";
    }).length,
    leave_requests: leaveRequests.length,
    approved_leave_requests: leaveRequests.filter(item => {
      return normalizeText(getCell(item.rowObj, "status")) === "approved";
    }).length,
    canceled_by_admin_leave_requests: leaveRequests.filter(item => {
      return normalizeText(getCell(item.rowObj, "status")) === "canceled_by_admin";
    }).length,
    paid_leave_grants: paidLeaveGrants.length
  };

  if (sheetData.leave_requests && sheetData.leave_requests.headerSet.has("status")) {
    report.summary.status_distribution = countBy(leaveRequests, "status");
  }

  if (sheetData.leave_requests && sheetData.leave_requests.headerSet.has("employee_id")) {
    report.summary.top_request_employees = topByEmployee(leaveRequests, "employee_id");
  }

  if (sheetData.paid_leave_grants && sheetData.paid_leave_grants.headerSet.has("employee_id")) {
    report.summary.top_grant_employees = topByEmployee(paidLeaveGrants, "employee_id");
  }

  addWarning("migration_attention", "Supabase移行前に、件数照合・ID重複・参照整合・残日数/FIFO照合を別途実施してください", {
    note: "この監査は読み取り専用の構造/品質チェックで、残日数の正解照合までは行いません。"
  });

  Logger.log("[SupabaseMigrationAudit] summary\n" + JSON.stringify(report.summary, null, 2));
  Logger.log("[SupabaseMigrationAudit] warnings\n" + JSON.stringify(report.warnings, null, 2));
  Logger.log("[SupabaseMigrationAudit] errors\n" + JSON.stringify(report.errors, null, 2));
  Logger.log("[SupabaseMigrationAudit] details\n" + JSON.stringify(report.details, null, 2));

  return report;
}
