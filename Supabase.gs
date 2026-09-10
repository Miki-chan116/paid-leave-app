/* =========================
   Supabase接続検証（読み取り専用）

   セキュリティ注意:
   - SUPABASE_URL / SUPABASE_ANON_KEY はコードに直書きせず、
     Apps Script の Script Properties に設定してください。
   - SERVICE_ROLE_KEY はGASクライアント検証では使用しません。
   - anon keyで読み取りを許可する場合は、Supabase側でRLS policyを
     読み取り専用かつ必要最小限に設計してください。
   - Supabase書き込みは USE_SUPABASE_WRITE=true の時だけ実行します。
     Spreadsheet保存成功後のDual Write検証用で、失敗してもSpreadsheetは
     ロールバックしません。
   - USE_SUPABASE_READS=true は読み取りSupabase / 書き込みSpreadsheet
     の混在検証用です。本番運用ONは、書き込み移行後に再判断してください。
   - admin_users はPINを含むため、Supabase anon keyでは読みません。
========================= */

function getSupabaseConfig_() {
  const props = PropertiesService.getScriptProperties();
  const url = String(props.getProperty("SUPABASE_URL") || "").trim().replace(/\/+$/, "");
  const anonKey = String(props.getProperty("SUPABASE_ANON_KEY") || "").trim();

  const missing = [];
  if (!url) missing.push("SUPABASE_URL");
  if (!anonKey) missing.push("SUPABASE_ANON_KEY");

  if (missing.length > 0) {
    throw new Error("Script Properties に " + missing.join(", ") + " を設定してください");
  }

  return {
    url: url,
    anonKey: anonKey
  };
}

function buildSupabaseQueryString_(params) {
  if (!params) return "";

  return Object.keys(params)
    .filter(key => params[key] !== undefined && params[key] !== null && params[key] !== "")
    .map(key => encodeURIComponent(key) + "=" + encodeURIComponent(String(params[key])))
    .join("&");
}

function supabaseGet_(tableName, params) {
  const config = getSupabaseConfig_();
  const queryString = buildSupabaseQueryString_(params);
  const endpoint = config.url + "/rest/v1/" + encodeURIComponent(tableName) +
    (queryString ? "?" + queryString : "");

  const response = UrlFetchApp.fetch(endpoint, {
    method: "get",
    muteHttpExceptions: true,
    headers: {
      apikey: config.anonKey,
      Authorization: "Bearer " + config.anonKey,
      Accept: "application/json",
      Prefer: "count=exact"
    }
  });

  const statusCode = response.getResponseCode();
  const body = response.getContentText();
  const headers = response.getAllHeaders();
  const contentRange = headers["Content-Range"] || headers["content-range"] || "";

  Logger.log("[SupabaseGET] table=" + tableName + " status=" + statusCode + " content_range=" + contentRange);

  if (statusCode < 200 || statusCode >= 300) {
    Logger.log("[SupabaseGET] error_body=" + body);
    throw new Error("Supabase GET failed: status=" + statusCode + " table=" + tableName);
  }

  try {
    return {
      statusCode: statusCode,
      contentRange: contentRange,
      data: body ? JSON.parse(body) : []
    };
  } catch (err) {
    Logger.log("[SupabaseGET] parse_error=" + err.message);
    Logger.log("[SupabaseGET] response_body=" + body);
    throw err;
  }
}

function shouldUseSupabaseWrites_() {
  const value = PropertiesService
    .getScriptProperties()
    .getProperty("USE_SUPABASE_WRITE");

  return String(value || "").trim().toLowerCase() === "true";
}

function supabaseInsert_(tableName, record) {
  const config = getSupabaseConfig_();
  const endpoint = config.url + "/rest/v1/" + encodeURIComponent(tableName);

  const response = UrlFetchApp.fetch(endpoint, {
    method: "post",
    muteHttpExceptions: true,
    contentType: "application/json",
    payload: JSON.stringify(record || {}),
    headers: {
      apikey: config.anonKey,
      Authorization: "Bearer " + config.anonKey,
      Accept: "application/json",
      Prefer: "return=representation"
    }
  });

  const statusCode = response.getResponseCode();
  const body = response.getContentText();

  Logger.log("[SupabaseINSERT] table=" + tableName + " status=" + statusCode);

  if (statusCode !== 201) {
    Logger.log("[SupabaseINSERT] error_body=" + body);
    throw new Error("Supabase INSERT failed: status=" + statusCode + " table=" + tableName);
  }

  return {
    statusCode: statusCode,
    data: body ? JSON.parse(body) : []
  };
}

// Spreadsheetを正DBとするdual write用の限定更新。時間休の時刻はtext/整数分のまま
// 送るため、ここでDateへの変換やタイムゾーン補正は行わない。
function supabasePatch_(tableName, filters, record) {
  const config = getSupabaseConfig_();
  const queryString = buildSupabaseQueryString_(filters || {});
  const endpoint = config.url + "/rest/v1/" + encodeURIComponent(tableName) +
    (queryString ? "?" + queryString : "");
  const response = UrlFetchApp.fetch(endpoint, {
    method: "patch",
    muteHttpExceptions: true,
    contentType: "application/json",
    payload: JSON.stringify(record || {}),
    headers: {
      apikey: config.anonKey,
      Authorization: "Bearer " + config.anonKey,
      Accept: "application/json",
      Prefer: "return=representation"
    }
  });
  const statusCode = response.getResponseCode();
  const body = response.getContentText();
  Logger.log("[SupabasePATCH] table=" + tableName + " status=" + statusCode);
  if (statusCode < 200 || statusCode >= 300) {
    Logger.log("[SupabasePATCH] error_body=" + body);
    throw new Error("Supabase PATCH failed: status=" + statusCode + " table=" + tableName);
  }
  return { statusCode: statusCode, data: body ? JSON.parse(body) : [] };
}

function shouldUseSupabaseReads_() {
  const value = PropertiesService
    .getScriptProperties()
    .getProperty("USE_SUPABASE_READS");

  return String(value || "").trim().toLowerCase() === "true";
}

function toSupabaseWriteDate_(value) {
  if (!value) return null;
  if (value instanceof Date) {
    return Utilities.formatDate(value, getAppTimeZone(), "yyyy-MM-dd");
  }
  const date = toSupabaseReadDate_(value);
  if (date instanceof Date) {
    return Utilities.formatDate(date, getAppTimeZone(), "yyyy-MM-dd");
  }
  return String(value || "").trim() || null;
}

function toSupabaseWriteTimestamp_(value) {
  if (!value) return null;
  if (value instanceof Date) return value.toISOString();
  const date = new Date(value);
  return isNaN(date.getTime()) ? String(value || "").trim() : date.toISOString();
}

function toSupabaseReadDate_(value) {
  if (!value) return "";
  if (value instanceof Date) return value;

  const text = String(value || "").trim();
  const ymd = text.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (ymd) {
    return new Date(Number(ymd[1]), Number(ymd[2]) - 1, Number(ymd[3]));
  }

  const date = new Date(text);
  return isNaN(date.getTime()) ? text : date;
}

function toSupabaseReadBoolean_(value) {
  if (value === true || value === false) return value;
  const text = String(value || "").trim().toLowerCase();
  if (!text) return false;
  return text === "true" || text === "1" || text === "yes" || text === "対象";
}

function toSupabaseReadNumber_(value, fallback) {
  if (value === "" || value == null) return fallback == null ? "" : fallback;
  const num = Number(value);
  return isFinite(num) ? num : (fallback == null ? "" : fallback);
}

function normalizeSupabaseRow_(row, options) {
  const opts = options || {};
  const dateColumns = opts.dateColumns || [];
  const booleanColumns = opts.booleanColumns || [];
  const numberColumns = opts.numberColumns || [];
  const result = {};

  Object.keys(row || {}).forEach(key => {
    result[key] = row[key] == null ? "" : row[key];
  });

  dateColumns.forEach(key => {
    result[key] = toSupabaseReadDate_(result[key]);
  });

  booleanColumns.forEach(key => {
    result[key] = toSupabaseReadBoolean_(result[key]);
  });

  numberColumns.forEach(key => {
    result[key] = toSupabaseReadNumber_(result[key]);
  });

  return result;
}

function supabaseGetAll_(tableName, params) {
  const query = params || {};
  if (!("limit" in query)) query.limit = 10000;

  const result = supabaseGet_(tableName, query);
  return Array.isArray(result.data) ? result.data : [];
}

function getEmployeesFromSupabase_() {
  return supabaseGetAll_("employees", {
    select: "*",
    order: "display_order.asc,employee_id.asc"
  }).map(row => normalizeSupabaseRow_(row, {
    dateColumns: ["hire_date", "leave_date", "created_at", "updated_at", "deleted_at"],
    booleanColumns: ["leave_management_target", "initial_grant_check_target", "is_driver"],
    numberColumns: ["work_days_per_week", "fiscal_start_month", "display_order"]
  }));
}

function getCompanyCalendarFromSupabase_() {
  return supabaseGetAll_("company_calendar", {
    select: "*",
    order: "date.asc"
  }).map(row => normalizeSupabaseRow_(row, {
    dateColumns: ["date", "created_at", "updated_at"]
  }));
}

function getPaidLeaveGrantsFromSupabase_() {
  return supabaseGetAll_("paid_leave_grants", {
    select: "*",
    order: "employee_id.asc,grant_date.asc,grant_id.asc"
  }).map(row => normalizeSupabaseRow_(row, {
    dateColumns: ["grant_date", "valid_from", "valid_to", "finalized_at", "created_at", "updated_at"],
    booleanColumns: ["is_finalized"],
    numberColumns: ["grant_days", "carry_over_days", "year"]
  }));
}

function getTimeLeaveSegmentsFromSupabase_() {
  return supabaseGetAll_("time_leave_segments", {
    select: "*",
    order: "employee_id.asc,leave_date.asc,time_leave_id.asc"
  }).map(row => normalizeSupabaseRow_(row, {
    dateColumns: ["leave_date", "created_at", "updated_at"],
    numberColumns: [
      "start_minute", "end_minute", "requested_minutes", "scheduled_minutes_per_day",
      "time_leave_unit_minutes", "work_start_minute", "work_end_minute"
    ]
  }));
}

function getLeaveRequestsFromSupabase_() {
  return supabaseGetAll_("leave_requests", {
    select: "*",
    order: "start_date.desc,request_date.desc,request_id.asc"
  }).map(row => normalizeSupabaseRow_(row, {
    dateColumns: ["request_date", "start_date", "end_date", "approved_at", "created_at", "updated_at"],
    numberColumns: ["days", "year"]
  }));
}

function getUsageLogsFromSupabase_() {
  return supabaseGetAll_("usage_logs", {
    select: "*",
    order: "action_date.desc,log_id.asc"
  }).map(row => {
    const normalized = normalizeSupabaseRow_(row, {
      dateColumns: ["action_date", "created_at", "updated_at"]
    });
    normalized.request_id = normalized.legacy_request_id ||
      normalized.target_id ||
      normalized.leave_request_id ||
      normalized.employee_id ||
      "";
    return normalized;
  });
}

function buildSupabaseLeaveRequestRecord_(rowObj) {
  return {
    request_id: String(rowObj.request_id || "").trim(),
    employee_id: String(rowObj.employee_id || "").trim(),
    request_date: toSupabaseWriteTimestamp_(rowObj.request_date),
    start_date: toSupabaseWriteDate_(rowObj.start_date),
    end_date: toSupabaseWriteDate_(rowObj.end_date),
    days: Number(rowObj.days || 0),
    type: String(rowObj.type || "paid_leave").trim() || "paid_leave",
    half_day: rowObj.half_day ? String(rowObj.half_day).trim() : null,
    reason: String(rowObj.reason || ""),
    reason_detail: String(rowObj.reason_detail || ""),
    status: String(rowObj.status || "pending").trim() || "pending",
    approver_id: String(rowObj.approver_id || ""),
    approver_name: String(rowObj.approver_name || ""),
    approved_at: toSupabaseWriteTimestamp_(rowObj.approved_at),
    rejected_reason: String(rowObj.rejected_reason || ""),
    cancel_reason: String(rowObj.cancel_reason || ""),
    year: rowObj.year === "" || rowObj.year == null ? null : Number(rowObj.year),
    created_at: toSupabaseWriteTimestamp_(rowObj.created_at),
    updated_at: toSupabaseWriteTimestamp_(rowObj.updated_at)
  };
}

function parseSupabaseBreakPeriods_(value) {
  if (Array.isArray(value)) return value;
  const text = String(value == null ? "" : value).trim();
  if (!text) return [];
  try {
    const parsed = JSON.parse(text);
    if (!Array.isArray(parsed)) throw new Error("配列ではありません");
    return parsed;
  } catch (error) {
    throw new Error("break_periods_json が不正です: " + error.message);
  }
}

function requireSupabaseInteger_(value, fieldName, minimum) {
  const number = Number(value);
  if (!Number.isInteger(number) || number < Number(minimum || 0)) {
    throw new Error(fieldName + " は" + Number(minimum || 0) + "以上の整数で指定してください");
  }
  return number;
}

// start_time/end_time はPostgreSQL text列へそのまま送る。日時に変換しないことで
// SpreadsheetのHH:mmとタイムゾーン非依存の意味を一致させる。
function buildSupabaseTimeLeaveSegmentRecord_(rowObj) {
  const startTime = String(rowObj.start_time || "").trim();
  const endTime = String(rowObj.end_time || "").trim();
  if (!/^([01][0-9]|2[0-3]):[0-5][0-9]$/.test(startTime) ||
      !/^([01][0-9]|2[0-3]):[0-5][0-9]$/.test(endTime)) {
    throw new Error("時間有給の開始・終了時刻はHH:mm形式で指定してください");
  }
  return {
    time_leave_id: String(rowObj.time_leave_id || "").trim(),
    request_id: String(rowObj.request_id || "").trim(),
    employee_id: String(rowObj.employee_id || "").trim(),
    company_code: String(rowObj.company_code || "").trim(),
    leave_date: toSupabaseWriteDate_(rowObj.leave_date),
    start_time: startTime,
    end_time: endTime,
    start_minute: requireSupabaseInteger_(rowObj.start_minute, "start_minute", 0),
    end_minute: requireSupabaseInteger_(rowObj.end_minute, "end_minute", 1),
    requested_minutes: requireSupabaseInteger_(rowObj.requested_minutes, "requested_minutes", 1),
    scheduled_minutes_per_day: requireSupabaseInteger_(rowObj.scheduled_minutes_per_day, "scheduled_minutes_per_day", 1),
    time_leave_unit_minutes: requireSupabaseInteger_(rowObj.time_leave_unit_minutes, "time_leave_unit_minutes", 1),
    work_start_minute: requireSupabaseInteger_(rowObj.work_start_minute, "work_start_minute", 0),
    work_end_minute: requireSupabaseInteger_(rowObj.work_end_minute, "work_end_minute", 1),
    break_periods_json: parseSupabaseBreakPeriods_(rowObj.break_periods_json),
    calculation_version: String(rowObj.calculation_version || "").trim(),
    created_at: toSupabaseWriteTimestamp_(rowObj.created_at),
    updated_at: toSupabaseWriteTimestamp_(rowObj.updated_at)
  };
}

function tryInsertLeaveRequestToSupabase_(rowObj) {
  if (!shouldUseSupabaseWrites_()) return { skipped: true, reason: "USE_SUPABASE_WRITE is not true" };

  try {
    const record = buildSupabaseLeaveRequestRecord_(rowObj);
    const result = supabaseInsert_("leave_requests", record);
    Logger.log("[SupabaseDualWrite] leave_requests inserted request_id=" + record.request_id);
    return {
      ok: true,
      request_id: record.request_id,
      statusCode: result.statusCode
    };
  } catch (err) {
    Logger.log("[SupabaseDualWrite] leave_requests insert failed: " + err.message);
    Logger.log("[SupabaseDualWrite] request=" + JSON.stringify({
      request_id: rowObj && rowObj.request_id,
      employee_id: rowObj && rowObj.employee_id
    }));
    return {
      ok: false,
      error: err.message
    };
  }
}

// 時間休はSpreadsheetの親・子保存が成功してから副DBへ送る。いずれかのSupabase失敗は
// Loggerへ残すだけでSpreadsheetをロールバックしない既存dual write方針を維持する。
function tryInsertTimeLeaveToSupabase_(parentRowObj, segmentRowObj) {
  if (!shouldUseSupabaseWrites_()) return { skipped: true, reason: "USE_SUPABASE_WRITE is not true" };
  try {
    const parent = buildSupabaseLeaveRequestRecord_(parentRowObj);
    supabaseInsert_("leave_requests", parent);
    const segment = buildSupabaseTimeLeaveSegmentRecord_(segmentRowObj);
    supabaseInsert_("time_leave_segments", segment);
    Logger.log("[SupabaseDualWrite] time_leave inserted request_id=" + parent.request_id + " time_leave_id=" + segment.time_leave_id);
    return { ok: true, request_id: parent.request_id, time_leave_id: segment.time_leave_id };
  } catch (err) {
    Logger.log("[SupabaseDualWrite] time_leave insert failed: " + err.message);
    Logger.log("[SupabaseDualWrite] request=" + JSON.stringify({
      request_id: parentRowObj && parentRowObj.request_id,
      time_leave_id: segmentRowObj && segmentRowObj.time_leave_id
    }));
    return { ok: false, error: err.message };
  }
}

function tryUpdateTimeLeaveToSupabase_(parentRowObj, segmentRowObj) {
  if (!shouldUseSupabaseWrites_()) return { skipped: true, reason: "USE_SUPABASE_WRITE is not true" };
  try {
    const parent = buildSupabaseLeaveRequestRecord_(parentRowObj);
    const segment = buildSupabaseTimeLeaveSegmentRecord_(segmentRowObj);
    supabasePatch_("leave_requests", { request_id: "eq." + parent.request_id }, parent);
    supabasePatch_("time_leave_segments", { time_leave_id: "eq." + segment.time_leave_id }, segment);
    Logger.log("[SupabaseDualWrite] time_leave updated request_id=" + parent.request_id + " time_leave_id=" + segment.time_leave_id);
    return { ok: true, request_id: parent.request_id, time_leave_id: segment.time_leave_id };
  } catch (err) {
    Logger.log("[SupabaseDualWrite] time_leave update failed: " + err.message);
    return { ok: false, error: err.message };
  }
}

// SpreadsheetやSupabaseへ接続しないpayload境界テスト。実DB書込みは禁止のため、
// Phase 5ではここで型・時刻・JSON・既存申請互換性を確認する。
function testSupabaseTimeLeaveMappingNoWrite_() {
  const now = new Date("2026-04-01T00:00:00.000Z");
  const parent = {
    request_id: "TEST-TIME-PARENT", employee_id: "MAIN-001", request_date: now,
    start_date: "2026-04-01", end_date: "2026-04-01", days: 0, type: "paid_leave",
    half_day: "", reason: "test", reason_detail: "", status: "pending", year: 2026,
    request_kind: "time_hourly", company_code_snapshot: "MAIN", policy_version: "v1",
    created_at: now, updated_at: now
  };
  const segment = {
    time_leave_id: "TEST-TIME-SEGMENT", request_id: parent.request_id, employee_id: parent.employee_id,
    company_code: "MAIN", leave_date: "2026-04-01", start_time: "08:00", end_time: "09:00",
    start_minute: 480, end_minute: 540, requested_minutes: 60,
    scheduled_minutes_per_day: 420, time_leave_unit_minutes: 60,
    work_start_minute: 480, work_end_minute: 1020,
    break_periods_json: "[{\"startMinute\":600,\"endMinute\":630}]",
    calculation_version: "time_leave_v1", created_at: now, updated_at: now
  };
  const updatedSegment = Object.assign({}, segment, {
    leave_date: "2026-04-02", start_time: "09:00", end_time: "10:00",
    start_minute: 540, end_minute: 600, updated_at: new Date("2026-04-01T01:00:00.000Z")
  });
  const legacyParent = buildSupabaseLeaveRequestRecord_(Object.assign({}, parent, {
    request_id: "TEST-LEGACY", days: 0.5, half_day: "am",
    request_kind: "", company_code_snapshot: "", policy_version: ""
  }));
  const parentRecord = buildSupabaseLeaveRequestRecord_(parent);
  const segmentRecord = buildSupabaseTimeLeaveSegmentRecord_(segment);
  const updatedRecord = buildSupabaseTimeLeaveSegmentRecord_(updatedSegment);
  const cases = [
    ["既存Supabase親payloadは未適用列を送らない", [
      Object.prototype.hasOwnProperty.call(parentRecord, "request_kind"),
      Object.prototype.hasOwnProperty.call(parentRecord, "company_code_snapshot"),
      Object.prototype.hasOwnProperty.call(parentRecord, "policy_version")
    ], [false, false, false]],
    ["子のHH:mmは変換しない", [segmentRecord.start_time, segmentRecord.end_time], ["08:00", "09:00"]],
    ["子の分値は整数", [segmentRecord.start_minute, segmentRecord.end_minute, segmentRecord.requested_minutes], [480, 540, 60]],
    ["制度スナップショット", [segmentRecord.scheduled_minutes_per_day, segmentRecord.time_leave_unit_minutes], [420, 60]],
    ["休憩JSONを保持", segmentRecord.break_periods_json, [{ startMinute: 600, endMinute: 630 }]],
    ["編集payloadの日時・時刻・分", [updatedRecord.leave_date, updatedRecord.start_time, updatedRecord.end_time, updatedRecord.start_minute, updatedRecord.end_minute], ["2026-04-02", "09:00", "10:00", 540, 600]],
    ["既存1日・半日親は追加列NULL", [legacyParent.request_kind, legacyParent.company_code_snapshot, legacyParent.policy_version, legacyParent.days, legacyParent.half_day], [null, null, null, 0.5, "am"]],
    ["繰越分のNULL/0は既存形式と互換", [toSupabaseReadNumber_(null, 0), toSupabaseReadNumber_(0, 0)], [0, 0]]
  ];
  const failures = cases.filter(item => JSON.stringify(item[1]) !== JSON.stringify(item[2]));
  if (failures.length) throw new Error("Supabase時間有給mappingテスト失敗: " + JSON.stringify(failures));
  return { ok: true, case_count: cases.length, cases: cases };
}

function testSupabaseConnection() {
  const result = supabaseGet_("employees", {
    select: "employee_id,name,company_code,employment_status",
    order: "employee_id.asc",
    limit: 5
  });
  const rows = Array.isArray(result.data) ? result.data : [];

  Logger.log("[SupabaseConnectionTest] employees limit=5 count=" + rows.length);
  Logger.log("[SupabaseConnectionTest] content_range=" + result.contentRange);
  Logger.log("[SupabaseConnectionTest] first_row=" + JSON.stringify(rows[0] || null, null, 2));
  Logger.log("[SupabaseConnectionTest] rows=" + JSON.stringify(rows, null, 2));

  return {
    ok: true,
    statusCode: result.statusCode,
    contentRange: result.contentRange,
    count: rows.length,
    firstRow: rows[0] || null
  };
}

function testSupabaseReadEmployees() {
  const rows = getEmployeesFromSupabase_();
  Logger.log("[SupabaseReadTest] employees count=" + rows.length);
  Logger.log("[SupabaseReadTest] employees first=" + JSON.stringify(rows[0] || null, null, 2));
  return { ok: true, table: "employees", count: rows.length, firstRow: rows[0] || null };
}

function testSupabaseReadLeaveRequests() {
  const rows = getLeaveRequestsFromSupabase_();
  Logger.log("[SupabaseReadTest] leave_requests count=" + rows.length);
  Logger.log("[SupabaseReadTest] leave_requests first=" + JSON.stringify(rows[0] || null, null, 2));
  return { ok: true, table: "leave_requests", count: rows.length, firstRow: rows[0] || null };
}

function testSupabaseReadPaidLeaveGrants() {
  const rows = getPaidLeaveGrantsFromSupabase_();
  Logger.log("[SupabaseReadTest] paid_leave_grants count=" + rows.length);
  Logger.log("[SupabaseReadTest] paid_leave_grants first=" + JSON.stringify(rows[0] || null, null, 2));
  return { ok: true, table: "paid_leave_grants", count: rows.length, firstRow: rows[0] || null };
}

function testSupabaseReadCompanyCalendar() {
  const rows = getCompanyCalendarFromSupabase_();
  Logger.log("[SupabaseReadTest] company_calendar count=" + rows.length);
  Logger.log("[SupabaseReadTest] company_calendar first=" + JSON.stringify(rows[0] || null, null, 2));
  return { ok: true, table: "company_calendar", count: rows.length, firstRow: rows[0] || null };
}

function testSupabaseReadAllCoreTables() {
  const result = {
    ok: true,
    use_supabase_reads: shouldUseSupabaseReads_(),
    employees: testSupabaseReadEmployees(),
    leave_requests: testSupabaseReadLeaveRequests(),
    paid_leave_grants: testSupabaseReadPaidLeaveGrants(),
    company_calendar: testSupabaseReadCompanyCalendar(),
    usage_logs: {
      table: "usage_logs",
      count: getUsageLogsFromSupabase_().length
    }
  };

  Logger.log("[SupabaseReadTest] all_core_tables=" + JSON.stringify(result, null, 2));
  return result;
}

function testSupabaseInsertLeaveRequest() {
  if (!shouldUseSupabaseWrites_()) {
    throw new Error("USE_SUPABASE_WRITE=true の時だけ実行できます");
  }

  const employees = getEmployeesFromSupabase_();
  const employee = employees.find(item => String(item.employee_id || "").trim());
  if (!employee) {
    throw new Error("Supabase employees にテスト用employee_idが見つかりません");
  }

  const now = new Date();
  const requestId = "TEST-SUPABASE-INSERT-" + Utilities.formatDate(
    now,
    getAppTimeZone(),
    "yyyyMMddHHmmss"
  );
  const startDate = new Date(now.getFullYear(), now.getMonth(), now.getDate());

  const rowObj = {
    request_id: requestId,
    employee_id: String(employee.employee_id || "").trim(),
    request_date: now,
    start_date: startDate,
    end_date: startDate,
    days: 1,
    type: "paid_leave",
    half_day: "",
    reason: "supabase_dual_write_test",
    reason_detail: "Supabase INSERT connectivity test",
    status: "pending",
    approver_id: "",
    approver_name: "",
    approved_at: "",
    rejected_reason: "",
    cancel_reason: "",
    year: getFiscalYearFromDate(startDate),
    created_at: now,
    updated_at: now
  };

  const record = buildSupabaseLeaveRequestRecord_(rowObj);
  const result = supabaseInsert_("leave_requests", record);
  Logger.log("[SupabaseInsertTest] inserted request_id=" + requestId);

  return {
    ok: true,
    request_id: requestId,
    employee_id: rowObj.employee_id,
    statusCode: result.statusCode
  };
}
