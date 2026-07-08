/* =========================
   Supabase接続検証（読み取り専用）

   セキュリティ注意:
   - SUPABASE_URL / SUPABASE_ANON_KEY はコードに直書きせず、
     Apps Script の Script Properties に設定してください。
   - SERVICE_ROLE_KEY はGASクライアント検証では使用しません。
   - anon keyで読み取りを許可する場合は、Supabase側でRLS policyを
     読み取り専用かつ必要最小限に設計してください。
   - このファイルの関数はGET専用です。既存のSpreadsheet処理、
     申請登録、承認、取消処理からは呼び出していません。
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

function shouldUseSupabaseReads_() {
  const value = PropertiesService
    .getScriptProperties()
    .getProperty("USE_SUPABASE_READS");

  return String(value || "").trim().toLowerCase() === "true";
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
