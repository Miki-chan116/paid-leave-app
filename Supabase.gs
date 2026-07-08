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
