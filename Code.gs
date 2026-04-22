const SS_ID = "1o7KbHHsPMiL684YJq_Fpzg6gHjD6HebjeL0BhQImkt4";

/* =========================
   ステータス定義
========================= */
const STATUS = {
  PENDING: "pending",
  APPROVED: "approved",
  REJECTED: "rejected"
};

/* =========================
   管理画面
========================= */
function doGet() {
  return HtmlService.createHtmlOutputFromFile("admin")
    .setTitle("Paid Leave Admin");
}

/* =========================
   シート取得
========================= */
function getSheet(name) {
  const ss = SpreadsheetApp.openById(SS_ID);
  const sheet = ss.getSheetByName(name);

  if (!sheet) {
    throw new Error(name + " シートが見つかりません");
  }

  return sheet;
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

  return Utilities.formatDate(
    date,
    Session.getScriptTimeZone(),
    "yyyy/MM/dd"
  );
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

  return {
    headers,
    map
  };
}

/* =========================
   必須ヘッダーチェック
========================= */
function requireHeaders(sheet, requiredHeaders) {
  const headerInfo = getHeaderMap(sheet);
  const missing = requiredHeaders.filter(h => !(h in headerInfo.map));

  if (missing.length > 0) {
    throw new Error(
      sheet.getName() + " に不足ヘッダーがあります: " + missing.join(", ")
    );
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

/* =========================
   社員名マップ
========================= */
function getEmployeeMap() {
  const employees = getEmployees();
  const map = {};

  employees.forEach(emp => {
    map[emp.id] = emp.name;
  });

  return map;
}

/* =========================
   土日除外
========================= */
function calculateLeaveDays(startDate, endDate) {
  const start = new Date(startDate);
  const end = new Date(endDate);

  let count = 0;

  while (start <= end) {
    const day = start.getDay();
    if (day !== 0 && day !== 6) {
      count++;
    }
    start.setDate(start.getDate() + 1);
  }

  return count;
}

/* =========================
   使用ログ追加
   usage_log:
   log_id, request_id, action_type, operator_id,
   operator_name, action_date, comment
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
   手動実行時の undefined ガード付き
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

  const start = new Date(data.start_date);
  const end = new Date(data.end_date);

  if (isNaN(start.getTime()) || isNaN(end.getTime())) {
    throw new Error("日付が不正です");
  }

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
  rowObj.year = start.getFullYear();
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

  return {
    ok: true,
    request_id: rowObj.request_id
  };
}

/* =========================
   申請一覧取得
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

  const result = data.slice(1)
    .map(row => {
      const rowObj = rowToObject(row, headerInfo.headers);
      const rowStatus = norm(rowObj.status);
      const employeeId = String(rowObj.employee_id || "").trim();

      return {
        request_id: String(rowObj.request_id || ""),
        employee_id: employeeId,
        employee_name: String(
          employeeMap[employeeId] || employeeId || "Unknown"
        ),
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
        status: rowStatus
      };
    })
    .filter(item => item.status === target);

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
   ステータス確認用
========================= */
function debugStatusValues() {
  const sheet = getSheet("leave_requests");
  const headerInfo = requireHeaders(sheet, [
    "request_id",
    "status"
  ]);

  const data = sheet.getDataRange().getValues();

  if (data.length <= 1) {
    Logger.log("データなし");
    return;
  }

  data.slice(1).forEach((row, i) => {
    const rowObj = rowToObject(row, headerInfo.headers);
    Logger.log(
      "row " + (i + 2) +
      " | request_id=" + rowObj.request_id +
      " | rawStatus=[" + rowObj.status + "]" +
      " | normalized=[" + norm(rowObj.status) + "]"
    );
  });
}

/* =========================
   手動確認用のテスト関数
   必要なときだけ使う
========================= */
function testSubmitLeaveRequest() {
  return submitLeaveRequest({
    employee_id: "TEST001",
    start_date: "2026-04-23",
    end_date: "2026-04-23",
    half_day: false,
    half_type: "",
    reason: "テスト申請",
    reason_detail: ""
  });
}