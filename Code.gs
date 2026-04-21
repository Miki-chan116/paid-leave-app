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
    .setTitle("有給申請管理画面");
}

/* =========================
   シート取得ヘルパー
========================= */
function getSheet(name) {
  const ss = SpreadsheetApp.openById(SS_ID);
  return ss.getSheetByName(name);
}

/* =========================
   文字正規化
   空白除去 + 小文字化
========================= */
function norm(v) {
  return String(v == null ? "" : v)
    .replace(/\s/g, "")
    .toLowerCase();
}

/* =========================
   日付表示用
========================= */
function formatDateValue(value) {
  if (!value) return "";

  const date = new Date(value);
  if (isNaN(date.getTime())) return String(value);

  return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy/MM/dd");
}

/* =========================
   社員取得
========================= */
function getEmployees() {
  const sheet = getSheet("employees");
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  return data.slice(1).map(r => ({
    id: String(r[0] || "").trim(),
    name: r[1] || r[0] || ""
  }));
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
   申請登録
========================= */
function submitLeaveRequest(data) {
  const sheet = getSheet("leave_requests");
  if (!sheet) {
    throw new Error("leave_requests シートが見つかりません");
  }

  const start = new Date(data.start_date);
  const end = new Date(data.end_date);

  const isHalf = data.half_day === true;
  const days = isHalf ? 0.5 : calculateLeaveDays(start, end);

  sheet.appendRow([
    Utilities.getUuid(),                // 0: request_id
    data.employee_id,                   // 1: employee_id
    new Date(),                         // 2: applied_at
    start,                              // 3: start_date
    end,                                // 4: end_date
    days,                               // 5: days
    "有給",                             // 6: leave_type
    isHalf ? data.half_type : "",       // 7: half_type
    data.reason || "",                  // 8: reason
    "",                                 // 9: manager_comment
    STATUS.PENDING,                     // 10: status
    "",                                 // 11: approver_id
    "",                                 // 12: approver_name
    "",                                 // 13: approved_at
    "",                                 // 14: reject_reason
    start.getFullYear(),                // 15: year
    new Date(),                         // 16: created_at
    new Date()                          // 17: updated_at
  ]);
}

/* =========================
   申請一覧取得
   K列(11列目 / index 10) を status として扱う
========================= */
function getRequestsByStatus(status) {
  try {
    const sheet = getSheet("leave_requests");
    const empSheet = getSheet("employees");

    if (!sheet) {
      throw new Error("leave_requests シートが見つかりません");
    }

    const reqData = sheet.getDataRange().getValues();
    if (reqData.length <= 1) {
      Logger.log("leave_requests にデータ行がありません");
      return [];
    }

    const empMap = {};
    if (empSheet) {
      const empData = empSheet.getDataRange().getValues();
      empData.slice(1).forEach(r => {
        const empId = String(r[0] || "").trim();
        const empName = r[1] || empId;
        if (empId) {
          empMap[empId] = empName;
        }
      });
    }

    const target = norm(status);
    Logger.log("=== getRequestsByStatus start ===");
    Logger.log("requested status = [" + target + "]");

    const result = [];

    reqData.slice(1).forEach((r, index) => {
      const rowNumber = index + 2;

      const requestId = r[0];
      const employeeId = String(r[1] || "").trim();
      const rawStartDate = r[3];
      const rawEndDate = r[4];
      const rawDays = r[5];
      const rawHalfType = r[7];
      const rawReason = r[8];
      const rawStatus = r[10];

      const normalizedStatus = norm(rawStatus);

      Logger.log(
        "row " + rowNumber +
        " | request_id=" + requestId +
        " | rawStatus=[" + rawStatus + "]" +
        " | normalized=[" + normalizedStatus + "]"
      );

      if (normalizedStatus === target) {
        result.push({
          request_id: requestId,
          employee_id: employeeId,
          employee_name: empMap[employeeId] || employeeId || "不明",
          start_date: formatDateValue(rawStartDate),
          end_date: formatDateValue(rawEndDate),
          date_label: formatDateValue(rawStartDate) + (
            formatDateValue(rawStartDate) !== formatDateValue(rawEndDate)
              ? " 〜 " + formatDateValue(rawEndDate)
              : ""
          ),
          days: rawDays || 0,
          half_day: rawHalfType || "",
          reason: rawReason || "",
          status: normalizedStatus
        });
      }
    });

    Logger.log("result length = " + result.length);
    Logger.log("=== getRequestsByStatus end ===");

    return result;

  } catch (e) {
    Logger.log("ERROR in getRequestsByStatus: " + e.toString());
    throw new Error("申請一覧取得でエラー: " + e.message);
  }
}

/* =========================
   承認
========================= */
function approveRequest(requestId) {
  const sheet = getSheet("leave_requests");
  if (!sheet) {
    throw new Error("leave_requests シートが見つかりません");
  }

  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(requestId)) {
      sheet.getRange(i + 1, 11).setValue(STATUS.APPROVED); // K列
      sheet.getRange(i + 1, 12).setValue("A001");          // L列
      sheet.getRange(i + 1, 13).setValue("管理者");         // M列
      sheet.getRange(i + 1, 14).setValue(new Date());      // N列
      sheet.getRange(i + 1, 18).setValue(new Date());      // R列
      return { ok: true };
    }
  }

  throw new Error("対象の申請が見つかりません");
}

/* =========================
   否認
========================= */
function rejectRequest(requestId, reason) {
  const sheet = getSheet("leave_requests");
  if (!sheet) {
    throw new Error("leave_requests シートが見つかりません");
  }

  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(requestId)) {
      sheet.getRange(i + 1, 11).setValue(STATUS.REJECTED); // K列
      sheet.getRange(i + 1, 15).setValue(reason || "");    // O列
      sheet.getRange(i + 1, 18).setValue(new Date());      // R列
      return { ok: true };
    }
  }

  throw new Error("対象の申請が見つかりません");
}

/* =========================
   ログ取得
========================= */
function getUsageLogs() {
  const sheet = getSheet("usage_log");
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  return data.slice(1).map(r => ({
    log_id: r[0],
    request_id: r[1],
    type: r[2],
    user_id: r[3],
    user_name: r[4],
    date: formatDateValue(r[5]),
    comment: r[6]
  })).sort((a, b) => new Date(b.date) - new Date(a.date));
}

/* =========================
   デバッグ用
   K列(status)の中身を確認
========================= */
function debugStatusValues() {
  const sheet = getSheet("leave_requests");
  if (!sheet) {
    Logger.log("leave_requests シートなし");
    return;
  }

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) {
    Logger.log("データなし");
    return;
  }

  data.slice(1).forEach((r, i) => {
    Logger.log(
      "row " + (i + 2) +
      " | request_id=" + r[0] +
      " | rawStatus=[" + r[10] + "]" +
      " | normalized=[" + norm(r[10]) + "]"
    );
  });
}