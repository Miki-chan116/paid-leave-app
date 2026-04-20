/**
 * 初期表示（管理画面）
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('admin');
}

/**
 * スプレッドシート取得
 */
function getSS() {
  return SpreadsheetApp.openById("1o7KbHHsPMiL684YJq_Fpzg6gHjD6HebjeL0BhQImkt4");
}

/**
 * 社員一覧取得
 */
function getEmployees() {
  try {
    const ss = getSS();
    const sheet = ss.getSheetByName("employees");
    if (!sheet) return [];

    const values = sheet.getDataRange().getValues();
    if (values.length <= 1) return [];

    return values.slice(1)
      .filter(r => r[0])
      .map(r => ({
        id: String(r[0]),
        name: String(r[1] || r[0])
      }));

  } catch (e) {
    Logger.log(e);
    return [];
  }
}

/**
 * 有給日数計算（土日除外）
 */
function calculateLeaveDays(startDate, endDate) {
  const start = new Date(startDate);
  const end = new Date(endDate);

  let count = 0;

  while (start <= end) {
    const day = start.getDay();
    if (day !== 0 && day !== 6) count++;
    start.setDate(start.getDate() + 1);
  }

  return count;
}

/**
 * 有給申請保存
 */
function submitLeaveRequest(data) {
  const ss = getSS();
  const sheet = ss.getSheetByName("leave_requests");
  if (!sheet) throw new Error("leave_requestsシートなし");

  const isHalf = data.half_day === true;

  const startDate = new Date(data.start_date);
  const endDate = new Date(data.end_date);

  const days = isHalf ? 0.5 : calculateLeaveDays(startDate, endDate);

  sheet.appendRow([
    Utilities.getUuid(),
    data.employee_id,
    new Date(),
    startDate,
    endDate,
    days,
    "有給",
    isHalf ? data.half_type : "",
    data.reason,
    "",
    "pending",
    "",
    "",
    "",
    "",
    startDate.getFullYear(),
    new Date(),
    new Date()
  ]);
}

/**
 * 承認待ち一覧取得（完全安定版）
 */
function getPendingRequests() {
  try {
    const ss = getSS();
    const reqSheet = ss.getSheetByName("leave_requests");
    const empSheet = ss.getSheetByName("employees");

    if (!reqSheet) return [];

    const reqData = reqSheet.getDataRange().getValues();

    // 社員マップ
    let empMap = {};
    if (empSheet) {
      const empData = empSheet.getDataRange().getValues();
      empData.slice(1).forEach(r => {
        empMap[r[0]] = r[1];
      });
    }

    return reqData.slice(1)
      .filter(r => String(r[10] || "").trim() === "pending")
      .map(r => ({
        request_id: String(r[0] || ""),
        employee_name: String(empMap[r[1]] || r[1] || ""),
        date: r[3] ? Utilities.formatDate(new Date(r[3]), Session.getScriptTimeZone(), "yyyy/MM/dd") : "",
        days: Number(r[5] || 0),
        half_day: String(r[7] || "")
      }));

  } catch (e) {
    Logger.log(e);
    return [];
  }
}

/**
 * 承認処理
 */
function approveRequest(requestId) {
  const ss = getSS();
  const sheet = ss.getSheetByName("leave_requests");
  if (!sheet) throw new Error("シートなし");

  const values = sheet.getDataRange().getValues();

  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === requestId) {
      sheet.getRange(i + 1, 11).setValue("approved");
      sheet.getRange(i + 1, 12).setValue("A001");
      sheet.getRange(i + 1, 13).setValue("管理者");
      sheet.getRange(i + 1, 14).setValue(new Date());
      sheet.getRange(i + 1, 18).setValue(new Date());
      return;
    }
  }

  throw new Error("対象なし");
}

/**
 * 否認処理
 */
function rejectRequest(requestId, reason) {
  const ss = getSS();
  const sheet = ss.getSheetByName("leave_requests");
  if (!sheet) throw new Error("シートなし");

  const values = sheet.getDataRange().getValues();

  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === requestId) {
      sheet.getRange(i + 1, 11).setValue("rejected");
      sheet.getRange(i + 1, 15).setValue(reason || "");
      sheet.getRange(i + 1, 18).setValue(new Date());
      return;
    }
  }

  throw new Error("対象なし");
}

/**
 * デバッグ
 */
function debugStatusFinal() {
  const ss = getSS();
  const sheet = ss.getSheetByName("leave_requests");

  const data = sheet.getDataRange().getValues();

  data.slice(1).forEach((r, i) => {
    Logger.log(i + " | [" + r[10] + "] type=" + typeof r[10]);
  });
}