/**
 * 初期表示（管理画面）
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('admin');
}

/**
 * 社員一覧取得
 */
function getEmployees() {
  try {
    const ss = SpreadsheetApp.openById("1o7KbHHsPMiL684YJq_Fpzg6gHjD6HebjeL0BhQImkt4");
    const sheet = ss.getSheetByName("employees");

    if (!sheet) return [];

    const values = sheet.getDataRange().getValues();

    if (values.length <= 1) return [];

    return values.slice(1)
      .filter(row => row[0])
      .map(row => ({
        name: row[0]
      }));

  } catch (e) {
    console.error(e);
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
  const ss = SpreadsheetApp.openById("1o7KbHHsPMiL684YJq_Fpzg6gHjD6HebjeL0BhQImkt4");
  const sheet = ss.getSheetByName("leave_requests");

  const isHalf = data.half_day === true;

  const startDate = new Date(data.start_date);
  const endDate = new Date(data.end_date);

  const days = isHalf ? 0.5 : calculateLeaveDays(startDate, endDate);

  const now = new Date();
  const year = startDate.getFullYear();
  const requestId = Utilities.getUuid();

  sheet.appendRow([
    requestId,
    data.employee_id,
    now,
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
    year,
    now,
    now
  ]);
}

/**
 * 承認待ち一覧取得（admin用）
 */
function getPendingRequests() {
  const ss = SpreadsheetApp.openById("1o7KbHHsPMiL684YJq_Fpzg6gHjD6HebjeL0BhQImkt4");
  const sheet = ss.getSheetByName("leave_requests");

  const data = sheet.getDataRange().getValues();

  if (data.length <= 1) return [];

  return data
    .slice(1)
    .filter(r => (r[10] || "").toString().trim() === "pending")
    .map(r => ({
      request_id: r[0],
      employee_name: r[1], // 仮
      date: r[3],
      days: r[5],
      half_day: r[7]
    }));
}

/**
 * 承認処理
 */
function approveRequest(requestId) {
  const ss = SpreadsheetApp.openById("1o7KbHHsPMiL684YJq_Fpzg6gHjD6HebjeL0BhQImkt4");
  const sheet = ss.getSheetByName("leave_requests");

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

  throw new Error("対象の申請が見つかりません");
}

/**
 * 否認処理
 */
function rejectRequest(requestId) {
  const ss = SpreadsheetApp.openById("1o7KbHHsPMiL684YJq_Fpzg6gHjD6HebjeL0BhQImkt4");
  const sheet = ss.getSheetByName("leave_requests");

  const values = sheet.getDataRange().getValues();

  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === requestId) {

      sheet.getRange(i + 1, 11).setValue("rejected");
      sheet.getRange(i + 1, 18).setValue(new Date());

      return;
    }
  }

  throw new Error("対象の申請が見つかりません");
}

/**
 * デバッグ用
 */
function debugPending() {
  const ss = SpreadsheetApp.openById("1o7KbHHsPMiL684YJq_Fpzg6gHjD6HebjeL0BhQImkt4");
  const sheet = ss.getSheetByName("leave_requests");

  const data = sheet.getDataRange().getValues();

  for (let i = 0; i < data.length; i++) {
    Logger.log(i + "行目 → ステータス：" + data[i][10]);
  }
}

/**
 * 生データ確認用
 */
function checkRawData() {
  const ss = SpreadsheetApp.openById("1o7KbHHsPMiL684YJq_Fpzg6gHjD6HebjeL0BhQImkt4");
  const sheet = ss.getSheetByName("leave_requests");

  const data = sheet.getDataRange().getValues();

  data.forEach((r, i) => {
    Logger.log(i + " → " + JSON.stringify(r));
  });
}

function checkScriptId() {
  Logger.log(ScriptApp.getScriptId());
}

function getPendingRequests() {
  const ss = SpreadsheetApp.openById("1o7KbHHsPMiL684YJq_Fpzg6gHjD6HebjeL0BhQImkt4");
  const reqSheet = ss.getSheetByName("leave_requests");
  const empSheet = ss.getSheetByName("employees");

  const reqData = reqSheet.getDataRange().getValues();
  const empData = empSheet.getDataRange().getValues();

  if (reqData.length <= 1) return [];

  // employee_id → name辞書作成
  const empMap = {};
  empData.slice(1).forEach(r => {
    empMap[r[0]] = r[1]; // ←もし名前がB列ならこれ
  });

  return reqData.slice(1)
    .filter(r => (r[10] || "").toString().trim() === "pending")
    .map(r => ({
      request_id: r[0],
      employee_name: empMap[r[1]] || r[1], // ←ここが重要
      date: r[3],
      days: r[5],
      half_day: r[7]
    }));
}