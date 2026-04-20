const SPREADSHEET_ID = "1o7KbHHsPMiL684YJq_Fpzg6gHjD6HebjeL0BhQImkt4";
const SHEET_NAME = "leave_requests";

/**
 * Webアプリ表示
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('admin');
}

/**
 * シート取得
 */
function getSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    throw new Error("シートが見つかりません: " + SHEET_NAME);
  }

  return sheet;
}

/**
 * 全データ取得（安全版）
 */
function getAllRequests() {
  const sheet = getSheet();
  const values = sheet.getDataRange().getValues();

  if (!values || values.length < 2) return [];

  const headers = values[0];
  const result = [];

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const obj = {};

    headers.forEach((h, index) => {
      let val = row[index];

      // Date対策（ここが今回の重要ポイント）
      if (val instanceof Date) {
        val = Utilities.formatDate(val, "Asia/Tokyo", "yyyy-MM-dd HH:mm:ss");
      }

      obj[h] = val;
    });

    result.push(obj);
  }

  return result;
}

/**
 * ステータス絞り込み（安全版）
 */
function getRequestsByStatus(status) {
  const data = getAllRequests();

  const target = String(status || "").trim().toLowerCase();

  return data.filter(r => {
    return String(r.status || "").trim().toLowerCase() === target;
  });
}

/**
 * デバッグ用
 */
function debugData() {
  const data = getAllRequests();
  Logger.log(data);
}

/**
 * 通信テスト
 */
function ping() {
  return "pong";
}

/**
 * テスト返却
 */
function testFinal() {
  return [{ ok: "WORKING" }];
}

function updateStatus(request_id, newStatus) {
  const sheet = getSheet();
  const values = sheet.getDataRange().getValues();
  const headers = values[0];

  const idCol = headers.indexOf("request_id");
  const statusCol = headers.indexOf("status");

  for (let i = 1; i < values.length; i++) {
    if (values[i][idCol] == request_id) {
      sheet.getRange(i + 1, statusCol + 1).setValue(newStatus);
      return { success: true };
    }
  }

  return { success: false, message: "not found" };
}