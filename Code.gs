function doGet() {
  return HtmlService
    .createTemplateFromFile('index')
    .evaluate();
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// 社員一覧取得
function getEmployees() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("社員マスター");
  const data = sheet.getDataRange().getValues();

  // 1行目（ヘッダー）削除
  data.shift();

  // 必要な形に整形
  return data.map(row => ({
    id: row[0],
    name: row[1]
  }));
}