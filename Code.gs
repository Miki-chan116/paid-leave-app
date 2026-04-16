function doGet() {
  return HtmlService.createTemplateFromFile('index').evaluate();
}

/**
 * 社員取得（安全・固定キー版）
 */
function getEmployees() {
  try {
    const ss = SpreadsheetApp.openById("【スプレッドシートID】");
    const sheet = ss.getSheetByName("社員マスター");

    if (!sheet) return [];

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];

    const headers = data[0];
    const rows = data.slice(1);

    return rows
      .filter(r => r[0])
      .map(row => {
        const obj = {};
        headers.forEach((h, i) => {
          obj[h] = row[i] ?? "";
        });

        return {
          name: obj["氏名"] || "",
          raw: obj
        };
      });

  } catch (e) {
    console.error(e);
    return [];
  }
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}