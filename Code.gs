/**
 * =================================================
 * 有給申請保存
 * =================================================
 */
function submitLeaveRequest(data) {
  try {
    const ss = SpreadsheetApp.openById("1o7KbHHsPMiL684YJq_Fpzg6gHjD6HebjeL0BhQImkt4");

    const leaveSheet = ss.getSheetByName("leave_requests");
    const calendarSheet = ss.getSheetByName("company_calendar");
    const logSheet = ss.getSheetByName("usage_log");

    if (!leaveSheet) throw new Error("leave_requests シートがありません");

    /*****************************************
     * ① 日数計算（営業日ベース）
     *****************************************/
    const start = new Date(data.start_date);
    const end = new Date(data.end_date);

    const calendar = calendarSheet.getDataRange().getValues();

    let days = 0;

    for (let i = 1; i < calendar.length; i++) {
      const date = new Date(calendar[i][0]);
      const type = calendar[i][1];

      if (date >= start && date <= end) {
        if (type === "営業日") {
          days++;
        }
      }
    }

    /*****************************************
     * ② ID採番（簡易）
     *****************************************/
    const requestId = "R" + new Date().getTime();

    /*****************************************
     * ③ leave_requestsに保存
     *****************************************/
    leaveSheet.appendRow([
      requestId,                 // request_id
      data.employee_id,         // employee_id
      new Date(),               // 申請日
      data.start_date,          // 開始日
      data.end_date,            // 終了日
      days,                     // 取得日数
      data.type || "1日",       // 区分
      data.half || "",          // 半日区分
      data.reason || "",        // 理由
      data.reason_detail || "", // 理由詳細
      "pending",               // ステータス
      "", "", "", "",           // 承認者系（空）
      new Date().getFullYear(), // 年度
      new Date(),              // 作成日時
      new Date()               // 更新日時
    ]);

    /*****************************************
     * ④ ログ記録（任意）
     *****************************************/
    if (logSheet) {
      logSheet.appendRow([
        "L" + new Date().getTime(),
        requestId,
        "申請",
        data.employee_id,
        data.name || "",
        new Date(),
        data.reason || ""
      ]);
    }

    /*****************************************
     * ⑤ レスポンス
     *****************************************/
    return {
      success: true,
      request_id: requestId,
      days: days
    };

  } catch (e) {
    console.error("submitLeaveRequest error:", e);

    return {
      success: false,
      message: e.message
    };
  }
}
function approveRequest(requestId) {
  try {
    const ss = SpreadsheetApp.openById("1o7KbHHsPMiL684YJq_Fpzg6gHjD6HebjeL0BhQImkt4");

    const sheet = ss.getSheetByName("leave_requests");
    const logSheet = ss.getSheetByName("usage_log");

    const data = sheet.getDataRange().getValues();

    const now = new Date();
    const approverEmail = Session.getActiveUser().getEmail();

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === requestId) {

        // ステータス更新
        sheet.getRange(i + 1, 11).setValue("approved"); // ステータス列
        sheet.getRange(i + 1, 12).setValue(approverEmail); // 承認者ID
        sheet.getRange(i + 1, 14).setValue(now); // 承認日時

        // ログ
        if (logSheet) {
          logSheet.appendRow([
            "L" + new Date().getTime(),
            requestId,
            "承認",
            approverEmail,
            approverEmail,
            now,
            ""
          ]);
        }

        return {
          success: true,
          message: "承認しました"
        };
      }
    }

    return {
      success: false,
      message: "対象データが見つかりません"
    };

  } catch (e) {
    return {
      success: false,
      message: e.message
    };
  }
}

function getPendingRequests() {
  const ss = SpreadsheetApp.openById("1o7KbHHsPMiL684YJq_Fpzg6gHjD6HebjeL0BhQImkt4");
  const sheet = ss.getSheetByName("leave_requests");

  const data = sheet.getDataRange().getValues();

  const headers = data[0];
  const rows = data.slice(1);

  const result = rows
    .filter(r => r[0] && r[10] === "pending") // status列
    .map(r => {
      return {
        request_id: r[0],
        employee_id: r[1],
        start_date: r[3],
        end_date: r[4],
        reason: r[8],
        status: r[10],
        name: r[1] // 仮（後で社員マスタ結合できる）
      };
    });

  return result;
}

function testAll() {
  const days = calculateLeaveDays("2026-04-01", "2026-04-10");
  Logger.log("日数：" + days);
}