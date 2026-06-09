/* =========================
   廃止済み年度切替API
   互換のため関数名だけ残し、実行時は即エラーにする
========================= */

function finalizeYearEndCarryOver(employeeId, fiscalYear, adminUser) {
  throw new Error("この年度切替確定処理は廃止されました。会社別年度切替パネルを使用してください。");
}

function finalizeSelectedYearEndCarryOver(employeeIds, fiscalYear, adminUser) {
  throw new Error("この年度切替確定処理は廃止されました。会社別年度切替パネルを使用してください。");
}

function executePartnerLeaveYearRollover2026() {
  throw new Error("2026専用の年度切替実行関数は廃止されました。会社別年度切替パネルを使用してください。");
}
