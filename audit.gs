/* =========================
   Audit utilities
   debug.gs から動作を変えずに移動
========================= */

function auditCarryOverGrantRowsForFifoMigration() {
  const sheet = getSheet("paid_leave_grants");
  const headerInfo = requireHeaders(sheet, [
    "grant_id",
    "employee_id",
    "grant_date",
    "grant_days",
    "carry_over_days",
    "valid_from",
    "valid_to",
    "grant_type",
    "year",
    "notes"
  ]);
  const data = sheet.getDataRange().getValues();
  const employeeMap = getEmployeeDetailMap();

  if (data.length <= 1) {
    return {
      row_count: 0,
      needs_review_count: 0,
      carry_over_days_total: 0,
      needs_review_carry_over_days_total: 0,
      rows: []
    };
  }

  const allRows = data.slice(1)
    .map(row => rowToObject(row, headerInfo.headers))
    .filter(rowObj => rowObj.employee_id && rowObj.grant_date)
    .map(rowObj => ({
      grant_id: String(rowObj.grant_id || ""),
      employee_id: String(rowObj.employee_id || "").trim(),
      grant_date_value: parseLocalDate(rowObj.grant_date),
      grant_date: formatDateValue(rowObj.grant_date),
      grant_days: Number(rowObj.grant_days || 0),
      carry_over_days: Number(rowObj.carry_over_days || 0),
      valid_from: formatDateValue(rowObj.valid_from || rowObj.grant_date),
      valid_to: formatDateValue(rowObj.valid_to || ""),
      grant_type: String(rowObj.grant_type || ""),
      year: rowObj.year || "",
      notes: String(rowObj.notes || "")
    }));
  const carryOverRows = allRows.filter(row => row.carry_over_days > 0);
  const resultRows = carryOverRows.map(row => {
    const earlierGrantRows = allRows.filter(candidate =>
      candidate.employee_id === row.employee_id &&
      candidate.grant_id !== row.grant_id &&
      candidate.grant_days > 0 &&
      candidate.grant_date_value < row.grant_date_value
    );
    const hasEarlierGrantDaysRecord = earlierGrantRows.length > 0;
    const isCarryOverOnlyRow = row.grant_days <= 0;
    const needsReview = isCarryOverOnlyRow || !hasEarlierGrantDaysRecord;
    const employee = employeeMap[row.employee_id] || {};

    return {
      employee_id: row.employee_id,
      name: String(employee.name || ""),
      display_name: String(employee.display_name || ""),
      grant_id: row.grant_id,
      grant_date: row.grant_date,
      grant_type: row.grant_type,
      year: row.year,
      grant_days: row.grant_days,
      carry_over_days: row.carry_over_days,
      valid_from: row.valid_from,
      valid_to: row.valid_to,
      notes: row.notes,
      is_carry_over_only_row: isCarryOverOnlyRow,
      has_earlier_grant_days_record: hasEarlierGrantDaysRecord,
      earlier_grant_ids: earlierGrantRows.map(candidate => candidate.grant_id),
      needs_review: needsReview,
      review_reason: isCarryOverOnlyRow
        ? "carry_over_days のみで残を持つ行です。FIFO除外前に権利の元データを確認してください。"
        : !hasEarlierGrantDaysRecord
          ? "繰越元となる過去の grant_days 行が確認できません。初期移行残または手入力残の可能性があります。"
          : "過去の grant_days 行があり、年度集計用 carry_over_days の可能性があります。"
    };
  });
  const needsReviewRows = resultRows.filter(row => row.needs_review);
  const result = {
    row_count: resultRows.length,
    needs_review_count: needsReviewRows.length,
    carry_over_days_total: resultRows.reduce(
      (sum, row) => sum + Number(row.carry_over_days || 0),
      0
    ),
    needs_review_carry_over_days_total: needsReviewRows.reduce(
      (sum, row) => sum + Number(row.carry_over_days || 0),
      0
    ),
    rows: resultRows
  };

  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

