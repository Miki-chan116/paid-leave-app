/* =========================
   有給ロット期限確認（管理画面）
   debug.gs から動作を変えずに移動
========================= */

function getPaidLeaveExpiryLotsForAdmin(options) {
  const opts = options || {};
  const asOfDate = opts.as_of_date ? parseLocalDate(opts.as_of_date) : parseLocalDate(new Date());
  const keyword = norm(opts.employee_keyword || "");
  const companyCode = String(opts.company_code || "").trim().toUpperCase();
  const expiredOnly = opts.expired_only === true;
  const expiringSoonOnly = opts.expiring_soon_only === true;
  const validityUnconfirmedOnly = opts.validity_unconfirmed_only === true;
  const page = normalizePagingOptions_(opts);
  const context = createFifoBalanceComparisonContext_(asOfDate);
  const rows = [];

  getEmployeesForAdmin()
    .filter(emp => isFifoBalanceCompareTargetEmployee_(emp))
    .filter(emp => {
      if (companyCode && String(emp.company_code || "").trim().toUpperCase() !== companyCode) {
        return false;
      }

      if (!keyword) return true;

      const targetText = norm(
        String(emp.employee_id || "") +
        String(emp.display_employee_id || "") +
        String(emp.name || "") +
        String(emp.display_name || "") +
        String(emp.name_kana || "")
      );
      return targetText.indexOf(keyword) !== -1;
    })
    .forEach(emp => {
      const employeeId = String(emp.employee_id || "").trim();
      const balance = calculateFifoBalanceWithOpeningBalanceFromContext_(
        employeeId,
        asOfDate,
        context
      );

      (balance.grant_details || []).forEach(lot => {
        const remainingDays = lot.is_expired
          ? Number(lot.expired_days || 0)
          : Number(lot.active_remaining_days || 0);
        if (remainingDays <= 0) return;

        const validTo = parseLocalDate(lot.valid_to);
        const daysUntilExpiry = Math.round(
          (validTo.getTime() - asOfDate.getTime()) / (24 * 60 * 60 * 1000)
        );
        const isOpeningBalance = lot.lot_type === "opening_balance_virtual_lot";
        let expiryStatus = "active";
        let expiryStatusLabel = "通常";

        if (lot.is_expired) {
          expiryStatus = "expired";
          expiryStatusLabel = "期限切れ";
        } else if (daysUntilExpiry <= 30) {
          expiryStatus = "within_30";
          expiryStatusLabel = "期限が近い（30日以内）";
        } else if (daysUntilExpiry <= 90) {
          expiryStatus = "within_90";
          expiryStatusLabel = "期限が近い（90日以内）";
        }

        rows.push({
          employee_id: employeeId,
          name: String(emp.name || ""),
          display_name: String(emp.display_name || ""),
          company_code: String(emp.company_code || ""),
          company_name: String(emp.company_name || ""),
          grant_id: String(lot.source_grant_id || lot.grant_id || ""),
          grant_date: String(lot.grant_date || ""),
          valid_from: String(lot.valid_from || ""),
          valid_to: String(lot.valid_to || ""),
          lot_type: isOpeningBalance ? "opening_balance" : "regular_grant",
          lot_type_label: isOpeningBalance ? "初期導入残高" : "通常付与",
          expiry_status: expiryStatus,
          expiry_status_label: expiryStatusLabel,
          granted_days: Number(lot.total_days || 0),
          used_days: Number(lot.used_days || 0),
          remaining_days: remainingDays,
          days_until_expiry: daysUntilExpiry,
          validity_needs_review: lot.validity_needs_review === true,
          validity_basis: String(lot.validity_basis || ""),
          validity_warning: lot.validity_needs_review === true
            ? "有効期限確認が必要です"
            : ""
        });
      });
    });

  const hasStatusFilter = expiredOnly || expiringSoonOnly || validityUnconfirmedOnly;
  const filteredRows = hasStatusFilter
    ? rows.filter(row =>
      (expiredOnly && row.expiry_status === "expired") ||
      (expiringSoonOnly && (
        row.expiry_status === "within_30" ||
        row.expiry_status === "within_90"
      )) ||
      (validityUnconfirmedOnly && row.validity_needs_review)
    )
    : rows;
  const priority = {
    expired: 0,
    within_30: 1,
    within_90: 2,
    active: 3
  };

  filteredRows.sort((a, b) => {
    if (priority[a.expiry_status] !== priority[b.expiry_status]) {
      return priority[a.expiry_status] - priority[b.expiry_status];
    }
    if (a.days_until_expiry !== b.days_until_expiry) {
      return a.days_until_expiry - b.days_until_expiry;
    }
    if (a.employee_id !== b.employee_id) {
      return a.employee_id.localeCompare(b.employee_id);
    }
    if (a.grant_date === b.grant_date) return 0;
    return a.grant_date < b.grant_date ? -1 : 1;
  });

  const pageRows = filteredRows.slice(page.offset, page.offset + page.limit);

  return {
    ok: true,
    as_of_date: formatDateValue(asOfDate),
    total_count: filteredRows.length,
    row_count: pageRows.length,
    offset: page.offset,
    limit: page.limit,
    has_prev: page.offset > 0,
    has_next: page.offset + page.limit < filteredRows.length,
    expired_count: filteredRows.filter(row => row.expiry_status === "expired").length,
    within_30_count: filteredRows.filter(row => row.expiry_status === "within_30").length,
    within_90_count: filteredRows.filter(row => row.expiry_status === "within_90").length,
    needs_review_count: filteredRows.filter(row => row.validity_needs_review).length,
    rows: pageRows
  };
}
