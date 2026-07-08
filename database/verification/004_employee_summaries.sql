-- =========================================================
-- Paid Leave App / Supabase Verification
-- 004_employee_summaries.sql
--
-- Purpose:
--   Compare per-employee paid leave usage and grant totals with GAS results.
-- =========================================================

-- Employee-level approved paid leave usage.
-- canceled, canceled_by_admin, rejected, and pending requests are excluded.
select
  e.employee_id,
  e.name,
  e.company_code,
  coalesce(sum(lr.days), 0) as approved_used_days,
  count(lr.request_id) as approved_request_count
from employees e
left join leave_requests lr
  on lr.employee_id = e.employee_id
  and lr.status = 'approved'
group by e.employee_id, e.name, e.company_code
order by e.employee_id;

-- Employee-level grant totals.
select
  e.employee_id,
  e.name,
  e.company_code,
  coalesce(sum(plg.grant_days), 0) as total_grant_days,
  coalesce(sum(plg.carry_over_days), 0) as total_carry_over_days,
  coalesce(sum(plg.grant_days + plg.carry_over_days), 0) as total_available_granted_days,
  count(plg.grant_id) as grant_count
from employees e
left join paid_leave_grants plg
  on plg.employee_id = e.employee_id
group by e.employee_id, e.name, e.company_code
order by e.employee_id;
