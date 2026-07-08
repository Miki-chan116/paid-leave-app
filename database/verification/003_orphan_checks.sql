-- =========================================================
-- Paid Leave App / Supabase Verification
-- 003_orphan_checks.sql
--
-- Purpose:
--   Confirm employee_id references resolve to employees.
--   With active foreign keys, these should normally return zero rows.
-- =========================================================

-- leave_requests.employee_id not found in employees
select
  'leave_requests.employee_id' as check_name,
  lr.employee_id,
  count(*) as row_count
from leave_requests lr
left join employees e
  on e.employee_id = lr.employee_id
where e.employee_id is null
group by lr.employee_id
order by lr.employee_id;

-- paid_leave_grants.employee_id not found in employees
select
  'paid_leave_grants.employee_id' as check_name,
  plg.employee_id,
  count(*) as row_count
from paid_leave_grants plg
left join employees e
  on e.employee_id = plg.employee_id
where e.employee_id is null
group by plg.employee_id
order by plg.employee_id;

-- usage_logs.employee_id not found in employees
select
  'usage_logs.employee_id' as check_name,
  ul.employee_id,
  count(*) as row_count
from usage_logs ul
left join employees e
  on e.employee_id = ul.employee_id
where ul.employee_id is not null
  and e.employee_id is null
group by ul.employee_id
order by ul.employee_id;
