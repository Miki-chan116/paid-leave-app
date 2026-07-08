-- =========================================================
-- Paid Leave App / Supabase Verification
-- 001_counts.sql
--
-- Purpose:
--   Compare imported Supabase row counts with Spreadsheet CSV row counts.
-- =========================================================

select
  'employees' as table_name,
  count(*) as row_count
from employees
union all
select
  'leave_requests' as table_name,
  count(*) as row_count
from leave_requests
union all
select
  'paid_leave_grants' as table_name,
  count(*) as row_count
from paid_leave_grants
union all
select
  'company_calendar' as table_name,
  count(*) as row_count
from company_calendar
union all
select
  'usage_logs' as table_name,
  count(*) as row_count
from usage_logs
union all
select
  'admin_users' as table_name,
  count(*) as row_count
from admin_users
order by table_name;
