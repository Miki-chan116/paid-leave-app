-- =========================================================
-- Paid Leave App / Supabase Verification
-- 002_distributions.sql
--
-- Purpose:
--   Compare key value distributions with Spreadsheet/GAS audit results.
-- =========================================================

-- leave_requests.status distribution
select
  status,
  count(*) as request_count
from leave_requests
group by status
order by status;

-- employees.company_code distribution
select
  company_code,
  count(*) as employee_count
from employees
group by company_code
order by company_code;

-- paid_leave_grants.grant_type distribution
select
  grant_type,
  count(*) as grant_count
from paid_leave_grants
group by grant_type
order by grant_type;
