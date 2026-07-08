-- =========================================================
-- Paid Leave App / Supabase Indexes
-- 002_indexes.sql
--
-- Keep indexes focused on current GAS query patterns:
--   - employee/status/date search
--   - yearly balance and grant lookups
--   - usage log search and migration reconciliation
-- =========================================================

begin;

create index if not exists idx_employees_company_code
  on employees(company_code);

create index if not exists idx_employees_employment_status
  on employees(employment_status);

create index if not exists idx_employees_leave_management_target
  on employees(leave_management_target);

create index if not exists idx_employees_deleted_at
  on employees(deleted_at);

create index if not exists idx_leave_requests_employee_id
  on leave_requests(employee_id);

create index if not exists idx_leave_requests_status
  on leave_requests(status);

create index if not exists idx_leave_requests_start_date
  on leave_requests(start_date);

create index if not exists idx_leave_requests_employee_status_start
  on leave_requests(employee_id, status, start_date);

create index if not exists idx_leave_requests_status_start_end
  on leave_requests(status, start_date, end_date);

create index if not exists idx_leave_requests_year
  on leave_requests(year);

create index if not exists idx_paid_leave_grants_employee_id
  on paid_leave_grants(employee_id);

create index if not exists idx_paid_leave_grants_employee_year
  on paid_leave_grants(employee_id, year);

create index if not exists idx_paid_leave_grants_employee_type_year
  on paid_leave_grants(employee_id, grant_type, year);

create index if not exists idx_paid_leave_grants_type
  on paid_leave_grants(grant_type);

create index if not exists idx_paid_leave_grants_valid_range
  on paid_leave_grants(valid_from, valid_to);

create index if not exists idx_usage_logs_action_date
  on usage_logs(action_date);

create index if not exists idx_usage_logs_target
  on usage_logs(target_type, target_id);

create index if not exists idx_usage_logs_legacy_request_id
  on usage_logs(legacy_request_id);

create index if not exists idx_usage_logs_leave_request_id
  on usage_logs(leave_request_id);

create index if not exists idx_usage_logs_employee_id
  on usage_logs(employee_id);

commit;
