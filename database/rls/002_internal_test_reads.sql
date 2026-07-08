-- =========================================================
-- Paid Leave App / Supabase RLS Draft
-- 002_internal_test_reads.sql
--
-- Purpose:
--   Internal test-only anon read policy candidates for GAS read verification.
--
-- Important:
--   - Review carefully before running in Supabase.
--   - These policies expose employee and leave-management data to anon role.
--   - Use only for limited internal verification, not production operation.
--   - admin_users is intentionally excluded and must not be exposed to anon.
--   - This file only creates SELECT policies. It does not allow writes.
-- =========================================================

alter table employees enable row level security;
alter table leave_requests enable row level security;
alter table paid_leave_grants enable row level security;
alter table usage_logs enable row level security;

drop policy if exists employees_read_for_gas_internal_test on employees;
create policy employees_read_for_gas_internal_test
on employees
for select
to anon
using (deleted_at is null);

drop policy if exists leave_requests_read_for_gas_internal_test on leave_requests;
create policy leave_requests_read_for_gas_internal_test
on leave_requests
for select
to anon
using (true);

drop policy if exists paid_leave_grants_read_for_gas_internal_test on paid_leave_grants;
create policy paid_leave_grants_read_for_gas_internal_test
on paid_leave_grants
for select
to anon
using (true);

drop policy if exists usage_logs_read_for_gas_internal_test on usage_logs;
create policy usage_logs_read_for_gas_internal_test
on usage_logs
for select
to anon
using (true);
