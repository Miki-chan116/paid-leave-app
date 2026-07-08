-- =========================================================
-- Paid Leave App / Supabase RLS Draft
-- 003_internal_test_writes.sql
--
-- Purpose:
--   Internal test-only INSERT policy candidate for GAS dual write verification.
--
-- Important:
--   - Review carefully before running in Supabase.
--   - This allows anon role to INSERT leave_requests.
--   - Use only for limited internal dual write verification.
--   - Do not add UPDATE / DELETE policies for this phase.
--   - admin_users is intentionally excluded and must not be exposed to anon.
-- =========================================================

alter table leave_requests enable row level security;

drop policy if exists leave_requests_insert_for_gas_internal_test on leave_requests;
create policy leave_requests_insert_for_gas_internal_test
on leave_requests
for insert
to anon
with check (
  request_id is not null
  and employee_id is not null
  and status = 'pending'
);
