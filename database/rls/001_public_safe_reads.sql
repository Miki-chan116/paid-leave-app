-- =========================================================
-- Paid Leave App / Supabase RLS Draft
-- 001_public_safe_reads.sql
--
-- Purpose:
--   Lower-risk anon read policy candidates for GAS read verification.
--
-- Important:
--   - Review before running in Supabase.
--   - This file intentionally excludes admin_users.
--   - This file only creates SELECT policies. It does not allow writes.
-- =========================================================

alter table company_calendar enable row level security;

drop policy if exists company_calendar_read_for_gas on company_calendar;
create policy company_calendar_read_for_gas
on company_calendar
for select
to anon
using (true);
