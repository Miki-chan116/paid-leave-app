-- =========================================================
-- Paid Leave App / Supabase Triggers
-- 003_triggers.sql
--
-- Automatically refresh updated_at on row updates.
-- =========================================================

begin;

create or replace function set_updated_at()
returns trigger
language plpgsql
as $$
begin
  new.updated_at = now();
  return new;
end;
$$;

drop trigger if exists trg_employees_updated_at on employees;
create trigger trg_employees_updated_at
before update on employees
for each row
execute function set_updated_at();

drop trigger if exists trg_leave_requests_updated_at on leave_requests;
create trigger trg_leave_requests_updated_at
before update on leave_requests
for each row
execute function set_updated_at();

drop trigger if exists trg_paid_leave_grants_updated_at on paid_leave_grants;
create trigger trg_paid_leave_grants_updated_at
before update on paid_leave_grants
for each row
execute function set_updated_at();

drop trigger if exists trg_company_calendar_updated_at on company_calendar;
create trigger trg_company_calendar_updated_at
before update on company_calendar
for each row
execute function set_updated_at();

drop trigger if exists trg_usage_logs_updated_at on usage_logs;
create trigger trg_usage_logs_updated_at
before update on usage_logs
for each row
execute function set_updated_at();

drop trigger if exists trg_admin_users_updated_at on admin_users;
create trigger trg_admin_users_updated_at
before update on admin_users
for each row
execute function set_updated_at();

commit;
