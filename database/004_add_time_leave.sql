-- =========================================================
-- Paid Leave App / Phase 5 time-based leave schema
-- 004_add_time_leave.sql
--
-- This migration is intentionally additive. It does not update existing
-- daily/half-day leave rows or legacy fractional carry_over_days values.
-- Run only after 001_initial_schema.sql through 003_triggers.sql.
-- =========================================================

begin;

alter table leave_requests
  add column if not exists request_kind text,
  add column if not exists company_code_snapshot text,
  add column if not exists policy_version text;

-- Text plus a narrow CHECK keeps the field extensible without modifying the
-- already-applied leave_request_type enum. NULL is required for legacy rows.
alter table leave_requests
  drop constraint if exists leave_requests_request_kind_check;
alter table leave_requests
  add constraint leave_requests_request_kind_check check (
    request_kind is null
    or request_kind in ('time_hourly')
  );

alter table paid_leave_grants
  add column if not exists carry_over_minutes integer;

alter table paid_leave_grants
  drop constraint if exists paid_leave_grants_carry_over_minutes_check;
alter table paid_leave_grants
  add constraint paid_leave_grants_carry_over_minutes_check check (
    carry_over_minutes is null
    or (carry_over_minutes >= 0 and carry_over_minutes < 420)
  );

create table if not exists time_leave_segments (
  time_leave_id text primary key,
  request_id text not null references leave_requests(request_id) on delete no action,
  employee_id text not null references employees(employee_id) on delete no action,
  company_code text not null,
  leave_date date not null,
  -- Wall-clock values deliberately use text, never timestamps.
  start_time text not null check (start_time ~ '^([01][0-9]|2[0-3]):[0-5][0-9]$'),
  end_time text not null check (end_time ~ '^([01][0-9]|2[0-3]):[0-5][0-9]$'),
  start_minute integer not null check (start_minute >= 0 and start_minute < 1440),
  end_minute integer not null check (end_minute > start_minute and end_minute <= 1440),
  requested_minutes integer not null check (requested_minutes > 0),
  scheduled_minutes_per_day integer not null check (scheduled_minutes_per_day > 0),
  time_leave_unit_minutes integer not null check (time_leave_unit_minutes > 0),
  work_start_minute integer not null check (work_start_minute >= 0 and work_start_minute < 1440),
  work_end_minute integer not null check (work_end_minute > work_start_minute and work_end_minute <= 1440),
  break_periods_json jsonb not null default '[]'::jsonb,
  calculation_version text not null,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  constraint time_leave_segments_requested_unit_check check (
    requested_minutes % time_leave_unit_minutes = 0
  )
);

create index if not exists idx_time_leave_segments_request_id
  on time_leave_segments(request_id);

create index if not exists idx_time_leave_segments_employee_leave_date
  on time_leave_segments(employee_id, leave_date);

create index if not exists idx_time_leave_segments_company_leave_date
  on time_leave_segments(company_code, leave_date);

drop trigger if exists trg_time_leave_segments_updated_at on time_leave_segments;
create trigger trg_time_leave_segments_updated_at
before update on time_leave_segments
for each row
execute function set_updated_at();

commit;
