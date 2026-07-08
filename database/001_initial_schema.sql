-- =========================================================
-- Paid Leave App / Supabase Initial Schema
-- 001_initial_schema.sql
--
-- Run order:
--   1. 001_initial_schema.sql
--   2. 002_indexes.sql
--   3. 003_triggers.sql
--
-- Notes:
--   - Existing Spreadsheet IDs are preserved as text primary keys.
--   - CSV import should convert empty date/timestamp values to NULL.
--   - usage_logs keeps legacy_request_id to preserve mixed historical values.
-- =========================================================

begin;

do $$
begin
  create type leave_request_status as enum (
    'pending',
    'approved',
    'rejected',
    'canceled',
    'canceled_by_admin'
  );
exception
  when duplicate_object then null;
end $$;

do $$
begin
  create type leave_request_type as enum (
    'paid_leave'
  );
exception
  when duplicate_object then null;
end $$;

do $$
begin
  create type leave_half_day_type as enum (
    'am',
    'pm'
  );
exception
  when duplicate_object then null;
end $$;

do $$
begin
  create type leave_grant_type as enum (
    'initial',
    'six_month',
    'six_month_processed',
    'six_month_skipped',
    'yearly'
  );
exception
  when duplicate_object then null;
end $$;

do $$
begin
  create type calendar_day_type as enum (
    'workday',
    'holiday',
    'no_leave'
  );
exception
  when duplicate_object then null;
end $$;

do $$
begin
  create type usage_log_target_type as enum (
    'leave_request',
    'employee',
    'legacy',
    'unknown'
  );
exception
  when duplicate_object then null;
end $$;

create table if not exists employees (
  employee_id text primary key,
  display_employee_id text unique,
  name text not null,
  display_name text,
  name_kana text,
  company_code text not null,
  company_name text,
  department text,
  employment_type text,
  employment_status text not null default 'active',
  hire_date date,
  leave_date date,
  work_days_per_week numeric(3,1) check (
    work_days_per_week is null
    or work_days_per_week between 0 and 7
  ),
  fiscal_start_month integer not null default 4 check (
    fiscal_start_month between 1 and 12
  ),
  leave_management_target boolean not null default false,
  initial_grant_check_target boolean not null default true,
  is_driver boolean not null default false,
  driver_type text,
  default_vehicle_id text,
  display_order integer check (
    display_order is null
    or display_order >= 0
  ),
  notes text,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  deleted_at timestamptz,
  constraint employees_leave_date_after_hire_date check (
    leave_date is null
    or hire_date is null
    or leave_date >= hire_date
  )
);

create table if not exists leave_requests (
  request_id text primary key,
  employee_id text not null references employees(employee_id),
  request_date timestamptz,
  start_date date not null,
  end_date date not null,
  days numeric(5,2) not null check (
    days >= 0
    and days <= 365
  ),
  type leave_request_type not null default 'paid_leave',
  half_day leave_half_day_type,
  reason text,
  reason_detail text,
  status leave_request_status not null default 'pending',
  approver_id text,
  approver_name text,
  approved_at timestamptz,
  rejected_reason text,
  cancel_reason text,
  year integer check (
    year is null
    or year between 2000 and 2100
  ),
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  constraint leave_requests_end_date_after_start_date check (
    end_date >= start_date
  ),
  constraint leave_requests_half_day_single_day check (
    half_day is null
    or start_date = end_date
  )
);

create table if not exists paid_leave_grants (
  grant_id text primary key,
  employee_id text not null references employees(employee_id),
  grant_date date not null,
  grant_days numeric(5,2) not null default 0 check (
    grant_days >= 0
    and grant_days <= 100
  ),
  carry_over_days numeric(5,2) not null default 0 check (
    carry_over_days >= 0
    and carry_over_days <= 200
  ),
  valid_from date,
  valid_to date,
  grant_type leave_grant_type not null,
  year integer check (
    year is null
    or year between 2000 and 2100
  ),
  notes text,
  is_finalized boolean not null default true,
  finalized_at timestamptz,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  constraint paid_leave_grants_valid_to_after_valid_from check (
    valid_to is null
    or valid_from is null
    or valid_to >= valid_from
  )
);

create table if not exists company_calendar (
  date date primary key,
  type calendar_day_type not null,
  notes text,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists usage_logs (
  log_id text primary key,
  legacy_request_id text,
  target_type usage_log_target_type not null default 'legacy',
  target_id text,
  leave_request_id text references leave_requests(request_id),
  employee_id text references employees(employee_id),
  action_type text not null,
  operator_id text,
  operator_name text,
  action_date timestamptz not null default now(),
  comment text,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  constraint usage_logs_target_consistency check (
    (target_type = 'leave_request' and leave_request_id is not null)
    or (target_type = 'employee' and employee_id is not null)
    or (target_type in ('legacy', 'unknown'))
  )
);

create table if not exists admin_users (
  admin_id text primary key,
  admin_name text not null,
  pin text,
  role text,
  notes text,
  is_active boolean not null default true,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

commit;
