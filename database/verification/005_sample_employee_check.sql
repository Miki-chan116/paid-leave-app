-- =========================================================
-- Paid Leave App / Supabase Verification
-- 005_sample_employee_check.sql
--
-- Purpose:
--   Check histories for representative employees against Spreadsheet/GAS.
--   Replace the EMP0001-EMP0010 values with actual representative IDs.
-- =========================================================

with target_employees(employee_id) as (
  values
    ('EMP0001'),
    ('EMP0002'),
    ('EMP0003'),
    ('EMP0004'),
    ('EMP0005'),
    ('EMP0006'),
    ('EMP0007'),
    ('EMP0008'),
    ('EMP0009'),
    ('EMP0010')
)
select
  te.employee_id,
  e.name,
  e.company_code,
  e.employment_status,
  e.hire_date,
  e.leave_date,
  coalesce(usage_summary.approved_used_days, 0) as approved_used_days,
  coalesce(usage_summary.approved_request_count, 0) as approved_request_count,
  coalesce(grant_summary.total_grant_days, 0) as total_grant_days,
  coalesce(grant_summary.total_carry_over_days, 0) as total_carry_over_days,
  coalesce(grant_summary.grant_count, 0) as grant_count
from target_employees te
left join employees e
  on e.employee_id = te.employee_id
left join (
  select
    employee_id,
    sum(days) as approved_used_days,
    count(*) as approved_request_count
  from leave_requests
  where status = 'approved'
  group by employee_id
) usage_summary
  on usage_summary.employee_id = te.employee_id
left join (
  select
    employee_id,
    sum(grant_days) as total_grant_days,
    sum(carry_over_days) as total_carry_over_days,
    count(*) as grant_count
  from paid_leave_grants
  group by employee_id
) grant_summary
  on grant_summary.employee_id = te.employee_id
order by te.employee_id;

with target_employees(employee_id) as (
  values
    ('EMP0001'),
    ('EMP0002'),
    ('EMP0003'),
    ('EMP0004'),
    ('EMP0005'),
    ('EMP0006'),
    ('EMP0007'),
    ('EMP0008'),
    ('EMP0009'),
    ('EMP0010')
)
select
  lr.employee_id,
  e.name,
  lr.request_id,
  lr.status,
  lr.request_date,
  lr.start_date,
  lr.end_date,
  lr.days,
  lr.half_day,
  lr.reason,
  lr.approver_name,
  lr.approved_at,
  lr.rejected_reason,
  lr.cancel_reason,
  lr.created_at,
  lr.updated_at
from leave_requests lr
join target_employees te
  on te.employee_id = lr.employee_id
left join employees e
  on e.employee_id = lr.employee_id
order by lr.employee_id, lr.start_date, lr.request_date, lr.request_id;

with target_employees(employee_id) as (
  values
    ('EMP0001'),
    ('EMP0002'),
    ('EMP0003'),
    ('EMP0004'),
    ('EMP0005'),
    ('EMP0006'),
    ('EMP0007'),
    ('EMP0008'),
    ('EMP0009'),
    ('EMP0010')
)
select
  plg.employee_id,
  e.name,
  plg.grant_id,
  plg.grant_type,
  plg.grant_date,
  plg.grant_days,
  plg.carry_over_days,
  plg.valid_from,
  plg.valid_to,
  plg.year,
  plg.is_finalized,
  plg.finalized_at,
  plg.notes,
  plg.created_at,
  plg.updated_at
from paid_leave_grants plg
join target_employees te
  on te.employee_id = plg.employee_id
left join employees e
  on e.employee_id = plg.employee_id
order by plg.employee_id, plg.grant_date, plg.grant_id;
