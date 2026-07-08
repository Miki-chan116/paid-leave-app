# Paid Leave App Database

This directory contains the initial Supabase schema for the paid leave management app.

## Files

Run the files in this order:

1. `001_initial_schema.sql`
2. `002_indexes.sql`
3. `003_triggers.sql`

## Initial Setup

1. Open the Supabase project.
2. Go to SQL Editor.
3. Paste and run `001_initial_schema.sql`.
4. Paste and run `002_indexes.sql`.
5. Paste and run `003_triggers.sql`.
6. Confirm that the following tables exist:
   - `employees`
   - `leave_requests`
   - `paid_leave_grants`
   - `company_calendar`
   - `usage_logs`
   - `admin_users`

## CSV Import Notes

- Convert empty date and timestamp cells to `NULL` before import.
- Keep existing Spreadsheet IDs as text:
  - `leave_requests.request_id`
  - `paid_leave_grants.grant_id`
  - `admin_users.admin_id`
- Import parent tables before child tables:
  1. `employees`
  2. `admin_users`
  3. `company_calendar`
  4. `leave_requests`
  5. `paid_leave_grants`
  6. `usage_logs`
- Exclude known test records such as `TEST-FIFO-001` from production import.
- Keep `usage_logs.legacy_request_id` exactly as exported from Spreadsheet.
- Populate `usage_logs.target_type`, `target_id`, `leave_request_id`, and `employee_id` only when the target can be classified safely.

## Important Design Notes

- `employees.deleted_at` is for logical deletion. Retirement should remain represented by `employment_status` and `leave_date`.
- `leave_requests.cancel_reason` is for user or admin cancellation reason text.
- `usage_logs` intentionally supports mixed legacy targets because Spreadsheet `usage_log.request_id` contains leave request IDs, employee IDs, and old test or legacy values.
- `grant_type` includes `initial` because existing production data may contain it.
- `updated_at` is maintained by triggers in `003_triggers.sql`.

## Update Rules

- Do not edit applied SQL files after they have been run in production.
- Add future schema changes as new numbered files, for example `004_add_xxx.sql`.
- Keep migrations reversible where practical, but do not use destructive SQL on production without a backup.
- Run schema changes in a staging Supabase project before production.
- After every import or migration, reconcile:
  - row counts
  - primary key counts
  - orphan references
  - leave request status distribution
  - paid leave balance/FIFO results for representative employees

## Safety

These SQL files create schema only. They do not import data and do not delete existing Spreadsheet data.
