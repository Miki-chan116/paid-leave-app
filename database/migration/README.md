# Spreadsheet CSV to Supabase CSV Migration

This tool converts exported Spreadsheet CSV files into CSV files matching the Supabase schema in `database/001_initial_schema.sql`.

It does not connect to Supabase, does not import data, and does not modify the original CSV files.

## Input

Put the exported Spreadsheet CSV files in:

```text
database/migration/input/
```

CSV files in this directory are ignored by Git.

Required input filenames:

```text
employees.csv
leave_requests.csv
paid_leave_grants.csv
company_calendar.csv
usage_log.csv
admin_users.csv
```

## Output

Converted files are written to:

```text
database/migration/output/
```

Generated CSV and report files in this directory are ignored by Git.

Output CSV files:

```text
employees.csv
leave_requests.csv
paid_leave_grants.csv
company_calendar.csv
usage_logs.csv
admin_users.csv
```

Report files:

```text
migration_report.json
migration_errors.json
migration_warnings.json
```

If conversion errors are found, the tool writes the report files and exits without writing the converted CSV files.

## Usage

From the repository root:

```bash
node database/migration/convert_spreadsheet_csv_to_supabase.js
```

With custom directories:

```bash
node database/migration/convert_spreadsheet_csv_to_supabase.js \
  --input /path/to/source_csv \
  --output /path/to/output_csv
```

## Conversion Rules

- Source CSV files are read only.
- Known test IDs such as `TEST-FIFO-001`, `TEST-001`, and IDs starting with `TEST-` are excluded.
- Empty strings are emitted as empty CSV fields, to be imported as `NULL` where appropriate.
- Dates are normalized to `YYYY-MM-DD`.
- Timestamps are normalized to ISO 8601.
- Boolean values are normalized to `true` or `false`.
- `grant_type = initial` is preserved.
- `usage_log.csv` is converted to `usage_logs.csv`.
- `usage_log.request_id` is copied to `usage_logs.legacy_request_id`.
- If `usage_log.request_id` matches `leave_requests.request_id`, `target_type` becomes `leave_request`.
- If `usage_log.request_id` matches `employees.employee_id`, `target_type` becomes `employee`.
- Otherwise, `target_type` remains `legacy`.

## Validation

The converter checks:

- Required input files exist.
- Primary IDs are present and unique.
- `leave_requests.employee_id` exists in `employees`.
- `paid_leave_grants.employee_id` exists in `employees`.
- Enums match the Supabase schema.
- Date/timestamp values can be parsed.
- Number values can be parsed and are not negative.
- Boolean values can be normalized.

## Import Order

After reviewing the generated CSV and reports, import into Supabase in this order:

1. `employees.csv`
2. `admin_users.csv`
3. `company_calendar.csv`
4. `leave_requests.csv`
5. `paid_leave_grants.csv`
6. `usage_logs.csv`

## Notes

- Do not import if `migration_errors.json` is not empty.
- Review `migration_warnings.json`, especially unclassified legacy usage log targets.
- Confirm row counts before and after import.
- Confirm representative employee balances and FIFO results after import.
