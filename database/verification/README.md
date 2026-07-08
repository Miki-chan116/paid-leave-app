# Supabase Data Verification SQL

Spreadsheet/GAS から Supabase へインポートした有給管理データを照合するためのSQLです。
このディレクトリのSQLは読み取り専用です。テーブル作成、更新、削除、importは行いません。

## 実行順

1. `001_counts.sql`
2. `002_distributions.sql`
3. `003_orphan_checks.sql`
4. `004_employee_summaries.sql`
5. `005_sample_employee_check.sql`

## SQL Editorでの実行方法

1. Supabase Dashboardで対象プロジェクトを開きます。
2. SQL Editorを開きます。
3. 上記の順番でSQLファイルの内容を貼り付けて実行します。
4. 実行結果をSpreadsheet側の監査ログまたは移行前CSV件数と比較します。

## 見るべきポイント

### 001_counts.sql

全テーブルの件数を確認します。

- `employees`
- `leave_requests`
- `paid_leave_grants`
- `company_calendar`
- `usage_logs`
- `admin_users`

移行前CSVの変換レポートと一致しているか確認してください。
テストデータ除外を行ったテーブルは、除外後件数と比較します。

### 002_distributions.sql

主要な区分値の分布を確認します。

- `leave_requests.status`
- `employees.company_code`
- `paid_leave_grants.grant_type`

`status` は `pending`, `approved`, `rejected`, `canceled`, `canceled_by_admin` の件数がGAS側と一致するか確認します。
`grant_type` は `initial` と `yearly` を含め、Spreadsheet側の分布と一致するか確認します。

### 003_orphan_checks.sql

`employees` に存在しない `employee_id` 参照がないか確認します。

- `leave_requests.employee_id`
- `paid_leave_grants.employee_id`
- `usage_logs.employee_id`

正常なら各SQLの結果は0件です。
Supabase schemaのForeign Keyが有効なため、通常はimport時点で不整合が止まります。

### 004_employee_summaries.sql

社員別の集計値を確認します。

- 承認済み有給取得日数
- 承認済み申請件数
- 付与日数合計
- 繰越日数合計
- 付与件数

有給取得日数は `leave_requests.status = 'approved'` のみ集計します。
`canceled`, `canceled_by_admin`, `rejected`, `pending` は除外します。

### 005_sample_employee_check.sql

代表社員10人のサマリ、申請履歴、付与履歴を確認します。

初期値は `EMP0001` から `EMP0010` です。
実データの代表社員IDに差し替えてから実行してください。
差し替える場合も `EMP0001` 形式の桁数に合わせる前提で確認します。

## 注意事項

- このSQLはSupabase側のデータ確認専用です。
- SpreadsheetやGASコードは変更しません。
- Supabaseへのimportや更新は行いません。
- 代表社員IDが存在しない場合、サマリSQLでは社員情報がNULLになります。
- `usage_logs.legacy_request_id` は過去データ保持用のため、申請ID以外の値が入っていても即エラーとは扱いません。
