# Supabase RLS Drafts

GASから `SUPABASE_ANON_KEY` で読み取り検証するためのRLS SQL案です。
このディレクトリのSQLはレビュー用で、Codex作業ではSupabaseに実行していません。

## ファイル構成

1. `001_public_safe_reads.sql`
   - `company_calendar` の読み取り候補です。
   - 比較的安全ですが、実行前に公開してよいカレンダー情報か確認してください。
2. `002_internal_test_reads.sql`
   - `employees`, `leave_requests`, `paid_leave_grants`, `usage_logs` の読み取り候補です。
   - 個人情報・申請理由・ログコメントを含むため、内部検証限定です。
   - 本番運用用としてそのまま使わないでください。

## 方針

- 読み取り専用の `select` policy だけを用意します。
- `insert` / `update` / `delete` policy はまだ作りません。
- `SERVICE_ROLE_KEY` はGAS側では使いません。
- `employees` は `deleted_at is null` の行だけ読める想定です。
- `admin_users` はanon公開禁止です。このディレクトリのRLS SQLには含めません。
- `USE_SUPABASE_READS=true` は、読み取りSupabase / 書き込みSpreadsheetの混在検証用です。
  本番運用ONは、申請登録・承認・否認・取消などの書き込み移行後に再判断してください。

## 注意点

`admin_users` には現在 `pin` が含まれます。
anon roleに `admin_users` のselectを許可すると、設計上は公開可能なanon keyでPIN列も読める状態になります。
そのため、`admin_users` はanon公開禁止です。管理者ログインは当面Spreadsheet読み取りのままにします。
production運用前に以下のどちらかを検討してください。

- `admin_users.pin` をハッシュ化し、照合方式を変更する
- 管理者認証だけはCloud RunなどのサーバーAPI経由にする
- Supabaseの権限設計を見直し、anon roleで `admin_users` を読まない

## 今回必要な読み取り対象

- 比較的安全: `company_calendar`
- 内部検証限定: `employees`, `leave_requests`, `paid_leave_grants`, `usage_logs`
- anon公開禁止: `admin_users`

## ロールバック

RLS policyを戻す場合は、対象policyを `drop policy if exists ... on ...;` で削除してください。
RLS自体を無効化するかどうかは、Supabase上の他用途への影響を確認してから判断します。
