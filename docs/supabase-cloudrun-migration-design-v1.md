# Supabase + Cloud Run移行設計書 v1.0

作成日: 2026-07-08  
対象システム: 有給管理システム  
対象範囲: GAS + Spreadsheet から Supabase + Cloud Run への段階移行設計  
ステータス: 設計レビュー版

---

## 0. 前提と設計方針

### 0.1 現在の実装状況

現在、有給管理システムは以下の移行段階にあります。

- 既存本番DB: Spreadsheet
- 新DB: Supabase
- 実行基盤: GAS
- Cloud Run: 未導入
- Supabase接続: GASからREST API経由
- Supabase認証: anon key + RLS
- `SERVICE_ROLE_KEY`: 未使用
- Feature Flag:
  - `USE_SUPABASE_READS`
  - `USE_SUPABASE_WRITE`
- 実装済み:
  - Supabaseスキーマ
  - CSV変換
  - Supabase import
  - GASからSupabase Read
  - Feature FlagによるDual Read
  - 有給申請登録のみDual Write検証
  - Spreadsheet保存成功後のみSupabase INSERT
  - Supabase失敗時はLogger出力のみ
  - Spreadsheetはロールバックしない

### 0.2 基本方針

本設計では、以下を基本方針とします。

1. 短期はSpreadsheetを正DBとして維持する
2. Supabaseはまず副DB・照合先として扱う
3. Dual Writeは移行ブリッジであり、恒久設計にしない
4. Cloud Run導入後、書き込み責務をCloud Runへ移す
5. 最終的にはSupabaseを正DBとし、Spreadsheetは参照・出力・バックアップ用途へ降格する
6. `admin_users` をanon keyで公開しない
7. `SERVICE_ROLE_KEY` はGASに置かない
8. 障害時に業務を止めないが、差分回復できる仕組みを必ず持つ

---

## 1. システム全体構成

### 1.1 現状構成

```text
┌──────────────┐
│ User / Admin │
└──────┬───────┘
       │
       ▼
┌──────────────┐
│     GAS      │
│ Web UI / API │
└──────┬───────┘
       │ read/write
       ▼
┌──────────────┐
│ Spreadsheet  │
│ 本番DB        │
└──────────────┘

補助的に:
┌──────────────┐
│  Supabase    │
│ 新DB / 移行先 │
└──────────────┘
```

#### 特徴

- Spreadsheetが唯一の正DB
- GASがUI、業務処理、DBアクセスをすべて担当
- Supabaseはまだ本番処理の正DBではない
- 既存運用は安定しているが、DBとしての整合性・拡張性・監査性に限界がある

### 1.2 移行中構成

```text
┌──────────────┐
│ User / Admin │
└──────┬───────┘
       │
       ▼
┌────────────────────────────┐
│            GAS             │
│                            │
│ - 既存画面                 │
│ - Spreadsheet read/write   │
│ - Supabase read検証        │
│ - 一部Dual Write検証       │
└──────┬───────────────┬─────┘
       │               │
       │ 正DB           │ 副DB / 照合先
       ▼               ▼
┌──────────────┐   ┌──────────────┐
│ Spreadsheet  │   │  Supabase    │
│ 本番DB        │   │ 新DB          │
└──────────────┘   └──────────────┘
```

#### 現在のFeature Flag

```text
USE_SUPABASE_READS=true
  → 一部読み取りをSupabaseへ切替

USE_SUPABASE_WRITE=true
  → 有給申請登録のみSpreadsheet保存後にSupabaseへINSERT
```

#### 注意点

この段階では、読み取りと書き込みの正が混在します。

```text
書き込み: Spreadsheet
読み取り: Spreadsheet または Supabase
```

そのため、`USE_SUPABASE_READS=true` を本番常時ONにするのはまだ早いです。

### 1.3 Cloud Run導入後の構成

```text
┌──────────────┐
│ User / Admin │
└──────┬───────┘
       │
       ▼
┌──────────────┐
│     GAS      │
│ UI / thin API│
└──────┬───────┘
       │ HTTPS
       ▼
┌────────────────────────┐
│       Cloud Run API     │
│                        │
│ - 認証                 │
│ - 入力検証             │
│ - 業務ルール           │
│ - 冪等性制御           │
│ - retry / sync管理     │
│ - audit log            │
└──────┬─────────────────┘
       │
       ▼
┌──────────────┐
│  Supabase    │
│ 新・正DB候補 │
└──────┬───────┘
       │ optional backup / export
       ▼
┌──────────────┐
│ Spreadsheet  │
│ 照合/出力/予備│
└──────────────┘
```

### 1.4 完成構成

```text
┌──────────────────────┐
│ User / Admin / Staff │
└──────────┬───────────┘
           │
           ▼
┌──────────────────────┐
│ GAS or Web Frontend  │
│ UI layer             │
└──────────┬───────────┘
           │ HTTPS
           ▼
┌──────────────────────┐
│ Cloud Run API        │
│ Business API         │
│                      │
│ - auth               │
│ - validation         │
│ - status transition  │
│ - idempotency        │
│ - audit log          │
│ - retry control      │
└──────────┬───────────┘
           │
           ▼
┌──────────────────────┐
│ Supabase PostgreSQL  │
│ Primary Database     │
└──────────┬───────────┘
           │
           ▼
┌──────────────────────┐
│ Supabase Realtime    │
│ notifications / UI   │
└──────────────────────┘

補助:
┌──────────────────────┐
│ Spreadsheet Archive  │
│ CSV / backup only    │
└──────────────────────┘
```

---

## 2. コンポーネント責務

### 2.1 Spreadsheet

#### 現在の責務

- 本番DB
- 有給申請データ保存
- 社員マスタ保存
- 有給付与データ保存
- 会社カレンダー保存
- 使用ログ保存
- 管理者情報保存
- CSV出力元
- 手動確認・監査用データ

#### 移行中の責務

- 正DB
- Supabase照合元
- Dual Write失敗時の復旧元
- CSV出力の安全な基準
- 管理者ログイン情報の保持

#### 完成後の責務

- 原則として本番DBから降格
- 必要に応じて以下のみ担当
  - CSVエクスポート先
  - バックアップ
  - 監査用スナップショット
  - 運用担当者向け閲覧用コピー

#### 責任範囲

| 項目 | 現在 | 完成後 |
|---|---|---|
| 正DB | はい | いいえ |
| 書き込み | はい | 原則停止 |
| CSV出力 | はい | 任意 |
| 監査 | 一部 | 補助 |
| 手動修正 | 可能 | 原則禁止 |

### 2.2 GAS

#### 現在の責務

- 画面表示
- フロントエンド処理
- Spreadsheet読み書き
- Supabase Read検証
- Feature Flag制御
- 有給申請登録のDual Write検証
- 管理画面処理
- CSV出力

#### 移行中の責務

- UI維持
- 既存運用維持
- Cloud Run導入までの中継役
- Feature Flagによる段階切替
- Spreadsheet正DB運用の継続

#### 完成後の責務

- 可能ならUI層へ縮小
- DBアクセスはCloud Run API経由にする
- 業務ロジックはCloud Runへ移管
- GASは以下に限定
  - 既存Google Workspace連携
  - 管理者向け軽量UI
  - CSV作成補助
  - 一部バックオフィス自動化

### 2.3 Cloud Run

#### 導入目的

Cloud Runは、GASとSupabaseの間に置く業務API層です。

Google Cloud RunではサービスごとにサービスIDやSecret Manager連携を使えるため、秘密情報をGASやコードに直書きせず、実行環境側で安全に扱えます。Cloud RunはSecret ManagerのSecretを環境変数やボリュームとして利用できる構成をサポートしています。  
参考: [Configure secrets for services | Cloud Run](https://cloud.google.com/run/docs/configuring/services/secrets)

また、Cloud RunサービスにはサービスIDを割り当て、Google Cloudリソースへのアクセス権限を制御できます。  
参考: [Introduction to service identity | Cloud Run](https://cloud.google.com/run/docs/securing/service-identity)

#### 責務

- Supabaseへの書き込み
- Supabase service role keyの安全な保持
- 業務ルールの検証
- 入力値検証
- 状態遷移制御
- 冪等性制御
- retry制御
- audit log作成
- API認証
- 将来のRealtime連携

#### 責任範囲

| 項目 | Cloud Runの責務 |
|---|---|
| DB書き込み | はい |
| 業務バリデーション | はい |
| 認証 | はい |
| 冪等性 | はい |
| 監査ログ | はい |
| Spreadsheet直接操作 | 原則いいえ |
| UI表示 | いいえ |

### 2.4 Supabase

#### 現在の責務

- 新DB
- 移行先データ保持
- Read検証
- 一部Dual Write先

#### 完成後の責務

- 正DB
- 有給管理データの永続化
- 検索・集計
- 状態履歴
- 監査ログ
- Realtime配信元
- バックアップ元

#### Supabaseで保持する主要テーブル

- `employees`
- `leave_requests`
- `paid_leave_grants`
- `company_calendar`
- `usage_logs`
- `admin_users`
- 将来:
  - `sync_queue`
  - `audit_logs`
  - `idempotency_keys`
  - `notification_events`

#### RLSの考え方

SupabaseのRLSはPostgreSQLのRow Level Securityを使い、テーブル単位で行アクセスを制御します。Supabase公式ドキュメントでも、RLSは有効化した上でpolicyを作成してアクセス可否を定義する設計になっています。  
参考: [Row Level Security | Supabase Docs](https://supabase.com/docs/guides/database/postgres/row-level-security)

本システムでは、完成形では以下とします。

- フロントから直接Supabaseに書かない
- 書き込みはCloud Run経由
- service role keyはCloud Runだけが保持
- anon keyは原則読み取り限定、または使わない
- `admin_users` はanon公開禁止

### 2.5 Realtime

#### 将来責務

Supabase RealtimeはPostgresの変更をもとに、クライアントへリアルタイム通知できます。Supabase公式では、RealtimeはBroadcast、Presence、Postgres Changesなどの機能を提供します。  
参考: [Realtime | Supabase Docs](https://supabase.com/docs/guides/realtime)

本システムでは以下に利用できます。

- 承認待ち申請のリアルタイム更新
- 申請ステータス変更通知
- 管理画面ダッシュボード更新
- 社員への承認/否認通知
- 締め処理進捗通知
- 監査ログ監視

---

## 3. データフロー

### 3.1 有給申請

#### 現在

```text
User
 ↓
GAS submitLeaveRequest()
 ↓
Spreadsheet leave_requests INSERT
 ↓
Spreadsheet usage_log INSERT
 ↓
if USE_SUPABASE_WRITE=true
    Supabase leave_requests INSERT
 ↓
Userへ成功応答
```

#### 特徴

- Spreadsheet保存成功が本処理の成功条件
- Supabase失敗はLoggerのみ
- Supabase失敗してもロールバックしない
- `request_id` はSpreadsheetとSupabaseで共通

#### 完成形

```text
User
 ↓
GAS / Web UI
 ↓
Cloud Run POST /leave-requests
 ↓
Cloud Run validation
 ↓
Supabase transaction
 ├─ leave_requests INSERT
 ├─ usage_logs INSERT
 └─ idempotency_keys UPSERT
 ↓
Response
 ↓
Realtime通知
```

#### 完成形の重要点

- Supabaseを正DBにする
- `request_id` を冪等キーにする
- `usage_logs` も同一transactionで保存
- 409 conflictは同一payloadなら成功扱い

### 3.2 承認

#### 現在

```text
Admin
 ↓
GAS approveRequest()
 ↓
Spreadsheet leave_requests UPDATE
 ↓
Spreadsheet usage_log INSERT
 ↓
clearAppCache()
```

#### 完成形

```text
Admin
 ↓
Cloud Run PATCH /leave-requests/{id}/approve
 ↓
状態検証:
  current_status == pending
 ↓
Supabase transaction
 ├─ leave_requests.status = approved
 ├─ approver_id
 ├─ approver_name
 ├─ approved_at
 ├─ updated_at
 └─ usage_logs INSERT
 ↓
Realtime通知
```

#### 状態遷移

```text
pending → approved
```

以下は不可です。

```text
approved → approved
canceled → approved
rejected → approved
canceled_by_admin → approved
```

### 3.3 否認

#### 現在

```text
Admin
 ↓
GAS rejectRequest()
 ↓
Spreadsheet leave_requests.status = rejected
 ↓
Spreadsheet rejected_reason 更新
 ↓
usage_log INSERT
```

#### 完成形

```text
Admin
 ↓
Cloud Run PATCH /leave-requests/{id}/reject
 ↓
状態検証:
  current_status == pending
 ↓
Supabase transaction
 ├─ leave_requests.status = rejected
 ├─ rejected_reason
 ├─ updated_at
 └─ usage_logs INSERT
```

### 3.4 取消

#### 本人取消 現在

```text
User
 ↓
GAS cancelPendingLeaveRequestForEmployee()
 ↓
Spreadsheet status = canceled
 ↓
usage_log INSERT
```

#### 管理者取消 現在

```text
Admin
 ↓
GAS cancelApprovedRequestByAdmin()
 ↓
Spreadsheet status = canceled_by_admin
 ↓
cancel_reason 更新
 ↓
usage_log INSERT
```

#### 完成形

```text
Cloud Run PATCH /leave-requests/{id}/cancel
Cloud Run PATCH /leave-requests/{id}/admin-cancel
```

#### 状態遷移

```text
pending  → canceled
approved → canceled_by_admin
```

#### 注意

承認済み取消は残日数計算に影響します。  
`canceled_by_admin` は使用日数から除外する必要があります。

### 3.5 有給付与

#### 現在

```text
Admin / 年度処理
 ↓
GAS
 ↓
Spreadsheet paid_leave_grants INSERT
 ↓
usage_log INSERT
```

#### 完成形

```text
Cloud Run POST /grants
 ↓
validation
 ↓
Supabase transaction
 ├─ paid_leave_grants INSERT
 └─ usage_logs INSERT
```

#### 必須検証

- `employee_id` が存在する
- `grant_type` がenum内
- `grant_days >= 0`
- `carry_over_days >= 0`
- `valid_to >= valid_from`
- 同一年・同種別の重複付与制御

### 3.6 社員編集

#### 現在

```text
Admin
 ↓
GAS updateEmployeeFromAdmin()
 ↓
Spreadsheet employees UPDATE
 ↓
usage_log INSERT
```

#### 完成形

```text
Cloud Run PATCH /employees/{id}
 ↓
validation
 ↓
Supabase transaction
 ├─ employees UPDATE
 └─ usage_logs INSERT
```

#### 注意

社員編集は有給計算に強く影響します。

特に以下は慎重に扱います。

- `hire_date`
- `leave_date`
- `employment_status`
- `leave_management_target`
- `fiscal_start_month`
- `company_code`

### 3.7 CSV出力

#### 現在

```text
GAS
 ↓
Spreadsheet / Output Spreadsheet
 ↓
CSV相当データ作成
```

#### 完成形

```text
Cloud Run GET /reports/monthly
Cloud Run GET /reports/yearly
 ↓
Supabase query
 ↓
CSV生成
 ↓
download or Spreadsheet export
```

#### 方針

- 完全移行後はSupabase基準
- 必要ならCloud RunがCSVを生成
- GASはダウンロードUIまたはGoogle Drive保存だけ担当

### 3.8 バックアップ

#### 現在

```text
Spreadsheet自体が本番DB兼バックアップ
Supabaseは移行先
```

#### 完成形

```text
Supabase primary
 ↓
scheduled export
 ↓
Cloud Storage / Spreadsheet snapshot / CSV archive
```

#### 推奨

- 日次CSV export
- 月次締めスナップショット
- 退職者・過年度データのアーカイブ
- `usage_logs` は長期保存

---

## 4. API設計

Cloud Runで作るAPI一覧です。

### 4.1 有給申請API

| Method | Path | 役割 |
|---|---|---|
| `POST` | `/leave-requests` | 有給申請登録 |
| `GET` | `/leave-requests/{id}` | 申請詳細取得 |
| `GET` | `/leave-requests` | 申請一覧検索 |
| `PATCH` | `/leave-requests/{id}` | 承認待ち申請の本人修正 |
| `PATCH` | `/leave-requests/{id}/approve` | 承認 |
| `PATCH` | `/leave-requests/{id}/reject` | 否認 |
| `PATCH` | `/leave-requests/{id}/cancel` | 本人取消 |
| `PATCH` | `/leave-requests/{id}/admin-cancel` | 管理者取消 |

#### `POST /leave-requests`

##### 入力

```json
{
  "request_id": "uuid-or-client-generated-id",
  "employee_id": "EMP0001",
  "start_date": "2026-07-10",
  "end_date": "2026-07-10",
  "days": 1,
  "type": "paid_leave",
  "half_day": null,
  "reason": "私用",
  "reason_detail": "",
  "idempotency_key": "same-as-request-id"
}
```

##### 出力

```json
{
  "ok": true,
  "request_id": "uuid-or-client-generated-id",
  "status": "pending"
}
```

##### 役割

- 日付検証
- 会社カレンダー検証
- 残日数の警告
- `leave_requests` INSERT
- `usage_logs` INSERT
- 冪等性制御

### 4.2 承認API

#### `PATCH /leave-requests/{id}/approve`

##### 入力

```json
{
  "admin_id": "admin001",
  "admin_name": "管理者",
  "idempotency_key": "approve-request-id-admin-id"
}
```

##### 出力

```json
{
  "ok": true,
  "request_id": "REQ001",
  "status": "approved"
}
```

##### 役割

- `pending` のみ承認可能
- `approved_at` 記録
- `approver_id` 記録
- audit log記録

### 4.3 否認API

#### `PATCH /leave-requests/{id}/reject`

##### 入力

```json
{
  "admin_id": "admin001",
  "admin_name": "管理者",
  "rejected_reason": "理由不備",
  "idempotency_key": "reject-request-id-admin-id"
}
```

##### 出力

```json
{
  "ok": true,
  "request_id": "REQ001",
  "status": "rejected"
}
```

### 4.4 取消API

| Method | Path | 役割 |
|---|---|---|
| `PATCH` | `/leave-requests/{id}/cancel` | 本人取消 |
| `PATCH` | `/leave-requests/{id}/admin-cancel` | 承認後管理者取消 |

#### `PATCH /leave-requests/{id}/admin-cancel`

##### 入力

```json
{
  "admin_id": "admin001",
  "admin_name": "管理者",
  "cancel_reason": "日付誤りのため",
  "idempotency_key": "admin-cancel-request-id-admin-id"
}
```

##### 出力

```json
{
  "ok": true,
  "request_id": "REQ001",
  "status": "canceled_by_admin"
}
```

### 4.5 社員API

| Method | Path | 役割 |
|---|---|---|
| `GET` | `/employees` | 社員一覧 |
| `GET` | `/employees/{id}` | 社員詳細 |
| `POST` | `/employees` | 社員追加 |
| `PATCH` | `/employees/{id}` | 社員編集 |
| `PATCH` | `/employees/{id}/retire` | 退職処理 |

### 4.6 有給付与API

| Method | Path | 役割 |
|---|---|---|
| `GET` | `/grants` | 付与一覧 |
| `POST` | `/grants` | 付与登録 |
| `GET` | `/employees/{id}/grants` | 社員別付与履歴 |
| `POST` | `/grants/yearly-rollover` | 年度付与処理 |
| `POST` | `/grants/six-month` | 6ヶ月付与処理 |

### 4.7 カレンダーAPI

| Method | Path | 役割 |
|---|---|---|
| `GET` | `/calendar` | 会社カレンダー取得 |
| `PUT` | `/calendar/{date}` | 日別更新 |
| `POST` | `/calendar/generate` | 年度カレンダー生成 |

### 4.8 残日数・FIFO API

| Method | Path | 役割 |
|---|---|---|
| `GET` | `/balance/{employee_id}` | 社員別残日数 |
| `GET` | `/balance` | 社員一覧残日数 |
| `GET` | `/fifo/{employee_id}` | FIFO詳細 |
| `GET` | `/fifo/compare` | Spreadsheet/Supabase比較用 |

### 4.9 ログ・監査API

| Method | Path | 役割 |
|---|---|---|
| `GET` | `/usage-logs` | 操作ログ検索 |
| `GET` | `/audit-logs` | 監査ログ検索 |
| `GET` | `/sync-queue` | 同期状態確認 |
| `POST` | `/sync-queue/retry` | 同期再送 |

### 4.10 CSV・レポートAPI

| Method | Path | 役割 |
|---|---|---|
| `GET` | `/reports/monthly` | 月間有給取得一覧 |
| `GET` | `/reports/yearly` | 年間有給取得一覧 |
| `GET` | `/reports/balance` | 残日数一覧 |
| `GET` | `/export/csv/{type}` | CSV出力 |

---

## 5. 認証設計

### 5.1 GAS

#### 現在保持するもの

| Secret | 用途 |
|---|---|
| `SUPABASE_URL` | Supabase REST API接続 |
| `SUPABASE_ANON_KEY` | Supabase検証用 |
| `USE_SUPABASE_READS` | Read切替 |
| `USE_SUPABASE_WRITE` | Write切替 |

Apps ScriptではPropertiesServiceによりScript Properties等を利用できます。公式リファレンスでも、スクリプト・ユーザー・ドキュメント単位のProperties取得APIが提供されています。  
参考: [Class PropertiesService | Apps Script](https://developers.google.com/apps-script/reference/properties/properties-service)

#### 方針

- GASに `SERVICE_ROLE_KEY` を置かない
- GASにDB管理権限を持たせない
- GASは将来的にCloud Run APIだけを呼ぶ
- `SUPABASE_ANON_KEY` は移行検証中のみ利用

### 5.2 Cloud Run

#### 保持するもの

| Secret | 保持場所 |
|---|---|
| `SUPABASE_URL` | Secret Manager / env |
| `SUPABASE_SERVICE_ROLE_KEY` | Secret Manager |
| API認証用secret | Secret Manager |
| webhook secret | Secret Manager |

#### 方針

- Supabase service role keyはCloud Runだけが保持
- Secret Managerを使う
- Cloud Run service accountを最小権限にする
- ログにsecretを出さない
- すべてのAPIに認証をかける

### 5.3 Supabase

#### 使用キー

| Key | 用途 |
|---|---|
| anon key | クライアント読み取り限定、または未使用 |
| service role key | Cloud Runのみ |
| JWT | 将来のユーザー認証 |

#### 方針

- service role keyはフロント/GASに置かない
- RLSは必ず有効化
- Cloud Run経由の書き込みを基本とする
- `admin_users` はanon公開禁止

### 5.4 OAuth / ID連携

将来的には以下を検討します。

- Google Workspaceログイン
- 管理者はGoogleアカウント認証
- 社員は社員ID + PINからGoogle認証へ移行
- Cloud RunでID token検証
- Supabase Auth利用も候補

---

## 6. データベース責務

### 6.1 Supabaseを正DBとする理由

- PostgreSQLによる制約管理
- 外部キー
- enum
- index
- transaction
- audit log
- Realtime
- API化しやすい
- 将来拡張しやすい

### 6.2 Spreadsheetの扱い

#### Phase中

- 正DB
- 照合元
- 障害時の復旧元
- 業務継続用

#### Supabase正DB後

- バックアップ
- CSV出力先
- 一部管理者向け閲覧
- 月次スナップショット

#### 最終

- アーカイブ
- 手動編集禁止
- 自動同期停止
- 必要時のみエクスポート

### 6.3 CSV出力

完成後はCloud RunまたはSupabase queryを基準にします。

```text
Supabase
 ↓
Cloud Run report API
 ↓
CSV
 ↓
Download / Drive / Spreadsheet
```

### 6.4 監査

監査ログはSpreadsheetではなくSupabaseを正にします。

推奨テーブル:

```text
audit_logs
- audit_id
- actor_type
- actor_id
- action
- entity_type
- entity_id
- before_json
- after_json
- request_id
- ip_address
- user_agent
- created_at
```

---

## 7. 同期設計

### 7.1 sync_queueの必要性

Dual Writeでは以下が必ず発生します。

```text
Spreadsheet成功
Supabase失敗
```

Loggerだけでは復旧できません。  
そのため、同期キューが必要です。

### 7.2 sync_queue設計

#### Spreadsheet版 sync_queue

短期ではSpreadsheetに `sync_queue` シートを追加します。

```text
sync_queue
- sync_id
- entity_type
- entity_id
- operation
- payload_json
- status
- attempt_count
- last_error
- created_at
- updated_at
- synced_at
```

#### Supabase版 sync_queue

Cloud Run導入後はSupabaseへ移します。

```sql
sync_queue
- sync_id text primary key
- source text
- entity_type text
- entity_id text
- operation text
- payload_json jsonb
- status text
- attempt_count integer
- last_error text
- next_retry_at timestamptz
- created_at timestamptz
- updated_at timestamptz
- synced_at timestamptz
```

### 7.3 状態遷移

```text
pending
  ↓
processing
  ↓
synced

pending
  ↓
processing
  ↓
failed
  ↓
retry_wait
  ↓
processing
  ↓
synced

failed
  ↓
dead_letter
```

### 7.4 retry設計

#### retry方針

- 初回失敗後すぐ再送
- 2回目以降は指数バックオフ
- 最大回数を超えたら `dead_letter`
- 手動再送APIを用意

#### retry間隔例

| attempt | retry after |
|---:|---|
| 1 | 1分 |
| 2 | 5分 |
| 3 | 15分 |
| 4 | 1時間 |
| 5 | 手動確認 |

### 7.5 reconcile設計

定期的にSpreadsheetとSupabaseを照合します。

対象:

- 件数
- `request_id`
- `status`
- `days`
- `start_date`
- `end_date`
- `updated_at`
- 社員別残日数
- FIFO結果

---

## 8. 障害設計

### 8.1 Spreadsheet成功 / Supabase失敗

#### 現状

```text
Spreadsheet成功
Supabase失敗
Logger.log
業務継続
```

#### 問題

- 差分が残る
- Loggerからの復旧が難しい
- 再送対象が管理しづらい

#### 推奨

```text
Spreadsheet成功
 ↓
sync_queue pending作成
 ↓
Supabase送信
 ↓
成功: synced
失敗: failed + last_error
```

### 8.2 Cloud Run失敗

#### 想定

- 5xx
- timeout
- cold start
- deploy不具合
- secret設定ミス

#### 対策

- GASはエラーを表示しすぎない
- sync_queueに記録
- Cloud Run retry
- Error Reporting
- rollback revision
- health check

### 8.3 通信タイムアウト

#### 問題

タイムアウト時、実際にはSupabase INSERTが成功している可能性があります。

#### 対策

- `request_id` をprimary keyにする
- 409 conflictを成功扱いにできる設計
- payload hashを保存
- idempotency keyを使う

### 8.4 409 conflict

#### 方針

同一 `request_id` でINSERT済みの場合:

| 条件 | 扱い |
|---|---|
| payload一致 | 成功扱い |
| payload不一致 | 不整合エラー |
| status違い | reconcile対象 |

### 8.5 重複送信

#### 対策

- `request_id` を共通IDにする
- `idempotency_keys` テーブルを作る
- API側で同一キーを検出する
- 同一結果を返す

### 8.6 再送

#### 再送方式

```text
sync_queue
 ↓
Cloud Run retry job
 ↓
Supabase upsert / insert
 ↓
synced
```

---

## 9. Feature Flag

### 9.1 現在のFeature Flag

| Flag | 役割 | デフォルト | 削除タイミング |
|---|---|---|---|
| `USE_SUPABASE_READS` | 読み取りをSupabaseへ切替 | false | Supabase正DB移行後 |
| `USE_SUPABASE_WRITE` | 有給申請登録のDual Write | false | Cloud Run API移行後 |

### 9.2 今後追加予定

| Flag | 役割 |
|---|---|
| `USE_CLOUD_RUN_API` | GASからCloud Runを呼ぶ |
| `USE_SUPABASE_PRIMARY` | Supabaseを正DBにする |
| `DUAL_WRITE_LEAVE_APPROVAL` | 承認Dual Write |
| `DUAL_WRITE_LEAVE_CANCEL` | 取消Dual Write |
| `DUAL_WRITE_GRANTS` | 付与Dual Write |
| `DUAL_WRITE_EMPLOYEES` | 社員編集Dual Write |
| `DISABLE_SPREADSHEET_WRITES` | Spreadsheet書き込み停止 |
| `ENABLE_REALTIME_NOTIFICATIONS` | Realtime通知ON |

### 9.3 削除方針

Feature Flagは移行後に残し続けない方がよいです。

削除条件:

- 1〜2締め期間の照合完了
- 未同期queueゼロ
- rollback手順確立
- 運用担当者承認
- production移行完了

---

## 10. 段階移行計画

### 10.1 Phase 1: 現在

```text
Spreadsheet正
Supabase副
GASからRead検証
有給申請のみDual Write検証
```

#### 目標

- Supabaseデータ整合性確認
- Readレスポンス互換性確認
- 有給申請INSERTの検証

#### 完了条件

- 代表社員で履歴一致
- 残日数一致
- FIFO一致
- 申請登録Dual Write成功
- 失敗時ログ確認

### 10.2 Phase 2: Cloud Run導入

```text
GAS
 ↓
Cloud Run
 ↓
Supabase
```

#### 目標

- 書き込みAPIをCloud Runへ移す
- service role keyをCloud Runに隔離
- sync_queue実装
- idempotency実装

#### 対象

1. 有給申請登録
2. 承認
3. 否認
4. 本人取消
5. 管理者取消

### 10.3 Phase 3: Supabase正

```text
Supabase primary
Spreadsheet secondary
```

#### 目標

- 読み取りをSupabaseへ統一
- 書き込みをCloud Run + Supabaseへ統一
- Spreadsheetはバックアップへ降格

### 10.4 Phase 4: Spreadsheet読み取り停止

```text
GAS / Cloud Run
 ↓
Supabase only
```

#### 目標

- Spreadsheetからの本番読み取り停止
- CSVもSupabase基準
- 管理画面もSupabase基準

### 10.5 Phase 5: Spreadsheetアーカイブ

```text
Spreadsheet
 ↓
Archive / Snapshot only
```

#### 目標

- 手動編集禁止
- 月次スナップショット用途
- 必要ならDriveにCSV保管

---

## 11. 将来拡張

### 11.1 勤怠管理

Supabaseに以下を追加します。

```text
attendance_records
- employee_id
- work_date
- clock_in
- clock_out
- break_minutes
- status
```

有給申請と勤怠実績を突合できます。

### 11.2 GPS

```text
gps_logs
- employee_id
- recorded_at
- latitude
- longitude
- source
```

用途:

- 出退勤位置
- 車両位置
- 現場到着確認

### 11.3 配車

```text
dispatch_orders
- dispatch_id
- employee_id
- vehicle_id
- date
- route
- status
```

有給・休職・退職社員を配車候補から除外できます。

### 11.4 車両管理

```text
vehicles
vehicle_assignments
vehicle_inspections
```

有給管理の `is_driver`, `driver_type`, `default_vehicle_id` と連携できます。

### 11.5 ICカード

```text
ic_card_events
- card_id
- employee_id
- event_time
- device_id
- event_type
```

出勤記録と有給申請の突合に使えます。

### 11.6 Realtime通知

用途:

- 承認待ち通知
- 承認完了通知
- 否認通知
- 管理者取消通知
- 年度付与完了通知

### 11.7 監査ログ

全操作をCloud Run経由で記録します。

```text
audit_logs
- actor
- action
- before
- after
- request_id
- created_at
```

---

## 12. ベストプラクティス

### 12.1 Google / Cloud Run

- SecretはSecret Managerへ置く
- Cloud Run service accountを最小権限にする
- revision rollbackを前提にする
- health checkを用意する
- Cloud Logging / Error Reportingを使う
- Cloud Run jobsでreconcileやretryを実行する

Cloud RunはSecret Managerとの連携、サービスIDによる権限制御を公式にサポートしています。

- [Configure secrets for services | Cloud Run](https://cloud.google.com/run/docs/configuring/services/secrets)
- [Introduction to service identity | Cloud Run](https://cloud.google.com/run/docs/securing/service-identity)

### 12.2 Supabase

- RLSを有効化する
- anon keyには最小権限
- service role keyはサーバーだけ
- admin系テーブルはanon公開しない
- DB制約で最終防衛線を作る
- enum / FK / check / uniqueを活用する
- Realtimeは必要なテーブルだけ有効化する

参考:

- [Row Level Security | Supabase Docs](https://supabase.com/docs/guides/database/postgres/row-level-security)
- [Realtime | Supabase Docs](https://supabase.com/docs/guides/realtime)

### 12.3 GAS

- Script Propertiesへ設定値を置く
- keyをコードに直書きしない
- GASはUIとGoogle連携に寄せる
- 長い処理はCloud Runへ逃がす
- ログだけに依存せず、状態管理シート/DBを持つ

参考:

- [Class PropertiesService | Apps Script](https://developers.google.com/apps-script/reference/properties/properties-service)

### 12.4 このシステムに最適な構成

最適解は以下です。

```text
GAS = UI / Google連携
Cloud Run = 業務API
Supabase = 正DB
Spreadsheet = アーカイブ / CSV / 照合
```

---

## 13. 完成アーキテクチャ

```text
┌──────────────────────────────────────────────┐
│                  Users                       │
│                                              │
│  Staff / Admin / Back Office                 │
└───────────────────────┬──────────────────────┘
                        │
                        ▼
┌──────────────────────────────────────────────┐
│                UI Layer                      │
│                                              │
│  GAS Web App or Future Web Frontend          │
│                                              │
│  Responsibilities:                           │
│  - screen rendering                          │
│  - input collection                          │
│  - Google Workspace integration              │
│  - download/export trigger                   │
└───────────────────────┬──────────────────────┘
                        │ HTTPS
                        ▼
┌──────────────────────────────────────────────┐
│              Cloud Run API                   │
│                                              │
│  Responsibilities:                           │
│  - authentication                            │
│  - authorization                             │
│  - validation                                │
│  - business rules                            │
│  - status transition                         │
│  - idempotency                               │
│  - audit logging                             │
│  - retry/reconcile                           │
│  - CSV/report generation                     │
└───────────────────────┬──────────────────────┘
                        │ service role
                        ▼
┌──────────────────────────────────────────────┐
│              Supabase PostgreSQL             │
│                                              │
│  Tables:                                     │
│  - employees                                 │
│  - leave_requests                            │
│  - paid_leave_grants                         │
│  - company_calendar                          │
│  - usage_logs                                │
│  - audit_logs                                │
│  - sync_queue                                │
│  - idempotency_keys                          │
└───────────────────────┬──────────────────────┘
                        │
                        ▼
┌──────────────────────────────────────────────┐
│              Supabase Realtime               │
│                                              │
│  - approval notifications                    │
│  - status updates                            │
│  - admin dashboard refresh                   │
└──────────────────────────────────────────────┘

Migration / Backup:
┌──────────────────────────────────────────────┐
│              Spreadsheet                     │
│                                              │
│  During migration: primary DB                │
│  After migration: archive / export / backup  │
└──────────────────────────────────────────────┘
```

---

## 14. 技術的負債

| 優先度 | 課題 | 重要度 | 対応方針 |
|---|---|---:|---|
| 1 | SpreadsheetとSupabaseの差分回復がLogger依存 | ★★★★★ | sync_queue実装 |
| 2 | anon key + RLSでINSERT検証中 | ★★★★★ | Cloud Runへ移行 |
| 3 | admin_usersがSpreadsheet依存 | ★★★★☆ | Cloud Run認証/API化 |
| 4 | Dual Read時に読み書きの正が分かれる | ★★★★☆ | 本番ON前に同期保証 |
| 5 | 承認/否認/取消がSupabase未対応 | ★★★★☆ | 状態遷移API設計 |
| 6 | usage_logの参照先混在 | ★★★☆☆ | target_type設計へ統一 |
| 7 | FIFO計算が複雑でDB移行影響大 | ★★★★★ | 代表社員で継続照合 |
| 8 | CSV出力がSpreadsheet依存 | ★★★☆☆ | Cloud Run report API化 |
| 9 | Feature Flagが増える可能性 | ★★★☆☆ | 移行後に削除計画 |
| 10 | 監査ログが業務DBと完全統合されていない | ★★★★☆ | audit_logs追加 |

---

## 15. 実装ロードマップ

### 15.1 Step 1: sync_queue設計・実装

| 項目 | 内容 |
|---|---|
| 難易度 | 中 |
| 工数目安 | 1〜2日 |
| リスク | 中 |
| 優先度 | ★★★★★ |

#### 内容

- Spreadsheetに `sync_queue` シート追加
- Dual Write失敗を記録
- retry対象を可視化
- 手動再送関数追加

### 15.2 Step 2: Supabase INSERT冪等化

| 項目 | 内容 |
|---|---|
| 難易度 | 中 |
| 工数目安 | 0.5〜1日 |
| リスク | 中 |
| 優先度 | ★★★★★ |

#### 内容

- 409 conflictを成功扱い可能にする
- `request_id` 重複時のpayload照合
- `idempotency_key` 設計

### 15.3 Step 3: Cloud Run API基盤

| 項目 | 内容 |
|---|---|
| 難易度 | 高 |
| 工数目安 | 3〜5日 |
| リスク | 中 |
| 優先度 | ★★★★★ |

#### 内容

- Cloud Run project scaffold
- Secret Manager連携
- Supabase client設定
- API認証
- health check
- logging

### 15.4 Step 4: 有給申請登録API化

| 項目 | 内容 |
|---|---|
| 難易度 | 中 |
| 工数目安 | 1〜2日 |
| リスク | 中 |
| 優先度 | ★★★★★ |

#### 内容

- `POST /leave-requests`
- GASからCloud Run呼び出し
- Supabase transaction
- `usage_logs` 保存
- Spreadsheet backup optional

### 15.5 Step 5: 承認/否認/取消API化

| 項目 | 内容 |
|---|---|
| 難易度 | 高 |
| 工数目安 | 3〜5日 |
| リスク | 高 |
| 優先度 | ★★★★★ |

#### 内容

- 状態遷移表作成
- API実装
- audit log
- 残日数影響確認
- 代表社員で照合

### 15.6 Step 6: 付与API化

| 項目 | 内容 |
|---|---|
| 難易度 | 高 |
| 工数目安 | 3〜5日 |
| リスク | 高 |
| 優先度 | ★★★★☆ |

#### 内容

- 6ヶ月付与
- 年次付与
- 友尚建設基準日
- 初期付与
- FIFO影響確認

### 15.7 Step 7: 社員編集API化

| 項目 | 内容 |
|---|---|
| 難易度 | 中〜高 |
| 工数目安 | 2〜4日 |
| リスク | 高 |
| 優先度 | ★★★★☆ |

#### 内容

- 社員追加
- 社員編集
- 退職処理
- 有給管理対象変更
- audit log

### 15.8 Step 8: Read完全Supabase化

| 項目 | 内容 |
|---|---|
| 難易度 | 中 |
| 工数目安 | 2〜3日 |
| リスク | 中 |
| 優先度 | ★★★★☆ |

#### 内容

- `USE_SUPABASE_READS=true` を標準化
- Spreadsheet read停止
- 管理画面/個人画面/CSV照合

### 15.9 Step 9: Spreadsheet書き込み停止

| 項目 | 内容 |
|---|---|
| 難易度 | 中 |
| 工数目安 | 1〜2日 |
| リスク | 高 |
| 優先度 | ★★★★★ |

#### 条件

- sync_queueゼロ
- 1〜2締め期間の照合完了
- rollback手順あり
- 運用承認済み

### 15.10 Step 10: Realtime / 通知

| 項目 | 内容 |
|---|---|
| 難易度 | 中 |
| 工数目安 | 2〜4日 |
| リスク | 低〜中 |
| 優先度 | ★★★☆☆ |

#### 内容

- 承認待ち通知
- ステータス変更通知
- 管理画面自動更新
- 将来Web UI対応

---

## 16. 最終提案

### 16.1 私ならこう設計する

短期では、現在のDual Writeは維持します。ただし対象は有給申請登録だけに限定し、次に必ず `sync_queue` を作ります。

中期では、Cloud Runを導入し、Supabaseへの書き込みをすべてCloud Runに集約します。GASはCloud Run APIを呼ぶだけにします。

長期では、Supabaseを正DBにし、SpreadsheetはCSV・監査・バックアップ用途へ降格します。

最終形は以下です。

```text
GAS / Web UI
    ↓
Cloud Run API
    ↓
Supabase PostgreSQL
    ↓
Supabase Realtime

Spreadsheet
    ↓
Archive / CSV / Backup only
```

この構成が最も安全です。

理由は以下です。

- GASに強いDB権限を置かない
- Supabase service role keyをCloud Runに閉じ込められる
- 業務ルールをCloud Runに集約できる
- retry / reconcile / idempotencyを設計しやすい
- 将来の勤怠・GPS・配車・車両管理へ拡張しやすい
- Spreadsheetを急に捨てず、安全に移行できる

---

## 17. 参考資料

- [Google Cloud Run: Configure secrets for services](https://cloud.google.com/run/docs/configuring/services/secrets)
- [Google Cloud Run: Introduction to service identity](https://cloud.google.com/run/docs/securing/service-identity)
- [Google Apps Script: PropertiesService](https://developers.google.com/apps-script/reference/properties/properties-service)
- [Supabase: Row Level Security](https://supabase.com/docs/guides/database/postgres/row-level-security)
- [Supabase: Realtime](https://supabase.com/docs/guides/realtime)
