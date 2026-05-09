# 📘 Paid Leave Management App（有給申請システム）

社内DXプロジェクトとして開発している  
Google Apps Script + スプレッドシートベースの有給申請システム

---

# 🎯 概要

このアプリは以下を目的としています：

- 紙の有給申請の廃止
- タブレットで誰でも簡単に申請
- 管理者の承認業務の効率化
- 有給取得状況の見える化
- 社員マスターの一元管理
- 有給付与漏れ防止

---

# 🧩 システム構成

- フロントエンド：HTML / CSS / JavaScript
- バックエンド：Google Apps Script（GAS）
- データベース：Google Spreadsheet
- 開発環境：VSCode + clasp + GitHub

---

# 🖥️ 画面構成

## 🏠 トップメニュー（menu.html）

- 有給申請
- 申請管理
- 社員マスター管理

各画面へカードUIで遷移

---

## 👤 申請画面（index.html）

### 主な機能

- 五十音で社員選択
- 有給残数表示
- 付与日数表示
- 5日取得義務進捗表示
- カレンダーUI
- 1日 / 半日 / 複数日申請
- 理由入力
- 年度内履歴表示

### 表示される情報

- 日付
- 休暇区分
- 日数
- 承認状態
  - 承認待ち
  - 承認済み
  - 否認

---

## 🛠 管理画面（admin.html）

### 主な機能

- 承認待ち一覧
- 承認済み一覧
- 否認一覧
- ワンクリック承認
- ワンクリック否認
- ログ確認
- 月間レポート出力
- 年間レポート出力

### 管理者認証

- 管理者選択
- PINログイン
- localStorageログイン保持
- 承認者名のログ保存

---

## 👥 社員マスター管理（employee-admin.html）

### 主な機能

- 社員追加
- 社員編集
- 退職処理
- 表示順整理
- ID自動採番
- 表示ID自動採番

### 管理項目

- 社員ID
- 表示ID
- 氏名
- ふりがな
- 会社区分
- 部署
- 雇用区分
- 在職状況
- 入社日
- 退職日
- 有給管理対象
- 有給年度開始月

---

# 📊 データ構造（主要シート）

## employees

- employee_id
- display_employee_id
- name
- name_kana
- company_code
- department
- employment_status
- hire_date
- leave_date

---

## leave_requests

- request_id
- employee_id
- start_date
- end_date
- days
- half_day
- reason
- status

---

## paid_leave_grants

- grant_id
- employee_id
- grant_days
- carry_over_days
- grant_type
- year

---

## company_calendar

- date
- type
  - workday
  - holiday
  - no_leave

---

## usage_log

- log_id
- action_type
- operator_name
- action_date
- comment

---

## admin_users

- admin_id
- admin_name
- pin
- is_active

---

# ⚙️ 主な機能

## 有給申請

- 1日申請
- 半日申請
- 複数日申請

---

## 営業日判定

company_calendar と連動

- 営業日のみ申請可能
- no_leave日は申請不可

---

## 有給残日数管理

- 年度別残数計算
- 繰越管理
- 消滅管理

---

## 有給付与

### 初回付与

- 入社6か月到達者抽出
- ワンクリック付与

### 年次付与

- 勤続年数自動判定
- 法定付与日数自動計算

---

## ログ管理

- 申請
- 承認
- 否認
- 社員追加
- 社員編集
- 退職処理

すべて usage_log に記録

---

# 🏢 マルチ会社対応

複数会社運用に対応

- MAIN
- PARTNER

会社ごとに：

- 有給年度開始月
- 出力レポート
- 社員管理

を分離可能

---

# 🎨 UI/UXの特徴

- タブレット前提UI
- iPad横置き最適化
- ブルー系デザイン
- 大型ボタン
- カレンダーUI
- モーダル編集UI
- 申請履歴表示
- ミス防止UI

---

# 📂 ファイル構成

```text
Code.gs
index.html
admin.html
employee-admin.html
menu.html
style.html
script.html
appsscript.json