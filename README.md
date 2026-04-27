# 📘 Paid Leave Management App（有給申請システム）

社内DXプロジェクトとして開発している  
**Google Apps Script + スプレッドシートベースの有給申請システム**

---

## 🎯 概要

このアプリは以下を目的としています：

- 紙の有給申請の廃止
- タブレットで誰でも簡単に申請
- 管理者の承認業務の効率化
- 有給取得状況の見える化

---

## 🧩 システム構成

- フロントエンド：HTML / CSS / JavaScript
- バックエンド：Google Apps Script（GAS）
- データベース：Google Spreadsheet
- 開発環境：VSCode + clasp + GitHub

---

## 🖥️ 画面構成

### 👤 申請画面（index.html）

- 五十音で社員選択
- 有給残数・付与日数表示
- 5日取得義務の進捗表示
- 日付選択・区分選択（1日 / 半日 / 複数日）
- 理由入力

### 🆕 追加機能

- 年度内の申請履歴表示（申請ボタン下）
  - 日付
  - 休暇種別（1日 / 半日）
  - 日数
  - ステータス（承認待ち / 承認済み / 否認）

---

### 🛠 管理画面（admin.html）

- 承認待ち / 承認済み / 否認の一覧表示
- ワンクリック承認・否認
- ログ確認
- 月間・年間レポート出力

---

## 📊 データ構造（主要シート）

### employees
- employee_id
- name
- name_kana

### leave_requests
- request_id
- employee_id
- start_date
- end_date
- days
- half_day
- reason
- status

### paid_leave_grants
- employee_id
- grant_days
- carry_over_days
- year

### company_calendar
- date
- type（workday / holiday / no_leave）

### usage_log
- 操作履歴

---

## ⚙️ 主な機能

- 有給申請（1日 / 半日 / 複数日）
- 営業日判定（company_calendar連動）
- 年度（4月開始）での残日数管理
- 承認ワークフロー
- 使用ログ記録
- 月次・年次レポート出力
- 社員別残日数自動計算

---

## 🎨 UI/UXの特徴

- タブレット前提の大きな操作UI
- 爽やかなブルー系デザイン
- 一覧はカードではなくコンパクトなリスト
- ミス防止（確認ダイアログ・ボタン制御）

---

## 🚀 セットアップ

### ① claspログイン
```bash
clasp login