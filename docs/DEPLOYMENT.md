# 本番Deployment運用情報

## 本番Web App

確認日: 2026-09-17

| 項目 | 値 |
| --- | --- |
| Deployment ID | `AKfycbw5etjnlUaElnuVeUIiWCM5yCJWf9YaPmyuzSMRpqoKDiydmMI9msxqa4M01PIOt8XD_w` |
| 現在の本番version | `382` |
| Version description | `Keep combined leave button on one line` |
| 種別 | Web App |

本番更新では、上記の既存Deployment IDを維持したまま新しいversionへ更新します。新規Deploymentは作成しません。Deployment IDを維持する限り、既存のWeb App URLも維持されます。

## 本番更新の基本手順

1. `git status --short` を確認する。
2. テストと `git diff --check` を実行する。
3. commitする。
4. `git push` する。
5. `clasp push` を実行する。
6. 新しいApps Script versionを作成する。
7. 上記の既存本番Deployment IDを新versionへ更新する。
8. `clasp deployments` でDeployment IDとversionを確認する。
9. 本番Web Appで実機確認する。

## 重要な注意

- 原則として新規Deploymentを作成しない。
- 本番更新では必ず上記Deployment IDを使用する。
- Deployment IDが変わると、現場で利用しているURLやブックマーク等へ影響する可能性がある。
- 更新前に、現在の本番versionを必ず確認する。
- この記録とGoogle Apps Script側の状態が食い違う場合は、この記録だけを信用して更新しない。実際の運用URLとApps ScriptのDeploymentを照合して本番Deploymentを特定する。
- 本番versionを更新したら、このドキュメントの「現在の本番version」とVersion descriptionを更新する。

このファイルには、パスワード、APIキー、アクセストークン、セッショントークン、private key、`.env` の秘密情報を記録しない。
