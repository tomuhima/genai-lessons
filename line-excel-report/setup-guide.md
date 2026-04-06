# LINE日報 → Excel月間稼働表 自動化 セットアップガイド

## システム概要

```
職人（LINE送信）
    ↓
LINE Messaging API（Webhook）
    ↓
n8n（自動処理）
    ├── メッセージ解析
    ├── 社員識別
    ├── OneDrive Excel更新
    └── LINE返信
```

## 日報送信フォーマット

職人さんがLINEで以下の形式で送信します：

```
日報
日付：2026/04/06
会社名：〇〇株式会社
物件名：渋谷タワービル
開始：09:00
終了：18:00
休憩：60
作業内容：電気配線工事、盤内配線確認
```

## Excelファイル構成

- **保存場所**: OneDrive `/稼働表/`
- **ファイル名**: `稼働表_YYYY年.xlsx`（例: `稼働表_2026年.xlsx`）
- **シート構成**: 社員ごとに1シート（自動作成）

| 日付 | 会社名 | 物件名 | 開始時間 | 終了時間 | 休憩(分) | 実働時間 | 作業内容 |
|------|--------|--------|----------|----------|----------|----------|----------|

---

## セットアップ手順

### 1. n8n環境構築（Docker）

```bash
# docker-compose.yml を編集して環境に合わせる
# WEBHOOK_URL を実際のドメインまたはngrokのURLに変更
# N8N_BASIC_AUTH_PASSWORD を安全なパスワードに変更

docker-compose up -d
```

### 2. LINE Messaging API設定

1. [LINE Developers Console](https://developers.line.biz/) でチャンネル作成
2. **Messaging API** チャンネルを選択
3. **Webhook URL** を設定:
   ```
   https://your-domain.com/webhook/line-webhook
   ```
4. **チャンネルアクセストークン（長期）** を発行・メモ

### 3. n8n設定

#### 3-1. Microsoft OneDriveクレデンシャル設定

1. n8n管理画面 → **Credentials** → **Add Credential**
2. `Microsoft OneDrive OAuth2 API` を選択
3. Azure App Registration で取得したクライアントID・シークレットを入力

#### 3-2. LINE認証設定

1. n8n管理画面 → **Credentials** → **Add Credential**
2. `Header Auth` を選択
3. Name: `Authorization`
4. Value: `Bearer {チャンネルアクセストークン}`

#### 3-3. ワークフローインポート

1. n8n管理画面 → **Workflows** → **Import from file**
2. `n8n-workflow.json` を選択してインポート

### 4. 社員マッピング設定

ワークフローの **「社員ID変換」ノード** を開き、`EMPLOYEE_MAP` を編集：

```javascript
const EMPLOYEE_MAP = {
  "U実際のLINEユーザーID1": "田中太郎",
  "U実際のLINEユーザーID2": "山田花子",
  // 社員追加時はここに追加
};
```

#### LINE UserIDの確認方法
1. ワークフローを有効化
2. 社員にLINEでテストメッセージを送ってもらう
3. n8nのWebhookノードの実行ログで `body.events[0].source.userId` を確認

### 5. OneDriveフォルダ設定

OneDrive に `/稼働表/` フォルダを作成しておきます。
（初回日報送信時に自動でExcelファイルが作成されます）

### 6. 動作確認

1. ワークフローを **Active** に設定
2. LINE DevelopersでWebhookの **「検証」** をクリック
3. 職人さんに日報フォーマットで送信してもらう
4. OneDriveのExcelファイルを確認

---

## トラブルシューティング

| 症状 | 原因 | 対処 |
|------|------|------|
| LINEからWebhookが来ない | Webhook URLが間違っている | LINE DeveloperでURLを確認 |
| 「登録されていないアカウント」エラー | UserIDが未登録 | EMPLOYEE_MAPにUserIDを追加 |
| Excelが更新されない | OneDriveクレデンシャルの問題 | 認証を再設定 |
| `Cannot find module 'exceljs'` | ExcelJSが未インストール | docker-compose.ymlの`NODE_FUNCTION_ALLOW_EXTERNAL`を確認 |

---

## 社員追加方法

1. 新しい社員にLINE公式アカウントを友達追加してもらう
2. 日報を1通送ってもらい、n8nのログでUserIDを確認
3. ワークフローの「社員ID変換」ノードの`EMPLOYEE_MAP`に追加
4. `employee-mapping.json`にも同様に追加（管理用）
