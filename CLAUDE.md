# プロジェクト作業ルール

## 大きなJSONファイルの編集ルール（分割方式）

JSONファイルを直接出力・編集する際は必ずPythonスクリプトを使う。
理由：大きなJSONをテキスト出力するとStream idle timeoutが発生する。

### やり方
```python
import json

# 読み込み
with open('path/to/file.json', 'r', encoding='utf-8') as f:
    w = json.load(f)

# 必要な箇所だけ修正
for node in w['nodes']:
    if node['name'] == '対象ノード名':
        node['parameters']['jsCode'] = '新しいコード'
        break

# 保存
with open('path/to/file.json', 'w', encoding='utf-8') as f:
    json.dump(w, f, ensure_ascii=False, indent=2)
```

### 絶対にやらないこと
- 大きなJSONファイルをそのままテキストで出力しない
- EditツールでJSONの複数行マッチングをしない（\nを含む文字列）
- Agentサブエージェントで大きなJSONを生成させない

---

## ドキュメントファイルの管理ルール

### 全チャットから参照できるようにする方法
ローカルブランチにコミットしただけでは他のチャットからファイルが見えない。
**重要なdocsファイルは必ずmainブランチにも直接プッシュする。**

#### 手順（成功パターン）
1. ローカルにファイルを作成・Read
2. `mcp__github__get_file_contents`でmainブランチに既存ファイルか確認（SHAを取得）
3. `mcp__github__create_or_update_file`でmainブランチに直接コミット
   - 新規ファイル：shaパラメータ不要
   - 既存ファイル：sha必須（ステップ2で取得したSHAを使う）

#### 絶対にやらないこと
- フィーチャーブランチにのみコミットして「保存した」と言わない
- docsの引き継ぎファイルをmain以外にだけ置かない

### docsディレクトリの現ファイル（mainブランチ）
| ファイル | 内容 |
|----------|------|
| `docs/daily_report_business_summary.md` | 日報システム引き継ぎ・ビジネス展開資料 |
| `docs/business_management_summary.md` | 経営管理システム引き継ぎ資料 |
| `docs/integration_design_notes.md` | 両システム連携設計メモ |

---

## n8nワークフロー作業ルール

- n8nはバージョン2.14.2（Self Hosted）、Dockerで動作
- サーバー：XServer VPS（Dockerコンテナ名：n8n-compose-n8n-1）
- ワークフローインポート手順：
  1. `python3 -c "import json; f=open('file.json'); w=json.load(f); f.close(); w['id']='100'; open('/tmp/fixed.json','w').write(json.dumps(w))"`
  2. `docker cp /tmp/fixed.json n8n-compose-n8n-1:/tmp/`
  3. `docker exec n8n-compose-n8n-1 n8n import:workflow --input=/tmp/fixed.json`

### n8n 2.14.2の注意事項
- Ifノードの文字列比較は `equal`（`equals`は不可）
- HTTP RequestノードでMicrosoft Graph APIを使う場合は `microsoftDriveOAuth2Api` を使用
- ExcelJSは `useSharedStrings: false` で読み込む
- Excel日付セルの判定：`typeof cell.getTime === 'function'`
- Excelファイルのダウンロードは `@microsoft.graph.downloadUrl` を使って Code ノード内で直接 httpRequest する（ファイルダウンロードノード不要）

---

## 日報システム（daily_report_workflow）

- ファイルパス：`/home/user/genai-lessons/workflows/daily_report_workflow_v2.json`
- Webhook URL：`https://n8n-light-mn.xvps.jp/webhook/daily-report`
- OneDriveパス：`/日報/稼働表_YYYY年MM月.xlsx`
- 請求集計パス：`/請求集計/請求集計_YYYY年MM月.xlsx`

### EMPLOYEE_MAP（LINE UserID → 名前）
```
Uc1e1e958fb952752c3d34628a17585a1 → 野添優
Uf662ab811bee1780522da7356c0efb9e → 丸田翔吾
Ua9a65f01bf77adcc825b05811950480d → 源地健史
U65929805c0527b9f75b5276fc250c304 → 服部秀一
Uadfb98160637cc7f96bccb051c4a51f9 → 梶原通信
Uf97fd2e35007f697094a9520e1baeafd → 井本貴士
```

### 請求集計対象会社（20日締め）
- 梶原通信（20日締め）
- その他は全て末締め

---

## 経費管理システム（business_management_workflow）

- ファイルパス：`/home/user/genai-lessons/workflows/business_management_workflow_v2.json`
- プレースホルダー：YOUR_EXPENSE_LINE_TOKEN, YOUR_CLAUDE_API_KEY, YOUR_USERID_NOZOE_YU, YOUR_USERID_NOZOE_AKIKO

---

## 動作確認済みコードパターン

### LINE返信（Codeノード）
```javascript
const lineToken = 'Bearer XXXXXXX'; // 実際のトークンに置き換え
const data = $input.first().json;
if (data.replyToken) {
  await this.helpers.httpRequest({
    method: 'POST',
    url: 'https://api.line.me/v2/bot/message/reply',
    headers: { 'Authorization': lineToken },
    body: { replyToken: data.replyToken, messages: [{ type: 'text', text: data.replyMessage }] },
    json: true
  });
}
return [{ json: { done: true } }];
```

### OneDriveファイル情報取得（HTTP Requestノード）
- 方法：GET
- URL：`=https://graph.microsoft.com/v1.0/me/drive/root:/日報/{{ $json.fileName }}`
- 認証：Predefined Credential Type → Microsoft Drive OAuth2 API

### OneDriveアップロード（HTTP Requestノード）
- 方法：PUT
- URL：`=https://graph.microsoft.com/v1.0/me/drive/items/{{ $('ファイル情報取得').first().json.id }}/content`
- 認証：Predefined Credential Type → Microsoft Drive OAuth2 API

### ExcelJSでファイル読み込み（Codeノード）
```javascript
const ExcelJS = require('exceljs');
const fileInfo = $('ファイル情報取得').first().json;
const downloadUrl = fileInfo['@microsoft.graph.downloadUrl'];
if (!downloadUrl) throw new Error('ダウンロードURLが取得できませんでした');
const res = await this.helpers.httpRequest({
  method: 'GET', url: downloadUrl, encoding: 'arraybuffer', returnFullResponse: false
});
const workbook = new ExcelJS.Workbook();
await workbook.xlsx.load(Buffer.from(res), { useSharedStrings: false });
```

### Excel日付セルの判定
```javascript
const getVal = (cell) => (cell && typeof cell === 'object' && 'result' in cell) ? cell.result : cell;
if (typeof dateCell.getTime === 'function') { /* 日付セル */ }
```

### ExcelJSでバイナリ出力してアップロード
```javascript
const out = await workbook.xlsx.writeBuffer({ useSharedStrings: false });
const bin = await this.helpers.prepareBinaryData(
  Buffer.from(out), fileName,
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
);
return [{ json: { ...data }, binary: { data: bin } }];
```

### 新規Excelファイル作成（パスベースPUT）
- URL：`=https://graph.microsoft.com/v1.0/me/drive/root:/請求集計/{{ $json.billingFileName }}:/content`
- 方法：PUT（ファイルが存在しなくても自動作成される）
