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
```

### 請求集計対象会社（20日締め）
- 梶原通信（20日締め）
- その他は全て末締め

---

## 経費管理システム（business_management_workflow）

- ファイルパス：`/home/user/genai-lessons/workflows/business_management_workflow_v2.json`
- プレースホルダー：YOUR_EXPENSE_LINE_TOKEN, YOUR_CLAUDE_API_KEY, YOUR_USERID_NOZOE_YU, YOUR_USERID_NOZOE_AKIKO
