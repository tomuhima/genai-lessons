#!/usr/bin/env python3
"""
月次ファイルが存在しない場合にテンプレートから新規作成する機能を追加。
1. テンプレートファイル情報取得 HTTP Requestノードを追加
2. Excel書き込み＋アップロード コードを修正（テンプレートフォールバック）
3. 接続を更新
"""
import json

path = '/home/user/genai-lessons/workflows/business_management_workflow_v3.json'
with open(path, 'r', encoding='utf-8') as f:
    w = json.load(f)

# ===== 1. Excel書き込み＋アップロード のコードを修正 =====
OLD_HEADER = (
    "const ExcelJS = require('exceljs');\n"
    "return await (async () => {\n"
    "  const data = $('LINEメッセージ解析').first().json;\n"
    "  const fileInfo = $('月次ファイル情報取得').first().json;\n"
    "  try {\n"
    "    const wb = new ExcelJS.Workbook();\n"
    "    const dlUrl = fileInfo['@microsoft.graph.downloadUrl'];\n"
    "    if (dlUrl) {\n"
    "      try {\n"
    "        const r = await this.helpers.httpRequest({ method:'GET', url:dlUrl, encoding:'arraybuffer', returnFullResponse:false });\n"
    "        await wb.xlsx.load(Buffer.from(r), { useSharedStrings: false });\n"
    "      } catch(e) {}\n"
    "    }"
)
NEW_HEADER = (
    "const ExcelJS = require('exceljs');\n"
    "return await (async () => {\n"
    "  const data = $('LINEメッセージ解析').first().json;\n"
    "  const fileInfo = $('月次ファイル情報取得').first().json;\n"
    "  const templateInfo = $('テンプレートファイル情報取得').first().json;\n"
    "  try {\n"
    "    const wb = new ExcelJS.Workbook();\n"
    "    // 月次ファイルが存在しない場合はテンプレートから読み込む\n"
    "    const dlUrl = fileInfo['@microsoft.graph.downloadUrl'] || templateInfo['@microsoft.graph.downloadUrl'];\n"
    "    if (dlUrl) {\n"
    "      try {\n"
    "        const r = await this.helpers.httpRequest({ method:'GET', url:dlUrl, encoding:'arraybuffer', returnFullResponse:false });\n"
    "        await wb.xlsx.load(Buffer.from(r), { useSharedStrings: false });\n"
    "      } catch(e) {}\n"
    "    }"
)

for node in w['nodes']:
    if node['name'] == 'Excel書き込み＋アップロード':
        code = node['parameters']['jsCode']
        assert OLD_HEADER in code, "置換対象が見つかりません"
        node['parameters']['jsCode'] = code.replace(OLD_HEADER, NEW_HEADER)
        print("✓ Excel書き込み＋アップロード: テンプレートフォールバックを追加")

# ===== 2. テンプレートファイル情報取得 ノードを追加 =====
template_node = {
    "parameters": {
        "method": "GET",
        "url": "=https://graph.microsoft.com/v1.0/me/drive/root:/業務管理システム/経費管理/経費管理_テンプレート.xlsx",
        "authentication": "predefinedCredentialType",
        "nodeCredentialType": "microsoftDriveOAuth2Api",
        "options": {}
    },
    "name": "テンプレートファイル情報取得",
    "type": "n8n-nodes-base.httpRequest",
    "typeVersion": 4,
    "position": [990, 500],
    "id": "bm-v3-template-info",
    "onError": "continueRegularOutput"
}
w['nodes'].append(template_node)
print("✓ テンプレートファイル情報取得 ノードを追加")

# ===== 3. 接続を更新 =====
# 変更前: 月次ファイル情報取得 → Excel書き込み＋アップロード
# 変更後: 月次ファイル情報取得 → テンプレートファイル情報取得 → Excel書き込み＋アップロード
conn = w['connections']
conn['月次ファイル情報取得'] = {
    "main": [[{"node": "テンプレートファイル情報取得", "type": "main", "index": 0}]]
}
conn['テンプレートファイル情報取得'] = {
    "main": [[{"node": "Excel書き込み＋アップロード", "type": "main", "index": 0}]]
}
print("✓ 接続を更新: 月次ファイル情報取得 → テンプレートファイル情報取得 → Excel書き込み＋アップロード")

# ===== 4. 保存 =====
with open(path, 'w', encoding='utf-8') as f:
    json.dump(w, f, ensure_ascii=False, indent=2)
print(f"\n✓ 保存完了: {path}")
