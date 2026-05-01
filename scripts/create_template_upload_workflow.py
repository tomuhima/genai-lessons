#!/usr/bin/env python3
"""
経費管理テンプレートを OneDrive にアップロードするワンショットワークフロー生成
アップロード先: /業務管理システム/経費管理/経費管理_テンプレート.xlsx
"""
import json

js_code = r"""
const fs = require('fs');
const fileName = '経費管理_テンプレート.xlsx';
const filePath = `/tmp/${fileName}`;

if (!fs.existsSync(filePath)) {
  throw new Error(`ファイルが見つかりません: ${filePath}\npython3 scripts/create_expense_template.py で生成してからコピーしてください`);
}

const fileBuffer = fs.readFileSync(filePath);
const bin = await this.helpers.prepareBinaryData(
  fileBuffer,
  fileName,
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
);
return [{ json: { fileName }, binary: { data: bin } }];
"""

workflow = {
    "id": "bm-template-upload-001",
    "name": "経費管理テンプレートアップロード",
    "nodes": [
        {
            "parameters": {},
            "name": "手動実行",
            "type": "n8n-nodes-base.manualTrigger",
            "typeVersion": 1,
            "position": [240, 300],
            "id": "tmpl-node-01"
        },
        {
            "parameters": {
                "jsCode": js_code.strip()
            },
            "name": "テンプレート読み込み",
            "type": "n8n-nodes-base.code",
            "typeVersion": 2,
            "position": [480, 300],
            "id": "tmpl-node-02"
        },
        {
            "parameters": {
                "method": "PUT",
                "url": "=https://graph.microsoft.com/v1.0/me/drive/root:/業務管理システム/経費管理/経費管理_テンプレート.xlsx:/content",
                "authentication": "predefinedCredentialType",
                "nodeCredentialType": "microsoftDriveOAuth2Api",
                "sendBody": True,
                "contentType": "binaryData",
                "inputDataFieldName": "data",
                "options": {}
            },
            "name": "OneDriveアップロード",
            "type": "n8n-nodes-base.httpRequest",
            "typeVersion": 4,
            "position": [720, 300],
            "id": "tmpl-node-03"
        },
        {
            "parameters": {
                "jsCode": "const r = $input.first().json;\nconsole.log('テンプレートアップロード完了:', r.name, r.id);\nreturn [{ json: { success: true, name: r.name, id: r.id } }];"
            },
            "name": "完了確認",
            "type": "n8n-nodes-base.code",
            "typeVersion": 2,
            "position": [960, 300],
            "id": "tmpl-node-04"
        }
    ],
    "connections": {
        "手動実行": {
            "main": [[{"node": "テンプレート読み込み", "type": "main", "index": 0}]]
        },
        "テンプレート読み込み": {
            "main": [[{"node": "OneDriveアップロード", "type": "main", "index": 0}]]
        },
        "OneDriveアップロード": {
            "main": [[{"node": "完了確認", "type": "main", "index": 0}]]
        }
    },
    "active": False,
    "settings": {}
}

out_path = '/home/user/genai-lessons/workflows/template_upload.json'
with open(out_path, 'w', encoding='utf-8') as f:
    json.dump(workflow, f, ensure_ascii=False, indent=2)

print(f"✓ ワークフロー生成: {out_path}")
print()
print("=== OneDriveにテンプレートをアップロードする手順 ===")
print()
print("1. テンプレートファイルを生成:")
print("   python3 scripts/create_expense_template.py")
print()
print("2. /tmp/ にテンプレートとしてコピー:")
print("   cp /tmp/経費管理_2026年04月.xlsx /tmp/経費管理_テンプレート.xlsx")
print()
print("3. n8nサーバーにコピー:")
print("   docker cp /tmp/経費管理_テンプレート.xlsx n8n-compose-n8n-1:/tmp/")
print()
print("4. ワークフローをn8nにインポート:")
print("   python3 -c \"import json; f=open('workflows/template_upload.json'); w=json.load(f); f.close(); w['id']='tmpl-001'; open('/tmp/tmpl.json','w').write(json.dumps(w))\"")
print("   docker cp /tmp/tmpl.json n8n-compose-n8n-1:/tmp/")
print("   docker exec n8n-compose-n8n-1 n8n import:workflow --input=/tmp/tmpl.json")
print()
print("5. n8n UIで「経費管理テンプレートアップロード」ワークフローを手動実行")
print()
print("6. OneDrive /業務管理システム/経費管理/経費管理_テンプレート.xlsx に保存されれば完了")
