#!/usr/bin/env python3
"""
経費管理テンプレートをOneDriveにアップロードする
n8n ワンショットワークフロー生成スクリプト
"""
import json

js_code = r"""
const fs = require('fs');
const year = 2026;
const month = 4;
const fileName = `経費管理_${year}年${String(month).padStart(2,'0')}月.xlsx`;
const filePath = `/tmp/${fileName}`;

if (!fs.existsSync(filePath)) {
  throw new Error(`ファイルが見つかりません: ${filePath}`);
}

const fileBuffer = fs.readFileSync(filePath);
const bin = await this.helpers.prepareBinaryData(
  fileBuffer,
  fileName,
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
);
return [{ json: { fileName, year, month }, binary: { data: bin } }];
"""

workflow = {
  "id": "bm-init-upload-001",
  "name": "経費管理テンプレート初期アップロード",
  "nodes": [
    {
      "parameters": {},
      "name": "手動実行",
      "type": "n8n-nodes-base.manualTrigger",
      "typeVersion": 1,
      "position": [240, 300],
      "id": "init-node-01"
    },
    {
      "parameters": {
        "jsCode": js_code.strip()
      },
      "name": "テンプレートファイル読み込み",
      "type": "n8n-nodes-base.code",
      "typeVersion": 2,
      "position": [480, 300],
      "id": "init-node-02"
    },
    {
      "parameters": {
        "method": "PUT",
        "url": "=https://graph.microsoft.com/v1.0/me/drive/root:/業務管理システム/経費管理/月次/{{ $json.fileName }}:/content",
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
      "id": "init-node-03"
    },
    {
      "parameters": {
        "jsCode": "const result = $input.first().json;\nconsole.log('アップロード完了:', JSON.stringify(result));\nreturn [{ json: { success: true, name: result.name, id: result.id } }];"
      },
      "name": "完了確認",
      "type": "n8n-nodes-base.code",
      "typeVersion": 2,
      "position": [960, 300],
      "id": "init-node-04"
    }
  ],
  "connections": {
    "手動実行": {
      "main": [[{"node": "テンプレートファイル読み込み", "type": "main", "index": 0}]]
    },
    "テンプレートファイル読み込み": {
      "main": [[{"node": "OneDriveアップロード", "type": "main", "index": 0}]]
    },
    "OneDriveアップロード": {
      "main": [[{"node": "完了確認", "type": "main", "index": 0}]]
    }
  },
  "active": False,
  "settings": {}
}

out_path = '/home/user/genai-lessons/workflows/business_management_init_upload.json'
with open(out_path, 'w', encoding='utf-8') as f:
    json.dump(workflow, f, ensure_ascii=False, indent=2)

print(f"✓ ワークフロー生成完了: {out_path}")
