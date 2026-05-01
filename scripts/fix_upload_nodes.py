#!/usr/bin/env python3
"""
Excel書き込み＋アップロード と 年次サマリー更新＋アップロード から
this.getCredentials() を除去し、HTTP Requestノードでアップロードする構成に変更。
また接続を更新して新しいノードを挿入する。
"""
import json

path = '/home/user/genai-lessons/workflows/business_management_workflow_v3.json'
with open(path, 'r', encoding='utf-8') as f:
    w = json.load(f)

# ===== 1. Excel書き込み＋アップロード: アップロード部分を削除してバイナリ出力に変更 =====
MONTHLY_UPLOAD_OLD = (
    "    // OneDriveアップロード\n"
    "    const creds = await this.getCredentials('microsoftDriveOAuth2Api');\n"
    "    const token = creds.oauthTokenData?.access_token;\n"
    "    if (!token) throw new Error('Microsoft認証トークンが取得できません。資格情報を再接続してください。');\n"
    "    const buf = await wb.xlsx.writeBuffer({ useSharedStrings: false });\n"
    "    const upUrl = fileInfo.id\n"
    "      ? `https://graph.microsoft.com/v1.0/me/drive/items/${fileInfo.id}/content`\n"
    "      : `https://graph.microsoft.com/v1.0/me/drive/root:/業務管理システム/経費管理/月次/${data.fileName}:/content`;\n"
    "    await this.helpers.httpRequest({\n"
    "      method:'PUT', url:upUrl,\n"
    "      headers:{ 'Authorization':`Bearer ${token}`, 'Content-Type':'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' },\n"
    "      body:Buffer.from(buf), returnFullResponse:false\n"
    "    });\n"
    "    return [{ json: { ...data, valid: true } }];"
)
MONTHLY_UPLOAD_NEW = (
    "    // バイナリ出力（OneDriveアップロード（月次）ノードがアップロード）\n"
    "    const buf = await wb.xlsx.writeBuffer({ useSharedStrings: false });\n"
    "    const bin = await this.helpers.prepareBinaryData(\n"
    "      Buffer.from(buf), data.fileName,\n"
    "      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'\n"
    "    );\n"
    "    return [{ json: { ...data, valid: true }, binary: { data: bin } }];"
)

for node in w['nodes']:
    if node['name'] == 'Excel書き込み＋アップロード':
        code = node['parameters']['jsCode']
        assert MONTHLY_UPLOAD_OLD in code, "月次アップロード置換対象が見つかりません"
        node['parameters']['jsCode'] = code.replace(MONTHLY_UPLOAD_OLD, MONTHLY_UPLOAD_NEW)
        print("✓ Excel書き込み＋アップロード: アップロード部分をバイナリ出力に変更")

# ===== 2. 年次サマリー更新＋アップロード: 同様の修正 =====
YEARLY_EARLY_RETURN_OLD = "    if (!dlUrl) return [{ json: { yearlyDone: false } }];"
YEARLY_EARLY_RETURN_NEW = "    if (!dlUrl) return [{ json: { ...data, yearlyDone: false, reason: '月次ファイルのダウンロードURLが取得できません' } }];"

YEARLY_UPLOAD_OLD = (
    "    // 年次ファイルをアップロード\n"
    "    const creds=await this.getCredentials('microsoftDriveOAuth2Api');\n"
    "    const token=creds.oauthTokenData?.access_token;\n"
    "    if(!token) return [{ json: { yearlyDone:false, reason:'token error' } }];\n"
    "    const buf=await yearlyWb.xlsx.writeBuffer({useSharedStrings:false});\n"
    "    const upUrl=yearlyFile.id\n"
    "      ?`https://graph.microsoft.com/v1.0/me/drive/items/${yearlyFile.id}/content`\n"
    "      :`https://graph.microsoft.com/v1.0/me/drive/root:/業務管理システム/経費管理/年次/年間サマリー_${data.targetYear}年.xlsx:/content`;\n"
    "    await this.helpers.httpRequest({\n"
    "      method:'PUT', url:upUrl,\n"
    "      headers:{'Authorization':`Bearer ${token}`,'Content-Type':'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'},\n"
    "      body:Buffer.from(buf), returnFullResponse:false\n"
    "    });\n"
    "    return [{ json: { yearlyDone:true } }];\n"
    "  } catch(e) {\n"
    "    return [{ json: { yearlyDone:false, reason:e.message } }];\n"
    "  }\n"
    "})();"
)
YEARLY_UPLOAD_NEW = (
    "    // バイナリ出力（OneDriveアップロード（年次）ノードがアップロード）\n"
    "    const buf=await yearlyWb.xlsx.writeBuffer({useSharedStrings:false});\n"
    "    const yfn=yearlyFile.name||`年間サマリー_${data.targetYear}年.xlsx`;\n"
    "    const bin=await this.helpers.prepareBinaryData(\n"
    "      Buffer.from(buf), yfn,\n"
    "      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'\n"
    "    );\n"
    "    return [{ json: { ...data, yearlyDone:true }, binary: { data: bin } }];\n"
    "  } catch(e) {\n"
    "    return [{ json: { ...data, yearlyDone:false, reason:e.message } }];\n"
    "  }\n"
    "})();"
)

for node in w['nodes']:
    if node['name'] == '年次サマリー更新＋アップロード':
        code = node['parameters']['jsCode']
        assert YEARLY_EARLY_RETURN_OLD in code, "年次early return置換対象が見つかりません"
        assert YEARLY_UPLOAD_OLD in code, "年次アップロード置換対象が見つかりません"
        code = code.replace(YEARLY_EARLY_RETURN_OLD, YEARLY_EARLY_RETURN_NEW)
        code = code.replace(YEARLY_UPLOAD_OLD, YEARLY_UPLOAD_NEW)
        node['parameters']['jsCode'] = code
        print("✓ 年次サマリー更新＋アップロード: アップロード部分をバイナリ出力に変更")

# ===== 3. 新規ノード追加 =====
new_nodes = [
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
        "name": "OneDriveアップロード（月次）",
        "type": "n8n-nodes-base.httpRequest",
        "typeVersion": 4,
        "position": [1340, 500],
        "id": "bm-v3-upload-monthly",
        "onError": "continueRegularOutput"
    },
    {
        "parameters": {
            "jsCode": (
                "const uploadResult = $input.first().json;\n"
                "const excelResult = $('Excel書き込み＋アップロード').first().json;\n"
                "if (!excelResult.valid) return [{ json: excelResult }];\n"
                "if (uploadResult.id) return [{ json: { ...excelResult, valid: true } }];\n"
                "const errMsg = (uploadResult.error && uploadResult.error.message) || JSON.stringify(uploadResult).slice(0, 100);\n"
                "return [{ json: { valid: false, replyToken: excelResult.replyToken, errorMessage: 'OneDriveアップロード失敗: ' + errMsg } }];"
            )
        },
        "name": "月次結果集約",
        "type": "n8n-nodes-base.code",
        "typeVersion": 2,
        "position": [1580, 500],
        "id": "bm-v3-monthly-agg"
    },
    {
        "parameters": {
            "method": "PUT",
            "url": "={{ $('年次ファイル情報取得').first().json.id ? 'https://graph.microsoft.com/v1.0/me/drive/items/' + $('年次ファイル情報取得').first().json.id + '/content' : 'https://graph.microsoft.com/v1.0/me/drive/root:/業務管理システム/経費管理/年次/年間サマリー_' + $json.targetYear + '年.xlsx:/content' }}",
            "authentication": "predefinedCredentialType",
            "nodeCredentialType": "microsoftDriveOAuth2Api",
            "sendBody": True,
            "contentType": "binaryData",
            "inputDataFieldName": "data",
            "options": {}
        },
        "name": "OneDriveアップロード（年次）",
        "type": "n8n-nodes-base.httpRequest",
        "typeVersion": 4,
        "position": [1820, 300],
        "id": "bm-v3-upload-yearly",
        "onError": "continueRegularOutput"
    },
    {
        "parameters": {
            "jsCode": (
                "// 年次サマリー更新＋アップロードの結果をそのまま通す（messageType/replyToken含む）\n"
                "return [{ json: $('年次サマリー更新＋アップロード').first().json }];"
            )
        },
        "name": "年次結果集約",
        "type": "n8n-nodes-base.code",
        "typeVersion": 2,
        "position": [2060, 300],
        "id": "bm-v3-yearly-agg"
    }
]

w['nodes'].extend(new_nodes)
print("✓ 新規ノード4つを追加: OneDriveアップロード（月次）, 月次結果集約, OneDriveアップロード（年次）, 年次結果集約")

# ===== 4. 接続の更新 =====
conn = w['connections']

# 月次: Excel書き込み＋アップロード → 書込み結果確認  を
#       Excel書き込み＋アップロード → OneDriveアップロード（月次）→ 月次結果集約 → 書込み結果確認 に変更
conn['Excel書き込み＋アップロード'] = {
    "main": [[{"node": "OneDriveアップロード（月次）", "type": "main", "index": 0}]]
}
conn['OneDriveアップロード（月次）'] = {
    "main": [[{"node": "月次結果集約", "type": "main", "index": 0}]]
}
conn['月次結果集約'] = {
    "main": [[{"node": "書込み結果確認", "type": "main", "index": 0}]]
}

# 年次: 年次サマリー更新＋アップロード → 成功メッセージ生成  を
#       年次サマリー更新＋アップロード → OneDriveアップロード（年次）→ 年次結果集約 → 成功メッセージ生成 に変更
conn['年次サマリー更新＋アップロード'] = {
    "main": [[{"node": "OneDriveアップロード（年次）", "type": "main", "index": 0}]]
}
conn['OneDriveアップロード（年次）'] = {
    "main": [[{"node": "年次結果集約", "type": "main", "index": 0}]]
}
conn['年次結果集約'] = {
    "main": [[{"node": "成功メッセージ生成", "type": "main", "index": 0}]]
}

print("✓ 接続を更新")

# ===== 5. 保存 =====
with open(path, 'w', encoding='utf-8') as f:
    json.dump(w, f, ensure_ascii=False, indent=2)

print(f"\n✓ 保存完了: {path}")
print("\n変更内容:")
print("  月次フロー: Excel書き込み→OneDriveアップロード（月次）→月次結果集約→書込み結果確認")
print("  年次フロー: 年次サマリー更新→OneDriveアップロード（年次）→年次結果集約→成功メッセージ生成")
