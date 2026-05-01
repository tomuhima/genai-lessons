# 日報システム v3 セットアップガイド

n8n 2.8.4 対応版。タスクランナー制限を回避する設計。

---

## 設計方針

| 設計要素 | 理由 |
|---------|------|
| OneDrive操作はHTTP Requestノード | Codeノード内で `httpRequestWithAuthentication` が使えないため |
| Mergeノードでデータ合流 | HTTP Requestノードは入力データを破棄するため |
| 全Codeノードが `$input.first().json` のみ使用 | クロスノード参照（`$('NodeName')`）はタスクランナーで動かない |

---

## インポート手順

### 1. ワークフローのインポート

n8nエディタで：

1. 左メニュー → **Workflows**
2. 右上 **Add workflow** → **Import from File**
3. `workflows/daily_report_workflow_v3.json` を選択

### 2. 認証情報の設定

#### Microsoft Drive OAuth2

1. **Credentials** → **Add credential**
2. **Microsoft OneDrive OAuth2 API** を選択
3. Client ID / Client Secret を入力
4. **Connect my account** で認証
5. ワークフロー内の以下のノードに認証情報を設定：
   - ファイル情報取得
   - アップロード
   - 請求集計Excelファイル情報取得
   - 請求集計アップロード
   - 請求集計ファイル情報取得（当月）
   - 請求集計ファイル情報取得（前月）

#### LINE Channel Access Token

ワークフロー内の以下のCodeノードを開いて、`YOUR_LINE_CHANNEL_ACCESS_TOKEN` を実際のトークンに置き換え：

- LINE返信（成功）
- LINE返信（エラー）
- LINE返信（請求集計）
- 野添通知

#### Anthropic API Key

`Claude API呼び出し` ノードを開いて、`YOUR_ANTHROPIC_API_KEY` を実際のAPIキーに置き換え。

### 3. LINE userId の登録

`LINEメッセージ解析` ノードを開いて、`EMPLOYEE_MAP` の `REPLACE_*_USERID` を実際のLINE userIDに置き換え：

```javascript
"REPLACE_IMOTO_USERID":  "井本貴史",
"REPLACE_LLS_USERID":    "LLS電気",
"REPLACE_TRUST_USERID":  "株式会社トラストテクノス",
"REPLACE_RISE_USERID":   "株式会社RISE"
```

### 4. Webhook URLの確認

ワークフロー保存後、`Webhook` ノードのURLを取得（例: `https://n8n-light-m-n.xvps.jp/webhook/daily-report`）。

LINE Developers Console → Messaging API設定 → Webhook URL に設定。

---

## ワークフロー構成

### 入口
```
Webhook → Webhook応答(OK) / LINEメッセージ解析
```

### LINE解析〜分岐
```
LINEメッセージ解析 → バリデーション確認 → メッセージタイプ判定
                                              ├ 請求集計クエリ
                                              └ 日報
```

### 日報パス
```
Claude API呼び出し → Claude解析結果処理 → 解析バリデーション
  ├ True → ファイル情報取得 ─→ データ合流1 → Excel書き込み → 書き込みバリデーション
  │                            ↑                                    ├ True → アップロード ─→ データ合流2 → 成功返信生成 → LINE返信(成功)
  │                            └─────────(直接)                                 ↑
  │                                                                              └─────────(直接)
  │                                                                                                                                                ・並行: 請求集計書き込み準備 → 請求集計チェック → 請求集計Excel書き込み → 請求集計アップロード → 野添通知
  └ False → LINE返信(エラー)
```

### 請求集計クエリパス
```
請求集計準備 → 請求集計ファイル情報取得（当月） → データ合流（当月） → 現当月URL保存
            ↗                                  ↑                       ↓
            └────────(直接)──────────────────┘                       請求集計ファイル情報取得（前月） → データ合流（前月） → 前月URL保存
                                                                                                          ↑                       ↓
                                                                                                          └────(直接)            請求集計処理 → LINE返信(請求集計)
```

---

## ファイル/フォルダ構造（OneDrive）

```
業務管理システム/
├── 日報/
│   └── 稼働表_YYYY年MM月.xlsx
└── 請求集計/
    └── 請求集計_YYYY年MM月.xlsx
```

### 稼働表のシート構造

シート名 = 従業員名（例：「野添優」「丸田翔吾」）

ヘッダーは2行目まで使用、3行目から実データ。

#### 通常社員列（A〜M）
1. 日付（文字列）
2. 区分（昼勤/夜勤）
3. 会社名
4. 物件名
5. 開始時刻（HH:MM）
6. 終了時刻（HH:MM）
7. 休憩（分）
8. 稼働時間（H:MM）
9. 作業内容
10. 高速料金
11. 距離（km）
12. 駐車場
13. 資材費

#### MULTI_PERSON社員列（A〜O）
通常列の8番目までは同じ。9番目以降が異なる：

10. 人員数
11. 人員名
12. 高速料金
13. 距離
14. 駐車場
15. 資材費

---

## 主要ルール

### 締日

| 会社 | 締日 |
|-----|------|
| 平成システム、梶原通信 | 20日 |
| その他 | 末日 |

### MULTI_PERSON

```
['梶原通信', 'LLS電気', '株式会社トラストテクノス', '株式会社RISE']
```

### 請求集計対象会社

```
港振興業、トラストテクノス、平成システム、FGE、mtr、
オークコミュニケーション、ライズ、千里スカイハイツ管理組合、
菊次、ページ、G-STYLE、梶原通信
```

### 会社名略称

| 略称 | 正式名称 |
|------|---------|
| 平成 | 平成システム |
| トラスト | トラストテクノス |
| オーク | オークコミュニケーション |
| 千里 | 千里スカイハイツ管理組合 |
| 港 / 港振 / コウシン | 港振興業 |
| 梶原 | 梶原通信 |
| ジースタイル / Gスタイル | G-STYLE |

### ライズ vs RISE

- **ライズ（日本語）** = 元請（請求集計対象）
- **RISE（英語）** = 外注会社（経費管理側）

---

## トラブルシューティング

### ファイル情報取得でエラー

→ Microsoft Drive OAuth2の認証情報が設定されていない。Credentialsを再確認。

### Excel書き込みで「シートが見つかりません」

→ 該当社員のシートがOneDriveのExcelに存在しない。先に空シートを作成する。

### data.sites is not iterable

→ データ合流1の Combine By が `Position` になっているか確認。`Matching Fields` だと正しく合流しない。

### LINE返信が届かない

→ `YOUR_LINE_CHANNEL_ACCESS_TOKEN` を実際のトークンに置き換えていない。

### Claude APIエラー

→ `YOUR_ANTHROPIC_API_KEY` を実際のAPIキーに置き換えていない。

---

## サーバー設定（必須）

n8n 2.8.4 の `--disallow-code-generation-from-strings` フラグを除去する必要があります（exceljs使用のため）。

### override.conf

```bash
sudo tee /etc/systemd/system/n8n.service.d/override.conf << 'EOF'
[Service]
Environment=NODE_FUNCTION_ALLOW_EXTERNAL=exceljs
EOF
sudo systemctl daemon-reload && sudo systemctl restart n8n
```

### task-runner-process-js.js のパッチ

```bash
sed -i "s/'--disallow-code-generation-from-strings', '--disable-proto=delete'/'--disable-proto=delete'/" /usr/lib/node_modules/n8n/dist/task-runners/task-runner-process-js.js
sudo systemctl restart n8n
```

---

*作成日：2026年5月2日*
