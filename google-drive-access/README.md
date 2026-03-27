# Google Drive AI環境アクセス設定

AI環境（.ai）からGoogle Drive内のファイル（Docs / Sheets / Slides 等）にアクセス・書き込みするためのセットアップガイドです。

## 対応サービス

| サービス | 読み取り | 書き込み | 作成 |
|----------|----------|----------|------|
| Google Docs | o | o | o |
| Google Sheets（Excel互換） | o | o | o |
| Google Slides | o | o | o |
| Google Drive（ファイル管理） | o | o | o |

## 前提条件

- Google Workspaceの管理者権限
- Google Cloud Projectへのアクセス
- Python 3.8以上

## セットアップ手順

### 1. Google Cloud ProjectでAPIを有効化

[Google Cloud Console](https://console.cloud.google.com/) で以下のAPIを有効化:

- **Google Drive API**
- **Google Docs API**
- **Google Sheets API**
- **Google Slides API**

### 2. サービスアカウントの作成

1. 「APIとサービス」>「認証情報」>「認証情報を作成」
2. 「サービスアカウント」を選択
3. 名前を入力（例: `ai-drive-access`）
4. 役割:「編集者」を付与
5. JSONキーを作成・ダウンロード
6. ダウンロードしたキーを `credentials/service-account.json` に配置

### 3. Google Workspaceドメイン全体の委任

メンバー全員がAI環境からアクセスできるようにするには:

1. Google Workspace管理コンソール > 「セキュリティ」>「APIの制御」>「ドメイン全体の委任」
2. サービスアカウントのクライアントIDを登録
3. 以下のスコープを追加:

```
https://www.googleapis.com/auth/drive
https://www.googleapis.com/auth/documents
https://www.googleapis.com/auth/spreadsheets
https://www.googleapis.com/auth/presentations
```

### 4. 環境変数の設定

```bash
cp .env.example .env
# .env ファイルを編集し、必要な値を設定
```

### 5. 依存パッケージのインストール

```bash
pip install -r requirements.txt
```

### 6. 動作確認

```bash
python scripts/verify_access.py
```

## 使い方

### Google Docsの操作

```python
from src.google_drive_client import GoogleDriveClient

client = GoogleDriveClient()

# ドキュメント読み取り
content = client.read_document("DOCUMENT_ID")

# ドキュメント書き込み
client.write_to_document("DOCUMENT_ID", "追記するテキスト")

# 新規ドキュメント作成
doc = client.create_document("タイトル")
```

### Google Sheets（Excel互換）の操作

```python
client = GoogleDriveClient()

# スプレッドシート読み取り
data = client.read_spreadsheet("SPREADSHEET_ID", range="Sheet1!A1:D10")

# スプレッドシート書き込み
client.write_to_spreadsheet("SPREADSHEET_ID", range="Sheet1!A1", values=[
    ["名前", "部署", "メール"],
    ["田中", "開発部", "tanaka@example.com"],
])

# 新規スプレッドシート作成
sheet = client.create_spreadsheet("売上レポート")
```

### Google Driveファイル管理

```python
client = GoogleDriveClient()

# ファイル一覧取得
files = client.list_files(folder_id="FOLDER_ID")

# ファイルアップロード
client.upload_file("local_file.xlsx", folder_id="FOLDER_ID")

# ファイルダウンロード
client.download_file("FILE_ID", "output_path.xlsx")

# Excel(.xlsx)をGoogle Sheetsに変換
client.convert_to_google_format("FILE_ID")
```

## ディレクトリ構成

```
google-drive-access/
├── README.md                # このファイル
├── .env.example             # 環境変数テンプレート
├── .gitignore               # Git除外設定
├── requirements.txt         # Python依存パッケージ
├── config.yaml              # アクセス設定
├── credentials/             # 認証情報（Git管理外）
│   └── .gitkeep
├── scripts/
│   └── verify_access.py     # 動作確認スクリプト
└── src/
    └── google_drive_client.py  # クライアントライブラリ
```

## トラブルシューティング

| 問題 | 対処法 |
|------|--------|
| 認証エラー (403) | サービスアカウントの権限とドメイン委任設定を確認 |
| API未有効エラー | Google Cloud Consoleで各APIが有効か確認 |
| ファイルが見つからない | ファイルIDが正しいか、共有設定を確認 |
| Excel変換エラー | Drive APIが有効か確認、ファイル形式を確認 |

## セキュリティに関する注意

- `credentials/` ディレクトリはGitにコミットしないこと（`.gitignore`に設定済み）
- `.env` ファイルもGitにコミットしないこと
- サービスアカウントキーは定期的にローテーションすること
- 最小権限の原則に従い、必要なスコープのみ付与すること
