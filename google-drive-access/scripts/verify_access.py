"""Google Drive APIへのアクセスを確認するスクリプト"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

from src.google_drive_client import GoogleDriveClient


def main():
    print("Google Drive APIアクセス確認を開始します...\n")

    try:
        client = GoogleDriveClient()
        print("[OK] 認証成功")
    except Exception as e:
        print(f"[NG] 認証失敗: {e}")
        sys.exit(1)

    # Drive API確認
    try:
        files = client.list_files(page_size=5)
        print(f"[OK] Drive API - {len(files)}件のファイルを取得")
        for f in files:
            print(f"     - {f['name']} ({f['mimeType']})")
    except Exception as e:
        print(f"[NG] Drive API: {e}")

    # Docs API確認
    try:
        doc = client.create_document("アクセス確認テスト")
        print(f"[OK] Docs API - テストドキュメント作成成功 (ID: {doc['id']})")
    except Exception as e:
        print(f"[NG] Docs API: {e}")

    # Sheets API確認
    try:
        sheet = client.create_spreadsheet("アクセス確認テスト")
        print(f"[OK] Sheets API - テストスプレッドシート作成成功 (ID: {sheet['id']})")
    except Exception as e:
        print(f"[NG] Sheets API: {e}")

    print("\n確認完了!")


if __name__ == "__main__":
    main()
