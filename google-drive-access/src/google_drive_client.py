"""Google Drive / Docs / Sheets / Slides クライアント（AI環境用）"""

import io
import os
from pathlib import Path

import yaml
from dotenv import load_dotenv
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload, MediaIoBaseDownload

load_dotenv()

CONFIG_PATH = Path(__file__).parent.parent / "config.yaml"


class GoogleDriveClient:
    """Google Drive APIを使ってDocs/Sheets/Slides/ファイルを操作するクライアント"""

    def __init__(self, config_path=None, delegated_user=None):
        self._config = self._load_config(config_path or CONFIG_PATH)
        self._delegated_user = delegated_user or os.getenv("GOOGLE_DELEGATED_USER")
        self._credentials = self._build_credentials()
        self._drive_service = build("drive", "v3", credentials=self._credentials)
        self._docs_service = build("docs", "v1", credentials=self._credentials)
        self._sheets_service = build("sheets", "v4", credentials=self._credentials)
        self._slides_service = build("slides", "v1", credentials=self._credentials)

    def _load_config(self, config_path):
        with open(config_path, encoding="utf-8") as f:
            return yaml.safe_load(f)

    def _build_credentials(self):
        creds_path = os.getenv(
            "GOOGLE_APPLICATION_CREDENTIALS", "credentials/service-account.json"
        )
        scopes = self._config.get("scopes", [])
        creds = service_account.Credentials.from_service_account_file(
            creds_path, scopes=scopes
        )
        if self._delegated_user:
            creds = creds.with_subject(self._delegated_user)
        return creds

    # --- Google Docs ---

    def read_document(self, document_id):
        """Google Docsの内容を読み取る"""
        doc = self._docs_service.documents().get(documentId=document_id).execute()
        content = []
        for element in doc.get("body", {}).get("content", []):
            paragraph = element.get("paragraph")
            if paragraph:
                for elem in paragraph.get("elements", []):
                    text_run = elem.get("textRun")
                    if text_run:
                        content.append(text_run.get("content", ""))
        return "".join(content)

    def write_to_document(self, document_id, text, index=None):
        """Google Docsにテキストを追記する"""
        if index is None:
            doc = self._docs_service.documents().get(documentId=document_id).execute()
            body_content = doc.get("body", {}).get("content", [])
            index = body_content[-1]["endIndex"] - 1 if body_content else 1

        requests = [{"insertText": {"location": {"index": index}, "text": text}}]
        self._docs_service.documents().batchUpdate(
            documentId=document_id, body={"requests": requests}
        ).execute()

    def create_document(self, title):
        """新しいGoogle Docsを作成する"""
        doc = self._docs_service.documents().create(body={"title": title}).execute()
        return {"id": doc["documentId"], "title": doc["title"]}

    # --- Google Sheets ---

    def read_spreadsheet(self, spreadsheet_id, range="Sheet1"):
        """Google Sheetsのデータを読み取る"""
        result = (
            self._sheets_service.spreadsheets()
            .values()
            .get(spreadsheetId=spreadsheet_id, range=range)
            .execute()
        )
        return result.get("values", [])

    def write_to_spreadsheet(self, spreadsheet_id, range, values):
        """Google Sheetsにデータを書き込む"""
        body = {"values": values}
        self._sheets_service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id,
            range=range,
            valueInputOption="USER_ENTERED",
            body=body,
        ).execute()

    def create_spreadsheet(self, title):
        """新しいGoogle Sheetsを作成する"""
        spreadsheet = (
            self._sheets_service.spreadsheets()
            .create(body={"properties": {"title": title}})
            .execute()
        )
        return {
            "id": spreadsheet["spreadsheetId"],
            "title": spreadsheet["properties"]["title"],
        }

    # --- Google Drive ファイル管理 ---

    def list_files(self, folder_id=None, page_size=100):
        """Google Driveのファイル一覧を取得する"""
        query = f"'{folder_id}' in parents" if folder_id else None
        results = (
            self._drive_service.files()
            .list(
                q=query,
                pageSize=page_size,
                fields="files(id, name, mimeType, modifiedTime)",
            )
            .execute()
        )
        return results.get("files", [])

    def upload_file(self, local_path, folder_id=None, convert=None):
        """ファイルをGoogle Driveにアップロードする"""
        if convert is None:
            convert = self._config.get("conversion", {}).get(
                "auto_convert_on_upload", True
            )

        file_metadata = {"name": Path(local_path).name}
        if folder_id:
            file_metadata["parents"] = [folder_id]

        # Excel/Word/PowerPointをGoogle形式に変換
        if convert:
            mime_map = {
                ".xlsx": "application/vnd.google-apps.spreadsheet",
                ".xls": "application/vnd.google-apps.spreadsheet",
                ".csv": "application/vnd.google-apps.spreadsheet",
                ".docx": "application/vnd.google-apps.document",
                ".doc": "application/vnd.google-apps.document",
                ".pptx": "application/vnd.google-apps.presentation",
                ".ppt": "application/vnd.google-apps.presentation",
            }
            ext = Path(local_path).suffix.lower()
            if ext in mime_map:
                file_metadata["mimeType"] = mime_map[ext]

        media = MediaFileUpload(local_path)
        file = (
            self._drive_service.files()
            .create(body=file_metadata, media_body=media, fields="id, name")
            .execute()
        )
        return {"id": file["id"], "name": file["name"]}

    def download_file(self, file_id, output_path):
        """Google Driveからファイルをダウンロードする"""
        request = self._drive_service.files().get_media(fileId=file_id)
        fh = io.BytesIO()
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while not done:
            _, done = downloader.next_chunk()

        with open(output_path, "wb") as f:
            f.write(fh.getvalue())

    def export_as(self, file_id, mime_type, output_path):
        """Google形式のファイルを指定フォーマットでエクスポートする"""
        request = self._drive_service.files().export_media(
            fileId=file_id, mimeType=mime_type
        )
        fh = io.BytesIO()
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while not done:
            _, done = downloader.next_chunk()

        with open(output_path, "wb") as f:
            f.write(fh.getvalue())
