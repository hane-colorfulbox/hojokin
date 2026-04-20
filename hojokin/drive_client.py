# -*- coding: utf-8 -*-
"""
Google Drive 連携モジュール

サービスアカウント経由でDriveフォルダからファイル一覧取得・ダウンロードを行う。
"""
from __future__ import annotations

import io
import logging
from pathlib import Path

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

logger = logging.getLogger(__name__)

SCOPES = ['https://www.googleapis.com/auth/drive.readonly']

# Google Docs Editors 形式 → Office 形式へのエクスポートマッピング
GOOGLE_EXPORT_MAP = {
    'application/vnd.google-apps.spreadsheet': (
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', '.xlsx',
    ),
    'application/vnd.google-apps.document': (
        'application/vnd.openxmlformats-officedocument.wordprocessingml.document', '.docx',
    ),
    'application/vnd.google-apps.presentation': (
        'application/vnd.openxmlformats-officedocument.presentationml.presentation', '.pptx',
    ),
}


class DriveClient:
    """Google Drive 読み取り専用クライアント"""

    def __init__(self, credentials_path: str | Path | None = None, credentials_dict: dict | None = None):
        """
        Args:
            credentials_path: サービスアカウントJSONファイルのパス
            credentials_dict: サービスアカウント情報のdict（Streamlit Secrets用）
        """
        if credentials_dict:
            creds = service_account.Credentials.from_service_account_info(
                credentials_dict, scopes=SCOPES,
            )
        elif credentials_path:
            creds = service_account.Credentials.from_service_account_file(
                str(credentials_path), scopes=SCOPES,
            )
        else:
            raise ValueError('credentials_path or credentials_dict is required')
        self.service = build('drive', 'v3', credentials=creds)
        logger.info('Drive接続完了')

    def list_folders(self, parent_id: str) -> list[dict]:
        """親フォルダ直下のサブフォルダ一覧を取得"""
        query = (
            f"'{parent_id}' in parents "
            "and mimeType='application/vnd.google-apps.folder' "
            "and trashed=false"
        )
        results = self.service.files().list(
            q=query,
            fields='files(id, name)',
            orderBy='name',
            pageSize=100,
            supportsAllDrives=True,
            includeItemsFromAllDrives=True,
        ).execute()
        folders = results.get('files', [])
        logger.info(f'フォルダ一覧: {len(folders)}件')
        return folders

    def list_files(self, folder_id: str, file_type: str | None = None) -> list[dict]:
        """
        フォルダ内のファイル一覧を取得

        Args:
            folder_id: DriveフォルダID
            file_type: フィルタ（'xlsx', 'pdf' 等）。Noneで全ファイル。
        """
        query = (
            f"'{folder_id}' in parents "
            "and mimeType!='application/vnd.google-apps.folder' "
            "and trashed=false"
        )

        mime_map = {
            'xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            'xls': 'application/vnd.ms-excel',
            'pdf': 'application/pdf',
        }
        if file_type and file_type in mime_map:
            query += f" and mimeType='{mime_map[file_type]}'"

        results = self.service.files().list(
            q=query,
            fields='files(id, name, mimeType, modifiedTime, size)',
            orderBy='modifiedTime desc',
            pageSize=100,
            supportsAllDrives=True,
            includeItemsFromAllDrives=True,
        ).execute()
        files = results.get('files', [])
        logger.info(f'ファイル一覧({folder_id}): {len(files)}件')
        return files

    def list_files_recursive(self, folder_id: str, file_type: str | None = None) -> list[dict]:
        """サブフォルダも含めて再帰的にファイルを検索"""
        all_files = self.list_files(folder_id, file_type)

        subfolders = self.list_folders(folder_id)
        for folder in subfolders:
            sub_files = self.list_files_recursive(folder['id'], file_type)
            for f in sub_files:
                f['folder_name'] = folder['name']
            all_files.extend(sub_files)

        return all_files

    def _build_download_request(self, file_id: str, mime_type: str | None):
        """mimeTypeに応じて get_media / export_media のリクエストを返す。

        Returns:
            (request, export_ext): Google形式なら補正すべき拡張子、それ以外はNone
        """
        if mime_type is None:
            meta = self.service.files().get(
                fileId=file_id, fields='mimeType',
                supportsAllDrives=True,
            ).execute()
            mime_type = meta.get('mimeType', '')

        if mime_type in GOOGLE_EXPORT_MAP:
            export_mime, ext = GOOGLE_EXPORT_MAP[mime_type]
            request = self.service.files().export_media(
                fileId=file_id, mimeType=export_mime,
            )
            return request, ext

        if mime_type.startswith('application/vnd.google-apps.'):
            raise ValueError(f'未対応のGoogle形式ファイルです: {mime_type}')

        request = self.service.files().get_media(
            fileId=file_id, supportsAllDrives=True,
        )
        return request, None

    def download_file(
        self, file_id: str, dest_path: str | Path, mime_type: str | None = None,
    ) -> Path:
        """ファイルをダウンロード。Google形式は自動でOffice形式にエクスポート。"""
        dest = Path(dest_path)
        dest.parent.mkdir(parents=True, exist_ok=True)

        request, export_ext = self._build_download_request(file_id, mime_type)
        if export_ext and dest.suffix.lower() != export_ext:
            dest = dest.with_suffix(export_ext)

        with open(dest, 'wb') as f:
            downloader = MediaIoBaseDownload(f, request)
            done = False
            while not done:
                _, done = downloader.next_chunk()

        logger.info(f'ダウンロード完了: {dest}')
        return dest

    def download_to_bytes(self, file_id: str, mime_type: str | None = None) -> bytes:
        """ファイルをバイト列としてダウンロード（一時ファイル不要）"""
        request, _ = self._build_download_request(file_id, mime_type)
        buffer = io.BytesIO()
        downloader = MediaIoBaseDownload(buffer, request)
        done = False
        while not done:
            _, done = downloader.next_chunk()
        return buffer.getvalue()

    def find_customer_folder(self, parent_id: str, customer_name: str) -> dict | None:
        """顧客名でフォルダを検索（部分一致）"""
        folders = self.list_folders(parent_id)
        for folder in folders:
            if customer_name in folder['name']:
                return folder
        return None
