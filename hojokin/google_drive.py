# -*- coding: utf-8 -*-
"""Google Drive API クライアント"""
from __future__ import annotations

import io
import logging
from pathlib import Path

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaFileUpload

logger = logging.getLogger(__name__)

SCOPES = ['https://www.googleapis.com/auth/drive']


class DriveClient:
    """Google Drive操作"""

    def __init__(self, credentials_path: str):
        creds = service_account.Credentials.from_service_account_file(
            credentials_path, scopes=SCOPES
        )
        self.service = build('drive', 'v3', credentials=creds)
        logger.info('Google Drive API 接続完了')

    def list_files(self, folder_id: str) -> list[dict]:
        """フォルダ内のファイル一覧を取得"""
        results = []
        page_token = None

        while True:
            response = self.service.files().list(
                q=f"'{folder_id}' in parents and trashed = false",
                fields='nextPageToken, files(id, name, mimeType, size, modifiedTime)',
                pageToken=page_token,
                pageSize=100,
            ).execute()

            results.extend(response.get('files', []))
            page_token = response.get('nextPageToken')
            if not page_token:
                break

        logger.info(f'フォルダ {folder_id}: {len(results)}件')
        return results

    def list_files_recursive(self, folder_id: str, prefix: str = '') -> list[dict]:
        """フォルダ内を再帰的に一覧取得"""
        files = self.list_files(folder_id)
        result = []

        for f in files:
            f['path'] = f'{prefix}{f["name"]}'
            if f['mimeType'] == 'application/vnd.google-apps.folder':
                result.extend(self.list_files_recursive(f['id'], f'{f["path"]}/'))
            else:
                result.append(f)

        return result

    def download_file(self, file_id: str, dest_path: Path) -> Path:
        """ファイルをダウンロード"""
        dest_path.parent.mkdir(parents=True, exist_ok=True)
        request = self.service.files().get_media(fileId=file_id)

        with open(dest_path, 'wb') as fh:
            downloader = MediaIoBaseDownload(fh, request)
            done = False
            while not done:
                _, done = downloader.next_chunk()

        logger.info(f'ダウンロード: {dest_path.name} ({dest_path.stat().st_size:,} bytes)')
        return dest_path

    def download_folder(self, folder_id: str, dest_dir: Path) -> list[Path]:
        """フォルダ内の全ファイルをダウンロード"""
        files = self.list_files_recursive(folder_id)
        downloaded = []

        for f in files:
            dest = dest_dir / f['path']
            self.download_file(f['id'], dest)
            downloaded.append(dest)

        return downloaded

    def upload_file(self, local_path: Path, folder_id: str, filename: str = None) -> str:
        """ファイルをアップロード。ファイルIDを返す"""
        name = filename or local_path.name
        mime_map = {
            '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            '.pdf': 'application/pdf',
            '.csv': 'text/csv',
        }
        mime_type = mime_map.get(local_path.suffix, 'application/octet-stream')

        file_metadata = {
            'name': name,
            'parents': [folder_id],
        }
        media = MediaFileUpload(str(local_path), mimetype=mime_type)
        file = self.service.files().create(
            body=file_metadata, media_body=media, fields='id'
        ).execute()

        file_id = file.get('id')
        logger.info(f'アップロード: {name} → {file_id}')
        return file_id

    def create_folder(self, name: str, parent_id: str = None) -> str:
        """フォルダを作成。フォルダIDを返す"""
        metadata = {
            'name': name,
            'mimeType': 'application/vnd.google-apps.folder',
        }
        if parent_id:
            metadata['parents'] = [parent_id]

        folder = self.service.files().create(
            body=metadata, fields='id'
        ).execute()

        folder_id = folder.get('id')
        logger.info(f'フォルダ作成: {name} → {folder_id}')
        return folder_id

    def find_or_create_subfolder(self, parent_id: str, name: str) -> str:
        """サブフォルダを検索し、なければ作成"""
        results = self.service.files().list(
            q=f"'{parent_id}' in parents and name = '{name}' and mimeType = 'application/vnd.google-apps.folder' and trashed = false",
            fields='files(id)',
        ).execute()

        files = results.get('files', [])
        if files:
            return files[0]['id']
        return self.create_folder(name, parent_id)
