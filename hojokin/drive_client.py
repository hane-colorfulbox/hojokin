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
from googleapiclient.http import MediaIoBaseDownload, MediaFileUpload

from .warnings_catalog import tag as warn_tag

logger = logging.getLogger(__name__)

SCOPES = ['https://www.googleapis.com/auth/drive']  # アップロード機能のため read/write 必要


class GoogleFormatNotSupportedError(ValueError):
    """対応外のGoogle形式ファイル (フォーム / 図面 / サイト / スクリプト等)

    呼出側で個別にスキップしてダウンロード処理を継続するために使う専用例外。
    ValueError を継承しているので、明示 catch しなければ従来通り例外として伝播する。
    """
    pass

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
        # エラー案内用（権限不足時に「どのアカウントに権限を付けるか」を提示する）
        self.service_account_email = getattr(creds, 'service_account_email', '')
        self.service = build('drive', 'v3', credentials=creds)
        logger.info('Drive接続完了')

    def can_write(self, folder_id: str) -> bool | None:
        """フォルダにファイルを追加できる権限があるか（処理前の事前チェック用）。

        Returns:
            True/False: capabilities.canAddChildren の値
            None: 判定不能（API エラー等）。呼出側は警告を出さずに続行してよい
        """
        try:
            meta = self.service.files().get(
                fileId=folder_id,
                fields='capabilities(canAddChildren)',
                supportsAllDrives=True,
            ).execute()
            return bool(meta.get('capabilities', {}).get('canAddChildren'))
        except Exception as e:
            logger.warning(f'書込権限の事前チェックに失敗（続行）: {e}')
            return None

    def upload_error_hint(self, exc: Exception) -> str:
        """アップロード失敗例外を、担当者が対処できる日本語メッセージに変換する。"""
        from googleapiclient.errors import HttpError

        sa = self.service_account_email or 'サービスアカウント'
        if isinstance(exc, HttpError):
            reason = ''
            try:
                err = exc.error_details[0] if exc.error_details else {}
                reason = err.get('reason', '') if isinstance(err, dict) else ''
            except Exception:
                pass
            status = exc.resp.status if exc.resp is not None else 0
            if reason == 'storageQuotaExceeded':
                return (
                    f'サービスアカウント（{sa}）は個人のマイドライブ配下に'
                    'ファイルを作成できません（容量0のため）。格納先フォルダを'
                    '共有ドライブ配下に移すか、管理者にご相談ください。'
                )
            if status == 403:
                return (
                    f'{warn_tag("W-DRIVE-001")}サービスアカウント（{sa}）にこのフォルダへの書き込み権限が'
                    'ありません。共有ドライブの管理者に、このアカウントを'
                    '「コンテンツ管理者」以上のメンバーとして追加してもらってください。'
                )
            if status == 404:
                return (
                    'アップロード先フォルダが見つかりません（削除・移動された'
                    '可能性）。フォルダを選び直してください。'
                )
        return f'{exc}'

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

    def list_files_recursive(
        self,
        folder_id: str,
        file_type: str | None = None,
        exclude_folder_names: set[str] | None = None,
    ) -> list[dict]:
        """サブフォルダも含めて再帰的にファイルを検索

        Args:
            folder_id: 起点フォルダID
            file_type: 拡張子フィルタ（'xlsx', 'pdf'等）。Noneで全ファイル。
            exclude_folder_names: 配下に降りないサブフォルダ名の集合。
                例: {'申請時使用'} を渡すと、その名前のサブフォルダは丸ごとスキップ。
                呼出し元（補助金 app.py 等）が用途別に除外ルールを管理する想定。
                drive_client.py 自体は汎用に保ち、業務固有ルールは外部から注入。
        """
        all_files = self.list_files(folder_id, file_type)

        subfolders = self.list_folders(folder_id)
        for folder in subfolders:
            if exclude_folder_names and folder['name'] in exclude_folder_names:
                logger.info(
                    f'サブフォルダ除外: {folder["name"]} '
                    f'(exclude_folder_names 指定)'
                )
                continue
            sub_files = self.list_files_recursive(
                folder['id'], file_type, exclude_folder_names
            )
            for f in sub_files:
                f['folder_name'] = folder['name']
            all_files.extend(sub_files)

        return all_files

    def _build_download_request(self, file_id: str, mime_type: str | None):
        """mimeTypeに応じて get_media / export_media のリクエストを返す。

        Returns:
            (request, export_ext): Google形式なら補正すべき拡張子、それ以外はNone
        Raises:
            GoogleFormatNotSupportedError: 対応外のGoogle形式（form, drawing, site等）
        """
        if mime_type is None:
            meta = self.service.files().get(
                fileId=file_id, fields='mimeType',
                supportsAllDrives=True,
            ).execute()
            mime_type = meta.get('mimeType', '')

        # ショートカットの解決：リンク先のファイルをたどる
        if mime_type == 'application/vnd.google-apps.shortcut':
            meta = self.service.files().get(
                fileId=file_id,
                fields='shortcutDetails',
                supportsAllDrives=True,
            ).execute()
            shortcut = meta.get('shortcutDetails') or {}
            target_id = shortcut.get('targetId')
            target_mime = shortcut.get('targetMimeType')
            if target_id:
                logger.info(
                    f'ショートカット解決: {file_id} → {target_id} '
                    f'(mime={target_mime})'
                )
                return self._build_download_request(target_id, target_mime)
            raise GoogleFormatNotSupportedError(
                f'ショートカットのリンク先が解決できません (id={file_id})'
            )

        if mime_type in GOOGLE_EXPORT_MAP:
            export_mime, ext = GOOGLE_EXPORT_MAP[mime_type]
            request = self.service.files().export_media(
                fileId=file_id, mimeType=export_mime,
            )
            return request, ext

        if mime_type.startswith('application/vnd.google-apps.'):
            raise GoogleFormatNotSupportedError(
                f'未対応のGoogle形式ファイルです: {mime_type}'
            )

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

    def upload_file(
        self,
        local_path: Path,
        parent_folder_id: str,
        mime_type: str | None = None,
        overwrite: bool = True,
    ) -> dict:
        """ローカルファイルを Drive にアップロード。

        同名ファイルが既にあれば overwrite=True なら上書き、False なら新規追加。
        サービスアカウント運用前提（共有ドライブまたは「編集者」権限が必要）。

        Args:
            local_path: アップロード元のローカルパス
            parent_folder_id: アップロード先 Drive フォルダ ID
            mime_type: 明示指定時のみ使う。None なら拡張子から自動推定
            overwrite: True なら同名ファイルを update、False なら create

        Returns:
            アップロード結果（id, name, webViewLink を含む dict）
        """
        local_path = Path(local_path)
        if not local_path.exists():
            raise FileNotFoundError(f'アップロード元ファイルが存在しません: {local_path}')

        if mime_type is None:
            ext = local_path.suffix.lower()
            mime_type = {
                '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                '.xls':  'application/vnd.ms-excel',
                '.csv':  'text/csv',
                '.pdf':  'application/pdf',
            }.get(ext, 'application/octet-stream')

        existing_id: str | None = None
        if overwrite:
            # 同名ファイル検索（ゴミ箱外、同フォルダ内）
            safe_name = local_path.name.replace("'", "\\'")
            query = (
                f"name='{safe_name}' "
                f"and '{parent_folder_id}' in parents "
                f"and trashed=false"
            )
            res = self.service.files().list(
                q=query, fields='files(id, name)', pageSize=1,
                supportsAllDrives=True, includeItemsFromAllDrives=True,
            ).execute()
            files = res.get('files', [])
            if files:
                existing_id = files[0]['id']

        media = MediaFileUpload(str(local_path), mimetype=mime_type, resumable=False)

        if existing_id:
            updated = self.service.files().update(
                fileId=existing_id,
                media_body=media,
                fields='id, name, webViewLink',
                supportsAllDrives=True,
            ).execute()
            logger.info(f'Drive更新: {updated.get("name")} (id={updated.get("id")})')
            return updated

        metadata = {'name': local_path.name, 'parents': [parent_folder_id]}
        created = self.service.files().create(
            body=metadata,
            media_body=media,
            fields='id, name, webViewLink',
            supportsAllDrives=True,
        ).execute()
        logger.info(f'Drive作成: {created.get("name")} (id={created.get("id")})')
        return created
