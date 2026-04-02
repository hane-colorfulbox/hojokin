# -*- coding: utf-8 -*-
"""Google Sheets API クライアント（管理画面用）"""
from __future__ import annotations

import logging
from datetime import datetime

from google.oauth2 import service_account
from googleapiclient.discovery import build

logger = logging.getLogger(__name__)

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# 管理シートの列定義
COL_COMPANY = 0     # A列: 顧客名
COL_TEMPLATE = 1    # B列: 申請枠
COL_FOLDER_ID = 2   # C列: DriveフォルダID
COL_STATUS = 3      # D列: 状態
COL_MESSAGE = 4     # E列: メッセージ
COL_UPDATED = 5     # F列: 更新日時
COL_OUTPUT = 6      # G列: 出力ファイル


class SheetsClient:
    """Google Sheets管理画面操作"""

    def __init__(self, credentials_path: str, spreadsheet_id: str):
        creds = service_account.Credentials.from_service_account_file(
            credentials_path, scopes=SCOPES
        )
        self.service = build('sheets', 'v4', credentials=creds)
        self.spreadsheet_id = spreadsheet_id
        self.sheet_name = '管理'
        logger.info(f'Google Sheets API 接続完了: {spreadsheet_id}')

    def _range(self, row: int, col_start: int, col_end: int) -> str:
        """セル範囲を文字列に変換"""
        cs = chr(65 + col_start)
        ce = chr(65 + col_end)
        return f'{self.sheet_name}!{cs}{row}:{ce}{row}'

    def get_all_companies(self) -> list[dict]:
        """全社のデータを取得"""
        result = self.service.spreadsheets().values().get(
            spreadsheetId=self.spreadsheet_id,
            range=f'{self.sheet_name}!A2:G100',
        ).execute()

        rows = result.get('values', [])
        companies = []
        for i, row in enumerate(rows):
            if not row or not row[0]:
                continue
            # 列が足りない場合に空文字で補完
            row += [''] * (7 - len(row))
            companies.append({
                'row_number': i + 2,  # シートの実行番号（1-indexed, ヘッダー除く）
                'company_name': row[COL_COMPANY],
                'template_type': row[COL_TEMPLATE],
                'folder_id': row[COL_FOLDER_ID],
                'status': row[COL_STATUS],
                'message': row[COL_MESSAGE],
                'updated': row[COL_UPDATED],
                'output': row[COL_OUTPUT],
            })
        return companies

    def get_pending_companies(self) -> list[dict]:
        """未処理の会社一覧を取得"""
        all_companies = self.get_all_companies()
        return [c for c in all_companies if c['status'] in ('未処理', '実行', '')]

    def update_status(self, row_number: int, status: str, message: str = '', output: str = ''):
        """ステータスを更新"""
        now = datetime.now().strftime('%Y-%m-%d %H:%M')
        values = [[status, message, now, output]]

        self.service.spreadsheets().values().update(
            spreadsheetId=self.spreadsheet_id,
            range=f'{self.sheet_name}!D{row_number}:G{row_number}',
            valueInputOption='RAW',
            body={'values': values},
        ).execute()

        logger.info(f'行{row_number} ステータス更新: {status}')

    def set_processing(self, row_number: int):
        self.update_status(row_number, '処理中', '処理を開始しました')

    def set_completed(self, row_number: int, output_files: list[str], empty_count: int = 0):
        msg = f'完了。空欄{empty_count}件' if empty_count else '完了'
        output = ', '.join(output_files)
        self.update_status(row_number, '完了', msg, output)

    def set_error(self, row_number: int, error_message: str):
        self.update_status(row_number, 'エラー', error_message[:200])
