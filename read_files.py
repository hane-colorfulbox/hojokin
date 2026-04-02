# -*- coding: utf-8 -*-
import sys, os, openpyxl
from pathlib import Path

sys.stdout.reconfigure(encoding='utf-8')

base = Path(__file__).parent / '1.交付申請_過去採択なし'
hearing_path = base / '資料' / 'ヒアリングシート_申請フォーマット2025_インボイス法人_オーサワジャパン_提出版.xlsx'
format_path = base / 'オーサワジャパン株式会社_インボイス枠_法人2025.xlsx'

print('=== ヒアリングシート ===')
print('exists:', hearing_path.exists())
wb = openpyxl.load_workbook(hearing_path)
print('シート:', wb.sheetnames)
for sn in wb.sheetnames:
    ws = wb[sn]
    print(f'\n--- {sn} ({ws.max_row}行 x {ws.max_column}列) ---')
    for row in ws.iter_rows(min_row=1, max_row=60, values_only=True):
        if any(v is not None for v in row):
            print(row)

print('\n\n=== 申請フォーマット ===')
print('exists:', format_path.exists())
wb2 = openpyxl.load_workbook(format_path)
print('シート:', wb2.sheetnames)
for sn in wb2.sheetnames:
    ws = wb2[sn]
    print(f'\n--- {sn} ({ws.max_row}行 x {ws.max_column}列) ---')
    for row in ws.iter_rows(min_row=1, max_row=60, values_only=True):
        if any(v is not None for v in row):
            print(row)
