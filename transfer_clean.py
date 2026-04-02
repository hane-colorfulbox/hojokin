# -*- coding: utf-8 -*-
"""
資料のみから申請フォーマットを作成（元ファイルの手入力データは使わない）
1. 元ファイルをコピー → 全手入力セルをクリア → 数式だけ残す
2. 資料のみから転記
"""
import sys, shutil, datetime, openpyxl
from pathlib import Path

sys.stdout.reconfigure(encoding='utf-8')

BASE = Path('c:/Users/user/projects/カラフルボックス/補助金/1.交付申請_過去採択なし')


def find_file(directory, name_contains):
    for p in directory.iterdir():
        if name_contains in p.name:
            return p
    return None


def clear_manual_cells(wb):
    """全シートの手入力値をクリア（数式は残す）"""
    cleared = 0

    # 転記シート: B列の値をクリア（A列のラベルは残す）
    ws_t = wb['転記']
    for row in ws_t.iter_rows(min_row=5, max_row=ws_t.max_row):
        cell = row[1]  # B列
        if cell.value is not None and not (isinstance(cell.value, str) and cell.value.startswith('=')):
            cell.value = None
            cleared += 1

    # 申請内容シート: C列の直接値をクリア（数式は残す）
    ws_s = wb['申請内容']
    for row in ws_s.iter_rows(min_row=16, max_row=110):
        cell_c = row[2]  # C列
        if cell_c.value is not None and not (isinstance(cell_c.value, str) and cell_c.value.startswith('=')):
            cell_c.value = None
            cleared += 1

    # 給与支給総額計算シート: 直接値をクリア
    ws_k = wb['給与支給総額計算 ']
    direct_cells = [
        (10, 2), (11, 2), (12, 2), (13, 2), (14, 2),  # B列: 売上〜減価償却
        (5, 5), (6, 5), (7, 5), (8, 5), (9, 5),        # E列: 給与関連
        (5, 6), (6, 6), (7, 6), (8, 6), (9, 6),        # F列: 原価報告書
    ]
    for r, c in direct_cells:
        cell = ws_k.cell(row=r, column=c)
        if cell.value is not None and not (isinstance(cell.value, str) and str(cell.value).startswith('=')):
            cell.value = None
            cleared += 1

    return cleared


def transfer_from_hearing(ws_hearing, ws_tenki):
    """ヒアリングシート基本情報 → 転記シート"""
    count = 0
    for row in ws_hearing.iter_rows(min_row=1, max_row=ws_hearing.max_row):
        row_num = row[0].row
        label = row[1].value
        value = row[2].value
        if label is None or value is None:
            continue
        format_label = ws_tenki.cell(row=row_num, column=1).value
        if format_label is not None:
            ws_tenki.cell(row=row_num, column=2).value = value
            count += 1
    return count


def transfer_from_pdfs(ws_shinsei, ws_kyuyo):
    """PDFから読み取ったデータのみを転記"""
    writes = []

    def ws(row, val, label):
        ws_shinsei.cell(row=row, column=3).value = val
        writes.append(f'申請内容 行{row:3d} [{label}]: {val!r}')

    def wk(row, col, val, label):
        ws_kyuyo.cell(row=row, column=col).value = val
        writes.append(f'給与計算  行{row:3d} {chr(64+col)}列 [{label}]: {val:,}')

    # ── 履歴事項全部証明書から読めるもの ──
    ws(20, 'オーサワジャパンカブシキガイシャ', '事業者名フリガナ')
    ws(21, '東京都目黒区東山三丁目１番６号', '本店所在地')
    ws(24, datetime.datetime(1969, 4, 17), '設立年月日')
    ws(25, 54_000_000, '資本金')
    ws(37, 4, '代表者・役員数(申請時)')
    ws(38, '代表取締役', '代表者役職')
    ws(41, '0367015900', '代表電話番号(ハイフンなし)')
    ws(42, '取締役', '役員(1)役職')
    ws(43, '尾賀\u3000健太朗', '役員(1)氏名')
    ws(45, '取締役', '役員(2)役職')
    ws(46, '左近\u3000一也', '役員(2)氏名')
    ws(47, 'サコン\u3000カズヤ', '役員(2)フリガナ')
    ws(48, '監査役', '役員(3)役職')
    ws(49, '岡本\u3000麻友子', '役員(3)氏名')
    ws(68, 6, '代表者・役員数(前期決算期末)')

    # ── 納税証明書から読めるもの ──
    ws(30, '3月', '決算月')

    # ── 決算書58期（直近）から読めるもの ──
    wk(10, 2, 4_220_720_145, '売上高')
    wk(11, 2, 1_180_062_647, '粗利益')
    wk(12, 2, 202_888_794, '営業利益')
    wk(13, 2, 188_243_068, '経常利益')
    wk(14, 2, 24_316_517, '減価償却費')
    wk(5, 5, 126_580_873, '給料手当(販管費)')
    wk(6, 5, 139_453_501, '雑給(販管費)')
    wk(7, 5, 43_931_100, '賞与手当(販管費)')
    wk(8, 5, 48_596_050, '役員報酬(販管費)')

    return writes


def check_empty_cells(wb):
    """転記後に空のままのセルを一覧表示"""
    ws = wb['申請内容']
    empty = []
    for row in ws.iter_rows(min_row=16, max_row=104):
        row_num = row[0].row
        label = row[1].value
        val = row[2].value
        if label is not None and val is None:
            # 数式参照先も空かもしれないが、ラベルがあって値がないもの
            empty.append(f'行{row_num:3d} [{label}]')
    return empty


def main():
    hearing_path = find_file(BASE / '資料', 'ヒアリングシート')
    format_path  = find_file(BASE, '株式会社')
    output_path  = BASE / 'オーサワジャパン_資料のみ.xlsx'

    print(f'ヒアリングシート: {hearing_path.name}')
    print(f'ベースファイル: {format_path.name}（構造のみ使用、値は全クリア）')
    print(f'出力: {output_path.name}')

    # コピー
    shutil.copy2(format_path, output_path)

    wb_h = openpyxl.load_workbook(hearing_path)
    wb_f = openpyxl.load_workbook(output_path)

    # STEP 1: 全手入力セルをクリア
    print('\n── STEP 1: 手入力データを全クリア ──')
    cleared = clear_manual_cells(wb_f)
    print(f'  {cleared}セル クリア完了')

    # STEP 2: ヒアリングシート → 転記
    print('\n── STEP 2: ヒアリングシート → 転記シート ──')
    count = transfer_from_hearing(wb_h['基本情報'], wb_f['転記'])
    print(f'  {count}件 転記完了')

    # STEP 3: PDF → 申請内容 + 給与計算
    print('\n── STEP 3: PDF資料 → 申請内容 + 給与計算 ──')
    writes = transfer_from_pdfs(wb_f['申請内容'], wb_f['給与支給総額計算 '])
    for w in writes:
        print(f'  {w}')

    # STEP 4: 空セル確認
    print('\n── STEP 4: 資料だけでは埋められなかったセル ──')
    empty = check_empty_cells(wb_f)
    for e in empty:
        print(f'  空: {e}')

    # バリデーション
    print('\n── バリデーション ──')
    print('  履歴事項: 取得日 2025/07/08 → 2025/10/06までに申請要')
    print('  納税証明書: その1 → OK')

    wb_f.save(output_path)
    print(f'\n出力完了: {output_path.name}')


if __name__ == '__main__':
    main()
