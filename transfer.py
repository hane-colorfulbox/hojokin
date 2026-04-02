# -*- coding: utf-8 -*-
"""
ヒアリングシート + PDF資料 → 申請フォーマット自動転記
オーサワジャパン株式会社（インボイス枠・2025）
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


# ── PDF読み取りデータ（履歴事項全部証明書・納税証明書・決算書） ──

PDF_DATA = {
    # 履歴事項全部証明書
    '本店所在地': '東京都目黒区東山三丁目１番６号',
    '設立年月日': datetime.datetime(1969, 4, 17),
    '資本金': 54_000_000,
    '代表者役職': '代表取締役',
    '事業者名フリガナ': 'オーサワジャパンカブシキガイシャ',
    '役員数_申請時': 4,
    '役員数_前期決算期末': 6,  # R7.5退任の程島・須藤を含む
    '役員': [
        {'役職': '取締役', '氏名': '尾賀\u3000健太朗', 'フリガナ': ''},
        {'役職': '取締役', '氏名': '左近\u3000一也', 'フリガナ': 'サコン\u3000カズヤ'},
        {'役職': '監査役', '氏名': '岡本\u3000麻友子', 'フリガナ': ''},
    ],
    '履歴事項取得日': datetime.datetime(2025, 7, 8),

    # 納税証明書
    '決算月': '3月',
    '納税証明書種別': 'その1',

    # 決算書 58期（直近: 2024/4〜2025/3）
    '売上高': 4_220_720_145,
    '粗利益': 1_180_062_647,
    '営業利益': 202_888_794,
    '経常利益': 188_243_068,
    '減価償却費': 24_316_517,
    '給料手当': 126_580_873,
    '雑給': 139_453_501,
    '賞与手当': 43_931_100,
    '役員報酬': 48_596_050,
}


def transfer_hearing_sheet(ws_hearing, ws_tenki):
    """ヒアリングシート基本情報 → 転記シート"""
    count = 0
    for row in ws_hearing.iter_rows(min_row=1, max_row=ws_hearing.max_row):
        row_num = row[0].row
        label = row[1].value   # B列 = 項目名
        value = row[2].value   # C列 = 値

        if label is None or value is None:
            continue

        format_label = ws_tenki.cell(row=row_num, column=1).value
        if format_label is not None:
            ws_tenki.cell(row=row_num, column=2).value = value
            count += 1

    return count


def transfer_pdf_data(ws_shinsei, ws_kyuyo):
    """PDF読み取りデータ → 申請内容シート + 給与支給総額計算シート"""
    d = PDF_DATA
    writes = []

    # ── 申請内容シート ──
    def w(row, val, label=''):
        ws_shinsei.cell(row=row, column=3).value = val
        writes.append(f'申請内容 行{row:3d}: {label} = {val!r}')

    w(20, d['事業者名フリガナ'], '事業者名フリガナ')
    w(21, d['本店所在地'], '本店所在地')
    w(24, d['設立年月日'], '設立年月日')
    w(25, d['資本金'], '資本金')
    w(30, d['決算月'], '決算月')
    w(37, d['役員数_申請時'], '代表者・役員数(申請時)')
    w(38, d['代表者役職'], '代表者役職')
    w(68, d['役員数_前期決算期末'], '代表者・役員数(前期決算期末)')

    # 役員情報
    role_rows = [(42, 43, 44), (45, 46, 47), (48, 49, 50)]
    for i, officer in enumerate(d['役員']):
        r_pos, r_name, r_kana = role_rows[i]
        w(r_pos, officer['役職'], f'役員({i+1})役職')
        w(r_name, officer['氏名'], f'役員({i+1})氏名')
        if officer['フリガナ']:
            w(r_kana, officer['フリガナ'], f'役員({i+1})フリガナ')

    # ── 給与支給総額計算シート ──
    def wk(row, col, val, label=''):
        ws_kyuyo.cell(row=row, column=col).value = val
        col_letter = chr(64 + col)
        writes.append(f'給与計算  行{row:3d} {col_letter}列: {label} = {val:,}')

    # 財務情報（B列）
    wk(10, 2, d['売上高'], '売上高')
    wk(11, 2, d['粗利益'], '粗利益')
    wk(12, 2, d['営業利益'], '営業利益')
    wk(13, 2, d['経常利益'], '経常利益')
    wk(14, 2, d['減価償却費'], '減価償却費')

    # 給与関連（E列 = 販管費）
    wk(5, 5, d['給料手当'], '給料手当')
    wk(6, 5, d['雑給'], '雑給')
    wk(7, 5, d['賞与手当'], '賞与手当')
    wk(8, 5, d['役員報酬'], '役員報酬')

    return writes


def validate(d):
    """資料のバリデーション"""
    issues = []

    # 履歴事項: 取得から3ヶ月以内か
    three_months = d['履歴事項取得日'] + datetime.timedelta(days=90)
    issues.append(f'履歴事項全部証明書 取得日: {d["履歴事項取得日"].strftime("%Y/%m/%d")} → '
                  f'{three_months.strftime("%Y/%m/%d")}までに申請が必要')

    # 納税証明書: その1 or その2
    if d['納税証明書種別'] in ('その1', 'その2'):
        issues.append(f'納税証明書: {d["納税証明書種別"]} → OK')
    else:
        issues.append(f'納税証明書: {d["納税証明書種別"]} → NG（その1/その2のみ可）')

    return issues


def main():
    hearing_path = find_file(BASE / '資料', 'ヒアリングシート')
    format_path  = find_file(BASE, '株式会社')
    output_path  = BASE / 'オーサワジャパン_出力.xlsx'

    if not hearing_path or not format_path:
        print('エラー: 必要なファイルが見つかりません')
        sys.exit(1)

    print(f'ヒアリングシート: {hearing_path.name}')
    print(f'申請フォーマット: {format_path.name}')
    print(f'出力: {output_path.name}')

    # コピーして作業
    shutil.copy2(format_path, output_path)

    wb_h = openpyxl.load_workbook(hearing_path)
    wb_f = openpyxl.load_workbook(output_path)

    # 1. ヒアリングシート → 転記
    print('\n── ヒアリングシート → 転記シート ──')
    count = transfer_hearing_sheet(wb_h['基本情報'], wb_f['転記'])
    print(f'  {count}件 転記完了')

    # 2. PDF → 申請内容 + 給与支給総額計算
    print('\n── PDF資料 → 申請内容 + 給与支給総額計算 ──')
    writes = transfer_pdf_data(wb_f['申請内容'], wb_f['給与支給総額計算 '])
    for w in writes:
        print(f'  {w}')

    # 3. バリデーション
    print('\n── バリデーション ──')
    issues = validate(PDF_DATA)
    for i in issues:
        print(f'  {i}')

    # 保存
    wb_f.save(output_path)
    print(f'\n出力完了: {output_path.name}')


if __name__ == '__main__':
    main()
