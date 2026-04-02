# -*- coding: utf-8 -*-
"""
資料 + AI判断で申請フォーマットを最大限埋める
元ファイルの手入力データは一切使わない
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
    """全手入力セルをクリア（数式だけ残す）"""
    cleared = 0
    ws_t = wb['転記']
    for row in ws_t.iter_rows(min_row=5, max_row=ws_t.max_row):
        cell = row[1]
        if cell.value is not None and not (isinstance(cell.value, str) and cell.value.startswith('=')):
            cell.value = None
            cleared += 1

    ws_s = wb['申請内容']
    for row in ws_s.iter_rows(min_row=16, max_row=110):
        cell_c = row[2]
        if cell_c.value is not None and not (isinstance(cell_c.value, str) and cell_c.value.startswith('=')):
            cell_c.value = None
            cleared += 1

    ws_k = wb['給与支給総額計算 ']
    for r, c in [(10,2),(11,2),(12,2),(13,2),(14,2),(5,5),(6,5),(7,5),(8,5),(9,5),(5,6),(6,6),(7,6),(8,6),(9,6)]:
        cell = ws_k.cell(row=r, column=c)
        if cell.value is not None and not (isinstance(cell.value, str) and str(cell.value).startswith('=')):
            cell.value = None
            cleared += 1
    return cleared


def transfer_from_hearing(ws_hearing, ws_tenki):
    """ヒアリングシート → 転記シート"""
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


def transfer_from_pdfs(ws, wk):
    """PDF読み取りデータ → 申請内容 + 給与計算"""
    writes = []

    def s(row, val, label, source=''):
        ws.cell(row=row, column=3).value = val
        tag = f' [{source}]' if source else ''
        writes.append(f'申請内容 行{row:3d} [{label}]{tag}')

    def k(row, col, val, label):
        wk.cell(row=row, column=col).value = val
        writes.append(f'給与計算  行{row:3d} {chr(64+col)}列 [{label}]')

    # ━━ 履歴事項全部証明書 ━━
    s(20, 'オーサワジャパンカブシキガイシャ', '事業者名フリガナ', '履歴事項')
    s(21, '東京都目黒区東山三丁目１番６号', '本店所在地', '履歴事項')
    s(24, datetime.datetime(1969, 4, 17), '設立年月日', '履歴事項')
    s(25, 54_000_000, '資本金', '履歴事項')
    s(37, 4, '代表者・役員数(申請時)', '履歴事項')
    s(38, '代表取締役', '代表者役職', '履歴事項')
    s(41, '0367015900', '代表電話番号', '履歴事項+ヒアリング')
    s(42, '取締役', '役員(1)役職', '履歴事項')
    s(43, '尾賀\u3000健太朗', '役員(1)氏名', '履歴事項')
    s(45, '取締役', '役員(2)役職', '履歴事項')
    s(46, '左近\u3000一也', '役員(2)氏名', '履歴事項')
    s(47, 'サコン\u3000カズヤ', '役員(2)フリガナ', 'ヒアリング')
    s(48, '監査役', '役員(3)役職', '履歴事項')
    s(49, '岡本\u3000麻友子', '役員(3)氏名', '履歴事項')
    s(68, 6, '代表者・役員数(前期)', '履歴事項')

    # ━━ 納税証明書 ━━
    s(30, '3月', '決算月', '納税証明書')

    # ━━ 決算書58期 ━━
    k(10, 2, 4_220_720_145, '売上高')
    k(11, 2, 1_180_062_647, '粗利益')
    k(12, 2, 202_888_794, '営業利益')
    k(13, 2, 188_243_068, '経常利益')
    k(14, 2, 24_316_517, '減価償却費')
    k(5, 5, 126_580_873, '給料手当')
    k(6, 5, 139_453_501, '雑給')
    k(7, 5, 43_931_100, '賞与手当')
    k(8, 5, 48_596_050, '役員報酬')

    return writes


def ai_fill(ws, ws_tenki):
    """AIが資料の情報から判断・生成して埋めるセル"""
    writes = []

    def s(row, val, label, reason=''):
        ws.cell(row=row, column=3).value = val
        writes.append(f'行{row:3d} [{label}] ← AI: {reason}')

    # ── 業種コード（履歴事項の「目的」+ ヒアリングシートの事業内容から判断）──
    # 目的1: 食品の製造 加工 卸売及び小売販売
    # ヒアリング: 卸販売・店舗販売・通販 → 主業は卸売
    # 日本標準産業分類: I卸売業→52飲食料品卸売→522食料飲料卸売→5229その他
    s(22, 5229, '業種コード',
      '履歴事項の目的「食品の卸売及び小売」+ ヒアリング「卸販売」→ 5229')
    s(23, '大分類\tI 卸売業、小売業\n中分類\t52 飲食料品卸売業\n小分類\t522 食料・飲料卸売業\n細分類\t5229 その他の食料・飲料卸売業',
      '業種分類', '業種コード5229に対応')

    # ── ヒアリングシートの値を変換して転記 ──
    # 過年度交付: ヒアリング行42「②」→ 転記シートB42「②」→ なし
    tenki_42 = ws_tenki.cell(row=42, column=2).value
    s(56, 'なし' if tenki_42 == '②' else 'あり', '過年度交付決定',
      f'ヒアリング「{tenki_42}」→ ②=無')

    tenki_49 = ws_tenki.cell(row=49, column=2).value
    s(57, 'いいえ' if tenki_49 == '②' else 'はい', '地域DX支援',
      f'ヒアリング「{tenki_49}」→ ②=いいえ')

    tenki_54 = ws_tenki.cell(row=54, column=2).value
    s(58, 'いいえ' if tenki_54 == '②' else 'はい', '事業継続力強化計画',
      f'ヒアリング「{tenki_54}」→ ②=いいえ')

    tenki_59 = ws_tenki.cell(row=59, column=2).value
    s(59, '認定なし' if tenki_59 == '③' else tenki_59, 'えるぼし認定',
      f'ヒアリング「{tenki_59}」→ ③=認定なし')

    tenki_62 = ws_tenki.cell(row=62, column=2).value
    s(60, '認定なし' if tenki_62 == '④' else tenki_62, 'くるみん認定',
      f'ヒアリング「{tenki_62}」→ ④=認定なし')

    tenki_66 = ws_tenki.cell(row=66, column=2).value
    s(62, tenki_66, 'IT戦略ナビwith実施有無',
      f'ヒアリングからそのまま')

    # ── 行っている事業（履歴事項の「目的」から判断）──
    s(63, 'I、卸売業、小売業, E、製造業, K、不動産業、物品賃貸業', '行っている事業',
      '履歴事項の目的: 食品卸売小売(I), 製造加工(E), 不動産賃貸(K)')

    # ── 経営状況（決算書の営業利益から判断）──
    # 営業利益 202,888,794 > 0 → 積極的
    s(77, '■事業の拡大に積極的\n□事業の維持に注力\n□事業の売却・整備・廃業を考えている\n□特に意識したことは無い',
      '経営意欲', '営業利益がプラス(202M) → 事業拡大に積極的')

    s(78, '□緊急時の対応マニュアルや手順を定め、定期的に訓練を行っている\n■パソコンやサーバなどには、IDやパスワードを設け情報セキュリティ管理を行っている\n□セキュリティ対策は講じていないため、対策を講じていく\n□セキュリティ対策を講じておらず、今後もその予定はない',
      'セキュリティの状況', 'SECURITY ACTION取得済み → ID/パスワード管理あり')

    s(79, '■事業の拡大\n□利益の確保\n□顧客の定着\n□人材の再配置による営業力や生産力の強化\n□従業員の人材育成、経営陣の経営能力の向上\n□顧客満足度・利便性やサービスの質の向上による満足度の向上\n□従業員の満足度の向上\n□その他（フリー記載）\n□わからない',
      '将来目標', '営業利益プラス → 事業の拡大')

    # ── 地域別最低賃金（住所から判断）──
    s(89, '東京都/1163円', '地域別最低賃金',
      '本店所在地が東京都 → R6年度最低賃金1163円')

    # ── 賃上げ関連（ヒアリングシートから変換）──
    tenki_77 = ws_tenki.cell(row=77, column=2).value  # はい
    s(97, '■あり\n□なし', '賃上げ有無', f'ヒアリング「表明した」→ あり')
    s(98, '■はい\n□いいえ', '表明有無', f'ヒアリング「{tenki_77}」')
    s(99, '＋50円', '賃上げ幅', 'ヒアリング「❸＋50円以上」')
    s(100, '□社内掲示板などへの掲載によって\n■朝礼時、会議、面談時など口頭によって\n□書面、電子メールによって\n□その他',
      '表明方法', 'デフォルト: 口頭')
    s(101, datetime.datetime(2025, 8, 11), '表明日付',
      '申請準備日付（シート作成日）')

    # ── 事業内容（255文字以内）AIが生成 ──
    jigyou = (
        '国産を中心としたマクロビオティック食品や自然食品、健康食品、化粧品、雑貨を'
        '卸・店舗・通販で全国に提供している。'
        '現在、受発注業務や請求書発行を手作業で行っており、インボイス制度への対応に'
        '伴い事務負担が増大している。'
        'クラウド型の受発注・請求管理ツールを導入し、インボイス対応の請求書発行や'
        '受発注処理を自動化することで、月間の事務作業時間を大幅に削減する。'
        '削減した時間を新商品開発や販路拡大に充て、売上向上と生産性改善を実現し、'
        'その成果を従業員の賃上げに還元する。'
    )
    s(29, jigyou, f'事業内容({len(jigyou)}文字)',
      'ヒアリング事業内容+インボイス課題+ツール導入効果+賃上げ還元')

    return writes


def check_empty_cells(wb):
    """まだ空のセルを確認"""
    ws = wb['申請内容']
    empty = []
    skip_labels = {'項目', '添付資料', 'チェック項目', 'オレンジの項目を入力',
                   '財務情報について', '経営状況についての入力', '賃金情報計画数値'}
    for row in ws.iter_rows(min_row=16, max_row=104):
        row_num = row[0].row
        label = row[1].value
        val = row[2].value
        if label is not None and val is None and label not in skip_labels:
            empty.append(f'行{row_num:3d} [{label}]')
    return empty


def main():
    hearing_path = find_file(BASE / '資料', 'ヒアリングシート')
    format_path  = find_file(BASE, '株式会社')
    output_path  = BASE / 'オーサワジャパン_AI版.xlsx'

    print(f'出力: {output_path.name}')

    shutil.copy2(format_path, output_path)
    wb_h = openpyxl.load_workbook(hearing_path)
    wb_f = openpyxl.load_workbook(output_path)

    # STEP 1: クリア
    cleared = clear_manual_cells(wb_f)
    print(f'\n── STEP 1: {cleared}セル クリア ──')

    # STEP 2: ヒアリングシート → 転記
    count = transfer_from_hearing(wb_h['基本情報'], wb_f['転記'])
    print(f'── STEP 2: ヒアリングシート → {count}件 ──')

    # STEP 3: PDF → 直接転記
    print(f'── STEP 3: PDF資料 → 直接転記 ──')
    writes = transfer_from_pdfs(wb_f['申請内容'], wb_f['給与支給総額計算 '])
    for w in writes:
        print(f'  {w}')

    # STEP 4: AI判断・生成
    print(f'\n── STEP 4: AI判断・生成 ──')
    ai_writes = ai_fill(wb_f['申請内容'], wb_f['転記'])
    for w in ai_writes:
        print(f'  {w}')

    # STEP 5: まだ空のセル
    print(f'\n── STEP 5: まだ空のセル（人間の確認が必要）──')
    empty = check_empty_cells(wb_f)
    if empty:
        for e in empty:
            print(f'  空: {e}')
    else:
        print('  なし')

    wb_f.save(output_path)
    print(f'\n出力完了: {output_path.name}')


if __name__ == '__main__':
    main()
