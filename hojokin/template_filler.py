# -*- coding: utf-8 -*-
"""申請書テンプレートへのデータ転記"""
from __future__ import annotations

import logging
import shutil
from pathlib import Path
import openpyxl

from .models import ExtractionResult
from .config import TemplateMapping, get_min_wage

logger = logging.getLogger(__name__)


def clear_manual_cells(wb: openpyxl.Workbook, mapping: TemplateMapping) -> int:
    """テンプレートの手入力セルをクリア（数式は残す）"""
    cleared = 0

    def is_formula(v):
        return isinstance(v, str) and v.startswith('=')

    # 転記シート: B列のテキスト項目範囲
    if '転記' in wb.sheetnames:
        ws_t = wb['転記']
        start, end = mapping.tenki_text_range
        for r in range(start, end):
            cell = ws_t.cell(row=r, column=2)
            if cell.value is not None and not is_formula(cell.value):
                cell.value = None
                cleared += 1

    # 申請内容シート: C列
    if '申請内容' in wb.sheetnames:
        ws_s = wb['申請内容']
        start, end = mapping.shinsei_clear_range
        for row in ws_s.iter_rows(min_row=start, max_row=end):
            cell_c = row[2] if len(row) > 2 else None
            if cell_c and cell_c.value is not None and not is_formula(cell_c.value):
                cell_c.value = None
                cleared += 1

    # 給与計算シート: マッピング対象セル
    if mapping.kyuyo_sheet_name in wb.sheetnames:
        ws_k = wb[mapping.kyuyo_sheet_name]
        for field_name, (row, col) in mapping.kyuyo.items():
            cell = ws_k.cell(row=row, column=col)
            if cell.value is not None and not is_formula(str(cell.value)):
                cell.value = None
                cleared += 1

    logger.info(f'{cleared}セル クリア完了')
    return cleared


def fill_shinsei_sheet(ws, mapping: TemplateMapping, data: ExtractionResult) -> list[str]:
    """申請内容シートにデータを転記。転記した項目のログリストを返す。"""
    writes = []
    m = mapping.shinsei
    co = data.company
    fi = data.financial
    ai = data.ai_judgment

    def write(field: str, value, label: str = ''):
        if field not in m:
            return
        if value is None:
            return
        # Excelが数式と誤認する文字列を防止
        if isinstance(value, str) and value.startswith('='):
            value = ' ' + value
        ws.cell(row=m[field], column=3).value = value
        writes.append(f'行{m[field]:3d} [{label or field}]: {str(value)[:50]}')

    # ── 履歴事項全部証明書 ──
    write('headquarters_address', co.address, '本店所在地')
    write('established_date', co.established_date, '設立年月日')
    write('capital', co.capital, '資本金')
    write('fiscal_month', fi.fiscal_month, '決算月')

    # 代表者
    officer_count = 1 + len(co.officers)
    write('officer_count', officer_count, '役員数(申請時)')
    write('officer_count_prev', officer_count, '役員数(前期)')
    write('rep_title', co.representative_title, '代表者役職')
    write('rep_name', co.representative_name, '代表者氏名')
    write('rep_kana', co.representative_kana, '代表者フリガナ')

    # 役員 (最大10名)
    for i, officer in enumerate(co.officers[:10]):
        idx = i + 1
        write(f'officer_{idx}_title', officer.get('title'), f'役員({idx})役職')
        write(f'officer_{idx}_name', officer.get('name'), f'役員({idx})氏名')
        write(f'officer_{idx}_kana', officer.get('kana'), f'役員({idx})フリガナ')

    # ── 認定・補助金系 ──
    write('past_subsidies', 'なし', '過年度交付決定')
    write('eruboshi', '認定なし', 'えるぼし')
    write('kurumin', '認定なし', 'くるみん')

    # ── AI判断項目 ──
    write('industry_code', ai.industry_code, '業種コード')
    write('industry_text', ai.industry_text, '業種分類')
    write('business_description', ai.business_description, '事業内容')
    write('management_intent', ai.management_intent, '経営意欲')
    write('future_goals', ai.future_goals, '将来目標')
    write('security_status', ai.security_status, 'セキュリティ')
    write('business_types', ai.business_types, '行っている事業')
    write('it_investment_status', ai.it_investment_status, 'IT投資状況')
    write('it_utilization_status', ai.it_utilization_status, 'IT活用状況')

    # ── インボイス枠特有の項目 ──
    write('it_utilization_scope', ai.it_utilization_scope, 'IT電子化範囲')
    write('invoice_related_work', ai.invoice_related_work, 'インボイス対応業務')

    # ── 最低賃金 ──
    min_wage = get_min_wage(co.address)
    if min_wage:
        write('min_wage', f'{min_wage[0]}/{min_wage[1]}円', '地域別最低賃金')
    elif ai.min_wage_text:
        write('min_wage', ai.min_wage_text, '地域別最低賃金')

    # ── 賃上げ関連（デフォルト値） ──
    write('wage_raise_declaration', '■はい\n□いいえ', '賃上げ表明')
    write('wage_raise_amount', '＋50円', '賃上げ幅')
    write('wage_raise_method',
          '□社内掲示板などへの掲載によって\n■朝礼時、会議、面談時など口頭によって\n□書面、電子メールによって\n□その他',
          '表明方法')

    # ── ツール名 ──
    if data.estimate.tool_name:
        write('tool_name', data.estimate.tool_name, 'ツール名')

    return writes


def fill_kyuyo_sheet(ws, mapping: TemplateMapping, data: ExtractionResult) -> list[str]:
    """給与計算シートに財務データを転記"""
    writes = []
    fi = data.financial
    m = mapping.kyuyo

    def write(field: str, value, label: str):
        if field not in m:
            return
        row, col = m[field]
        ws.cell(row=row, column=col).value = value
        col_letter = chr(64 + col)
        writes.append(f'給与計算 行{row:3d} {col_letter}列 [{label}]: {value:,}')

    write('revenue', fi.revenue, '売上高')
    write('gross_profit', fi.gross_profit, '粗利益')
    write('operating_profit', fi.operating_profit, '営業利益')
    write('ordinary_profit', fi.ordinary_profit, '経常利益')
    write('depreciation', fi.depreciation, '減価償却費')
    write('salary', fi.salary, '給料手当')
    write('misc_wages', fi.misc_wages, '雑給')
    write('bonus', fi.bonus, '賞与手当')
    write('officer_comp', fi.officer_compensation, '役員報酬')

    return writes


def check_empty_cells(wb: openpyxl.Workbook) -> list[str]:
    """申請内容シートで空のままのセルを一覧表示"""
    ws = wb['申請内容']
    empty = []

    skip_keywords = {
        # 操作手順・ボタン
        '次へ', 'クリック', '宣誓', 'ファイル添付', 'アンケート',
        '計画数値入力', '書類添付', '交付申請情報', '申請要件確認',
        '事務局へ提出', '提出完了', '認証コード', '最終確認',
        '内容確認', '注意！',
        # セクションヘッダ・ラベル
        '項目', '添付資料', 'チェック項目', 'オレンジ',
        '財務情報', '経営状況', '賃金情報',
        '基本情報入力', '申請類型選択', '支援事業者入力',
        '申請要件に関する確認', '⇩必要に応じて',
        # gBizID自動取得項目（手入力不要）
        '法人番号', '事業者名', '事業者名フリガナ', '郵便番号',
        # 転記シートから手動コピーする項目
        '店舗事業所数', '事業者URL', '主な事業内容',
        '強み', '時間がかかっている', '月間何時間', 'どの機能',
        '何％', '浮いた時間', '売上目標', '属性の取引先',
        '担当部署', '担当者氏名', '担当者メールアドレス',
        '担当者電話番号', '担当者携帯番号', '代表電話番号',
        # 外部サイト確認項目
        'SECURITY ACTION照合', 'SECURITY ACTION自己宣言',
        'IT戦略ナビ', '省力化ナビ',
        # 別添資料（ファイル添付）
        '履歴事項全部証明書', '納税証明書', '決算書', 'その他資料',
        # 給与計画（別途実装予定）
        '給与支給総額', '従業員数（全期間', '賃上げを行いますか',
        '事業計画期間における', '計画数値',
        # 賃金状況関連（別途実装予定）
        '賃金状況', '最低賃金近傍', '最低賃金未満',
        '事業実施年度内', '交付申請の直近月',
        # 従業員がいない場合の項目
        '従業員がいない場合', '従業員を雇用する場合',
        # ここまで入力確認
        'ここまで入力',
        # プロンプト
        'プロンプト',
        # 補助事業者登録（手動確認項目）
        '補助事業者登録',
        # 代表者フリガナ・代表電話番号（転記シートから）
        '代表者氏名（フリガナ）',
    }

    # 使われていない役員枠を除外（役員(N)で値がないもの）
    def is_empty_officer_slot(label_str, row_num):
        """役員(N)のラベルだが値が空の場合True"""
        import re
        return bool(re.match(r'役員（[0-9０-９]+）', label_str))

    for row in ws.iter_rows(min_row=35, max_row=250):
        row_num = row[0].row
        label = row[1].value if len(row) > 1 else None
        value = row[2].value if len(row) > 2 else None

        if label is None or value is not None:
            continue

        label_str = str(label).strip()
        if any(kw in label_str for kw in skip_keywords):
            continue

        # 使われていない役員枠はスキップ
        if is_empty_officer_slot(label_str, row_num):
            continue

        empty.append(f'行{row_num:3d} [{label_str[:60]}]')

    return empty


def fill_template(
    template_path: Path,
    output_path: Path,
    mapping: TemplateMapping,
    hearing_data: dict,
    extraction: ExtractionResult,
    tenki_texts: dict[int, str] | None = None,
) -> list[str]:
    """
    テンプレートをコピーし、全データを転記して保存。
    空セルのリストを返す。
    """
    from .hearing_reader import transfer_hearing_to_tenki

    # テンプレートコピー
    shutil.copy2(template_path, output_path)
    logger.info(f'テンプレートコピー: {template_path.name} → {output_path.name}')

    wb = openpyxl.load_workbook(output_path)

    # STEP 1: サンプルデータクリア
    cleared = clear_manual_cells(wb, mapping)
    logger.info(f'STEP 1: {cleared}セル クリア')

    # STEP 2: ヒアリング → 転記
    count = 0
    if '転記' in wb.sheetnames and hearing_data:
        count = transfer_hearing_to_tenki(hearing_data, wb['転記'], mapping.hearing_to_tenki)
    logger.info(f'STEP 2: ヒアリング → {count}件転記')

    # テキスト項目（転記シートの行17-25等）
    if tenki_texts and '転記' in wb.sheetnames:
        ws_t = wb['転記']
        for row, text in tenki_texts.items():
            ws_t.cell(row=row, column=2).value = text

    # STEP 3: PDF → 申請内容 + 給与計算
    if '申請内容' in wb.sheetnames:
        shinsei_writes = fill_shinsei_sheet(wb['申請内容'], mapping, extraction)
        for w in shinsei_writes:
            logger.info(f'STEP 3: {w}')

    if mapping.kyuyo_sheet_name in wb.sheetnames:
        kyuyo_writes = fill_kyuyo_sheet(wb[mapping.kyuyo_sheet_name], mapping, extraction)
        for w in kyuyo_writes:
            logger.info(f'STEP 3: {w}')

    # STEP 4: 空セル確認
    empty = check_empty_cells(wb)
    logger.info(f'STEP 4: 空セル {len(empty)}件')

    # 保存
    wb.save(output_path)
    wb.close()
    logger.info(f'保存完了: {output_path}')

    return empty
