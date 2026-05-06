# -*- coding: utf-8 -*-
"""申請書テンプレートへのデータ転記"""
from __future__ import annotations

import logging
import shutil
from pathlib import Path
import openpyxl

from openpyxl.cell.cell import MergedCell

from .models import ExtractionResult
from .config import TemplateMapping, get_min_wage

logger = logging.getLogger(__name__)


def _safe_write_cell(ws, row: int, col: int, value):
    """結合セルに対応した安全な書き込み。

    openpyxl では結合範囲の左上以外のセルに書き込むと
    ``'MergedCell' object attribute 'value' is read-only`` が出る。
    そのため、対象セルが MergedCell の場合は結合範囲の左上セルに書き込む。
    """
    cell = ws.cell(row=row, column=col)
    if isinstance(cell, MergedCell):
        for merged_range in ws.merged_cells.ranges:
            if cell.coordinate in merged_range:
                ws.cell(
                    row=merged_range.min_row,
                    column=merged_range.min_col,
                ).value = value
                return
    cell.value = value


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
        preserve = set(mapping.preserve_rows)
        for row in ws_s.iter_rows(min_row=start, max_row=end):
            cell_c = row[2] if len(row) > 2 else None
            if cell_c is None or cell_c.value is None:
                continue
            if cell_c.row in preserve:
                continue
            if is_formula(cell_c.value):
                continue
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
        _safe_write_cell(ws, m[field], 3, value)
        writes.append(f'行{m[field]:3d} [{label or field}]: {str(value)[:50]}')

    # ── 履歴事項全部証明書 or 本人確認資料 ──
    if mapping.is_kojin:
        # 個人事業主は履歴事項がないため固定値で埋める。
        # 現住所・氏名・生年月日は担当者が本人確認資料から手書き記入する前提。
        write('headquarters_address', co.address, '現在住所')
        write('established_date', co.established_date, '事業開始年月日')
        write('capital', 0, '資本金')
        write('fin_capital', 0, '資本金(財務)')
        write('fiscal_month', '12月', '決算月')
        write('officer_count_prev', 1, '役員数(前期)')
        write('rep_name', co.representative_name, '代表者氏名')
        write('rep_kana', co.representative_kana, '代表者氏名(フリガナ)')
    else:
        write('headquarters_address', co.address, '本店所在地')
        write('established_date', co.established_date, '設立年月日')
        write('capital', co.capital, '資本金')
        write('fiscal_month', fi.fiscal_month, '決算月')

        # 代表者（法人）
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
    # ── プルダウン選択型項目（番号付き選択肢）──
    write('weakness', ai.weakness, '弱み')
    write('it_investment_process', ai.it_investment_process, 'IT投資プロセス')
    write('improvement_process', ai.improvement_process, '改善プロセス')
    write('expected_effect_dept', ai.expected_effect_dept, '強化部門')
    write('expected_effect', ai.expected_effect, '期待効果')

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

    # ── 財務情報（数式参照を直接値で上書き）──
    # Phase 2: 低信頼項目は空欄+警告（write 関数を信頼度対応に拡張）
    fi_conf = getattr(fi, 'confidence', None) or {}

    def write_fin(field: str, value, label: str, conf_key: str):
        """財務系の write: confidence['xxx'].level == 'low' なら空欄+警告"""
        if field not in mapping.shinsei:
            return
        c = fi_conf.get(conf_key)
        if c and getattr(c, 'level', 'high') == 'low':
            # mapping.shinsei は dict[str, int] (行番号のみ、C列固定)。
            # mapping.kyuyo の dict[str, tuple] と混同しないこと。
            row = mapping.shinsei[field]
            writes.append(
                f'⚠ 申請内容 行{row:3d} C列 [{label}]: 低信頼のため空欄 '
                f'(理由: {getattr(c, "reason", "")})'
            )
            return
        write(field, value, label)

    write_fin('fin_revenue', fi.revenue, '売上高', 'revenue')
    write_fin('fin_gross_profit', fi.gross_profit, '粗利益', 'gross_profit')
    write_fin('fin_operating_profit', fi.operating_profit, '営業利益', 'operating_profit')
    write_fin('fin_ordinary_profit', fi.ordinary_profit, '経常利益', 'ordinary_profit')
    write_fin('fin_depreciation', fi.depreciation, '減価償却費', 'depreciation')
    # 人件費合計の信頼度: salary/misc_wages/bonus/travel いずれかが low なら全体 low
    personnel = (fi.salary or 0) + (fi.misc_wages or 0) + (fi.bonus or 0) + (fi.travel_expense or 0)
    personnel_low = any(
        fi_conf.get(k) and getattr(fi_conf[k], 'level', 'high') == 'low'
        for k in ('salary', 'misc_wages', 'bonus', 'travel_expense')
    )
    if personnel_low and 'fin_personnel' in mapping.shinsei:
        # mapping.shinsei は int（行番号のみ）、C列固定
        row = mapping.shinsei['fin_personnel']
        writes.append(
            f'⚠ 申請内容 行{row:3d} C列 [人件費]: 低信頼項目を含むため空欄'
        )
    else:
        write('fin_personnel', personnel, '人件費')
    # fin_capital は個人事業主の場合、上部で0固定済み。法人のみ co.capital を転記
    if not mapping.is_kojin:
        write('fin_capital', co.capital, '資本金(財務)')

    # ── 1人当たり給与支給総額の計画値（賃金台帳から算出時のみ）──
    # wage_plan は fill_template() から渡される場合のみ有効
    # （この関数のスコープ外で処理される）

    return writes


def fill_kyuyo_sheet(ws, mapping: TemplateMapping, data: ExtractionResult) -> list[str]:
    """給与計算シートに財務データを転記。

    Phase 2: financial.confidence を見て、低信頼項目は空欄+警告マーカー扱い。
    申請書側の処理タスクの empty_cells に追加されるよう writes に '⚠' マーカー付き行を返す。
    """
    writes = []
    fi = data.financial
    conf = getattr(fi, 'confidence', None) or {}
    m = mapping.kyuyo

    def write(field: str, value, label: str, conf_key: str = ''):
        if field not in m:
            return
        # 信頼度チェック: low なら空欄+警告マーカー（書込スキップ）
        c = conf.get(conf_key) if conf_key else None
        if c and getattr(c, 'level', 'high') == 'low':
            row, col = m[field]
            col_letter = chr(64 + col)
            writes.append(
                f'⚠ 給与計算 行{row:3d} {col_letter}列 [{label}]: 低信頼のため空欄 '
                f'(理由: {getattr(c, "reason", "")})'
            )
            return
        row, col = m[field]
        _safe_write_cell(ws, row, col, value)
        col_letter = chr(64 + col)
        writes.append(f'給与計算 行{row:3d} {col_letter}列 [{label}]: {value:,}')

    write('revenue', fi.revenue, '売上高', 'revenue')
    write('gross_profit', fi.gross_profit, '粗利益', 'gross_profit')
    write('operating_profit', fi.operating_profit, '営業利益', 'operating_profit')
    write('ordinary_profit', fi.ordinary_profit, '経常利益', 'ordinary_profit')
    write('depreciation', fi.depreciation, '減価償却費', 'depreciation')
    write('salary', fi.salary, '給料手当', 'salary')
    write('misc_wages', fi.misc_wages, '雑給', 'misc_wages')
    write('bonus', fi.bonus, '賞与手当', 'bonus')
    write('officer_comp', fi.officer_compensation, '役員報酬', 'officer_compensation')
    write('travel_expense', fi.travel_expense, '旅費交通費', 'travel_expense')

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
        # 個人事業主の別添資料
        '身分証明書', '確定申告書', '収支内訳書', '青色申告',
        # 給与計画（賃金台帳がある場合に自動入力、なければスキップ）
        '給与支給総額', '従業員数（全期間', '賃上げを行いますか',
        '事業計画期間における', '計画数値',
        # 賃金状況関連（手動確認）
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

    # 「⬇︎従業員がいない場合」セクション配下は空セルを報告しない
    # （従業員がいる場合は上部の賃上げ項目を埋めれば十分）
    in_no_employee_section = False

    for row in ws.iter_rows(min_row=35, max_row=250):
        row_num = row[0].row
        label = row[1].value if len(row) > 1 else None
        value = row[2].value if len(row) > 2 else None

        if label is not None:
            label_str_raw = str(label).strip()
            if '従業員がいない場合' in label_str_raw:
                in_no_employee_section = True
            elif '賃金状況' in label_str_raw or '最低賃金近傍' in label_str_raw:
                in_no_employee_section = False

        if label is None or value is not None:
            continue

        label_str = str(label).strip()
        if any(kw in label_str for kw in skip_keywords):
            continue

        if in_no_employee_section:
            continue

        # 使われていない役員枠はスキップ
        if is_empty_officer_slot(label_str, row_num):
            continue

        # 数値だけのラベル（テンプレート上で「数値=ラベル」になってしまっているセル）はスキップ
        # 例: 行149 のラベルが「67146344」（粗利の数値そのまま）になっているケース
        # これは check 対象としては誤検出（実際は計算式参照セルで、表示用の数値）
        ascii_only = label_str.replace(',', '').replace('-', '').replace('.', '').replace(' ', '').replace('　', '')
        if ascii_only.isdigit():
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
    wage_plan: dict[str, float] | None = None,
) -> list[str]:
    """
    テンプレートをコピーし、全データを転記して保存。
    空セルのリストを返す。

    wage_plan: 1人当たり給与支給総額の計画値
        {'year_0': 基準年, 'year_1': 1年目, 'year_2': 2年目, 'year_3': 3年目}
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
            _safe_write_cell(ws_t, row, 2, text)

    # STEP 3: PDF → 申請内容 + 給与計算
    if '申請内容' in wb.sheetnames:
        shinsei_writes = fill_shinsei_sheet(wb['申請内容'], mapping, extraction)
        for w in shinsei_writes:
            logger.info(f'STEP 3: {w}')

    if mapping.kyuyo_sheet_name in wb.sheetnames:
        kyuyo_writes = fill_kyuyo_sheet(wb[mapping.kyuyo_sheet_name], mapping, extraction)
        for w in kyuyo_writes:
            logger.info(f'STEP 3: {w}')

    # STEP 3.5: 給与支給総額の計画値を申請内容シートに転記
    if wage_plan and '申請内容' in wb.sheetnames:
        ws_shinsei = wb['申請内容']
        m = mapping.shinsei
        # 従業員数（FTE換算）
        if 'employee_count_fte' in m and 'employee_count_fte' in wage_plan:
            fte = wage_plan['employee_count_fte']
            _safe_write_cell(ws_shinsei, m['employee_count_fte'], 3, round(fte, 1))
            logger.info(f'STEP 3.5: 行{m["employee_count_fte"]:3d} [従業員数FTE]: {fte:.1f}人')
        # 給与支給総額（基準年 + 3年計画）
        plan_fields = [
            ('wage_total_base', 'wage_total_base', '給与支給総額(基準年)'),
            ('wage_total_y1', 'wage_total_y1', '給与支給総額(1年目)'),
            ('wage_total_y2', 'wage_total_y2', '給与支給総額(2年目)'),
            ('wage_total_y3', 'wage_total_y3', '給与支給総額(3年目)'),
        ]
        for field, plan_key, label in plan_fields:
            if field in m and plan_key in wage_plan:
                val = round(wage_plan[plan_key])
                _safe_write_cell(ws_shinsei, m[field], 3, val)
                logger.info(f'STEP 3.5: 行{m[field]:3d} [{label}]: {val:,}円')

    # STEP 3.6: 生産性指標シートの総労働時間(B40)を賃金台帳実績で上書き（通常枠）
    # 通常枠の「事業者あたりの総労働時間」は B38*B39（人数×平均時間）の式だが、
    # 賃金台帳の実績がある場合は基準年だけ直接値で上書きする。C40-E40(計画年次)は
    # 既存の `=C38*C39` 等の式を保持する。
    if (
        wage_plan
        and 'total_annual_hours' in wage_plan
        and '生産性指標給与支給総額計算' in wb.sheetnames
    ):
        ws_prod = wb['生産性指標給与支給総額計算']
        hours = wage_plan['total_annual_hours']
        _safe_write_cell(ws_prod, 40, 2, hours)
        logger.info(f'STEP 3.6: 生産性指標 B40 [総労働時間]: {hours:,.1f}時間')

    # STEP 4: 空セル確認
    empty = check_empty_cells(wb)
    logger.info(f'STEP 4: 空セル {len(empty)}件')

    # 保存
    wb.save(output_path)
    wb.close()
    logger.info(f'保存完了: {output_path}')

    return empty
