# -*- coding: utf-8 -*-
"""
給与支給総額計算Excel生成 + 1人当たり給与支給総額算出

2026年度要件:
- 指標: 1人当たり給与支給総額（非常勤を含む全従業員）
- 計算: 給与支給総額（役員報酬除く）÷ 従業員数（パートは正社員換算）
- 年平均成長率: 3%以上

対象給与: 給料、賃金、賞与、各種手当（残業手当、休日出勤手当、
         職務手当、地域手当、家族手当、住宅手当）等
除外: 役員報酬、福利厚生費、法定福利費、退職金

対象従業員: 全月分の給与を受けた従業員のみ（中途・退職者はその年度除外）
パート換算: 正社員の所定労働時間で換算
"""
from __future__ import annotations

import logging
from dataclasses import dataclass, field as dc_field
from pathlib import Path
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter

from .models import FinancialData, MonthlyWageData
from .config import STANDARD_ANNUAL_HOURS

logger = logging.getLogger(__name__)


# ── 1人当たり給与支給総額計算（2026年新要件）──

@dataclass
class PayrollEmployee:
    """賃金台帳から読み取った従業員1名のデータ"""
    name: str
    employment_type: str  # 正社員, 契約社員, パート, アルバイト, 役員
    monthly_salary: list[float] = dc_field(default_factory=list)  # 12ヶ月分の総支給額
    monthly_hours: list[float] = dc_field(default_factory=list)   # 12ヶ月分の労働時間
    is_officer: bool = False
    is_excluded: bool = False      # 産休・育休等で除外
    full_year: bool = True         # 全月分の給与を受けたか


@dataclass
class PerCapitaWageResult:
    """1人当たり給与支給総額の計算結果"""
    total_salary: float = 0.0              # 給与支給総額（役員報酬除く）
    employee_count_fte: float = 0.0        # 従業員数（正社員換算）
    per_person_salary: float = 0.0         # 1人当たり給与支給総額
    officer_compensation: float = 0.0      # 役員報酬合計
    regular_annual_hours: float = 0.0      # 正社員の年間所定労働時間
    included: list[PayrollEmployee] = dc_field(default_factory=list)
    excluded_names: list[str] = dc_field(default_factory=list)

    GROWTH_RATE = 0.03  # 3%

    def plan_values(self) -> dict[str, float]:
        """3年分の計画数値（3%成長）"""
        b = self.per_person_salary
        r = self.GROWTH_RATE
        return {
            'year_0': b,
            'year_1': b * (1 + r),
            'year_2': b * (1 + r) ** 2,
            'year_3': b * (1 + r) ** 3,
        }


def is_full_time_employment(employment_type: str | None) -> bool:
    """雇用区分が「正規雇用相当（フルタイム）」か判定する共通 predicate。

    `_calc_fte` / 給与計算の人数集計 / 詳細表示の3箇所で同じ判定を使うことで、
    「契約社員」が場所によって正規雇用扱いになったりパート扱いになったりする
    不整合を避ける（Codex Round 4 指摘）。

    判定: '正社員' か '契約社員' を文字列に含む（provenance/修飾付きも許容）。
    例: '正社員', '正社員(推定)', '契約社員', '契約社員(臨時)' → True
        'パート', 'アルバイト', '日雇い', '役員' → False
    """
    et = employment_type or ''
    return '正社員' in et or '契約社員' in et


def _calc_fte(emp: PayrollEmployee, annual_hours: float) -> float:
    """パート・アルバイトを正社員換算。フルタイム雇用は FTE 1.0。"""
    if is_full_time_employment(emp.employment_type):
        return 1.0
    if not emp.monthly_hours:
        return 1.0
    return sum(emp.monthly_hours) / annual_hours


def calculate_per_capita_wage(
    employees: list[PayrollEmployee],
    regular_annual_hours: float = STANDARD_ANNUAL_HOURS,
) -> PerCapitaWageResult:
    """従業員リストから1人当たり給与支給総額を算出"""
    result = PerCapitaWageResult(regular_annual_hours=regular_annual_hours)

    for emp in employees:
        if emp.is_officer:
            result.officer_compensation += sum(emp.monthly_salary)
            continue
        if emp.is_excluded or not emp.full_year:
            result.excluded_names.append(emp.name)
            continue

        annual = sum(emp.monthly_salary)
        result.total_salary += annual
        fte = _calc_fte(emp, regular_annual_hours)
        result.employee_count_fte += fte
        result.included.append(emp)

    if result.employee_count_fte > 0:
        result.per_person_salary = result.total_salary / result.employee_count_fte

    logger.info(
        f'1人当たり計算: {result.total_salary:,.0f}円 / '
        f'{result.employee_count_fte:.1f}人 = {result.per_person_salary:,.0f}円'
    )
    return result

# ── スタイル定義 ──
TITLE_FONT = Font(name='游ゴシック', size=14, bold=True)
HEADER_FONT = Font(name='游ゴシック', size=10, bold=True)
NORMAL_FONT = Font(name='游ゴシック', size=10)
SMALL_FONT = Font(name='游ゴシック', size=9)
HEADER_FONT_WHITE = Font(name='游ゴシック', size=10, bold=True, color='FFFFFF')
BOLD_FONT = Font(name='游ゴシック', size=10, bold=True)
RESULT_FONT = Font(name='游ゴシック', size=12, bold=True, color='C00000')
NUMBER_FMT = '#,##0'
PCT_FMT = '0.00%'
THIN_BORDER = Border(
    left=Side(style='thin'), right=Side(style='thin'),
    top=Side(style='thin'), bottom=Side(style='thin'),
)
FILL_HEADER = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
FILL_BLUE = PatternFill(start_color='D6E4F0', end_color='D6E4F0', fill_type='solid')
FILL_YELLOW = PatternFill(start_color='FFF2CC', end_color='FFF2CC', fill_type='solid')
FILL_GREEN = PatternFill(start_color='E2EFDA', end_color='E2EFDA', fill_type='solid')
FILL_GRAY = PatternFill(start_color='F2F2F2', end_color='F2F2F2', fill_type='solid')
FILL_HEADER_DARK = PatternFill(start_color='2F5496', end_color='2F5496', fill_type='solid')
# AI抽出値（決算書PDFから Claude が読み取った値）の視認用。誤読の可能性があるため
# 人間チェックの対象として目立たせる。色は控えめなオレンジ系。
FILL_AI_EXTRACTED = PatternFill(start_color='FCE4D6', end_color='FCE4D6', fill_type='solid')


def _cell(ws, row, col, value, font=NORMAL_FONT, fmt=None, fill=None, border=THIN_BORDER):
    """セルに値とスタイルをまとめて設定"""
    c = ws.cell(row=row, column=col, value=value)
    c.font = font
    if fmt:
        c.number_format = fmt
    if fill:
        c.fill = fill
    if border:
        c.border = border
    return c


def create_wage_calculation(
    output_path: Path,
    company_name: str,
    fiscal_year_label: str,
    financial: FinancialData,
    seishain_count: int,
    part_count: int,
    yakuin_count: int,
    yakuin_hoshu_3m: int,
    employees_detail: list[dict] | None = None,
    source_files: dict[str, str] | None = None,
) -> Path:
    """
    給与支給総額計算Excelを作成。

    employees_detail: [{'no': 1, 'name': '氏名', 'type': '正社員',
                        'm1': 基本給, 'm2': 基本給, 'm3': 基本給,
                        'hr': 時給, 'monthly_hours': 月間時間, 'judge': '対象外'}, ...]
    source_files: 各データソースのファイル名（人間チェックの突合用）。
                  キー: 'pl' / 'wage_ledger' / 'wage_report' / 'registry'
                  値: ファイル名（パス除外）。AI抽出値の出所追跡用に各セクション
                  ヘッダー下に表示する。None または欠落キーは「未取得」表示。
    """
    sources = source_files or {}
    pl_source = sources.get('pl', '') or '（不明）'
    ledger_source = sources.get('wage_ledger', '') or '（不明）'
    wage_report_source = sources.get('wage_report', '')
    registry_source = sources.get('registry', '') or '（不明）'
    # 値→ページ番号の逆引き結果（pipeline 側で機械的に特定済み）
    pl_value_pages: dict[str, list[int]] = sources.get('pl_value_pages', {}) or {}

    def _page_tag(key: str) -> str:
        """財務値が決算書PDFの何ページに見つかったかをタグ文字列化。

        例: [1] → ' (PDF p.1で確認済)'
            [3,5] → ' (PDF p.3,5で確認済)'
            [] → ' (⚠PDFに値なし／AI誤読の可能性)'
            キー未登録 → ''（ページ番号特定機能が動かなかった想定。何も表示しない）
        """
        if key not in pl_value_pages:
            return ''
        pages = pl_value_pages[key]
        if not pages:
            return ' (⚠PDFに値なし／AI誤読の可能性)'
        return f' (PDF p.{",".join(map(str, pages))}で確認済)'
    wb = openpyxl.Workbook()

    # 計算用定数
    total_wage_pl = financial.salary + financial.misc_wages + financial.bonus
    yakuin_annual = yakuin_hoshu_3m * 4
    wage_excl_yakuin = total_wage_pl - yakuin_annual
    total_emp = seishain_count + part_count
    standard_monthly = STANDARD_ANNUAL_HOURS / 12

    # FTE 計算（R215「FTE換算従業員数」用）
    # 公募要領（IT2026 通常枠 p.9-10）原文:
    #   「基準年度および算出対象年度に全月分の給与等の支給を受けた従業員のみ算定対象。
    #    パート・非常勤含む。中途入退社者は除外。」
    # → 中途者は **除外**（按分ではない）。docs/補助金_実務知識ベース.md:51-52, 203 参照
    # - 12ヶ月在籍正社員: 1.0
    # - 12ヶ月在籍パート: 月間労働時間/標準月間労働時間
    # - 中途入退社者（full_year=False）: 完全除外
    fte_seishain = 0.0
    fte_part = 0.0
    excluded_midyear = 0  # 中途者として除外した人数（凡例表示用）
    if employees_detail:
        for e in employees_detail:
            if not e.get('full_year', True):
                excluded_midyear += 1
                continue
            if is_full_time_employment(e.get('type')):
                fte_seishain += 1.0
            else:
                monthly_h = e.get('monthly_hours', 0)
                if monthly_h <= 0:
                    continue
                fte_part += monthly_h / standard_monthly
        fte_adjusted = fte_seishain + fte_part
    else:
        # employees_detail が無い場合（賃金台帳・賃金状況報告シートともに取得失敗）の
        # フォールバック: 雇用区分の単純集計（中途者特定不可、警告表示）
        fte_seishain = float(seishain_count)
        fte_adjusted = seishain_count + fte_part

    # ===== Sheet 1: 給与支給総額計算 =====
    ws1 = wb.active
    ws1.title = '給与支給総額計算'

    _cell(ws1, 2, 2, '給与支給総額計算書', TITLE_FONT, border=None)
    _cell(ws1, 3, 2, f'株式会社 {company_name}', Font(name='游ゴシック', size=12), border=None)
    # 事業年度ラベルは PL 由来（AI抽出）or 賃金台帳期間からの導出。誤読リスクのあるセル
    _cell(ws1, 4, 2, f'事業年度: {fiscal_year_label}', NORMAL_FONT,
          fill=FILL_AI_EXTRACTED, border=None)
    _cell(ws1, 4, 4, f'出所: {pl_source}', SMALL_FONT, border=None)

    # 凡例（AI抽出セルの視認説明）— 計算結果より前に置いてユーザーに認識させる
    r = 5
    _cell(ws1, r, 2,
          '※薄オレンジ色のセル＝AI（Claude）が決算書PDFから抽出した値です。'
          '誤読の可能性があるため、決算書原本と必ず照合してください。',
          SMALL_FONT, border=None)
    ws1.cell(r, 2).fill = FILL_AI_EXTRACTED

    # P/Lデータ
    # 販管費・原価部の合算が実施されたかでセクションヘッダーを切替（建設業・製造業対応）
    breakdown = getattr(financial, 'breakdown', {}) or {}
    has_cost_section_merge = any(
        (breakdown.get(k, {}) or {}).get('cost_section', 0) > 0
        for k in ('salary', 'misc_wages', 'bonus', 'legal_welfare', 'welfare')
    )
    pl_section_header = (
        '【損益計算書データ（販管費＋製造原価報告書 合算）】'
        if has_cost_section_merge else '【損益計算書データ（販管費）】'
    )

    r = 7
    _cell(ws1, r, 2, pl_section_header, HEADER_FONT, border=None)
    _cell(ws1, r, 3, f'出所: {pl_source}', SMALL_FONT, border=None)
    r += 1
    for i, h in enumerate(['科目', '金額（円）', '備考']):
        _cell(ws1, r, 2 + i, h, HEADER_FONT_WHITE, fill=FILL_HEADER)

    # 各勘定科目に対応する販管費側ラベル・原価部側ラベル
    # （cost_section_label には「給料手当/賃金/給料/労務費」など複数勘定が含まれるため
    #  代表ラベルとして "賃金等" のような総称を採用）
    PL_ITEM_LABELS = {
        'salary':        ('給料手当', '賃金等（賃金/給料/労務費）'),
        'misc_wages':    ('雑給', '雑給'),
        'bonus':         ('賞与', '賞与'),
        'legal_welfare': ('法定福利費', '法定福利費'),
        'welfare':       ('福利厚生費', '福利厚生費'),
    }

    def _ai_source_tag(key: str) -> str:
        """AI抽出マーカー + ページ番号の表示文字列を生成（「決算書PDF」の重複を回避）。

        - ページ番号取得済み: 'AI抽出：決算書PDF p.3,5'
        - PDF テキストに値なし: 'AI抽出：決算書PDF（⚠PDFに値なし／AI誤読の可能性）'
        - ページ番号未取得（キー未登録）: 'AI抽出：決算書PDF'
        """
        if key not in pl_value_pages:
            return 'AI抽出：決算書PDF'
        pages = pl_value_pages[key]
        if not pages:
            return 'AI抽出：決算書PDF（⚠PDFに値なし／AI誤読の可能性）'
        return f'AI抽出：決算書PDF p.{",".join(map(str, pages))}'

    def _build_pl_note(key: str, default_pl_note: str = '', excluded: bool = False) -> str:
        """販管費＋原価部の内訳を含む備考文を生成。

        - 合算あり（販管費>0 かつ 原価部>0）: "内訳: 販管費「X」200,000 + 製造原価「Y」13,870,373 ..."
        - 販管費のみ: "販管費「X」より ..."（金額は C列に出るため重複表示しない）
        - 原価部のみ: "製造原価「Y」より ..."
        - 内訳情報なし: 旧来の default_pl_note を返す（レガシー経路・テスト互換）
        """
        bd = (breakdown.get(key) or {}) if isinstance(breakdown, dict) else {}
        pl_v = int(bd.get('pl_section') or 0)
        cost_v = int(bd.get('cost_section') or 0)
        excluded_tag = '｜※給与支給総額から除外' if excluded else ''
        ai_tag = _ai_source_tag(key)
        pl_label, cost_label = PL_ITEM_LABELS.get(key, (key, key))

        if pl_v > 0 and cost_v > 0:
            return (
                f'内訳: 販管費「{pl_label}」{pl_v:,}円 ＋ '
                f'製造原価「{cost_label}」{cost_v:,}円{excluded_tag}（{ai_tag}）'
            )
        if pl_v > 0 and cost_v == 0:
            return f'販管費「{pl_label}」より{excluded_tag}（{ai_tag}）'
        if pl_v == 0 and cost_v > 0:
            return f'製造原価「{cost_label}」より{excluded_tag}（{ai_tag}）'
        # breakdown 情報なし → 旧来表示にフォールバック（決算書構造の固定マッピング表現）
        if default_pl_note:
            return f'{default_pl_note}（{ai_tag}）'
        return f'販管費「{pl_label}」より{excluded_tag}（{ai_tag}）'

    items = [
        ('給料手当', financial.salary,
            _build_pl_note('salary', default_pl_note='販管費より｜正社員給与')),
        ('雑給', financial.misc_wages,
            _build_pl_note('misc_wages', default_pl_note='販管費より｜パート・アルバイト給与')),
        ('賞与', financial.bonus,
            _build_pl_note('bonus', default_pl_note='販管費より')),
        ('法定福利費', financial.legal_welfare,
            _build_pl_note('legal_welfare', excluded=True)),
        ('福利厚生費', financial.welfare,
            _build_pl_note('welfare', excluded=True)),
    ]
    # 各科目のセル位置を後段の Excel 計算式で参照できるよう記録
    item_rows: dict[str, int] = {}
    item_keys = ['salary', 'misc_wages', 'bonus', 'legal_welfare', 'welfare']
    for (name, amount, note), key in zip(items, item_keys):
        r += 1
        _cell(ws1, r, 2, name)
        # AI抽出値はオレンジで塗る（除外行は灰色優先）
        is_excluded = '除外' in note
        cell_fill = FILL_GRAY if is_excluded else FILL_AI_EXTRACTED
        _cell(ws1, r, 3, amount, fmt=NUMBER_FMT, fill=cell_fill)
        _cell(ws1, r, 4, note, SMALL_FONT, fill=cell_fill if is_excluded else None)
        if is_excluded:
            ws1.cell(r, 2).fill = FILL_GRAY
        item_rows[key] = r

    r += 1
    a_row = r  # (A) の行を覚えておく → 後段で参照
    # 給与関連合計 (A) を Excel 式で表現（給料手当+雑給+賞与）。
    # ユーザーが Excel 上で値を編集しても合計が自動更新される。
    a_formula = (
        f'=C{item_rows["salary"]}+C{item_rows["misc_wages"]}+C{item_rows["bonus"]}'
    )
    _cell(ws1, r, 2, '給与関連合計（A）', BOLD_FONT, fill=FILL_BLUE)
    _cell(ws1, r, 3, a_formula, BOLD_FONT, NUMBER_FMT, FILL_BLUE)
    _cell(ws1, r, 4,
          f'機械計算: C{item_rows["salary"]}+C{item_rows["misc_wages"]}+C{item_rows["bonus"]}（給料手当+雑給+賞与）',
          SMALL_FONT, fill=FILL_BLUE)

    # 役員報酬
    r += 2
    _cell(ws1, r, 2, '【役員報酬の控除】', HEADER_FONT, border=None)
    # ソース表示: 賃金状況報告シート優先、未取得時は PL から推定
    yakuin_source_is_ai = (
        financial.officer_compensation > 0
        and yakuin_hoshu_3m == int(financial.officer_compensation / 4)
    )
    if yakuin_source_is_ai:
        yakuin_source_label = f'出所: {pl_source}（決算書PDF）'
    elif wage_report_source:
        yakuin_source_label = f'出所: {wage_report_source}'
    else:
        yakuin_source_label = '出所: （不明）'
    _cell(ws1, r, 3, yakuin_source_label, SMALL_FONT, border=None)
    r += 1
    _cell(ws1, r, 2, '役員報酬（3ヶ月合計）')
    yakuin_cell_fill = FILL_AI_EXTRACTED if yakuin_source_is_ai else None
    yakuin_note = (
        '※AI抽出：決算書PDFの役員報酬を÷4で推定'
        if yakuin_source_is_ai else '賃金状況報告シートより'
    )
    # 役員報酬の備考に「販管費の役員報酬欄から」を明示
    if yakuin_source_is_ai:
        yakuin_note = (
            f'販管費「役員報酬」より｜※AI抽出：決算書PDFの年額'
            f'÷4で推定{_page_tag("officer_compensation")}'
        )
    _cell(ws1, r, 3, yakuin_hoshu_3m, fmt=NUMBER_FMT, fill=yakuin_cell_fill)
    _cell(ws1, r, 4, yakuin_note, SMALL_FONT, fill=yakuin_cell_fill)
    yakuin_3m_row = r  # 3ヶ月合計の行を覚えておく
    r += 1
    b_row = r  # (B) の行を覚えておく
    # (B) を Excel 式で表現（3ヶ月合計 × 4）
    _cell(ws1, r, 2, '役員報酬（年間概算）（B）', BOLD_FONT, fill=FILL_YELLOW)
    _cell(ws1, r, 3, f'=C{yakuin_3m_row}*4', BOLD_FONT, NUMBER_FMT, FILL_YELLOW)
    _cell(ws1, r, 4, f'機械計算: C{yakuin_3m_row}×4（3ヶ月合計 × 4）',
          SMALL_FONT, fill=FILL_YELLOW)

    # 給与支給総額
    r += 2
    _cell(ws1, r, 2, '【給与支給総額の算定】', HEADER_FONT, border=None)
    r += 1
    # 役員報酬込 = (A) を Excel 式で再掲
    _cell(ws1, r, 2, '給与支給総額（役員報酬込）')
    _cell(ws1, r, 3, f'=C{a_row}', fmt=NUMBER_FMT)
    _cell(ws1, r, 4, f'機械計算: =C{a_row}（=A の再掲、テンプレートE13相当）',
          SMALL_FONT)
    r += 1
    # 役員報酬除外 = (A) - (B) を Excel 式で
    _cell(ws1, r, 2, '給与支給総額（役員報酬除外）')
    _cell(ws1, r, 3, f'=C{a_row}-C{b_row}', fmt=NUMBER_FMT)
    _cell(ws1, r, 4, f'機械計算: =C{a_row}-C{b_row}（=A-B 賃上げ計算用）',
          SMALL_FONT)

    # 従業員数
    r += 2
    _cell(ws1, r, 2, '【従業員数と1人当たり給与支給総額】', HEADER_FONT, border=None)
    # 雇用区分の人数集計: 賃金状況報告シートがあればそちら、なければ賃金台帳から逆算
    headcount_source = wage_report_source or ledger_source
    _cell(ws1, r, 3, f'出所: {headcount_source}（役員数は{registry_source}）',
          SMALL_FONT, border=None)
    r += 1
    for i, h in enumerate(['項目', '人数/金額', '備考']):
        _cell(ws1, r, 2 + i, h, HEADER_FONT_WHITE, fill=FILL_HEADER)

    for name, val, note in [
        ('正規雇用従業員', f'{seishain_count}人', ''),
        ('契約社員', '0人', ''),
        ('パート・アルバイト', f'{part_count}人', ''),
        ('役員', f'{yakuin_count}人', '※従業員数に含まず'),
    ]:
        r += 1
        _cell(ws1, r, 2, name)
        _cell(ws1, r, 3, val)
        _cell(ws1, r, 4, note, SMALL_FONT)

    r += 1
    _cell(ws1, r, 2, '従業員合計（C）', BOLD_FONT, fill=FILL_BLUE)
    _cell(ws1, r, 3, f'{total_emp}人', BOLD_FONT, fill=FILL_BLUE)
    _cell(ws1, r, 4, '', fill=FILL_BLUE)

    # FTE
    if employees_detail:
        r += 2
        _cell(ws1, r, 2, '【FTE換算（12ヶ月在籍者のみ／中途入退社は除外）】',
              HEADER_FONT, border=None)
        _cell(ws1, r, 3,
              f'出所: IT2026 通常枠公募要領 p.9-10（中途入退社者は算定対象外）',
              SMALL_FONT, border=None)
        r += 1
        _cell(ws1, r, 2, '標準年間労働時間')
        _cell(ws1, r, 3, f'{STANDARD_ANNUAL_HOURS}時間')
        _cell(ws1, r, 4, '40h/週 x 52週', SMALL_FONT)
        r += 1
        seishain_fte_row = r
        _cell(ws1, r, 2, '正社員FTE合計（12ヶ月在籍のみ）')
        _cell(ws1, r, 3, round(fte_seishain, 2), fmt='0.00')
        _cell(ws1, r, 4,
              f'12ヶ月在籍正社員のみカウント（中途者は公募要領により除外）',
              SMALL_FONT)
        r += 1
        part_fte_row = r
        _cell(ws1, r, 2, 'パートFTE換算合計（12ヶ月在籍のみ）')
        _cell(ws1, r, 3, round(fte_part, 2), fmt='0.00')
        _cell(ws1, r, 4,
              f'12ヶ月在籍パートのみ。月間労働時間/標準月間時間 で正社員換算',
              SMALL_FONT)
        if excluded_midyear:
            r += 1
            _cell(ws1, r, 2, '中途入退社で除外した人数')
            _cell(ws1, r, 3, excluded_midyear)
            _cell(ws1, r, 4,
                  '※R215 算定対象外（公募要領「全月分の給与等の支給を受けた従業員」）',
                  SMALL_FONT)
            ws1.cell(r, 2).fill = FILL_GRAY
            ws1.cell(r, 3).fill = FILL_GRAY
            ws1.cell(r, 4).fill = FILL_GRAY
        r += 1
        # FTE換算後（D）を Excel 関数化（正社員FTE + パートFTE）
        _cell(ws1, r, 2, 'FTE換算後従業員数（D）', BOLD_FONT, fill=FILL_GREEN)
        _cell(ws1, r, 3, f'=C{seishain_fte_row}+C{part_fte_row}',
              BOLD_FONT, '0.00', FILL_GREEN)
        _cell(ws1, r, 4,
              f'機械計算: C{seishain_fte_row}+C{part_fte_row}（正社員FTE + パートFTE）',
              SMALL_FONT, fill=FILL_GREEN)

    # 1人当たり計算
    r += 2
    _cell(ws1, r, 2, '【1人当たり給与支給総額】', Font(name='游ゴシック', size=12, bold=True), border=None)
    r += 1
    for i, h in enumerate(['算出方法', '金額', '']):
        _cell(ws1, r, 2 + i, h, HEADER_FONT_WHITE, fill=FILL_HEADER_DARK)

    calc_methods = [
        ('(A)÷(C) 頭数割り', total_wage_pl / total_emp if total_emp else 0),
        ('(A-B)÷(C) 役員除外・頭数', wage_excl_yakuin / total_emp if total_emp else 0),
    ]
    if employees_detail and fte_adjusted > 0:
        calc_methods.extend([
            ('(A)÷(D) FTE換算', total_wage_pl / fte_adjusted),
            ('(A-B)÷(D) 役員除外・FTE（推奨）', wage_excl_yakuin / fte_adjusted),
        ])

    for i, (label, amount) in enumerate(calc_methods):
        r += 1
        is_last = (i == len(calc_methods) - 1)
        _cell(ws1, r, 2, label, BOLD_FONT if is_last else NORMAL_FONT,
              fill=FILL_GREEN if is_last else None)
        _cell(ws1, r, 3, round(amount), RESULT_FONT if is_last else NORMAL_FONT,
              NUMBER_FMT, FILL_GREEN if is_last else None)

    # テンプレート転記用（全項目 AI 抽出：決算書PDF由来）
    r += 2
    _cell(ws1, r, 2, '【2026テンプレート転記用】', HEADER_FONT, border=None)
    _cell(ws1, r, 3, f'出所: {pl_source}', SMALL_FONT, border=None)
    r += 1
    _cell(ws1, r, 2,
          '※下記すべてAI抽出値（決算書PDF由来）。テンプレ転記前に決算書原本と照合してください。',
          SMALL_FONT, border=None)
    ws1.cell(r, 2).fill = FILL_AI_EXTRACTED
    # 決算書のどの表に載っている値かをセル右側の備考に明示する
    # （会計指針で標準化されているため固定マッピングで 100% 正確）
    # 加えて PDF テキストから逆引きしたページ番号も付記する
    template_items = [
        ('給料手当（販管費E5）', financial.salary, '販売費及び一般管理費', 'salary'),
        ('雑給（販管費E6）', financial.misc_wages, '販売費及び一般管理費', 'misc_wages'),
        ('賞与手当（販管費E7）', financial.bonus, '販売費及び一般管理費', 'bonus'),
        ('売上高（B10）', financial.revenue, '損益計算書', 'revenue'),
        ('粗利益（B11）', financial.gross_profit, '損益計算書（売上総利益）', 'gross_profit'),
        ('営業利益（B12）', financial.operating_profit, '損益計算書', 'operating_profit'),
        ('経常利益（B13）', financial.ordinary_profit, '損益計算書', 'ordinary_profit'),
        ('減価償却費（B14）', financial.depreciation,
         '販売費及び一般管理費 or 製造原価報告書', 'depreciation'),
    ]
    for name, val, where, page_key in template_items:
        r += 1
        _cell(ws1, r, 2, name)
        _cell(ws1, r, 3, val, fmt=NUMBER_FMT, fill=FILL_AI_EXTRACTED)
        _cell(ws1, r, 4, f'決算書「{where}」より{_page_tag(page_key)}', SMALL_FONT)

    ws1.column_dimensions['A'].width = 2
    ws1.column_dimensions['B'].width = 38
    ws1.column_dimensions['C'].width = 20
    ws1.column_dimensions['D'].width = 40

    # ===== Sheet 2: 従業員別明細 =====
    if employees_detail:
        ws2 = wb.create_sheet('従業員別明細')
        _cell(ws2, 2, 2, '従業員別給与明細（直近3ヶ月）', TITLE_FONT, border=None)
        _cell(ws2, 2, 4, f'出所: {ledger_source}', SMALL_FONT, border=None)
        _cell(ws2, 3, 2,
              '※氏名・雇用形態・月給・時給はAIが賃金台帳から読み取った値です。'
              '誤読の可能性があるため賃金台帳原本と照合してください。',
              SMALL_FONT, border=None)
        ws2.cell(3, 2).fill = FILL_AI_EXTRACTED

        headers = ['No', '氏名', '雇用形態', '1月基本給', '2月基本給', '3月基本給',
                   '3ヶ月平均', '時給', '月間平均時間', 'FTE', '最低賃金判定', '備考']
        r = 4
        for i, h in enumerate(headers):
            _cell(ws2, r, 2 + i, h, HEADER_FONT_WHITE, fill=FILL_HEADER)
            ws2.cell(r, 2 + i).alignment = Alignment(horizontal='center', wrap_text=True)

        # 中途入退社社員のチェック視認性向上のため、行全体を灰色塗りする
        FILL_INCOMPLETE = PatternFill(start_color='DDDDDD', end_color='DDDDDD', fill_type='solid')

        for e in employees_detail:
            r += 1
            m_vals = [e.get('m1', 0), e.get('m2', 0), e.get('m3', 0)]
            # 在籍月のみで3ヶ月平均を算出（0月を分母に入れると過小評価される）
            in_service = [v for v in m_vals if v > 0]
            avg3 = sum(in_service) / len(in_service) if in_service else 0

            is_seishain = is_full_time_employment(e.get('type'))
            full_year = e.get('full_year', True)
            tenure_months = e.get('tenure_months', 12)
            # 在籍月数を反映した FTE（中途入退社は分母12を按分）
            tenure_factor = min(tenure_months, 12) / 12 if tenure_months > 0 else 0
            if is_seishain:
                fte = 1.0 * tenure_factor
            else:
                monthly_h = e.get('monthly_hours', 0)
                fte = (monthly_h / standard_monthly) * tenure_factor if standard_monthly else 0

            # 備考: 中途入退社の表示 + 実際の月並びの提示（誤読防止）
            note_parts = []
            if not full_year:
                note_parts.append(f'中途入退社（在籍{tenure_months}ヶ月）')
                # 実際の在籍月ラベルが分かっていれば表示
                labels = [l for l in e.get('last_three_labels', []) if l]
                if labels:
                    note_parts.append(f'実体: {"/".join(labels)}')
            note = ' '.join(note_parts)

            # 最低賃金判定: 賃金状況報告シート由来の judge があれば優先、
            # 無ければ「-」（このシートでは時給・都道府県情報が揃わないため判定不能）
            judge_val = e.get('judge') or '-'

            vals = [e['no'], e['name'], e['type'],
                    e.get('m1', 0), e.get('m2', 0), e.get('m3', 0),
                    round(avg3), e.get('hr', 0), round(e.get('monthly_hours', 0), 1),
                    round(fte, 2), judge_val, note]

            for i, v in enumerate(vals):
                fmt = NUMBER_FMT if i in (3, 4, 5, 6) else ('0.00' if i == 9 else None)
                # 行の塗り分け（優先順位: 中途入退社 > 非正規 > 通常）
                if not full_year:
                    fill = FILL_INCOMPLETE
                elif not is_seishain:
                    fill = FILL_GRAY
                else:
                    fill = None
                _cell(ws2, r, 2 + i, v, fmt=fmt, fill=fill)

        # ── 合計行（全員 / 12ヶ月在籍のみの2段）────────────────────────
        # 後段の検算・他ファイルとの突合用。SUM 式で書いておくと行追加・編集後も
        # 自動再計算される。
        if any(e.get('m1') or e.get('m2') or e.get('m3') for e in employees_detail):
            first_data_row = 5  # ヘッダー r=4 の直下が最初のデータ行
            last_data_row = r
            r += 1
            FILL_SUBTOTAL_ALL = PatternFill(start_color='B4C7E7', end_color='B4C7E7', fill_type='solid')
            FILL_SUBTOTAL_TARGET = PatternFill(start_color='C6E0B4', end_color='C6E0B4', fill_type='solid')

            _cell(ws2, r, 2, '', BOLD_FONT, fill=FILL_SUBTOTAL_ALL)
            _cell(ws2, r, 3, '合計（全員）', BOLD_FONT, fill=FILL_SUBTOTAL_ALL)
            _cell(ws2, r, 4, '', fill=FILL_SUBTOTAL_ALL)
            # 1月/2月/3月/3ヶ月平均 の列合計
            for col_idx in (5, 6, 7, 8):
                col_letter = get_column_letter(col_idx)
                _cell(
                    ws2, r, col_idx,
                    f'=SUM({col_letter}{first_data_row}:{col_letter}{last_data_row})',
                    BOLD_FONT, fmt=NUMBER_FMT, fill=FILL_SUBTOTAL_ALL,
                )
            # 残り列はブランク（合計対象外）
            for col_idx in (9, 10, 11, 12, 13):
                _cell(ws2, r, col_idx, '', fill=FILL_SUBTOTAL_ALL)

            # 12ヶ月在籍のみ（R216の母数になる集計対象）
            target_rows = [
                first_data_row + i for i, e in enumerate(employees_detail)
                if e.get('full_year', True)
            ]
            r += 1
            _cell(ws2, r, 2, '', BOLD_FONT, fill=FILL_SUBTOTAL_TARGET)
            _cell(ws2, r, 3, '合計（12ヶ月在籍のみ）', BOLD_FONT, fill=FILL_SUBTOTAL_TARGET)
            _cell(ws2, r, 4, '', fill=FILL_SUBTOTAL_TARGET)
            for col_idx in (5, 6, 7, 8):
                col_letter = get_column_letter(col_idx)
                if target_rows:
                    parts = '+'.join(f'{col_letter}{tr}' for tr in target_rows)
                    formula = f'={parts}'
                else:
                    formula = 0
                _cell(
                    ws2, r, col_idx, formula,
                    BOLD_FONT, fmt=NUMBER_FMT, fill=FILL_SUBTOTAL_TARGET,
                )
            for col_idx in (9, 10, 11, 12, 13):
                _cell(ws2, r, col_idx, '', fill=FILL_SUBTOTAL_TARGET)

        # 凡例
        r += 2
        _cell(ws2, r, 2,
              '※灰色（濃）行＝直近事業年度に12ヶ月在籍していない社員（中途入社・退職含む）。',
              SMALL_FONT, border=None)
        r += 1
        _cell(ws2, r, 2,
              '　1月/2月/3月の列見出しは便宜表示で、中途者の列は実在籍月の時系列順（備考の「実体」欄を参照）。',
              SMALL_FONT, border=None)
        r += 1
        _cell(ws2, r, 2,
              '※灰色（薄）行＝非正規雇用（パート・アルバイト）。',
              SMALL_FONT, border=None)
        r += 1
        _cell(ws2, r, 2,
              '※最低賃金判定の「-」＝このシートでは判定なし（加点判定タスクで都道府県を指定した場合に判定されます）。',
              SMALL_FONT, border=None)

        for i, w in enumerate([4, 5, 14, 12, 12, 12, 12, 12, 8, 13, 8, 12, 30]):
            ws2.column_dimensions[get_column_letter(i + 1)].width = w

    # ===== Sheet 3: 賃上げ計画 =====
    # 機械計算セルは全部 Excel 関数化（ユーザーが C5 を編集すると D5-F5 が自動再計算）
    ws3 = wb.create_sheet('賃上げ計画')
    _cell(ws3, 2, 2, '賃上げ計画シミュレーション', TITLE_FONT, border=None)

    r = 4
    for i, h in enumerate(['', '直近決算期\n(実績値)', '1年目計画', '2年目計画', '3年目計画']):
        _cell(ws3, r, 2 + i, h, HEADER_FONT_WHITE, fill=FILL_HEADER)
        ws3.cell(r, 2 + i).alignment = Alignment(horizontal='center', wrap_text=True)

    r += 1
    basis_row = r  # 給与支給総額 行（C列 = 直近実績、D-F 列 = 計画値）
    _cell(ws3, r, 2, '給与支給総額', BOLD_FONT)
    # C5: 直近実績（給与支給総額計算シートの (A) と連動）
    # シート間参照式でリンク → (A) を編集すると賃上げ計画も自動再計算
    _cell(ws3, r, 3, f"='給与支給総額計算'!C{a_row}", fmt=NUMBER_FMT)
    # D5, E5, F5: 前年 × 1.03（年率3%）。ROUND で円単位に丸め
    prev_col = 'C'
    for i in range(1, 4):
        col_letter = get_column_letter(3 + i)  # D, E, F
        _cell(ws3, r, 3 + i, f'=ROUND({prev_col}{r}*1.03, 0)', fmt=NUMBER_FMT)
        prev_col = col_letter

    r += 1
    _cell(ws3, r, 2, '増加率（対基準年）')
    _cell(ws3, r, 3, '-')
    # D6, E6, F6: (D5 - C5) / C5 → 累積増加率
    for i in range(1, 4):
        col_letter = get_column_letter(3 + i)
        _cell(ws3, r, 3 + i,
              f'=({col_letter}{basis_row}-$C${basis_row})/$C${basis_row}',
              fmt=PCT_FMT)

    ws3.column_dimensions['B'].width = 28
    for col in ['C', 'D', 'E', 'F']:
        ws3.column_dimensions[col].width = 18

    # 保存
    wb.save(str(output_path))
    wb.close()
    logger.info(f'給与支給総額計算 保存: {output_path}')
    return output_path
