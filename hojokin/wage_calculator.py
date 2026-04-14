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


def _calc_fte(emp: PayrollEmployee, annual_hours: float) -> float:
    """パート・アルバイトを正社員換算"""
    if emp.employment_type in ('正社員', '契約社員'):
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
) -> Path:
    """
    給与支給総額計算Excelを作成。

    employees_detail: [{'no': 1, 'name': '氏名', 'type': '正社員',
                        'm1': 基本給, 'm2': 基本給, 'm3': 基本給,
                        'hr': 時給, 'monthly_hours': 月間時間, 'judge': '対象外'}, ...]
    """
    wb = openpyxl.Workbook()

    # 計算用定数
    total_wage_pl = financial.salary + financial.misc_wages + financial.bonus
    yakuin_annual = yakuin_hoshu_3m * 4
    wage_excl_yakuin = total_wage_pl - yakuin_annual
    total_emp = seishain_count + part_count
    standard_monthly = STANDARD_ANNUAL_HOURS / 12

    # パートFTE計算
    fte_part = 0
    if employees_detail:
        for e in employees_detail:
            if e['type'] != '正社員' and e.get('monthly_hours', 0) > 0:
                fte_part += e['monthly_hours'] / standard_monthly
    fte_adjusted = seishain_count + fte_part

    # ===== Sheet 1: 給与支給総額計算 =====
    ws1 = wb.active
    ws1.title = '給与支給総額計算'

    _cell(ws1, 2, 2, '給与支給総額計算書', TITLE_FONT, border=None)
    _cell(ws1, 3, 2, f'株式会社 {company_name}', Font(name='游ゴシック', size=12), border=None)
    _cell(ws1, 4, 2, f'事業年度: {fiscal_year_label}', NORMAL_FONT, border=None)

    # P/Lデータ
    r = 6
    _cell(ws1, r, 2, '【損益計算書データ（販管費）】', HEADER_FONT, border=None)
    r += 1
    for i, h in enumerate(['科目', '金額（円）', '備考']):
        _cell(ws1, r, 2 + i, h, HEADER_FONT_WHITE, fill=FILL_HEADER)

    items = [
        ('給料手当', financial.salary, '正社員給与'),
        ('雑給', financial.misc_wages, 'パート・アルバイト給与'),
        ('賞与', financial.bonus, ''),
        ('法定福利費', financial.legal_welfare, '※給与支給総額から除外'),
        ('福利厚生費', financial.welfare, '※給与支給総額から除外'),
    ]
    for name, amount, note in items:
        r += 1
        _cell(ws1, r, 2, name)
        _cell(ws1, r, 3, amount, fmt=NUMBER_FMT)
        _cell(ws1, r, 4, note, SMALL_FONT)
        if '除外' in note:
            for c in range(2, 5):
                ws1.cell(r, c).fill = FILL_GRAY

    r += 1
    _cell(ws1, r, 2, '給与関連合計（A）', BOLD_FONT, fill=FILL_BLUE)
    _cell(ws1, r, 3, total_wage_pl, BOLD_FONT, NUMBER_FMT, FILL_BLUE)
    _cell(ws1, r, 4, '給料手当 + 雑給 + 賞与', SMALL_FONT, fill=FILL_BLUE)

    # 役員報酬
    r += 2
    _cell(ws1, r, 2, '【役員報酬の控除】', HEADER_FONT, border=None)
    r += 1
    _cell(ws1, r, 2, '役員報酬（3ヶ月合計）')
    _cell(ws1, r, 3, yakuin_hoshu_3m, fmt=NUMBER_FMT)
    _cell(ws1, r, 4, '賃金状況報告シートより', SMALL_FONT)
    r += 1
    _cell(ws1, r, 2, '役員報酬（年間概算）（B）', BOLD_FONT, fill=FILL_YELLOW)
    _cell(ws1, r, 3, yakuin_annual, BOLD_FONT, NUMBER_FMT, FILL_YELLOW)
    _cell(ws1, r, 4, '3ヶ月合計 x 4', SMALL_FONT, fill=FILL_YELLOW)

    # 給与支給総額
    r += 2
    _cell(ws1, r, 2, '【給与支給総額の算定】', HEADER_FONT, border=None)
    r += 1
    _cell(ws1, r, 2, '給与支給総額（役員報酬込）')
    _cell(ws1, r, 3, total_wage_pl, fmt=NUMBER_FMT)
    _cell(ws1, r, 4, '(A) テンプレートE13相当', SMALL_FONT)
    r += 1
    _cell(ws1, r, 2, '給与支給総額（役員報酬除外）')
    _cell(ws1, r, 3, wage_excl_yakuin, fmt=NUMBER_FMT)
    _cell(ws1, r, 4, '(A) - (B) 賃上げ計算用', SMALL_FONT)

    # 従業員数
    r += 2
    _cell(ws1, r, 2, '【従業員数と1人当たり給与支給総額】', HEADER_FONT, border=None)
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
        _cell(ws1, r, 2, '【パートFTE換算（参考）】', HEADER_FONT, border=None)
        r += 1
        _cell(ws1, r, 2, '標準年間労働時間')
        _cell(ws1, r, 3, f'{STANDARD_ANNUAL_HOURS}時間')
        _cell(ws1, r, 4, '40h/週 x 52週', SMALL_FONT)
        r += 1
        _cell(ws1, r, 2, 'パートFTE換算合計')
        _cell(ws1, r, 3, round(fte_part, 2), fmt='0.00')
        r += 1
        _cell(ws1, r, 2, 'FTE換算後従業員数（D）', BOLD_FONT, fill=FILL_GREEN)
        _cell(ws1, r, 3, round(fte_adjusted, 2), BOLD_FONT, '0.00', FILL_GREEN)
        _cell(ws1, r, 4, f'正社員{seishain_count} + パートFTE{round(fte_part, 2)}', SMALL_FONT, fill=FILL_GREEN)

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

    # テンプレート転記用
    r += 2
    _cell(ws1, r, 2, '【2026テンプレート転記用】', HEADER_FONT, border=None)
    for name, val in [
        ('給料手当（販管費E5）', financial.salary),
        ('雑給（販管費E6）', financial.misc_wages),
        ('賞与手当（販管費E7）', financial.bonus),
        ('売上高（B10）', financial.revenue),
        ('粗利益（B11）', financial.gross_profit),
        ('営業利益（B12）', financial.operating_profit),
        ('経常利益（B13）', financial.ordinary_profit),
        ('減価償却費（B14）', financial.depreciation),
    ]:
        r += 1
        _cell(ws1, r, 2, name)
        _cell(ws1, r, 3, val, fmt=NUMBER_FMT)

    ws1.column_dimensions['A'].width = 2
    ws1.column_dimensions['B'].width = 38
    ws1.column_dimensions['C'].width = 20
    ws1.column_dimensions['D'].width = 40

    # ===== Sheet 2: 従業員別明細 =====
    if employees_detail:
        ws2 = wb.create_sheet('従業員別明細')
        _cell(ws2, 2, 2, '従業員別給与明細（直近3ヶ月）', TITLE_FONT, border=None)

        headers = ['No', '氏名', '雇用形態', '1月基本給', '2月基本給', '3月基本給',
                   '3ヶ月平均', '時給', '月間平均時間', 'FTE', '最低賃金判定']
        r = 4
        for i, h in enumerate(headers):
            _cell(ws2, r, 2 + i, h, HEADER_FONT_WHITE, fill=FILL_HEADER)
            ws2.cell(r, 2 + i).alignment = Alignment(horizontal='center', wrap_text=True)

        for e in employees_detail:
            r += 1
            avg3 = (e.get('m1', 0) + e.get('m2', 0) + e.get('m3', 0)) / 3
            fte = e.get('monthly_hours', 0) / standard_monthly if e['type'] != '正社員' else 1.0

            vals = [e['no'], e['name'], e['type'],
                    e.get('m1', 0), e.get('m2', 0), e.get('m3', 0),
                    round(avg3), e.get('hr', 0), round(e.get('monthly_hours', 0), 1),
                    round(fte, 2), e.get('judge', '')]

            for i, v in enumerate(vals):
                fmt = NUMBER_FMT if i in (3, 4, 5, 6) else ('0.00' if i == 9 else None)
                fill = FILL_GRAY if e['type'] != '正社員' else None
                _cell(ws2, r, 2 + i, v, fmt=fmt, fill=fill)

        for i, w in enumerate([4, 5, 14, 12, 12, 12, 12, 12, 8, 13, 8, 12]):
            ws2.column_dimensions[get_column_letter(i + 1)].width = w

    # ===== Sheet 3: 賃上げ計画 =====
    ws3 = wb.create_sheet('賃上げ計画')
    _cell(ws3, 2, 2, '賃上げ計画シミュレーション', TITLE_FONT, border=None)

    r = 4
    for i, h in enumerate(['', '直近決算期\n(実績値)', '1年目計画', '2年目計画', '3年目計画']):
        _cell(ws3, r, 2 + i, h, HEADER_FONT_WHITE, fill=FILL_HEADER)
        ws3.cell(r, 2 + i).alignment = Alignment(horizontal='center', wrap_text=True)

    growth = 0.03
    projections = [total_wage_pl]
    for _ in range(3):
        projections.append(round(projections[-1] * (1 + growth)))

    r += 1
    _cell(ws3, r, 2, '給与支給総額', BOLD_FONT)
    for i, p in enumerate(projections):
        _cell(ws3, r, 3 + i, p, fmt=NUMBER_FMT)

    r += 1
    _cell(ws3, r, 2, '増加率（対基準年）')
    _cell(ws3, r, 3, '-')
    for i in range(1, 4):
        _cell(ws3, r, 3 + i, (projections[i] - projections[0]) / projections[0] if projections[0] else 0, fmt=PCT_FMT)

    ws3.column_dimensions['B'].width = 28
    for col in ['C', 'D', 'E', 'F']:
        ws3.column_dimensions[col].width = 18

    # 保存
    wb.save(str(output_path))
    wb.close()
    logger.info(f'給与支給総額計算 保存: {output_path}')
    return output_path
