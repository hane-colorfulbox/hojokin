# -*- coding: utf-8 -*-
"""
賃金台帳Excel読み取り + 加点措置判定

3つのフォーマットに対応:
  1. 集計表型: 1行1人、列=月（No, 氏名, 雇用形態, 月間平均時間, 時給, 1月～12月）
  2. 月別行型: 1人×12行、列=項目（従業員番号, 氏名, 対象年月, 基本給, 所定労働時間...）
  3. 個人台帳型: 行=項目、列=月、1人1ブロック（給与ソフト出力形式）

加点措置の判定ロジック:
  ①用（インボイス枠/セキュリティ枠）:
    R6年10月～R7年9月の間で、地域別最低賃金以上かつR7年度改定後未満で
    雇用していた従業員が全従業員の30%以上いる月が3か月以上あるか
  ②用（共通）:
    交付申請の直近月における事業場内最低賃金が、
    R7年7月の事業場内最低賃金+63円以上の水準か
"""
from __future__ import annotations

import logging
import re
from dataclasses import dataclass, field
from pathlib import Path

import openpyxl

from .config import MIN_WAGE_MAP

logger = logging.getLogger(__name__)

# R6年度の最低賃金（加点措置①の下限判定に使用）
MIN_WAGE_R6 = {
    '北海道': 1010, '青森県': 953, '岩手県': 952, '宮城県': 973,
    '秋田県': 951, '山形県': 955, '福島県': 955, '東京都': 1163,
    '茨城県': 1005, '栃木県': 1004, '群馬県': 985, '埼玉県': 1078,
    '千葉県': 1076, '神奈川県': 1162, '新潟県': 985, '富山県': 998,
    '石川県': 984, '福井県': 984, '山梨県': 988, '長野県': 998,
    '岐阜県': 1001, '静岡県': 1034, '愛知県': 1077, '京都府': 1058,
    '大阪府': 1114, '三重県': 1023, '滋賀県': 1017, '兵庫県': 1052,
    '奈良県': 986, '和歌山県': 980, '鳥取県': 957, '島根県': 962,
    '岡山県': 982, '広島県': 1020, '山口県': 979, '徳島県': 980,
    '香川県': 970, '愛媛県': 956, '高知県': 952, '福岡県': 992,
    '佐賀県': 956, '長崎県': 953, '大分県': 954, '熊本県': 952,
    '宮崎県': 952, '鹿児島県': 953, '沖縄県': 952,
}

BONUS_THRESHOLD_YEN = 63

MONTH_NAMES = ['1月', '2月', '3月', '4月', '5月', '6月',
               '7月', '8月', '9月', '10月', '11月', '12月']


# ============================================================
# データ構造
# ============================================================

@dataclass
class WageEmployee:
    """賃金台帳から読み取った従業員"""
    no: int
    name: str
    employment_type: str  # 正社員 / パート・アルバイト
    monthly_avg_hours: float
    hourly_rate: float  # 代表的な時給（フォーマット1用、他は月別から算出）
    monthly_wages: list[float | None]  # 12か月分の支給合計
    monthly_hourly_rates: list[float | None] = field(
        default_factory=lambda: [None] * 12
    )

    @property
    def is_full_year(self) -> bool:
        return all(w is not None for w in self.monthly_wages)

    def months_with_data(self) -> list[int]:
        return [i for i, w in enumerate(self.monthly_wages) if w is not None]

    def get_hourly_for_month(self, month_idx: int) -> float | None:
        """指定月の時給を取得（月別データ優先、なければ代表時給）"""
        if self.monthly_hourly_rates[month_idx] is not None:
            return self.monthly_hourly_rates[month_idx]
        return self.hourly_rate if self.hourly_rate > 0 else None


@dataclass
class BonusPointResult:
    """加点措置の判定結果"""
    bonus1_eligible: bool = False
    bonus1_months_met: list[str] = field(default_factory=list)
    bonus1_details: list[dict] = field(default_factory=list)

    bonus2_eligible: bool = False
    bonus2_min_wage_july: float = 0.0
    bonus2_min_wage_latest: float = 0.0
    bonus2_diff: float = 0.0

    employees: list[WageEmployee] = field(default_factory=list)
    prefecture: str = ''
    min_wage_r6: int = 0
    min_wage_r7: int = 0


# ============================================================
# フォーマット自動検出
# ============================================================

def _detect_format(wb: openpyxl.Workbook) -> str:
    """
    賃金台帳のフォーマットを自動検出

    Returns:
        'summary'      - 集計表型（1行1人）
        'monthly_rows' - 月別行型（1人×12行）
        'individual'   - 個人台帳型（行=項目、列=月）
    """
    for ws in wb.worksheets:
        for r in range(1, min(10, ws.max_row + 1)):
            for c in range(1, min(30, ws.max_column + 1)):
                val = str(ws.cell(r, c).value or '')
                if '対象年月' in val:
                    return 'monthly_rows'
                if '月度給与' in val:
                    return 'individual'
    return 'summary'


# ============================================================
# フォーマット1: 集計表型（既存）
# ============================================================

# 列オフセット（B列=2始まり）
_F1_COL_NO = 2       # B
_F1_COL_NAME = 3     # C
_F1_COL_TYPE = 4     # D
_F1_COL_HOURS = 5    # E: 月間平均時間
_F1_COL_HOURLY = 6   # F: 時給
_F1_COL_M1 = 7       # G: 1月


def _read_summary_table(wb: openpyxl.Workbook) -> list[WageEmployee]:
    """フォーマット1: 1行1人、列=月"""
    # シート選択
    ws = None
    for name in wb.sheetnames:
        if '従業員' in name or '明細' in name or '給与' in name:
            ws = wb[name]
            break
    if ws is None:
        ws = wb[wb.sheetnames[0]]

    # ヘッダー行を探す
    header_row = None
    col_offset = 0
    for row_idx in range(1, min(10, ws.max_row + 1)):
        for col_idx in range(1, min(10, ws.max_column + 1)):
            val = ws.cell(row=row_idx, column=col_idx).value
            if val == 'No':
                header_row = row_idx
                col_offset = col_idx - _F1_COL_NO
                break
        if header_row:
            break

    if header_row is None:
        logger.warning('ヘッダー行が見つかりません（集計表型）')
        return []

    employees = []
    for row_idx in range(header_row + 1, ws.max_row + 1):
        no_val = ws.cell(row=row_idx, column=_F1_COL_NO + col_offset).value
        if no_val is None or not isinstance(no_val, (int, float)):
            continue

        name = str(ws.cell(row=row_idx, column=_F1_COL_NAME + col_offset).value or '')
        emp_type = str(ws.cell(row=row_idx, column=_F1_COL_TYPE + col_offset).value or '')
        hours = ws.cell(row=row_idx, column=_F1_COL_HOURS + col_offset).value or 0
        hourly = ws.cell(row=row_idx, column=_F1_COL_HOURLY + col_offset).value or 0

        monthly = []
        for m in range(12):
            val = ws.cell(row=row_idx, column=_F1_COL_M1 + col_offset + m).value
            monthly.append(float(val) if val is not None else None)

        name = name.replace('\u3000', ' ').strip()
        hourly_f = float(hourly)

        employees.append(WageEmployee(
            no=int(no_val),
            name=name,
            employment_type=emp_type,
            monthly_avg_hours=float(hours),
            hourly_rate=hourly_f,
            monthly_wages=monthly,
            # 集計表型はF列の時給を全月に適用
            monthly_hourly_rates=[hourly_f if w is not None else None for w in monthly],
        ))

    return employees


# ============================================================
# フォーマット2: 月別行型
# ============================================================

def _parse_year_month(text: str) -> int | None:
    """'2024年4月' や '2025年10月' → 月インデックス (0=1月)"""
    m = re.search(r'(\d+)月', str(text))
    if m:
        month = int(m.group(1))
        if 1 <= month <= 12:
            return month - 1
    return None


def _read_monthly_rows(wb: openpyxl.Workbook) -> list[WageEmployee]:
    """フォーマット2: 1人×12行、列=項目"""
    # 全シートからデータを集める（後のシートが優先）
    emp_data: dict[str, dict] = {}

    for ws in wb.worksheets:
        # ヘッダー行を探す
        header_map = {}
        header_row = None
        for r in range(1, min(10, ws.max_row + 1)):
            for c in range(1, min(30, ws.max_column + 1)):
                val = str(ws.cell(r, c).value or '').strip()
                if val in ('氏名', '従業員番号', '対象年月', '基本給',
                           '所定労働時間', '雇用形態', '支給合計'):
                    header_map[val] = c
                    header_row = r

        if header_row is None or '氏名' not in header_map:
            continue

        col_id = header_map.get('従業員番号', 1)
        col_name = header_map['氏名']
        col_type = header_map.get('雇用形態')
        col_month = header_map.get('対象年月')
        col_base = header_map.get('基本給')
        col_hours = header_map.get('所定労働時間')
        col_total = header_map.get('支給合計')

        for r in range(header_row + 1, ws.max_row + 1):
            name_val = ws.cell(r, col_name).value
            if not name_val:
                continue

            name = str(name_val).replace('\u3000', ' ').strip()
            emp_id = str(ws.cell(r, col_id).value or '').strip()
            key = f"{emp_id}_{name}"

            if key not in emp_data:
                emp_type = ''
                if col_type:
                    emp_type = str(ws.cell(r, col_type).value or '')
                emp_data[key] = {
                    'emp_id': emp_id,
                    'name': name,
                    'employment_type': emp_type,
                    'monthly_wages': [None] * 12,
                    'monthly_hourly_rates': [None] * 12,
                    'monthly_hours': [None] * 12,
                }

            # 月を特定
            month_idx = None
            if col_month:
                month_idx = _parse_year_month(
                    str(ws.cell(r, col_month).value or '')
                )
            if month_idx is None:
                continue

            # 支給合計
            if col_total:
                total = ws.cell(r, col_total).value
                if total is not None:
                    emp_data[key]['monthly_wages'][month_idx] = float(total)

            # 時給計算: 基本給 / 所定労働時間
            if col_base and col_hours:
                base = ws.cell(r, col_base).value
                hours = ws.cell(r, col_hours).value
                if base and hours:
                    base_f = float(base)
                    hours_f = float(hours)
                    emp_data[key]['monthly_hours'][month_idx] = hours_f
                    if hours_f > 0:
                        emp_data[key]['monthly_hourly_rates'][month_idx] = (
                            base_f / hours_f
                        )

    # WageEmployeeリストに変換
    employees = []
    for i, (key, data) in enumerate(emp_data.items()):
        hourly_rates = [h for h in data['monthly_hourly_rates'] if h is not None]
        avg_hourly = sum(hourly_rates) / len(hourly_rates) if hourly_rates else 0
        hours_list = [h for h in data['monthly_hours'] if h is not None]
        avg_hours = sum(hours_list) / len(hours_list) if hours_list else 0

        employees.append(WageEmployee(
            no=i + 1,
            name=data['name'],
            employment_type=data['employment_type'],
            monthly_avg_hours=round(avg_hours, 1),
            hourly_rate=round(avg_hourly, 1),
            monthly_wages=data['monthly_wages'],
            monthly_hourly_rates=data['monthly_hourly_rates'],
        ))

    return employees


# ============================================================
# フォーマット3: 個人台帳型（給与ソフト出力）
# ============================================================

def _parse_hours_str(val) -> float:
    """時間を数値に変換: 248, '248:00', '168:30' → float"""
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip()
    m = re.match(r'(\d+):(\d+)', s)
    if m:
        return int(m.group(1)) + int(m.group(2)) / 60
    try:
        return float(s)
    except ValueError:
        return 0.0


def _extract_name_from_cell(text: str) -> str:
    """'007\\n嘉口澪\\xa0(女)' → '嘉口澪'"""
    # 改行で分割して最後の部分（名前部分）を取得
    parts = str(text).split('\n')
    name_part = parts[-1] if len(parts) > 1 else parts[0]
    # 先頭の番号を除去
    name_part = re.sub(r'^\d+\s*', '', name_part)
    # 性別マーカーを除去: (女) (男) （女） （男）
    name_part = re.sub(r'\s*[\(（][男女][\)）]\s*$', '', name_part)
    # 不要な空白を整理
    name_part = name_part.replace('\xa0', ' ').replace('\u3000', ' ').strip()
    return name_part


def _parse_month_from_header(text: str) -> int | None:
    """'令和 7年\\n1月度給与' → 0 (1月=index0)"""
    m = re.search(r'(\d+)月度給与', str(text))
    if m:
        month = int(m.group(1))
        if 1 <= month <= 12:
            return month - 1
    return None


def _read_individual_ledger(wb: openpyxl.Workbook) -> list[WageEmployee]:
    """フォーマット3: 行=項目、列=月、1人1ブロック"""
    employees = []

    for ws in wb.worksheets:
        # シート内のブロックを探す（「賃金台帳」を含むセルが区切り）
        blocks = _find_individual_blocks(ws)

        for block_start, block_end in blocks:
            emp = _parse_individual_block(ws, block_start, block_end)
            if emp:
                emp.no = len(employees) + 1
                employees.append(emp)

    return employees


def _find_individual_blocks(ws) -> list[tuple[int, int]]:
    """個人台帳のブロック開始・終了行を特定"""
    blocks = []
    block_start = None

    for r in range(1, ws.max_row + 1):
        val = str(ws.cell(r, 1).value or '')
        # 「賃金台帳」または「頁」を含む行がブロック開始
        if '賃金台帳' in val or '頁' in val:
            if block_start is not None:
                blocks.append((block_start, r - 1))
            block_start = r

    # 最後のブロック
    if block_start is not None:
        blocks.append((block_start, ws.max_row))

    # ブロックが見つからなかった場合、シート全体を1ブロックとする
    if not blocks:
        blocks = [(1, ws.max_row)]

    return blocks


def _parse_individual_block(ws, start_row: int, end_row: int) -> WageEmployee | None:
    """個人台帳の1ブロックを解析"""
    # 名前を探す（開始行付近のA列）
    name = ''
    for r in range(start_row, min(start_row + 5, end_row + 1)):
        val = str(ws.cell(r, 1).value or '')
        # 番号+改行+名前のパターン、または名前を含む行
        if '\n' in val and re.search(r'\d+\n', val):
            name = _extract_name_from_cell(val)
            break

    if not name:
        return None

    # 月列のマッピングを構築（ヘッダー行から）
    month_cols: dict[int, int] = {}  # month_index → column
    for r in range(start_row, min(start_row + 5, end_row + 1)):
        for c in range(2, ws.max_column + 1):
            val = str(ws.cell(r, c).value or '')
            m_idx = _parse_month_from_header(val)
            if m_idx is not None:
                month_cols[m_idx] = c

    if not month_cols:
        return None

    # 行ラベルのマッピングを構築
    row_labels: dict[str, int] = {}
    for r in range(start_row, end_row + 1):
        label = str(ws.cell(r, 1).value or '').strip()
        if label:
            row_labels[label] = r

    # 基本給の行を特定
    base_wage_row = row_labels.get('基本給')
    # 所定労働時間の行を特定
    hours_row = row_labels.get('所定労働時間')
    # 支給合計の行を特定（候補順）
    total_row = (
        row_labels.get('課税支給合計')
        or row_labels.get('支給合計')
        or row_labels.get('差引支給合計')
    )
    # 基本給(時給)があれば時給ベースの判別に使える
    hourly_base_row = row_labels.get('基本給(時給)')

    # 月別データを抽出
    monthly_wages = [None] * 12
    monthly_hourly = [None] * 12

    for m_idx, col in month_cols.items():
        # 支給合計
        if total_row:
            val = ws.cell(total_row, col).value
            if val is not None:
                try:
                    monthly_wages[m_idx] = float(val)
                except (ValueError, TypeError):
                    pass

        # 時給計算
        if base_wage_row and hours_row:
            base = ws.cell(base_wage_row, col).value
            hours_val = ws.cell(hours_row, col).value
            if base is not None and hours_val is not None:
                try:
                    base_f = float(base)
                    hours_f = _parse_hours_str(hours_val)
                    if hours_f > 0:
                        monthly_hourly[m_idx] = base_f / hours_f
                except (ValueError, TypeError):
                    pass

    # 代表時給を算出
    valid_hourly = [h for h in monthly_hourly if h is not None]
    avg_hourly = sum(valid_hourly) / len(valid_hourly) if valid_hourly else 0

    # 平均労働時間
    total_hours = []
    if hours_row:
        for m_idx, col in month_cols.items():
            val = ws.cell(hours_row, col).value
            if val is not None:
                h = _parse_hours_str(val)
                if h > 0:
                    total_hours.append(h)
    avg_hours = sum(total_hours) / len(total_hours) if total_hours else 0

    # 雇用形態の推定（基本給(時給)行にデータがあればパート系）
    emp_type = ''
    if hourly_base_row:
        hourly_vals = [
            ws.cell(hourly_base_row, col).value
            for col in month_cols.values()
        ]
        has_hourly = any(v and float(v) > 0 for v in hourly_vals
                        if v is not None)
        if has_hourly:
            emp_type = 'パート・アルバイト'

    return WageEmployee(
        no=0,
        name=name,
        employment_type=emp_type,
        monthly_avg_hours=round(avg_hours, 1),
        hourly_rate=round(avg_hourly, 1),
        monthly_wages=monthly_wages,
        monthly_hourly_rates=monthly_hourly,
    )


# ============================================================
# メイン読み取り関数
# ============================================================

def read_wage_ledger(file_path: Path) -> list[WageEmployee]:
    """
    賃金台帳Excelを読み取り（フォーマット自動検出）

    対応フォーマット:
      1. 集計表型: 1行1人、列=月
      2. 月別行型: 1人×12行、列=項目
      3. 個人台帳型: 行=項目、列=月
    """
    wb = openpyxl.load_workbook(str(file_path), data_only=True)

    fmt = _detect_format(wb)
    logger.info(f'賃金台帳フォーマット検出: {fmt}')

    if fmt == 'monthly_rows':
        employees = _read_monthly_rows(wb)
    elif fmt == 'individual':
        employees = _read_individual_ledger(wb)
    else:
        employees = _read_summary_table(wb)

    wb.close()
    logger.info(f'賃金台帳読み取り完了: {len(employees)}名 ({fmt})')
    return employees


# ============================================================
# 加点措置判定
# ============================================================

def judge_bonus_points(
    employees: list[WageEmployee],
    prefecture: str,
    latest_month_idx: int | None = None,
) -> BonusPointResult:
    """
    加点措置①②の判定を行う

    Args:
        employees: 従業員リスト
        prefecture: 事業場の都道府県
        latest_month_idx: 直近月のインデックス（0=1月, 11=12月）。
                          Noneの場合は最新のデータがある月を使用。
    """
    result = BonusPointResult(
        employees=employees,
        prefecture=prefecture,
        min_wage_r6=MIN_WAGE_R6.get(prefecture, 0),
        min_wage_r7=MIN_WAGE_MAP.get(prefecture, 0),
    )

    if not result.min_wage_r6 or not result.min_wage_r7:
        logger.warning(f'最低賃金が見つかりません: {prefecture}')
        return result

    mw_r6 = result.min_wage_r6
    mw_r7 = result.min_wage_r7

    # ── 加点措置① ──
    target_months = list(range(0, 12))
    months_meeting_criteria = []

    for m_idx in target_months:
        total_emps = 0
        under_r7_emps = 0
        month_detail = {
            'month': MONTH_NAMES[m_idx],
            'total': 0,
            'under_r7': 0,
            'ratio': 0.0,
            'meets_30pct': False,
            'employees': [],
        }

        for emp in employees:
            if emp.monthly_wages[m_idx] is None:
                continue

            hourly = emp.get_hourly_for_month(m_idx)
            if hourly is None or hourly <= 0:
                continue

            total_emps += 1
            is_under_r7 = mw_r6 <= hourly < mw_r7

            if is_under_r7:
                under_r7_emps += 1

            month_detail['employees'].append({
                'name': emp.name,
                'hourly': round(hourly),
                'is_target': is_under_r7,
            })

        month_detail['total'] = total_emps
        month_detail['under_r7'] = under_r7_emps

        if total_emps > 0:
            ratio = under_r7_emps / total_emps
            month_detail['ratio'] = ratio
            month_detail['meets_30pct'] = ratio >= 0.30

            if month_detail['meets_30pct']:
                months_meeting_criteria.append(MONTH_NAMES[m_idx])

        result.bonus1_details.append(month_detail)

    result.bonus1_months_met = months_meeting_criteria
    result.bonus1_eligible = len(months_meeting_criteria) >= 3

    logger.info(
        f'加点措置①: {len(months_meeting_criteria)}か月が条件達成 '
        f'→ {"対象" if result.bonus1_eligible else "対象外"}'
    )

    # ── 加点措置② ──
    july_idx = 6

    july_hourly_rates = [
        emp.get_hourly_for_month(july_idx)
        for emp in employees
        if emp.monthly_wages[july_idx] is not None
        and emp.get_hourly_for_month(july_idx) is not None
        and emp.get_hourly_for_month(july_idx) > 0
    ]

    if latest_month_idx is None:
        for m in range(11, -1, -1):
            if any(emp.monthly_wages[m] is not None for emp in employees):
                latest_month_idx = m
                break
        if latest_month_idx is None:
            latest_month_idx = 11

    latest_hourly_rates = [
        emp.get_hourly_for_month(latest_month_idx)
        for emp in employees
        if emp.monthly_wages[latest_month_idx] is not None
        and emp.get_hourly_for_month(latest_month_idx) is not None
        and emp.get_hourly_for_month(latest_month_idx) > 0
    ]

    if july_hourly_rates and latest_hourly_rates:
        result.bonus2_min_wage_july = min(july_hourly_rates)
        result.bonus2_min_wage_latest = min(latest_hourly_rates)
        result.bonus2_diff = result.bonus2_min_wage_latest - result.bonus2_min_wage_july
        result.bonus2_eligible = result.bonus2_diff >= BONUS_THRESHOLD_YEN

    logger.info(
        f'加点措置②: 7月={result.bonus2_min_wage_july:.0f}円 → '
        f'直近={result.bonus2_min_wage_latest:.0f}円 '
        f'(差額{result.bonus2_diff:.0f}円) '
        f'→ {"対象" if result.bonus2_eligible else "対象外"}'
    )

    return result


# ============================================================
# 加点措置シートへの自動入力
# ============================================================

def fill_bonus_sheet_1(
    template_path: Path,
    output_path: Path,
    result: BonusPointResult,
    selected_months: list[int] | None = None,
) -> Path:
    """
    加点措置①用シートに従業員データを入力

    加点措置①のシート構成:
      3つの賃金計算期間（3か月分）を横に並べて入力
      期間①: B-K列, 期間②: M-U列, 期間③: W-AE列
      データ行は18行目から
    """
    wb = openpyxl.load_workbook(str(template_path))
    ws = wb[wb.sheetnames[0]]

    if selected_months is None:
        if result.bonus1_months_met:
            month_name_to_idx = {f'{i+1}月': i for i in range(12)}
            selected_months = [
                month_name_to_idx[m] for m in result.bonus1_months_met[:3]
            ]
        else:
            all_months = [d for d in result.bonus1_details if d['total'] > 0]
            selected_months = [
                MONTH_NAMES.index(d['month']) for d in all_months[:3]
            ]

    period_cols = [
        {'no': 2, 'name': 3, 'pref': 4, 'wage': 8, 'hourly': 9},
        {'no': 13, 'name': 14, 'pref': 15, 'wage': 18, 'hourly': 19},
        {'no': 23, 'name': 24, 'pref': 25, 'wage': 28, 'hourly': 29},
    ]

    DATA_START_ROW = 18

    for period_idx, m_idx in enumerate(selected_months[:3]):
        cols = period_cols[period_idx]

        active_emps = [
            e for e in result.employees
            if e.monthly_wages[m_idx] is not None
            and e.get_hourly_for_month(m_idx) is not None
            and e.get_hourly_for_month(m_idx) > 0
        ]

        for i, emp in enumerate(active_emps):
            row = DATA_START_ROW + i
            wage = emp.monthly_wages[m_idx]
            hourly = emp.get_hourly_for_month(m_idx)

            ws.cell(row=row, column=cols['no'], value=i + 1)
            ws.cell(row=row, column=cols['name'], value=emp.name)
            ws.cell(row=row, column=cols['pref'], value=result.prefecture)
            ws.cell(row=row, column=cols['wage'], value=wage)
            ws.cell(row=row, column=cols['hourly'], value=round(hourly))

    wb.save(str(output_path))
    wb.close()
    logger.info(f'加点措置①シート保存: {output_path}')
    return output_path


def fill_bonus_sheet_2(
    template_path: Path,
    output_path: Path,
    result: BonusPointResult,
    july_month_idx: int = 6,
    latest_month_idx: int = 11,
) -> Path:
    """
    加点措置②用シートに従業員データを入力

    加点措置②のシート構成:
      2つの賃金計算期間を横に並べて入力
      データ行は17行目から
    """
    wb = openpyxl.load_workbook(str(template_path))
    ws = wb[wb.sheetnames[0]]

    period_cols = [
        {'no': 2, 'name': 3, 'pref': 4, 'wage': 6, 'hourly': 7},
        {'no': 10, 'name': 11, 'pref': 12, 'wage': 14, 'hourly': 15},
    ]

    DATA_START_ROW = 17

    for period_idx, m_idx in enumerate([july_month_idx, latest_month_idx]):
        cols = period_cols[period_idx]

        active_emps = [
            e for e in result.employees
            if e.monthly_wages[m_idx] is not None
            and e.get_hourly_for_month(m_idx) is not None
            and e.get_hourly_for_month(m_idx) > 0
        ]

        for i, emp in enumerate(active_emps):
            row = DATA_START_ROW + i
            wage = emp.monthly_wages[m_idx]
            hourly = emp.get_hourly_for_month(m_idx)

            ws.cell(row=row, column=cols['no'], value=i + 1)
            ws.cell(row=row, column=cols['name'], value=emp.name)
            ws.cell(row=row, column=cols['pref'], value=result.prefecture)
            ws.cell(row=row, column=cols['wage'], value=wage)
            ws.cell(row=row, column=cols['hourly'], value=round(hourly))

    wb.save(str(output_path))
    wb.close()
    logger.info(f'加点措置②シート保存: {output_path}')
    return output_path
