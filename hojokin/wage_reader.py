# -*- coding: utf-8 -*-
"""
賃金台帳Excel読み取り + 加点措置判定

賃金台帳サンプル構成:
  シート「従業員別明細」
  B列: No, C列: 氏名, D列: 雇用形態, E列: 月間平均時間, F列: 時給
  G～R列: 1月～12月の月別給与（課税対象額合計）

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

# 加点措置②の閾値
BONUS_THRESHOLD_YEN = 63

# 賃金台帳の列オフセット（B列=2始まり）
COL_NO = 2       # B
COL_NAME = 3     # C
COL_TYPE = 4     # D
COL_HOURS = 5    # E: 月間平均時間
COL_HOURLY = 6   # F: 時給
COL_M1 = 7       # G: 1月
# H=2月, I=3月 ... R=12月


@dataclass
class WageEmployee:
    """賃金台帳から読み取った従業員"""
    no: int
    name: str
    employment_type: str  # 正社員 / パート・アルバイト
    monthly_avg_hours: float
    hourly_rate: float
    monthly_wages: list[float | None]  # 12か月分（None=データなし）

    @property
    def is_full_year(self) -> bool:
        """12か月分すべてデータがあるか"""
        return all(w is not None for w in self.monthly_wages)

    def months_with_data(self) -> list[int]:
        """データがある月のインデックスリスト（0=1月, 11=12月）"""
        return [i for i, w in enumerate(self.monthly_wages) if w is not None]

    def calc_hourly_for_month(self, month_idx: int) -> float | None:
        """指定月の時間換算給与を計算"""
        wage = self.monthly_wages[month_idx]
        if wage is None or self.monthly_avg_hours <= 0:
            return None
        return wage / self.monthly_avg_hours


@dataclass
class BonusPointResult:
    """加点措置の判定結果"""
    # 加点措置①（30%・3か月条件）
    bonus1_eligible: bool = False
    bonus1_months_met: list[str] = field(default_factory=list)
    bonus1_details: list[dict] = field(default_factory=list)

    # 加点措置②（+63円条件）
    bonus2_eligible: bool = False
    bonus2_min_wage_july: float = 0.0
    bonus2_min_wage_latest: float = 0.0
    bonus2_diff: float = 0.0

    # 全従業員データ
    employees: list[WageEmployee] = field(default_factory=list)
    prefecture: str = ''
    min_wage_r6: int = 0
    min_wage_r7: int = 0


def read_wage_ledger(file_path: Path) -> list[WageEmployee]:
    """賃金台帳Excelを読み取り"""
    wb = openpyxl.load_workbook(str(file_path), data_only=True)

    # シート名の候補
    target_sheet = None
    for name in wb.sheetnames:
        if '従業員' in name or '明細' in name or '給与' in name:
            target_sheet = wb[name]
            break
    if target_sheet is None:
        target_sheet = wb[wb.sheetnames[0]]

    ws = target_sheet
    employees = []

    # ヘッダー行を探す（「No」「氏名」を含む行）
    header_row = None
    for row_idx in range(1, min(10, ws.max_row + 1)):
        for col_idx in range(1, min(10, ws.max_column + 1)):
            val = ws.cell(row=row_idx, column=col_idx).value
            if val == 'No':
                header_row = row_idx
                # No列の位置を基準にオフセットを調整
                col_offset = col_idx - COL_NO
                break
        if header_row:
            break

    if header_row is None:
        logger.warning('ヘッダー行が見つかりません')
        wb.close()
        return []

    col_offset = col_offset if 'col_offset' in dir() else 0

    # データ行を読み取り
    for row_idx in range(header_row + 1, ws.max_row + 1):
        no_val = ws.cell(row=row_idx, column=COL_NO + col_offset).value
        if no_val is None or not isinstance(no_val, (int, float)):
            # 例行（"例"）やヘッダーはスキップ
            continue

        name = ws.cell(row=row_idx, column=COL_NAME + col_offset).value or ''
        emp_type = ws.cell(row=row_idx, column=COL_TYPE + col_offset).value or ''
        hours = ws.cell(row=row_idx, column=COL_HOURS + col_offset).value or 0
        hourly = ws.cell(row=row_idx, column=COL_HOURLY + col_offset).value or 0

        # 12か月分の給与を読み取り
        monthly = []
        for m in range(12):
            val = ws.cell(row=row_idx, column=COL_M1 + col_offset + m).value
            monthly.append(float(val) if val is not None else None)

        # 全角スペースを半角に
        name = str(name).replace('\u3000', ' ').strip()

        employees.append(WageEmployee(
            no=int(no_val),
            name=name,
            employment_type=str(emp_type),
            monthly_avg_hours=float(hours),
            hourly_rate=float(hourly),
            monthly_wages=monthly,
        ))

    wb.close()
    logger.info(f'賃金台帳読み取り: {len(employees)}名')
    return employees


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

    mw_r6 = result.min_wage_r6   # R6改定後（下限）
    mw_r7 = result.min_wage_r7   # R7改定後（上限）

    # ── 加点措置① ──
    # 判定にはF列の時給を使用する（月給/時間の計算は勤務時間変動で不正確）
    # 正社員: F列の時給 = 月給相当額/月間所定時間
    # パート: F列の時給 = 契約上の時給
    target_months = list(range(0, 12))
    month_names = ['1月', '2月', '3月', '4月', '5月', '6月',
                   '7月', '8月', '9月', '10月', '11月', '12月']

    months_meeting_criteria = []

    for m_idx in target_months:
        total_emps = 0
        under_r7_emps = 0
        month_detail = {
            'month': month_names[m_idx],
            'total': 0,
            'under_r7': 0,
            'ratio': 0.0,
            'meets_30pct': False,
            'employees': [],
        }

        for emp in employees:
            # その月にデータがない従業員はスキップ
            if emp.monthly_wages[m_idx] is None:
                continue
            if emp.hourly_rate <= 0:
                continue

            total_emps += 1
            hourly = emp.hourly_rate

            # 時給がR6改定後以上かつR7改定後未満 = 加点対象
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
                months_meeting_criteria.append(month_names[m_idx])

        result.bonus1_details.append(month_detail)

    result.bonus1_months_met = months_meeting_criteria
    result.bonus1_eligible = len(months_meeting_criteria) >= 3

    logger.info(
        f'加点措置①: {len(months_meeting_criteria)}か月が条件達成 '
        f'→ {"対象" if result.bonus1_eligible else "対象外"}'
    )

    # ── 加点措置② ──
    # 事業場内最低賃金 = 全従業員のうち最も低い時給（F列）
    # R7年7月時点と直近月を比較して+63円以上かどうか
    july_idx = 6  # 7月

    # 7月にデータがある従業員の時給
    july_hourly_rates = [
        emp.hourly_rate for emp in employees
        if emp.monthly_wages[july_idx] is not None and emp.hourly_rate > 0
    ]

    # 直近月を決定
    if latest_month_idx is None:
        for m in range(11, -1, -1):
            if any(emp.monthly_wages[m] is not None for emp in employees):
                latest_month_idx = m
                break
        if latest_month_idx is None:
            latest_month_idx = 11

    latest_hourly_rates = [
        emp.hourly_rate for emp in employees
        if emp.monthly_wages[latest_month_idx] is not None and emp.hourly_rate > 0
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
      各期間: No, 氏名, 事業場所在地, 最低賃金(R6), 最低賃金(R7), 基本給, 時間換算給与, 判定, 備考
      データ行は18行目から

    Args:
        selected_months: 判定に使う3か月のインデックス（0=1月）。
                         Noneの場合は条件達成月から選ぶ。
    """
    wb = openpyxl.load_workbook(str(template_path))
    ws = wb[wb.sheetnames[0]]

    # 入力する3か月を決定
    if selected_months is None:
        if result.bonus1_months_met:
            month_name_to_idx = {f'{i+1}月': i for i in range(12)}
            selected_months = [month_name_to_idx[m] for m in result.bonus1_months_met[:3]]
        else:
            # 条件未達でもデータがある月を3つ選ぶ
            all_months = [d for d in result.bonus1_details if d['total'] > 0]
            selected_months = [
                ['1月','2月','3月','4月','5月','6月',
                 '7月','8月','9月','10月','11月','12月'].index(d['month'])
                for d in all_months[:3]
            ]

    # 3つの期間の開始列（1始まり）
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
            if e.monthly_wages[m_idx] is not None and e.hourly_rate > 0
        ]

        for i, emp in enumerate(active_emps):
            row = DATA_START_ROW + i
            wage = emp.monthly_wages[m_idx]

            ws.cell(row=row, column=cols['no'], value=i + 1)
            ws.cell(row=row, column=cols['name'], value=emp.name)
            ws.cell(row=row, column=cols['pref'], value=result.prefecture)
            ws.cell(row=row, column=cols['wage'], value=wage)
            ws.cell(row=row, column=cols['hourly'], value=round(emp.hourly_rate))

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
      期間①(R7年7月): B-H列, 期間②(直近月): J-P列
      各期間: No, 氏名, 事業場所在地, 最低賃金, 基本給, 時間換算給与, 備考
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
            if e.monthly_wages[m_idx] is not None and e.hourly_rate > 0
        ]

        for i, emp in enumerate(active_emps):
            row = DATA_START_ROW + i
            wage = emp.monthly_wages[m_idx]

            ws.cell(row=row, column=cols['no'], value=i + 1)
            ws.cell(row=row, column=cols['name'], value=emp.name)
            ws.cell(row=row, column=cols['pref'], value=result.prefecture)
            ws.cell(row=row, column=cols['wage'], value=wage)
            ws.cell(row=row, column=cols['hourly'], value=round(emp.hourly_rate))

    wb.save(str(output_path))
    wb.close()
    logger.info(f'加点措置②シート保存: {output_path}')
    return output_path
