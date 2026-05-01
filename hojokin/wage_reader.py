# -*- coding: utf-8 -*-
"""
賃金台帳Excel読み取り + 加点措置判定

対応方針:
  ヘッダー別名辞書で正規化マッチ → 月の並び方を値ベースで自動判定する
  柔軟パーサー(_read_flexible) を採用。複数の実フォーマット差異を1本で吸収する。
  配置パターン:
    (a) 1月〜12月が列見出し   … 集計表型（1行1人）
    (b) 対象年月/給与年月列あり … 月別行型（1人×N行、月は明示列）
    (c) 先頭列に YYYYMM 値     … YYYYMM月次型（給与ソフト出力・1ファイル1人のケース）
  個人台帳型（行=項目、列=月、月度給与ブロック）は別ルートで温存。

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
import unicodedata
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
    monthly_hours: list[float | None] = field(
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
# 柔軟パーサー（集計表型 / 月別行型 / YYYYMM月次型を統一処理）
# ============================================================

# ヘッダー別名辞書（正規化後に部分一致 or 完全一致で判定）
_HEADER_ALIASES = {
    'name':       ['氏名', '従業員氏名', '社員氏名', '名前'],
    'emp_id':     ['従業員番号', '従業員コード', '社員番号', 'no', 'ＮＯ', 'Ｎｏ'],
    'emp_type':   ['雇用形態', '区分', '従業員区分'],
    'base_wage':  ['基本給'],
    'hourly_wage':['基本給(時給)', '時給', '時間給'],
    'hours':      ['所定労働時間', '労働時間', '月間平均時間', '平均時間'],
    'total':      ['支給合計額', '支給合計', '総支給額', '総支給',
                   '課税支給合計', '差引支給合計'],
    'paid_date':  ['支給日', '支払日'],
    'month_col':  ['対象年月', '給与年月', '支給年月', '年月'],
}


def _norm(val) -> str:
    """文字列を正規化（NFKC・空白除去・小文字化）"""
    if val is None:
        return ''
    s = unicodedata.normalize('NFKC', str(val))
    s = s.replace('\u3000', '').replace(' ', '').strip()
    return s.lower()


def _match_alias(val: str, aliases: list[str]) -> bool:
    v = _norm(val)
    if not v:
        return False
    for a in aliases:
        na = _norm(a)
        if v == na or (len(na) >= 2 and na in v):
            return True
    return False


def _detect_field_map(ws, header_row: int) -> dict[str, int]:
    """指定行をヘッダーと見なし、各フィールドの列番号を割り出す"""
    fmap: dict[str, int] = {}
    month_cols: dict[int, int] = {}
    for c in range(1, min(ws.max_column + 1, 80)):
        val = ws.cell(header_row, c).value
        s = _norm(val)
        if not s:
            continue
        # 1月〜12月 → 集計表型
        m = re.fullmatch(r'(\d{1,2})月', s)
        if m:
            idx = int(m.group(1)) - 1
            if 0 <= idx <= 11:
                month_cols[idx] = c
                continue
        for key, aliases in _HEADER_ALIASES.items():
            if key in fmap:
                continue
            if _match_alias(val, aliases):
                fmap[key] = c
                break
    if month_cols:
        fmap['_month_cols'] = month_cols  # type: ignore[assignment]
    return fmap


def _find_header_rows(ws) -> list[tuple[int, dict]]:
    """シート内のヘッダー行を全て発見（給与/賞与セクション両方を取るため）"""
    rows = []
    for r in range(1, min(ws.max_row + 1, 40)):
        fmap = _detect_field_map(ws, r)
        has_name = 'name' in fmap
        has_total = 'total' in fmap
        has_month_cols = '_month_cols' in fmap
        if has_name and (has_total or has_month_cols):
            rows.append((r, fmap))
    return rows


def _parse_month(val, paid_date_val=None) -> int | None:
    """セル値から月インデックス(0-11)を抽出。YYYYMM数値/'〇年〇月'/支給日まで対応"""
    if val is not None:
        # YYYYMM 数値（例: 202503 → 3月=index2）
        if isinstance(val, (int, float)):
            n = int(val)
            if 100000 <= n <= 999999:
                month = n % 100
                if 1 <= month <= 12:
                    return month - 1
        s = str(val)
        # '2025年3月' 等
        m = re.search(r'(\d{4})[年/\-](\d{1,2})', s)
        if m:
            month = int(m.group(2))
            if 1 <= month <= 12:
                return month - 1
        # '3月' 単独
        m = re.search(r'(\d{1,2})月', s)
        if m:
            month = int(m.group(1))
            if 1 <= month <= 12:
                return month - 1
        # 純粋なYYYYMM文字列
        m = re.fullmatch(r'\d{6}', s.strip())
        if m:
            month = int(s.strip()) % 100
            if 1 <= month <= 12:
                return month - 1
    # フォールバック: 支給日（例: 2025/07/10）
    if paid_date_val is not None:
        s = str(paid_date_val)
        m = re.search(r'\d{4}[/\-年](\d{1,2})', s)
        if m:
            month = int(m.group(1))
            if 1 <= month <= 12:
                return month - 1
    return None


def _to_float(val) -> float | None:
    if val is None:
        return None
    try:
        f = float(val)
        return f
    except (ValueError, TypeError):
        return None


def _new_emp_record(name: str, emp_type: str = '') -> dict:
    return {
        'name': name,
        'employment_type': emp_type,
        'monthly_wages': [None] * 12,
        'monthly_hourly_rates': [None] * 12,
        'monthly_hours': [None] * 12,
        'hourly_rate_flat': 0.0,
        'avg_hours_flat': 0.0,
    }


def _parse_section_rowwise(ws, header_row: int, end_row: int,
                           fmap: dict, emp_data: dict) -> None:
    """月別行型 or YYYYMM月次型 のデータ行を処理（月=行方向）"""
    col_name = fmap['name']
    col_total = fmap.get('total')
    col_type = fmap.get('emp_type')
    col_month = fmap.get('month_col')
    col_base = fmap.get('base_wage')
    col_hours = fmap.get('hours')
    col_paid = fmap.get('paid_date')

    for r in range(header_row + 1, end_row):
        name_val = ws.cell(r, col_name).value
        if not name_val:
            continue
        name = str(name_val).replace('\u3000', ' ').strip()
        if not name:
            continue

        # 月を特定: month_col > 先頭列YYYYMM > 支給日
        month_idx = None
        if col_month:
            month_idx = _parse_month(ws.cell(r, col_month).value)
        if month_idx is None:
            # 先頭列がYYYYMM(例: 202503)
            month_idx = _parse_month(ws.cell(r, 1).value)
        if month_idx is None and col_paid:
            month_idx = _parse_month(None, ws.cell(r, col_paid).value)
        if month_idx is None:
            continue

        if name not in emp_data:
            et = ''
            if col_type:
                et = str(ws.cell(r, col_type).value or '')
            emp_data[name] = _new_emp_record(name, et)

        rec = emp_data[name]

        if col_total:
            t = _to_float(ws.cell(r, col_total).value)
            if t is not None:
                # 給与＋賞与セクション両方が来たら加算（同月の別セクション）
                existing = rec['monthly_wages'][month_idx]
                rec['monthly_wages'][month_idx] = (existing or 0) + t

        if col_base and col_hours:
            base = _to_float(ws.cell(r, col_base).value)
            hours = _to_float(ws.cell(r, col_hours).value)
            if base is not None and hours is not None and hours > 0:
                rec['monthly_hours'][month_idx] = hours
                rec['monthly_hourly_rates'][month_idx] = base / hours


def _parse_section_summary(ws, header_row: int, fmap: dict,
                           emp_data: dict) -> None:
    """集計表型（列=月）のデータ行を処理"""
    col_name = fmap['name']
    col_type = fmap.get('emp_type')
    col_hours = fmap.get('hours')
    col_hourly = fmap.get('hourly_wage')
    month_cols: dict[int, int] = fmap['_month_cols']  # type: ignore[assignment]

    for r in range(header_row + 1, ws.max_row + 1):
        name_val = ws.cell(r, col_name).value
        if not name_val:
            continue
        name = str(name_val).replace('\u3000', ' ').strip()
        if not name:
            continue

        if name not in emp_data:
            et = ''
            if col_type:
                et = str(ws.cell(r, col_type).value or '')
            emp_data[name] = _new_emp_record(name, et)
        rec = emp_data[name]

        if col_hours:
            h = _to_float(ws.cell(r, col_hours).value)
            if h is not None:
                rec['avg_hours_flat'] = h
        if col_hourly:
            hr = _to_float(ws.cell(r, col_hourly).value)
            if hr is not None:
                rec['hourly_rate_flat'] = hr

        for midx, c in month_cols.items():
            v = _to_float(ws.cell(r, c).value)
            if v is not None:
                existing = rec['monthly_wages'][midx]
                rec['monthly_wages'][midx] = (existing or 0) + v
                if rec['hourly_rate_flat'] > 0:
                    rec['monthly_hourly_rates'][midx] = rec['hourly_rate_flat']


def _read_flexible(wb: openpyxl.Workbook,
                   emp_data: dict | None = None) -> dict:
    """柔軟パーサー本体（emp_dataに蓄積）"""
    if emp_data is None:
        emp_data = {}

    for ws in wb.worksheets:
        header_rows = _find_header_rows(ws)
        if not header_rows:
            continue

        for i, (hr, fmap) in enumerate(header_rows):
            end = (header_rows[i + 1][0]
                   if i + 1 < len(header_rows) else ws.max_row + 1)
            if '_month_cols' in fmap:
                _parse_section_summary(ws, hr, fmap, emp_data)
            else:
                _parse_section_rowwise(ws, hr, end, fmap, emp_data)

    return emp_data


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
    monthly_hours_list: list[float | None] = [None] * 12

    for m_idx, col in month_cols.items():
        # 支給合計
        if total_row:
            val = ws.cell(total_row, col).value
            if val is not None:
                try:
                    monthly_wages[m_idx] = float(val)
                except (ValueError, TypeError):
                    pass

        # 月別の労働時間
        if hours_row:
            hours_val = ws.cell(hours_row, col).value
            if hours_val is not None:
                h = _parse_hours_str(hours_val)
                if h > 0:
                    monthly_hours_list[m_idx] = h

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

    # 平均労働時間（月別データがあればそこから算出）
    valid_hours = [h for h in monthly_hours_list if h is not None and h > 0]
    avg_hours = sum(valid_hours) / len(valid_hours) if valid_hours else 0

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
        monthly_hours=monthly_hours_list,
    )


# ============================================================
# メイン読み取り関数
# ============================================================

def _is_individual_ledger(wb: openpyxl.Workbook) -> bool:
    """個人台帳型（月度給与ブロック）かどうか判定"""
    for ws in wb.worksheets:
        for r in range(1, min(ws.max_row + 1, 30)):
            for c in range(1, min(ws.max_column + 1, 30)):
                val = str(ws.cell(r, c).value or '')
                if '月度給与' in val:
                    return True
    return False


def _emp_dict_to_list(emp_data: dict) -> list[WageEmployee]:
    """内部dict表現 → WageEmployeeリスト変換"""
    employees = []
    for i, (name, data) in enumerate(emp_data.items()):
        hourly_rates = [h for h in data['monthly_hourly_rates'] if h is not None]
        hours_list = [h for h in data['monthly_hours'] if h is not None]
        if hourly_rates:
            avg_hourly = sum(hourly_rates) / len(hourly_rates)
        else:
            avg_hourly = data.get('hourly_rate_flat', 0.0)
        if hours_list:
            avg_hours = sum(hours_list) / len(hours_list)
        else:
            avg_hours = data.get('avg_hours_flat', 0.0)

        employees.append(WageEmployee(
            no=i + 1,
            name=data['name'],
            employment_type=data['employment_type'],
            monthly_avg_hours=round(avg_hours, 1),
            hourly_rate=round(avg_hourly, 1),
            monthly_wages=data['monthly_wages'],
            monthly_hourly_rates=data['monthly_hourly_rates'],
            monthly_hours=data['monthly_hours'],
        ))
    return employees


def read_wage_ledger(file_path: Path) -> list[WageEmployee]:
    """
    単一の賃金台帳Excelを読み取る。
    個人台帳型（月度給与ブロック）は専用パーサー、それ以外は柔軟パーサーで統一処理。
    """
    wb = openpyxl.load_workbook(str(file_path), data_only=True)

    if _is_individual_ledger(wb):
        employees = _read_individual_ledger(wb)
        fmt = 'individual'
    else:
        emp_data = _read_flexible(wb)
        employees = _emp_dict_to_list(emp_data)
        fmt = 'flexible'

    wb.close()
    logger.info(f'賃金台帳読み取り完了: {file_path.name} → {len(employees)}名 ({fmt})')
    return employees


def _workbook_to_tsv(wb: openpyxl.Workbook, file_label: str) -> str:
    """ワークブック全シートをTSV文字列に変換（AI入力用）。"""
    parts: list[str] = [f'### ファイル: {file_label} ###']
    for ws in wb.worksheets:
        parts.append(f'\n--- シート: {ws.title} ---')
        for row in ws.iter_rows(values_only=True):
            # 末尾の None だけのセルは無視して圧縮
            cells = list(row)
            while cells and cells[-1] is None:
                cells.pop()
            if not cells:
                continue
            line = '\t'.join('' if v is None else str(v) for v in cells)
            parts.append(line)
    return '\n'.join(parts)


def _validate_ai_employee(emp: dict) -> tuple[bool, str]:
    """AI抽出した1従業員データの妥当性チェック。(OK?, エラー理由)"""
    name = emp.get('name')
    if not name or not isinstance(name, str):
        return False, 'name が空または文字列でない'
    monthly_wages = emp.get('monthly_wages')
    monthly_hours = emp.get('monthly_hours')
    if not isinstance(monthly_wages, list) or len(monthly_wages) != 12:
        return False, f'monthly_wages が12要素のリストでない (len={len(monthly_wages) if isinstance(monthly_wages, list) else "N/A"})'
    if not isinstance(monthly_hours, list) or len(monthly_hours) != 12:
        return False, f'monthly_hours が12要素のリストでない'
    # 金額の現実的範囲チェック (0〜1000万円/月)
    for i, w in enumerate(monthly_wages):
        if w is None:
            continue
        if not isinstance(w, (int, float)) or w < 0 or w > 10_000_000:
            return False, f'{i+1}月の給与額が異常: {w}'
    # 労働時間の現実的範囲チェック (0〜400時間/月)
    for i, h in enumerate(monthly_hours):
        if h is None:
            continue
        if not isinstance(h, (int, float)) or h < 0 or h > 400:
            return False, f'{i+1}月の労働時間が異常: {h}'
    return True, ''


def _ai_data_to_wage_employees(ai_data: list[dict]) -> list[WageEmployee]:
    """AI抽出データを WageEmployee リストに変換（バリデーション付き）。"""
    employees: list[WageEmployee] = []
    for i, emp in enumerate(ai_data):
        if not isinstance(emp, dict):
            logger.warning(f'AI抽出: index={i} が辞書でないためスキップ: {type(emp).__name__}')
            continue
        ok, reason = _validate_ai_employee(emp)
        if not ok:
            logger.warning(f'AI抽出: index={i} ({emp.get("name", "?")}) バリデーション失敗: {reason}')
            continue

        monthly_wages = [
            float(w) if w is not None else None for w in emp['monthly_wages']
        ]
        monthly_hours = [
            float(h) if h is not None else None for h in emp['monthly_hours']
        ]
        # 月平均労働時間（None除外で平均）
        valid_hours = [h for h in monthly_hours if h is not None and h > 0]
        avg_hours = sum(valid_hours) / len(valid_hours) if valid_hours else 0.0
        # 月別時給は AI 出力に含めない方針 → monthly_wages / monthly_hours から逆算可能だが現状は空でOK
        employees.append(WageEmployee(
            no=i + 1,
            name=str(emp['name']).strip(),
            employment_type=str(emp.get('employment_type', '') or '').strip(),
            monthly_avg_hours=round(avg_hours, 1),
            hourly_rate=0.0,
            monthly_wages=monthly_wages,
            monthly_hourly_rates=[None] * 12,
            monthly_hours=monthly_hours,
        ))
    return employees


def read_wage_ledgers_with_ai(
    file_paths: list[Path],
    extractor,
    fiscal_period_hint: str | None = None,
) -> list[WageEmployee]:
    """
    AI による賃金台帳読み取り。
    全ファイルを TSV に変換して1回の API 呼び出しで全従業員を抽出する。
    バリデーション失敗時は空リストを返す（呼び出し側で fallback 判断）。
    """
    if not file_paths:
        return []

    tsv_blocks: list[str] = []
    for path in file_paths:
        try:
            wb = openpyxl.load_workbook(str(path), data_only=True)
        except Exception as e:
            logger.warning(f'賃金台帳読込失敗(AI経路): {path.name} ({e})')
            continue
        tsv_blocks.append(_workbook_to_tsv(wb, path.name))
        wb.close()

    if not tsv_blocks:
        return []

    combined_tsv = '\n\n'.join(tsv_blocks)
    logger.info(
        f'AI抽出開始: {len(file_paths)}ファイル '
        f'→ TSV {len(combined_tsv):,}文字'
        + (f' (前事業年度ヒント: {fiscal_period_hint})' if fiscal_period_hint else '')
    )

    try:
        ai_data = extractor.extract_wage_ledger(combined_tsv, fiscal_period_hint)
    except Exception as e:
        logger.error(f'AI抽出例外: {e}', exc_info=True)
        return []

    employees = _ai_data_to_wage_employees(ai_data)
    logger.info(
        f'AI抽出結果: 入力{len(ai_data)}名 → 妥当{len(employees)}名'
    )
    return employees


def read_wage_ledgers(
    file_paths: list[Path],
    extractor=None,
    fiscal_period_hint: str | None = None,
) -> list[WageEmployee]:
    """
    複数の賃金台帳ファイルを読み、同名の従業員をマージして返す。
    1人1ファイル運用（給与ソフト出力）と、1ファイルに全員を入れる運用の両方に対応。

    extractor が渡され、かつ環境変数 USE_AI_WAGE_EXTRACTION が有効な場合は
    AI 抽出を優先し、結果が空なら決定論パーサーにフォールバックする。
    """
    if not file_paths:
        return []

    # AI 経路（extractor がある場合）
    if extractor is not None:
        from .config import USE_AI_WAGE_EXTRACTION
        if USE_AI_WAGE_EXTRACTION:
            ai_employees = read_wage_ledgers_with_ai(
                file_paths, extractor, fiscal_period_hint
            )
            if ai_employees:
                logger.info(f'賃金台帳合算結果(AI): {len(ai_employees)}名 ({len(file_paths)}ファイル)')
                return ai_employees
            logger.warning('AI抽出が0件を返したため、決定論パーサーにフォールバック')

    # 決定論パーサー経路（フォールバック or extractor なし）
    individual_ledger_paths = []
    merged_emp_data: dict = {}

    for path in file_paths:
        try:
            wb = openpyxl.load_workbook(str(path), data_only=True)
        except Exception as e:
            logger.warning(f'賃金台帳読込失敗: {path.name} ({e})')
            continue

        if _is_individual_ledger(wb):
            individual_ledger_paths.append(path)
            wb.close()
            continue

        before = len(merged_emp_data)
        _read_flexible(wb, merged_emp_data)
        wb.close()
        logger.info(
            f'賃金台帳読み取り: {path.name} '
            f'→ 追加/更新 {len(merged_emp_data) - before}名 (累計{len(merged_emp_data)}名)'
        )

    employees = _emp_dict_to_list(merged_emp_data)

    # 個人台帳型ファイルは別途パース（統合が複雑なためファイル単位で結合）
    for path in individual_ledger_paths:
        wb = openpyxl.load_workbook(str(path), data_only=True)
        extra = _read_individual_ledger(wb)
        wb.close()
        logger.info(f'賃金台帳読み取り(個人台帳型): {path.name} → {len(extra)}名')
        for e in extra:
            e.no = len(employees) + 1
            employees.append(e)

    logger.info(f'賃金台帳合算結果(決定論): {len(employees)}名 ({len(file_paths)}ファイル)')
    return employees


# ============================================================
# 賃金台帳一覧Excel出力（チェック用）
# ============================================================

def export_wage_ledger_summary(
    employees: list[WageEmployee],
    output_path: Path,
    company_name: str = '',
) -> Path:
    """
    賃金台帳から読み取ったデータを一覧Excelに出力（チェック用）

    出力内容:
      左ブロック  : 月別課税対象額（12か月）+ 年間合計賃金
      右ブロック  : 月別労働時間（12か月）+ 年間合計時間 + 月平均労働時間
    """
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
    from openpyxl.utils import get_column_letter

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '賃金台帳一覧'

    # スタイル定義
    header_fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
    group_fill = PatternFill(start_color='8FAADC', end_color='8FAADC', fill_type='solid')
    header_font_white = Font(bold=True, size=10, color='FFFFFF')
    number_fmt = '#,##0'
    hours_fmt = '#,##0.0'
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin'),
    )

    # タイトル行
    title = '賃金台帳 読取データ一覧'
    if company_name:
        title = f'{company_name} — {title}'
    ws.cell(row=1, column=1, value=title).font = Font(bold=True, size=12)
    ws.cell(row=2, column=1, value='※この一覧は賃金台帳から機械的に読み取ったデータです（AI生成ではありません）')
    ws.cell(row=2, column=1).font = Font(size=9, color='666666')

    # 列レイアウト
    # 1: No, 2: 従業員名, 3: 雇用形態,
    # 4-15: 1月〜12月 賃金, 16: 年間合計賃金,
    # 17-28: 1月〜12月 時間, 29: 年間合計時間, 30: 月平均労働時間
    wage_start = 4
    wage_total_col = wage_start + 12  # 16
    hours_start = wage_total_col + 1  # 17
    hours_total_col = hours_start + 12  # 29
    avg_hours_col = hours_total_col + 1  # 30

    # グループヘッダー（4行目）
    group_row = 4
    ws.cell(row=group_row, column=wage_start, value='月別課税対象額（円）')
    ws.merge_cells(start_row=group_row, start_column=wage_start,
                   end_row=group_row, end_column=wage_total_col)
    ws.cell(row=group_row, column=hours_start, value='月別労働時間')
    ws.merge_cells(start_row=group_row, start_column=hours_start,
                   end_row=group_row, end_column=avg_hours_col)
    for c in (wage_start, hours_start):
        cell = ws.cell(row=group_row, column=c)
        cell.font = header_font_white
        cell.fill = group_fill
        cell.alignment = Alignment(horizontal='center')
        cell.border = thin_border

    # 列ヘッダー（5行目）
    header_row = 5
    headers = (
        ['No', '従業員名', '雇用形態']
        + MONTH_NAMES + ['年間合計']
        + MONTH_NAMES + ['年間合計', '月平均']
    )
    for c, h in enumerate(headers, 1):
        cell = ws.cell(row=header_row, column=c, value=h)
        cell.font = header_font_white
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center')
        cell.border = thin_border

    # データ行
    for i, emp in enumerate(employees):
        r = header_row + 1 + i
        ws.cell(row=r, column=1, value=emp.no).border = thin_border
        ws.cell(row=r, column=2, value=emp.name).border = thin_border
        ws.cell(row=r, column=3, value=emp.employment_type).border = thin_border

        annual_wage = 0.0
        for m in range(12):
            cell = ws.cell(row=r, column=wage_start + m)
            cell.border = thin_border
            val = emp.monthly_wages[m]
            if val is not None:
                cell.value = val
                cell.number_format = number_fmt
                annual_wage += val
        wage_total_cell = ws.cell(row=r, column=wage_total_col, value=annual_wage)
        wage_total_cell.number_format = number_fmt
        wage_total_cell.font = Font(bold=True)
        wage_total_cell.border = thin_border

        annual_hours = 0.0
        has_any_hours = False
        for m in range(12):
            cell = ws.cell(row=r, column=hours_start + m)
            cell.border = thin_border
            val = emp.monthly_hours[m] if m < len(emp.monthly_hours) else None
            if val is not None and val > 0:
                cell.value = val
                cell.number_format = hours_fmt
                annual_hours += val
                has_any_hours = True

        # 年間合計時間（月別データが無ければ 月平均×月数 で代用）
        hours_total_cell = ws.cell(row=r, column=hours_total_col)
        hours_total_cell.border = thin_border
        hours_total_cell.number_format = hours_fmt
        hours_total_cell.font = Font(bold=True)
        if has_any_hours:
            hours_total_cell.value = round(annual_hours, 1)
        elif emp.monthly_avg_hours > 0:
            months_with_wage = sum(
                1 for w in emp.monthly_wages if w is not None
            )
            hours_total_cell.value = round(
                emp.monthly_avg_hours * months_with_wage, 1
            )

        avg_hours_cell = ws.cell(row=r, column=avg_hours_col,
                                 value=emp.monthly_avg_hours)
        avg_hours_cell.number_format = hours_fmt
        avg_hours_cell.border = thin_border

    # 列幅調整
    ws.column_dimensions['A'].width = 5
    ws.column_dimensions['B'].width = 14
    ws.column_dimensions['C'].width = 14
    for c in range(wage_start, avg_hours_col + 1):
        ws.column_dimensions[get_column_letter(c)].width = 11
    ws.column_dimensions[get_column_letter(wage_total_col)].width = 13
    ws.column_dimensions[get_column_letter(hours_total_col)].width = 13
    ws.column_dimensions[get_column_letter(avg_hours_col)].width = 11

    wb.save(str(output_path))
    wb.close()
    logger.info(f'賃金台帳一覧出力: {output_path} ({len(employees)}名)')
    return output_path


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
