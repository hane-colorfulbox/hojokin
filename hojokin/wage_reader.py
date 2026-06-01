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
import pandas as pd

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
    # データソース（抽出根拠の元ファイル名）。複数ファイル統合時は '統合(Nファイル)'
    source_file: str = ''
    # 年間通勤手当（非課税分のみ、円）。AI が抽出できた場合のみ > 0。
    # R216 算定対象から除外するため、ツール側で在籍月数で均等割して各月から減算する。
    # 既存呼出は 0.0（未設定）で動作、賃金台帳の作成タスクで S列に書き出す用途に使う。
    annual_transport_allowance: float = 0.0

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
    # 加点措置②で使用した「直近月」のインデックス（0=1月, 11=12月）。
    # judge_bonus_points が動的決定した月を保持し、fill_bonus_sheet_2 で
    # 同じ月を出力シートに反映するために使う（画面判定とExcelの不整合防止）。
    latest_month_idx: int | None = None


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
    # 非課税額（通勤費の非課税分など）。total(gross) 採用時に差し引いて課税額へ補正。
    # 公募要領 p.10「給与所得として課税対象」+ 国税庁 No.2585（通勤手当は限度額まで非課税）。
    # ★宣言順を total_taxable / total より前に置く理由★: _match_alias は部分一致を許すため
    #   「非課税支給合計」が「課税支給合計」(total_taxable) や「支給合計」(total) の別名に
    #   部分一致してしまう。nontax を先に走査して『非課税◯◯』を確実に nontax へ確定させ、
    #   課税列・支給合計列への誤割当（R216 過小計上）を防ぐ。
    'nontax':     ['非課税額', '非課税合計', '非課税支給合計', '非課税分'],
    # 課税支給合計（R216 が求める「給与所得として課税対象」の合計）。最優先で採用。
    # 給与ソフトにより「課税分合計」「課税給与合計」等の表記揺れがあるため広めに拾う。
    # 注: 「課税対象額」は社保控除後の所得税課税ベースを指すソフトもあり ambiguous なので含めない。
    'total_taxable': ['課税支給合計', '課税分合計', '課税給与合計', '課税支給額'],
    # 支給合計（非課税通勤費等を含む可能性あり）。total_taxable が無い時のフォールバック。
    # この値を採用する場合は nontax 列があれば差し引いて課税額に補正する。
    # 注: 「差引支給合計」は社保・税控除後の手取りであり給与支給総額ではない（R216 過小計上に
    #     なる）ため、ここには含めない。
    'total':      ['支給合計額', '支給合計', '総支給額', '総支給'],
    'paid_date':  ['支給日', '支払日'],
    'month_col':  ['対象年月', '給与年月', '支給年月', '年月'],
    # 年間通勤手当（非課税分）。集計表型のみ有効。在籍月数で均等割して
    # monthly_wages から減算する（R216 公募要領 p.10 の課税給与定義に揃える）
    'transport_annual': ['年間通勤手当', '年間通勤費',
                         '通勤手当(年間)', '通勤費(年間)',
                         '非課税通勤手当(年間)'],
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
                # 部分一致の落とし穴対策（R216 母数の取り違え防止）:
                #   - '差引支給合計' は「支給合計」に部分一致するが手取り(控除後)なので総額系に割り当てない
                #   - '非課税支給合計' は「課税支給合計」「支給合計」に部分一致するが非課税分なので同様
                # nontax を先に走査しているので非課税列は通常そちらで確定するが、多重防御として弾く。
                nval = _norm(val)
                if key in ('total', 'total_taxable') and ('差引' in nval or '非課税' in nval):
                    continue
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
        # 課税支給合計（total_taxable）も合計列として認める。これが無いと、
        # 課税支給合計のみ（支給合計列なし・月列なし）の行型台帳でヘッダーが
        # 検出されず当該従業員の給与が丸ごと欠落し R216=0 になる。
        has_total = ('total' in fmap) or ('total_taxable' in fmap)
        has_month_cols = '_month_cols' in fmap
        if has_name and (has_total or has_month_cols):
            rows.append((r, fmap))
    return rows


def _parse_month(val, paid_date_val=None) -> int | None:
    """セル値から月インデックス(0-11)を抽出。YYYYMM数値/'〇年〇月'/支給日まで対応。

    給与ソフトの特殊コード対応:
      YYYYMM の月部分が 21 → 7月（夏季賞与） Index 6
      YYYYMM の月部分が 22 → 12月（冬季賞与） Index 11
      21,22 は一部医療機関向け給与ソフト出力で観測された賞与識別子。

    重要: 21/22 マッピングは給与ソフト依存のため、paid_date_val（支給日）が
    与えられている場合はそちらを優先する。これにより「21=春賞与/22=夏賞与」のように
    異なる慣習を持つソフトでも実支給月で正しく月マッピングできる。
    """
    def _normal_month_to_idx(m: int) -> int | None:
        """通常の月（1〜12）のみ受け付ける。21/22 は除外"""
        if 1 <= m <= 12:
            return m - 1
        return None

    def _bonus_code_to_idx(m: int) -> int | None:
        """賞与識別コード（21/22）のフォールバック。paid_date が無いときのみ使う"""
        if m == 21:
            return 6   # 夏季賞与 → 7月に加算（給与ソフト慣習。実 paid_date があれば paid_date 優先）
        if m == 22:
            return 11  # 冬季賞与 → 12月に加算
        return None

    # ── ステップ1: 通常の月（1〜12）の判定 ──
    bonus_code: int | None = None  # 21/22 を後回しにするため記録
    if val is not None:
        # YYYYMM 数値
        if isinstance(val, (int, float)):
            n = int(val)
            if 100000 <= n <= 999999:
                month = n % 100
                idx = _normal_month_to_idx(month)
                if idx is not None:
                    return idx
                if month in (21, 22):
                    bonus_code = month
        s = str(val)
        # '2025年3月' 等
        m = re.search(r'(\d{4})[年/\-](\d{1,2})', s)
        if m:
            idx = _normal_month_to_idx(int(m.group(2)))
            if idx is not None:
                return idx
        # '3月' 単独
        m = re.search(r'(\d{1,2})月', s)
        if m:
            idx = _normal_month_to_idx(int(m.group(1)))
            if idx is not None:
                return idx
        # 純粋なYYYYMM文字列
        m = re.fullmatch(r'\d{6}', s.strip())
        if m:
            month = int(s.strip()) % 100
            idx = _normal_month_to_idx(month)
            if idx is not None:
                return idx
            if month in (21, 22) and bonus_code is None:
                bonus_code = month
    # ── ステップ2: 支給日（paid_date）から月を取る（21/22 ボーナスコードより優先） ──
    # paid_date があれば、それが「実際に支給された月」なので最も信頼できる
    if paid_date_val is not None:
        s = str(paid_date_val)
        m = re.search(r'\d{4}[/\-年](\d{1,2})', s)
        if m:
            month = int(m.group(1))
            if 1 <= month <= 12:
                return month - 1

    # ── ステップ3: 21/22 賞与コードのフォールバック ──
    # paid_date が無い場合のみ、給与ソフト慣習の固定マッピングを使う
    if bonus_code is not None:
        idx = _bonus_code_to_idx(bonus_code)
        if idx is not None:
            return idx
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
    col_total_taxable = fmap.get('total_taxable')
    col_total = fmap.get('total')
    col_nontax = fmap.get('nontax')
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

        # 月を特定: 各 _parse_month 呼出に paid_date_val を渡す。
        # _parse_month は内部で「通常月(1-12) > paid_date > 21/22 賞与コード」の優先順序で
        # 月を決定するため、paid_date を毎回渡しておけば 21/22 のような ambiguous な
        # ボーナスコードに対しても、実支給月（paid_date）が優先される。
        paid_val = ws.cell(r, col_paid).value if col_paid else None
        month_idx = None
        if col_month:
            month_idx = _parse_month(ws.cell(r, col_month).value, paid_val)
        if month_idx is None:
            # 先頭列がYYYYMM(例: 202503)
            month_idx = _parse_month(ws.cell(r, 1).value, paid_val)
        if month_idx is None:
            continue

        if name not in emp_data:
            et = ''
            if col_type:
                et = str(ws.cell(r, col_type).value or '')
            emp_data[name] = _new_emp_record(name, et)

        rec = emp_data[name]

        # 月の給与額（課税支給合計）を決定。優先順（公募要領 p.10「課税対象となる経費」準拠）:
        #   1) 課税支給合計/課税分合計 列をそのまま採用（最も確実）
        #   2) 支給合計 − 非課税額（非課税通勤費等を差し引いて課税額へ補正）
        #   3) 支給合計 のみ（非課税分を含む可能性。雇用形態欠落警告で人手確認を促す）
        t = None
        if col_total_taxable:
            t = _to_float(ws.cell(r, col_total_taxable).value)
        elif col_total:
            gross = _to_float(ws.cell(r, col_total).value)
            if gross is not None:
                nontax = _to_float(ws.cell(r, col_nontax).value) if col_nontax else None
                t = gross - (nontax or 0)
                # 非課税額 > 支給合計（入力ミス・列誤マッピング・部分月行など）で
                # 負値になった場合は、負の給与を R216 に混入させず gross にフォールバック。
                if t < 0:
                    logger.warning(
                        f'非課税額が支給合計を超過（{name}）: 支給合計{gross:,.0f} '
                        f'< 非課税{(nontax or 0):,.0f} → 支給合計を採用し要確認'
                    )
                    t = gross
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
    col_transport = fmap.get('transport_annual')
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

        # \u5e74\u9593\u901a\u52e4\u624b\u5f53\uff08\u975e\u8ab2\u7a0e\u5206\uff09\u3092\u5728\u7c4d\u6708\u6570\u3067\u5747\u7b49\u5272\u3057\u3066 monthly_wages \u304b\u3089\u6e1b\u7b97\u3002
        # R216 \u7b97\u5165\u984d\u3092\u300c\u7d66\u4e0e\u6240\u5f97\u3068\u3057\u3066\u8ab2\u7a0e\u5bfe\u8c61\u300d\u3060\u3051\u306b\u63c3\u3048\u308b\u305f\u3081\u306e\u88dc\u6b63\u3002
        # \u901a\u52e4\u624b\u5f53\u306f\u56fd\u7a0e\u5e81 No.2585 \u3067\u670815\u4e07\u5186\u307e\u3067\u975e\u8ab2\u7a0e\u306e\u305f\u3081\u7d66\u4e0e\u6240\u5f97\u306b\u542b\u307e\u308c\u306a\u3044\u3002
        # \uff08\u516c\u52df\u8981\u9818 \u7b2c6\u56de p.10 / docs/\u88dc\u52a9\u91d1_\u5b9f\u52d9\u77e5\u8b58\u30d9\u30fc\u30b9.md \u53c2\u7167\uff09
        if col_transport:
            ta = _to_float(ws.cell(r, col_transport).value)
            if ta is not None and ta > 0:
                non_null = [
                    m for m in range(12)
                    if rec['monthly_wages'][m] is not None
                ]
                if non_null:
                    per_month = ta / len(non_null)
                    for m in non_null:
                        rec['monthly_wages'][m] = max(
                            0.0, rec['monthly_wages'][m] - per_month
                        )
                    logger.info(
                        f'\u901a\u52e4\u624b\u5f53\u6e1b\u7b97: {name} \u5e74\u9593{ta:,.0f}\u5186 \u00f7 '
                        f'\u5728\u7c4d{len(non_null)}\u30f6\u6708 = \u6708{per_month:,.0f}\u5186\u6e1b'
                    )


def _read_csv(path: Path, emp_data: dict | None = None) -> None:
    """CSV ファイルの簡易フォールバック読込み。

    通常は AI 経路（read_wage_ledgers_with_ai → _csv_to_tsv）で処理されるため
    この関数が呼ばれるのは USE_AI_WAGE_EXTRACTION=false または AI 抽出失敗時のみ。
    対応範囲：氏名列＋年間給与列がある単純な集計表型 CSV のみ。
    医療機関の月別行型（YYYYMM 1行=1月）には対応しない（AI 経路で処理）。
    """
    if emp_data is None:
        emp_data = {}

    # gzip CSV (magic 0x1F 0x8B) なら自動解凍
    raw_bytes: bytes | None = None
    try:
        with open(path, 'rb') as f:
            head = f.read(2)
        if head == b'\x1f\x8b':
            import gzip
            with gzip.open(path, 'rb') as gz:
                raw_bytes = gz.read()
    except Exception:
        head = b''

    last_error: Exception | None = None
    df = None
    for encoding in ('utf-8-sig', 'utf-8', 'cp932', 'shift_jis'):
        try:
            if raw_bytes is not None:
                import io as _io
                df = pd.read_csv(_io.BytesIO(raw_bytes), header=0,
                                 na_values=['', 'N/A', 'null'], encoding=encoding)
            else:
                df = pd.read_csv(path, header=0, na_values=['', 'N/A', 'null'],
                                 encoding=encoding)
            break
        except (UnicodeDecodeError, UnicodeError) as e:
            last_error = e
            continue
        except Exception as e:
            # encoding 以外の理由（パースエラー等）は早期返却
            level = _csv_decode_warning_level(path)
            msg = f'CSV 読込失敗: {path.name} ({e})'
            if level == 'info':
                logger.info(msg + ' → 同名 xlsx/xlsm を優先使用します')
            else:
                logger.warning(msg)
            return

    if df is None:
        # 全エンコーディング失敗 → 同名 xlsx/xlsm があれば info 降格
        level = _csv_decode_warning_level(path)
        msg = (
            f'CSV読込失敗（対応外形式の可能性、先頭バイト=0x{head.hex() if head else ""}）: '
            f'{path.name} ({last_error})'
        )
        if level == 'info':
            logger.info(msg + ' → 同名 xlsx/xlsm を優先使用します')
        else:
            logger.warning(msg)
        return

    # 列の正規化（日本語テキストの NFD→NFC 変換）
    df.columns = [unicodedata.normalize('NFC', str(col)) for col in df.columns]

    # 氏名列を探す（複数パターン対応）
    name_cols = [c for c in df.columns if '氏名' in c or '名前' in c or 'name' in c.lower()]
    if not name_cols:
        logger.warning(f'CSV に氏名列が見つかりません: {path.name}')
        return
    name_col = name_cols[0]

    # 給与列を探す（複数パターン対応）
    wage_col_patterns = ['支給', '給与', '合計', '給', 'wage', 'total']
    wage_cols = [c for c in df.columns if any(p in c for p in wage_col_patterns)]
    if not wage_cols:
        logger.warning(f'CSV に給与列が見つかりません: {path.name}')
        return

    # 各行を処理（従業員追加）
    for idx, row in df.iterrows():
        name_val = row[name_col]
        if pd.isna(name_val) or not str(name_val).strip():
            continue

        name = str(name_val).replace('　', ' ').strip()
        if not name or name in emp_data:
            continue

        # 初期化（_new_emp_record と同じ構造）
        emp_type = ''
        if '雇用形態' in df.columns:
            et = row['雇用形態']
            if not pd.isna(et):
                emp_type = str(et)
        emp_data[name] = {
            'name': name,
            'employment_type': emp_type,
            'monthly_wages': [None] * 12,
            'monthly_hours': [None] * 12,
            'hourly_rate_flat': 0.0,
            'avg_hours_flat': 0.0,
            'monthly_hourly_rates': [None] * 12,
        }

        # 給与合計を取得（最初の給与列を使用）
        if wage_cols:
            total_wage = row[wage_cols[0]]
            if not pd.isna(total_wage):
                try:
                    annual_total = float(total_wage)
                    # 年間総額を12で除算して月別平均を算出
                    monthly_avg = annual_total / 12
                    emp_data[name]['monthly_wages'] = [monthly_avg] * 12
                except (ValueError, TypeError):
                    pass

        logger.debug(f'CSV から従業員を追加: {name}')


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
    # 支給合計の行を特定（候補順）。R216 は課税支給合計を最優先（公募要領 p.10）。
    total_row = (
        row_labels.get('課税支給合計')
        or row_labels.get('課税分合計')
        or row_labels.get('課税給与合計')
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


def _pdf_to_tsv(path: Path) -> str:
    """PDFからテキストを抽出してTSV風文字列に変換（AI入力用）。
    テキスト層が薄い場合は RuntimeError を送出 → 呼出し側でバイナリ送信にフォールバックする。
    """
    import fitz  # PyMuPDF
    MIN_CHARS_PER_PAGE = 50   # これ未満なら画像PDF扱いとしてフォールバック
    MAX_TOTAL_CHARS = 500_000  # 超大量テキストへのガード（約125Kトークン相当）

    doc = fitz.open(str(path))
    page_count = len(doc)  # close前に保持
    parts: list[str] = [f'### ファイル: {path.name} ###']
    total_chars = 0
    try:
        for page_num in range(page_count):
            text = doc[page_num].get_text('text')
            if text.strip():
                parts.append(f'\n--- ページ {page_num + 1} ---')
                parts.append(text)
                total_chars += len(text)
                if total_chars > MAX_TOTAL_CHARS:
                    logger.warning(
                        f'PDF テキスト量上限超過({total_chars:,}文字)でページ{page_num+1}以降を切り捨て: {path.name}'
                    )
                    break
    finally:
        doc.close()

    avg_chars = total_chars / max(page_count, 1)
    if avg_chars < MIN_CHARS_PER_PAGE:
        raise RuntimeError(
            f'PDF テキスト層が薄すぎます ({avg_chars:.0f}文字/ページ): {path.name}'
        )
    return '\n'.join(parts)


def _csv_decode_warning_level(path: Path) -> str:
    """CSV decode 失敗時に warning と info のどちらで出力するかを判定する。

    同じディレクトリに同名の .xlsx / .xlsm が存在し正常そうなら、
    そちら経由で情報を拾えるので info 降格。それ以外は warning 維持。
    """
    stem = path.stem
    for ext in ('.xlsx', '.xlsm'):
        sibling = path.parent / f'{stem}{ext}'
        if sibling.exists() and sibling.stat().st_size > 0:
            return 'info'
    return 'warning'


def _csv_to_tsv(path: Path) -> str:
    """CSVファイルをTSV文字列に変換（AI入力用）。
    ヘッダー位置・列構成は AI に解釈させるため、全行を文字列のまま渡す。
    複数エンコーディング（UTF-8 / CP932）を順次試す。
    gzip圧縮 CSV (magic 0x1F 0x8B) の場合は自動解凍してから decode。
    """
    # gzip magic を見て解凍を試みる
    try:
        with open(path, 'rb') as f:
            head = f.read(2)
    except Exception:
        head = b''

    raw_bytes: bytes | None = None
    if head == b'\x1f\x8b':
        # gzip CSV — 解凍してから渡す
        try:
            import gzip
            with gzip.open(path, 'rb') as gz:
                raw_bytes = gz.read()
            logger.info(f'CSV(gzip): {path.name} を解凍しました ({len(raw_bytes):,}バイト)')
        except Exception as e:
            logger.warning(f'CSV(gzip) 解凍失敗: {path.name} ({e})')

    last_error: Exception | None = None
    for encoding in ('utf-8-sig', 'utf-8', 'cp932', 'shift_jis'):
        try:
            if raw_bytes is not None:
                import io as _io
                df = pd.read_csv(
                    _io.BytesIO(raw_bytes),
                    header=None,
                    dtype=str,
                    keep_default_na=False,
                    encoding=encoding,
                )
            else:
                df = pd.read_csv(
                    path,
                    header=None,
                    dtype=str,
                    keep_default_na=False,
                    encoding=encoding,
                )
            break
        except (UnicodeDecodeError, UnicodeError) as e:
            last_error = e
            continue
    else:
        # decode 全敗 — 同名 xlsx があれば info 降格、なければ warning
        level = _csv_decode_warning_level(path)
        msg = (
            f'CSV読込失敗（対応外形式の可能性、先頭バイト=0x{head.hex() if head else ""}）: '
            f'{path.name}'
        )
        if level == 'info':
            msg += ' → 同名 xlsx/xlsm を優先使用します'
            logger.info(msg)
        else:
            logger.warning(msg)
        raise last_error or RuntimeError(f'CSV decode failed: {path.name}')

    parts: list[str] = [f'### ファイル: {path.name} ###']
    for row in df.itertuples(index=False, name=None):
        cells = [str(c) if c is not None else '' for c in row]
        while cells and cells[-1] == '':
            cells.pop()
        if not cells:
            continue
        parts.append('\t'.join(cells))
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
    # monthly_work_days は任意フィールド（無くても可）。あれば妥当性チェック
    monthly_work_days = emp.get('monthly_work_days')
    if monthly_work_days is not None:
        if not isinstance(monthly_work_days, list) or len(monthly_work_days) != 12:
            return False, 'monthly_work_days が12要素のリストでない'
        for i, d in enumerate(monthly_work_days):
            if d is None:
                continue
            if not isinstance(d, (int, float)) or d < 0 or d > 31:
                return False, f'{i+1}月の労働日数が異常: {d}'
    return True, ''


def _ai_data_to_wage_employees(ai_data: list[dict]) -> list[WageEmployee]:
    """AI抽出データを WageEmployee リストに変換（バリデーション付き）。

    労働時間が「ない / 異常に少ない（残業時間と誤認の疑い）」場合は、
    労働日数×8時間で補完する。役員は労働時間補完の対象外。
    """
    HOURS_PER_DAY = 8.0
    SUSPICIOUS_AVG_HOURS = 50.0  # 役員/パート以外で月平均がこれ未満なら誤認の疑い

    employees: list[WageEmployee] = []
    for i, emp in enumerate(ai_data):
        if not isinstance(emp, dict):
            logger.warning(f'AI抽出: index={i} が辞書でないためスキップ: {type(emp).__name__}')
            continue
        ok, reason = _validate_ai_employee(emp)
        if not ok:
            logger.warning(f'AI抽出: index={i} ({emp.get("name", "?")}) バリデーション失敗: {reason}')
            continue

        emp_name = str(emp['name']).strip()
        emp_type = str(emp.get('employment_type', '') or '').strip()
        is_officer = '役員' in emp_type
        is_part = 'パート' in emp_type or 'アルバイト' in emp_type

        monthly_wages = [
            float(w) if w is not None else None for w in emp['monthly_wages']
        ]
        monthly_hours: list[float | None] = [
            float(h) if h is not None else None for h in emp['monthly_hours']
        ]
        monthly_work_days = emp.get('monthly_work_days') or [None] * 12

        valid_hours = [h for h in monthly_hours if h is not None and h > 0]
        avg_hours = sum(valid_hours) / len(valid_hours) if valid_hours else 0.0

        # 月給制（正社員 = 役員でもパートでもない）で時間データが部分的（12ヶ月のうち
        # PARTIAL_HOURS_THRESHOLD 未満）の場合は、月給制と判断して時間情報を全 null にする。
        # → 賃金台帳に「平日普通4時間、休日普通15時間」等の勤怠内訳が部分記録された月だけ
        #    AI が拾ってしまい、後段で時給を誤計算する問題を回避。
        PARTIAL_HOURS_THRESHOLD = 3  # 3ヶ月未満なら「月給制で部分記録」と判定
        is_monthly_paid_full_time = not is_officer and not is_part
        if is_monthly_paid_full_time and 0 < len(valid_hours) < PARTIAL_HOURS_THRESHOLD:
            logger.info(
                f'AI抽出補正: {emp_name} (正社員) は時間データが {len(valid_hours)}ヶ月のみで '
                f'月給制と判断 → monthly_hours を全 null に補正'
            )
            monthly_hours = [None] * 12
            valid_hours = []
            avg_hours = 0.0

        # 労働時間が無い、または役員/パート以外で異常に少ない場合は労働日数×8hで補完
        needs_fallback = not valid_hours or (
            not is_officer
            and not is_part
            and avg_hours < SUSPICIOUS_AVG_HOURS
        )
        if needs_fallback and not is_officer:
            valid_days = [
                d for d in monthly_work_days
                if d is not None and isinstance(d, (int, float)) and d > 0
            ]
            # 正社員で労働日数も部分的（PARTIAL_HOURS_THRESHOLD 未満）の場合は補完しない
            # → 中途入社月給制で、入社月だけ勤怠記録が残るケース対策
            if is_monthly_paid_full_time and 0 < len(valid_days) < PARTIAL_HOURS_THRESHOLD:
                logger.info(
                    f'AI抽出補正: {emp_name} (正社員) は労働日数も {len(valid_days)}ヶ月のみで '
                    '月給制と判断 → 労働日数×8h補完をスキップ'
                )
            elif valid_days:
                old_avg = avg_hours
                monthly_hours = [
                    float(d) * HOURS_PER_DAY if (d is not None and isinstance(d, (int, float)) and d > 0) else None
                    for d in monthly_work_days
                ]
                valid_hours = [h for h in monthly_hours if h is not None and h > 0]
                avg_hours = sum(valid_hours) / len(valid_hours) if valid_hours else 0.0
                logger.info(
                    f'AI抽出補完: {emp_name} の労働時間を労働日数×{HOURS_PER_DAY}hで再計算 '
                    f'(月平均 {old_avg:.1f}h → {avg_hours:.1f}h)'
                )

        # 時給算出: 月別 monthly_wages / monthly_hours から計算
        # 加点判定 (judge_bonus_points) と給与計算シートの時給表示で使用
        # 役員は時給を出さない（月給制で意味がないため）
        monthly_hourly_rates: list[float | None] = []
        for w, h in zip(monthly_wages, monthly_hours):
            if not is_officer and w is not None and h is not None and h > 0:
                monthly_hourly_rates.append(w / h)
            else:
                monthly_hourly_rates.append(None)
        valid_hourly = [h for h in monthly_hourly_rates if h is not None and h > 0]
        avg_hourly = sum(valid_hourly) / len(valid_hourly) if valid_hourly else 0.0

        # 年間通勤手当（AI が抽出できた場合のみ取得、なければ 0）
        atransport = emp.get('annual_transport_allowance', 0)
        try:
            atransport_val = float(atransport) if atransport is not None else 0.0
        except (TypeError, ValueError):
            atransport_val = 0.0
        if atransport_val < 0:
            atransport_val = 0.0  # 負値は無視

        # 非課税通勤手当が monthly_wages（支給合計＝通勤込み）に含まれている場合の課税額補正。
        # AI プロンプトは「monthly_wages に課税支給合計を入れた場合は通勤費除外済み→atransport=0」
        # を指示しているが、課税列が無い台帳で AI が通勤込みを入れ atransport>0 を返すことがある。
        # その場合は在籍月数で均等割して monthly_wages から減算し、課税額（R216 母数）に揃える
        # （決定論パーサー _parse_section_summary の S列減算と同一ロジック）。
        # 減算後は atransport_val=0 とする（消費済み）。これにより賃金台帳作成タスクで S列に
        # 二重計上→再読込時に二重減算、という不整合を防ぐ。不変条件「monthly_wages＝課税額」を担保。
        if atransport_val > 0:
            _non_null = [m for m in range(12) if monthly_wages[m] is not None]
            if _non_null:
                _per_month = atransport_val / len(_non_null)
                monthly_wages = [
                    (max(0.0, w - _per_month) if w is not None else None)
                    for w in monthly_wages
                ]
                logger.info(
                    f'AI抽出: {emp_name} の年間通勤手当{atransport_val:,.0f}円を'
                    f'在籍{len(_non_null)}ヶ月で均等割し monthly_wages から減算（課税額補正）'
                )
            atransport_val = 0.0

        employees.append(WageEmployee(
            no=i + 1,
            name=emp_name,
            employment_type=emp_type,
            monthly_avg_hours=round(avg_hours, 1),
            hourly_rate=round(avg_hourly, 1),
            monthly_wages=monthly_wages,
            monthly_hourly_rates=monthly_hourly_rates,
            monthly_hours=monthly_hours,
            annual_transport_allowance=atransport_val,
        ))
    return employees


def read_wage_ledgers_with_ai(
    file_paths: list[Path],
    extractor,
    fiscal_period_hint: str | None = None,
    *,
    disable_image_fallback: bool = False,
) -> list[WageEmployee]:
    """
    AI による賃金台帳読み取り。
    Excel(.xlsx/.xlsm)・CSV は TSV に変換、PDF はそのまま添付して1回の API 呼出しで抽出する。
    各フォーマット混在も可。バリデーション失敗時は空リストを返す（呼出し側で fallback 判断）。

    Args:
        disable_image_fallback: True の場合、Document AI 失敗時に Sonnet 画像経路へ
            落ちる前に ImageFallbackBlockedError が送出される（伝播してくる）。
            「賃金台帳の作成」タスク向け。
    """
    if not file_paths:
        return []

    tsv_blocks: list[str] = []
    pdf_files: list[tuple[str, bytes]] = []  # テキスト抽出失敗時のバイナリフォールバック用

    for path in file_paths:
        ext = path.suffix.lower()
        if ext == '.pdf':
            try:
                tsv_blocks.append(_pdf_to_tsv(path))
                logger.info(f'賃金台帳PDF→テキスト変換: {path.name}')
            except RuntimeError as e:
                # テキスト層が薄い（画像PDF）→ バイナリ送信にフォールバック
                logger.warning(f'{e} → バイナリ送信にフォールバック')
                try:
                    pdf_files.append((path.name, path.read_bytes()))
                except Exception as e2:
                    logger.warning(f'賃金台帳PDF読込失敗(AI経路): {path.name} ({e2})')
            except Exception as e:
                logger.warning(f'賃金台帳PDF読込失敗(AI経路): {path.name} ({e})')
            continue
        if ext == '.csv':
            try:
                tsv_blocks.append(_csv_to_tsv(path))
            except Exception as e:
                # _csv_to_tsv 内で詳細ログ済み。ここでは降格判定のみ
                # （同名 xlsx/xlsm が存在すればそちら経由で読めるので info で十分）
                level = _csv_decode_warning_level(path)
                msg = f'賃金台帳CSV読込失敗(AI経路): {path.name} ({e})'
                if level == 'info':
                    logger.info(msg)
                else:
                    logger.warning(msg)
            continue
        try:
            wb = openpyxl.load_workbook(str(path), data_only=True)
        except Exception as e:
            logger.warning(f'賃金台帳読込失敗(AI経路): {path.name} ({e})')
            continue
        tsv_blocks.append(_workbook_to_tsv(wb, path.name))
        wb.close()

    if not tsv_blocks and not pdf_files:
        return []

    combined_tsv = '\n\n'.join(tsv_blocks) if tsv_blocks else ''
    pdf_count_binary = len(pdf_files)
    pdf_total_mb = sum(len(p[1]) for p in pdf_files) / 1_000_000
    logger.info(
        f'AI抽出開始: TSV{len(tsv_blocks)}ブロック({len(combined_tsv):,}文字)'
        + (f' + PDFバイナリ{pdf_count_binary}件({pdf_total_mb:.2f}MB)' if pdf_files else '')
        + (f' (前事業年度ヒント: {fiscal_period_hint})' if fiscal_period_hint else '')
    )

    try:
        ai_data = extractor.extract_wage_ledger(
            combined_tsv,
            fiscal_period_hint,
            pdf_files=pdf_files if pdf_files else None,
            disable_image_fallback=disable_image_fallback,
        )
    except Exception as e:
        # API残高切れ・画像フォールバック禁止例外は pipeline で全体停止する必要があるので再 raise
        from .ai_extractor import APICreditExhaustedError, ImageFallbackBlockedError
        if isinstance(e, (APICreditExhaustedError, ImageFallbackBlockedError)):
            raise
        logger.error(f'AI抽出例外: {e}', exc_info=True)
        return []

    employees = _ai_data_to_wage_employees(ai_data)
    logger.info(
        f'AI抽出結果: 入力{len(ai_data)}名 → 妥当{len(employees)}名'
    )
    return employees


# OCR が混同しやすい異体字・類似字形の正規化辞書。
# 同一人物判定の取りこぼしを防ぐため、代表字に統一する（読みやすさより一貫性優先）。
# ここに足すかどうかの基準: 字形が似ていて OCR が両方候補に挙げる + 同一人物のはずが
# 別人扱いされる典型例を実観測している こと。
_OCR_VARIANT_MAP = {
    # 旧字 ⇔ 新字（実観測: 吉田 壽 ⇔ 吉田 靖）
    '壽': '寿',
    # 旧字柳 / 櫛 の混乱（実観測: 栁井 ⇔ 櫛井）
    '栁': '柳',
    '櫛': '柳',
    # 嶋 / 崎 / 島 の OCR 混乱（実観測: 大嶋 ⇔ 大崎、宮嶋 ⇔ 宮崎）
    '嶋': '島',
    '崎': '島',
    # 高 / 髙 の混乱（実観測: 髙橋 ⇔ 高橋）
    '髙': '高',
    # 斉 / 齊 / 斎 / 齋 の混乱（実観測: 斉藤 ⇔ 斎藤）
    '齊': '斉',
    '齋': '斉',
    '斎': '斉',
    # 渡辺 / 渡邉 / 渡邊 の混乱
    '邉': '辺',
    '邊': '辺',
    # 沢 / 澤 の混乱
    '澤': '沢',
    # 浜 / 濱 の混乱
    '濱': '浜',
}


def _apply_variant_normalization(s: str) -> str:
    """OCR 異体字を代表字に置換"""
    return ''.join(_OCR_VARIANT_MAP.get(c, c) for c in s)


def _normalize_name_key(name: str) -> str:
    """同一人物判定用の正規化キー。

    - NFKC正規化（全角英数→半角、半角カナ→全角等）
    - 全角・半角の空白をすべて除去
    - OCR 異体字を代表字に置換（壽→寿、栁→柳、嶋→島 等）
    - 大文字小文字統一はしない（漢字氏名の運用なので影響小）

    例: '長塚 典子' / '長塚　典子' / '長塚典子' は同じキー '長塚典子' になる。
    例: '吉田 壽' と '吉田 靖' は字形違いで別人扱いされやすいが、
        '壽' は '寿' に正規化される。'靖' は別字なので統合されないが、
        '吉田寿' に近い候補として後段（_merge_similar_names）で統合判定される。
    """
    if not name:
        return ''
    n = unicodedata.normalize('NFKC', str(name))
    n = re.sub(r'[\s　]+', '', n)
    n = _apply_variant_normalization(n)
    return n


def _merge_two_employees(a: WageEmployee, b: WageEmployee) -> WageEmployee:
    """同一人物と判定された2件のWageEmployeeを統合する。

    同月衝突ポリシー:
        - 片方null → もう片方を採用（補完）
        - 両方値あり・差 < 1円(時間 < 0.1h) → そのまま採用
        - 両方値あり・差大 → WARNING ログ + **大きい方を採用**
          （理由: 部分入力・欠損で値が小さくなる典型的ミスを救う。
           給与は「完全データ ≥ 部分データ ≥ 0」なので max が最も完全な値を選ぶ）
    """
    new_wages: list[float | None] = []
    new_hours: list[float | None] = []
    for i in range(12):
        wa = a.monthly_wages[i] if i < len(a.monthly_wages) else None
        wb = b.monthly_wages[i] if i < len(b.monthly_wages) else None
        if wa is None:
            new_wages.append(wb)
        elif wb is None:
            new_wages.append(wa)
        elif abs(wa - wb) < 1.0:
            new_wages.append(wa)
        else:
            chosen = max(wa, wb)
            logger.warning(
                f'同名統合: "{a.name}" の{i+1}月給与に差分があります '
                f'({wa:,.0f} vs {wb:,.0f}) → 大きい方 {chosen:,.0f} を採用'
            )
            new_wages.append(chosen)

        ha = a.monthly_hours[i] if i < len(a.monthly_hours) else None
        hb = b.monthly_hours[i] if i < len(b.monthly_hours) else None
        if ha is None:
            new_hours.append(hb)
        elif hb is None:
            new_hours.append(ha)
        elif abs(ha - hb) < 0.1:
            new_hours.append(ha)
        else:
            # 時間も同じ理由で max（残業時間との誤判定で部分値が出るケースを救う）
            new_hours.append(max(ha, hb))

    # 月平均/時給を統合後の値で再計算
    valid_h = [h for h in new_hours if h is not None and h > 0]
    avg_h = sum(valid_h) / len(valid_h) if valid_h else 0.0

    is_officer = '役員' in (a.employment_type or '')
    new_rates: list[float | None] = []
    for w, h in zip(new_wages, new_hours):
        if not is_officer and w is not None and h is not None and h > 0:
            new_rates.append(w / h)
        else:
            new_rates.append(None)
    valid_r = [r for r in new_rates if r is not None and r > 0]
    avg_r = sum(valid_r) / len(valid_r) if valid_r else 0.0

    # source_file の統合: 両方あれば連結（重複排除）、片方なら採用
    sa = (a.source_file or '').strip()
    sb = (b.source_file or '').strip()
    if sa and sb:
        if sa == sb:
            merged_source = sa
        else:
            merged_source = f'{sa} + {sb}'
    else:
        merged_source = sa or sb

    return WageEmployee(
        no=a.no,
        name=a.name,
        employment_type=a.employment_type or b.employment_type,
        monthly_avg_hours=round(avg_h, 1),
        hourly_rate=round(avg_r, 1),
        monthly_wages=new_wages,
        monthly_hourly_rates=new_rates,
        monthly_hours=new_hours,
        source_file=merged_source,
    )


def _dedupe_employees_by_normalized_name(
    employees: list[WageEmployee],
) -> list[WageEmployee]:
    """正規化キー（NFKC + 全空白除去）で同一人物の重複を統合する。

    例: '長塚典子' (xlsx由来) と '長塚 典子' (ファイル名由来) を同一として統合。
    """
    if not employees:
        return employees

    merged: dict[str, WageEmployee] = {}
    order: list[str] = []
    for emp in employees:
        key = _normalize_name_key(emp.name)
        if not key:
            continue
        if key not in merged:
            merged[key] = emp
            order.append(key)
        else:
            existing = merged[key]
            if existing.name != emp.name:
                logger.info(
                    f'同名統合: "{existing.name}" と "{emp.name}" を同一人物として統合'
                )
            merged[key] = _merge_two_employees(existing, emp)

    result = [merged[k] for k in order]
    for i, emp in enumerate(result):
        emp.no = i + 1
    return result


def _reconcile_midyear_positions_with_deterministic(
    ai_employees: list[WageEmployee],
    file_paths: list[Path],
    fiscal_period_hint: str | None,
) -> list[WageEmployee]:
    """AI 抽出された中途者の monthly_wages 位置を決定論パーサーで補正。

    背景:
        AI 抽出（Claude）は短期在籍者（在籍 1〜数ヶ月）の monthly_wages を
        誤った位置（暦月とズレた index）に格納する事例が観測されている。
        値そのものは正しいが index がズレるので、出力備考の月ラベルが
        実体と食い違う（例: 賃金台帳の暦4月の値が AI 出力では index 5＝6月扱い）。
        さらに同一ファイルでも実行ごとに変動するため不安定。

    対処:
        - 「データ位置だけ違って値の集合は一致」しているケースに限定して補正
          → 月給合算など値が変わる処理を AI が施した行は触らない（安全側）
        - AI 側 or 決定論側どちらかが「12ヶ月未満」であれば候補（中途者扱い）
          → AI が誤って full_year=True を返した場合の取りこぼし対策
        - 名前は既存の _normalize_name_key で照合（OCR異体字・空白対応）
        - 同名 collision は曖昧として両方とも突合対象から除外（誤上書き防止）
        - 値の上書きは「中身が有効な月が1つ以上ある」場合のみ
          → 空リストで AI の有効データを潰さない

    決定論パーサーは API を呼ばないので追加コストは CPU のみ。
    PDF のみの賃金台帳（Excel/CSV なし）では決定論パーサーが失敗するため
    その場合は AI 結果をそのまま返す。

    Returns:
        補正後の ai_employees（破壊的に書き換え、戻り値は同一リスト）。
    """
    # 突合候補: AI 側で is_full_year=False の人だけでなく、AI が誤って
    # 全埋めしてきた可能性も考えるため、判定は後段で決定論側との比較で行う。
    # ここでは空集合の枝刈りだけ
    if not ai_employees:
        return ai_employees

    # 決定論パーサーで全件読む（CPU のみ、API ゼロ）
    try:
        det_employees = read_wage_ledgers(
            file_paths, extractor=None, fiscal_period_hint=fiscal_period_hint,
        )
    except Exception as e:
        logger.warning(
            f'位置突合用の決定論パーサー読込に失敗: {e} → AI 結果のままで継続'
        )
        return ai_employees

    if not det_employees:
        # PDF のみ等で決定論パーサーが読めないケース
        logger.info('決定論パーサーで0件 → AI 結果をそのまま採用（突合スキップ）')
        return ai_employees

    # 既存の名前正規化を使う（OCR異体字・空白除去を共有）
    # 同名 collision を検出して、その名前は両側とも除外（安全側）
    det_by_name: dict[str, WageEmployee] = {}
    ambiguous_names: set[str] = set()
    for de in det_employees:
        key = _normalize_name_key(de.name)
        if not key:
            continue
        if key in det_by_name:
            ambiguous_names.add(key)
            continue
        det_by_name[key] = de
    for k in ambiguous_names:
        det_by_name.pop(k, None)  # 曖昧キーは突合対象外
        logger.warning(
            f'同名 collision を検出: 正規化キー "{k}" が決定論側に複数あり、'
            '位置突合をスキップ（誤上書き回避）'
        )

    def _data_indices(monthly: list) -> set[int]:
        return {i for i, w in enumerate(monthly or []) if w is not None}

    def _data_multiset(monthly: list) -> tuple:
        """値の multiset。位置非依存で比較するための tuple（順序付け済み）。"""
        vals = [w for w in (monthly or []) if w is not None]
        return tuple(sorted(vals))

    fixed_count = 0
    for ai_emp in ai_employees:
        det = det_by_name.get(_normalize_name_key(ai_emp.name))
        if det is None:
            continue
        ai_months = _data_indices(ai_emp.monthly_wages)
        det_months = _data_indices(det.monthly_wages)
        # 突合候補は「どちらかが12ヶ月未満」かつ「インデックス集合が違う」
        if len(ai_months) == 12 and len(det_months) == 12:
            continue
        if ai_months == det_months:
            continue
        # 値の集合（multiset）が一致するか確認
        # → 一致なら「位置だけ違う、値は同じ」ことが確定。安全に上書き
        # → 不一致なら値そのものが違う可能性が高いので、警告だけ出して上書きしない
        ai_vals = _data_multiset(ai_emp.monthly_wages)
        det_vals = _data_multiset(det.monthly_wages)
        if ai_vals != det_vals:
            logger.warning(
                f'位置不一致だが値集合も不一致: {ai_emp.name} '
                f'AI 値集合={ai_vals} 決定論 値集合={det_vals} '
                f'→ 上書きせず AI 結果のまま（要目視確認）'
            )
            continue
        # 位置だけ違い、値は完全一致 → 決定論側で位置を上書き
        logger.warning(
            f'位置のみ不一致を検出: {ai_emp.name} '
            f'AI={sorted(ai_months)} 決定論={sorted(det_months)} '
            f'値集合は一致 → 決定論側で位置上書き'
        )
        ai_emp.monthly_wages = list(det.monthly_wages)
        # monthly_hours / monthly_hourly_rates は「有効な要素が1つ以上ある」場合のみ上書き
        # （[None]*12 のような空リストで AI の有効データを潰さない）
        if det.monthly_hours and any(h is not None for h in det.monthly_hours):
            ai_emp.monthly_hours = list(det.monthly_hours)
        if det.monthly_hourly_rates and any(
            h is not None for h in det.monthly_hourly_rates
        ):
            ai_emp.monthly_hourly_rates = list(det.monthly_hourly_rates)
        fixed_count += 1

    if fixed_count:
        logger.warning(
            f'AI抽出 vs 決定論パーサーの突合: {fixed_count}名の monthly_wages '
            f'位置を決定論側で上書き（中途者の備考月ラベルを正しく出すため）'
        )

    return ai_employees


def read_wage_ledgers(
    file_paths: list[Path],
    extractor=None,
    fiscal_period_hint: str | None = None,
    *,
    disable_image_fallback: bool = False,
) -> list[WageEmployee]:
    """
    複数の賃金台帳ファイルを読み、同名の従業員をマージして返す。
    1人1ファイル運用（給与ソフト出力）と、1ファイルに全員を入れる運用の両方に対応。

    extractor が渡され、かつ環境変数 USE_AI_WAGE_EXTRACTION が有効な場合は
    AI 抽出を優先し、結果が空なら決定論パーサーにフォールバックする。

    最終結果は NFKC 正規化キー（全空白除去）で同一人物の重複を統合する。
    'A' / 'A ' / 'A　' / 'Ａ' のような表記揺れを 1 人として扱う。

    Args:
        disable_image_fallback: True の場合、Document AI 失敗時に Sonnet 画像経路への
            フォールバックを禁止する（ImageFallbackBlockedError を送出）。
            「賃金台帳の作成」タスクで Document AI 一本に絞る用途。
    """
    if not file_paths:
        return []

    # AI 経路（extractor がある場合）
    if extractor is not None:
        from .config import USE_AI_WAGE_EXTRACTION
        if USE_AI_WAGE_EXTRACTION:
            ai_employees = read_wage_ledgers_with_ai(
                file_paths, extractor, fiscal_period_hint,
                disable_image_fallback=disable_image_fallback,
            )
            if ai_employees:
                ai_employees = _dedupe_employees_by_normalized_name(ai_employees)
                # employment_type 補完: provenance を残す形式「正社員(推定)」で埋める。
                # check_employment_type_missing は「(推定)」を含む値もカウントするため、
                # 補完しても警告は出る（人間チェック可能）。
                for emp in ai_employees:
                    if not (emp.employment_type or '').strip():
                        emp.employment_type = '正社員(推定)'
                # データソース（抽出根拠ファイル名）の補完。
                # AI 経路は複数 PDF をまとめて投入するため、ファイル単位の分離は困難。
                # ファイル名一覧として記録（全員同じ値になる）
                _assign_source_files(ai_employees, file_paths)
                # AI 抽出の中途者は monthly_wages の位置が暦月とズレる事例が観測
                # されている（実行ごとに +1〜+3ヶ月変動）。決定論パーサーは
                # Excel/CSV を直接読むので位置は確実。中途者だけ位置を突合補正する。
                ai_employees = _reconcile_midyear_positions_with_deterministic(
                    ai_employees, file_paths, fiscal_period_hint,
                )
                logger.info(f'賃金台帳合算結果(AI): {len(ai_employees)}名 ({len(file_paths)}ファイル)')
                return ai_employees
            logger.warning('AI抽出が0件を返したため、決定論パーサーにフォールバック')

    # 決定論パーサー経路（フォールバック or extractor なし）
    individual_ledger_paths = []
    merged_emp_data: dict = {}

    for path in file_paths:
        try:
            # CSV の場合は pandas で読込、以降の処理は Excel と同じ
            if path.suffix.lower() == '.csv':
                before = len(merged_emp_data)
                _read_csv(path, merged_emp_data)
                logger.info(
                    f'賃金台帳読み取り(CSV): {path.name} '
                    f'→ 追加/更新 {len(merged_emp_data) - before}名 (累計{len(merged_emp_data)}名)'
                )
                continue

            # Excel の場合は既存処理
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
            # 個人台帳型は 1ファイル=1〜数名なので、source_file を直接記録
            if not e.source_file:
                e.source_file = path.name
            employees.append(e)

    employees = _dedupe_employees_by_normalized_name(employees)

    # employment_type の補完: provenance を残す形式「正社員(推定)」で埋める。
    # 給与計算で「正社員」「パート」のどちらにもカウントされないと人数 0 になるため、
    # 補完は必須だが、推定値であることを明示することで wage_validator や UI が
    # 「人間チェック対象」として扱える。
    for emp in employees:
        if not (emp.employment_type or '').strip():
            emp.employment_type = '正社員(推定)'

    # データソース（抽出根拠ファイル名）の補完
    # 統合読み込み（_read_flexible / _read_csv）由来の従業員は path 単位で分離できないので
    # 「統合」表記。個人台帳型は _read_individual_ledger 内で既に設定済み。
    _assign_source_files(employees, file_paths)

    logger.info(f'賃金台帳合算結果(決定論): {len(employees)}名 ({len(file_paths)}ファイル)')
    return employees


def _assign_source_files(
    employees: list[WageEmployee],
    file_paths: list[Path],
) -> None:
    """source_file 未設定の従業員に既定値を割り当てる。

    - 1ファイルのみ: そのファイル名
    - 複数ファイル: '統合(Nファイル)' とファイル名先頭3つを併記
    """
    if not employees or not file_paths:
        return
    if len(file_paths) == 1:
        default = file_paths[0].name
    else:
        # 多すぎる場合は先頭3つだけ列挙（セル幅対策）
        names = [p.name for p in file_paths[:3]]
        suffix = f' 他{len(file_paths) - 3}件' if len(file_paths) > 3 else ''
        default = f'統合({len(file_paths)}ファイル): ' + ', '.join(names) + suffix
    for emp in employees:
        if not emp.source_file:
            emp.source_file = default


# ============================================================
# 賃金台帳一覧Excel出力（チェック用）
# ============================================================

def _is_excluded_from_wage_total(emp: WageEmployee) -> bool:
    """給与支給総額（R216）の集計から除外される従業員か判定。

    除外条件（公募要領準拠）:
      - 役員（employment_type に「役員」を含む）
      - 基準年度に全月分の給与支給を受けていない（中途入退社等）
    """
    if '役員' in (emp.employment_type or ''):
        return True
    if not emp.is_full_year:
        return True
    return False


def export_wage_ledger_summary(
    employees: list[WageEmployee],
    output_path: Path,
    company_name: str = '',
    extraction_method: str = '',
) -> Path:
    """
    賃金台帳から読み取ったデータを一覧Excelに出力（チェック用）

    出力内容:
      シート1「賃金台帳一覧」:
        左ブロック  : 月別課税対象額（12か月）+ 年間合計賃金
        右ブロック  : 月別労働時間（12か月）+ 年間合計時間 + 月平均労働時間
        右端       : データソース（抽出根拠の元ファイル名）
      シート2「算定根拠」:
        - 採用列・含む経費・除外経費・役員除外ルール
        - 抽出経路（AI / 決定論）
        - IT導入補助金 2026 通常枠 公募要領 p.10 原文引用
        - データソースとなった賃金台帳ファイル一覧（重複排除）

    集計対象外（役員 or 非全月在籍）の行は薄いグレーで塗り、目視で除外行を判別可能にする。

    引数:
      extraction_method: 抽出経路の説明（'AI抽出（Claude Sonnet 4.6）' / '決定論パーサー' 等）。
                        空文字なら算定根拠シートでは「未指定」表示。
    """
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
    from openpyxl.utils import get_column_letter

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '賃金台帳一覧'

    # スタイル定義
    header_fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
    group_fill = PatternFill(start_color='8FAADC', end_color='8FAADC', fill_type='solid')
    excluded_fill = PatternFill(  # 集計対象外行（役員 / 非全月在籍）のグレー塗り
        start_color='E7E6E6', end_color='E7E6E6', fill_type='solid',
    )
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

    # 抽出経路の明示（AI抽出かどうかを最上部に大きく表示）
    is_ai_extraction = 'AI' in (extraction_method or '')
    extraction_label = extraction_method or '未指定'
    if is_ai_extraction:
        warn_msg = (
            f'⚠ 抽出経路: {extraction_label} '
            '— AI（Claude）による読み取りのため誤読の可能性があります。'
            '賃金台帳原本と必ず照合してください。'
        )
    else:
        warn_msg = f'抽出経路: {extraction_label}'
    cell_warn = ws.cell(row=2, column=1, value=warn_msg)
    cell_warn.font = Font(bold=True, size=10, color='C00000' if is_ai_extraction else '333333')
    if is_ai_extraction:
        # 視認性のため薄オレンジで強調（給与支給総額計算.xlsx の AI 色と統一）
        cell_warn.fill = PatternFill(
            start_color='FCE4D6', end_color='FCE4D6', fill_type='solid',
        )

    ws.cell(
        row=3, column=1,
        value=(
            '※採用列: 賃金台帳の「課税支給合計」（給与所得として課税対象となる経費。'
            '非課税通勤手当・社保等控除前）。'
            '出典: IT導入補助金 2026 通常枠 公募要領 p.10。'
            '抽出経路と公募要領原文の引用は「算定根拠」シートを参照。'
        ),
    )
    ws.cell(row=3, column=1).font = Font(size=9, color='666666')
    ws.cell(
        row=4, column=1,
        value='※グレー行は給与支給総額（R216）の集計対象外 — 役員 or 非全月在籍（中途入退社）',
    )
    ws.cell(row=4, column=1).font = Font(size=9, color='666666')

    # 列レイアウト
    # 1: No, 2: 従業員名, 3: 雇用形態,
    # 4-15: 1月〜12月 賃金, 16: 年間合計賃金,
    # 17-28: 1月〜12月 時間, 29: 年間合計時間, 30: 月平均労働時間,
    # 31: データソース
    wage_start = 4
    wage_total_col = wage_start + 12  # 16
    hours_start = wage_total_col + 1  # 17
    hours_total_col = hours_start + 12  # 29
    avg_hours_col = hours_total_col + 1  # 30
    source_col = avg_hours_col + 1  # 31

    # グループヘッダー（5行目）— 抽出経路メッセージを 2行目に追加したため1行ずれる
    group_row = 5
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

    # 列ヘッダー（6行目）
    header_row = 6
    headers = (
        ['No', '従業員名', '雇用形態']
        + MONTH_NAMES + ['年間合計']
        + MONTH_NAMES + ['年間合計', '月平均']
        + ['データソース']
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
        is_excluded = _is_excluded_from_wage_total(emp)

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

        # データソース（抽出根拠の元ファイル名）
        source_cell = ws.cell(row=r, column=source_col, value=emp.source_file)
        source_cell.border = thin_border
        source_cell.font = Font(size=9, color='666666')
        source_cell.alignment = Alignment(horizontal='left', vertical='center')

        # 集計対象外行（役員 or 非全月在籍）はグレーに塗る
        if is_excluded:
            for c in range(1, source_col + 1):
                ws.cell(row=r, column=c).fill = excluded_fill

    # ── 合計行（全員 / 集計対象のみの2段）────────────────────────────
    # 後段の給与支給総額計算や申請書 R216 との突合用。SUM 式で記述しておくと
    # 行追加・編集後も自動再計算されるので、人間チェックの再利用性が上がる。
    if employees:
        from openpyxl.styles import Font as _Font  # 上で import 済みだが明示
        subtotal_fill_all = PatternFill(
            start_color='B4C7E7', end_color='B4C7E7', fill_type='solid',
        )
        subtotal_fill_target = PatternFill(
            start_color='C6E0B4', end_color='C6E0B4', fill_type='solid',
        )
        first_data_row = header_row + 1
        last_data_row = header_row + len(employees)
        total_all_row = last_data_row + 1
        total_target_row = last_data_row + 2

        ws.cell(row=total_all_row, column=2, value='合計（全員）').font = Font(bold=True, size=10)
        ws.cell(row=total_target_row, column=2,
                value='合計（集計対象のみ）').font = Font(bold=True, size=10)

        # 集計対象行のインデックス（行番号は first_data_row 基準）
        target_row_nums = [
            first_data_row + i for i, emp in enumerate(employees)
            if not _is_excluded_from_wage_total(emp)
        ]

        def _set_total_cell(row: int, col: int, formula: str, fmt: str, fill):
            c = ws.cell(row=row, column=col, value=formula)
            c.font = Font(bold=True)
            c.number_format = fmt
            c.fill = fill
            c.border = thin_border

        # 賃金: 1月〜12月 + 年間合計
        for col in list(range(wage_start, wage_start + 12)) + [wage_total_col]:
            col_letter = get_column_letter(col)
            # 全員
            _set_total_cell(
                total_all_row, col,
                f'=SUM({col_letter}{first_data_row}:{col_letter}{last_data_row})',
                number_fmt, subtotal_fill_all,
            )
            # 集計対象のみ（個別セルの加算式）
            if target_row_nums:
                parts = '+'.join(f'{col_letter}{r}' for r in target_row_nums)
                target_formula = f'={parts}'
            else:
                target_formula = 0
            _set_total_cell(
                total_target_row, col, target_formula,
                number_fmt, subtotal_fill_target,
            )

        # 労働時間: 1月〜12月 + 年間合計
        for col in list(range(hours_start, hours_start + 12)) + [hours_total_col]:
            col_letter = get_column_letter(col)
            _set_total_cell(
                total_all_row, col,
                f'=SUM({col_letter}{first_data_row}:{col_letter}{last_data_row})',
                hours_fmt, subtotal_fill_all,
            )
            if target_row_nums:
                parts = '+'.join(f'{col_letter}{r}' for r in target_row_nums)
                target_formula = f'={parts}'
            else:
                target_formula = 0
            _set_total_cell(
                total_target_row, col, target_formula,
                hours_fmt, subtotal_fill_target,
            )

        # 備考: 「集計対象のみ」の年間合計は R216 の母数（給与支給総額・役員除外）
        wage_total_letter = get_column_letter(wage_total_col)
        note_cell = ws.cell(
            row=total_target_row, column=source_col,
            value=f'※「合計（集計対象のみ）」の年間合計（{wage_total_letter}列）'
                  f'＝ R216 給与支給総額の母数',
        )
        note_cell.font = Font(size=9, color='666666')
        note_cell.fill = subtotal_fill_target
        note_cell.border = thin_border
        note_cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)

        # 全員行の備考: グレー行（役員・中途）も含む合計である旨を明示
        note_all_cell = ws.cell(
            row=total_all_row, column=source_col,
            value='※役員・中途入退社を含む全行合計（R216 母数には未調整）',
        )
        note_all_cell.font = Font(size=9, color='666666')
        note_all_cell.fill = subtotal_fill_all
        note_all_cell.border = thin_border
        note_all_cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)

    # 列幅調整
    ws.column_dimensions['A'].width = 5
    ws.column_dimensions['B'].width = 14
    ws.column_dimensions['C'].width = 14
    for c in range(wage_start, avg_hours_col + 1):
        ws.column_dimensions[get_column_letter(c)].width = 11
    ws.column_dimensions[get_column_letter(wage_total_col)].width = 13
    ws.column_dimensions[get_column_letter(hours_total_col)].width = 13
    ws.column_dimensions[get_column_letter(avg_hours_col)].width = 11
    ws.column_dimensions[get_column_letter(source_col)].width = 40

    # 算定根拠シート（採用列・除外ルール・公募要領原文・データソース一覧）
    _write_calculation_basis_sheet(
        wb,
        employees=employees,
        extraction_method=extraction_method,
        thin_border=thin_border,
    )

    wb.save(str(output_path))
    wb.close()
    logger.info(f'賃金台帳一覧出力: {output_path} ({len(employees)}名)')
    return output_path


def _write_calculation_basis_sheet(
    wb,
    employees: list[WageEmployee],
    extraction_method: str,
    thin_border,
) -> None:
    """『算定根拠』シートを追加。R216 算定の根拠条文・採用列・データソースを記録。"""
    from openpyxl.styles import Font, Alignment, PatternFill

    ws = wb.create_sheet(title='算定根拠')

    title_font = Font(bold=True, size=12)
    section_font = Font(bold=True, size=10, color='FFFFFF')
    section_fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
    label_font = Font(bold=True, size=10)
    quote_font = Font(size=10, italic=True, color='333333')
    note_font = Font(size=9, color='666666')

    ws.cell(row=1, column=1, value='給与支給総額（R216）の算定根拠').font = title_font
    ws.cell(row=2, column=1, value='IT導入補助金 2026 通常枠 公募要領 準拠').font = note_font

    rows: list[tuple[str, str]] = [
        ('対象補助金', 'IT導入補助金 2026（デジタル化・AI導入補助金2026、通常枠／インボイス対応類型）'),
        ('採用列', '賃金台帳の「課税支給合計」（給与所得として課税対象となる経費）'),
        (
            '含む経費',
            '給料・賃金・賞与・各種手当（残業手当／休日出勤手当／'
            '職務手当／地域手当／家族(扶養)手当／住宅手当）等',
        ),
        (
            '含まない経費',
            '福利厚生費・法定福利費・退職金・'
            '非課税通勤手当（限度額内分、国税庁 No.2585 により給与所得に含まれない）',
        ),
        ('役員の扱い', '集計対象外（IT導入補助金 2026 通常枠 公募要領 p.10「役員報酬…は除く」）。ただし従業員0名の法人のみ役員で読み替え可'),
        (
            '中途入退社者の扱い',
            '集計対象外（基準年度に全月分の給与等の支給を受けていない従業員）',
        ),
        ('賞与の合算', '月別の課税給与に合算（賞与シートが別ファイルの場合は対応月に加算）'),
        ('抽出経路', extraction_method or '未指定'),
        (
            '決定論パーサーの列優先順',
            '① 課税支給合計 → ② 支給合計 → ③ 差引支給合計（手取り）',
        ),
    ]

    start_row = 4
    ws.cell(row=start_row, column=1, value='項目').font = section_font
    ws.cell(row=start_row, column=1).fill = section_fill
    ws.cell(row=start_row, column=2, value='値・説明').font = section_font
    ws.cell(row=start_row, column=2).fill = section_fill
    ws.cell(row=start_row, column=1).border = thin_border
    ws.cell(row=start_row, column=2).border = thin_border

    for i, (label, value) in enumerate(rows, 1):
        r = start_row + i
        c1 = ws.cell(row=r, column=1, value=label)
        c1.font = label_font
        c1.alignment = Alignment(vertical='top', wrap_text=True)
        c1.border = thin_border
        c2 = ws.cell(row=r, column=2, value=value)
        c2.alignment = Alignment(vertical='top', wrap_text=True)
        c2.border = thin_border

    # IT導入補助金 2026 通常枠 公募要領 p.10 原文
    quote_row = start_row + len(rows) + 3
    ws.cell(row=quote_row, column=1, value='IT導入補助金 2026 通常枠 公募要領 p.10 原文').font = label_font
    ws.merge_cells(start_row=quote_row + 1, start_column=1,
                   end_row=quote_row + 1, end_column=2)
    quote_cell = ws.cell(
        row=quote_row + 1, column=1,
        value=(
            '算定対象となる給与等は、給料、賃金、賞与、各種手当'
            '（残業手当、休日出勤手当、職務手当、地域手当、家族（扶養）手当、住宅手当）等、'
            '給与所得として課税対象となる経費を指す。'
            '役員報酬、福利厚生費、法定福利費や退職金は除く。'
        ),
    )
    quote_cell.font = quote_font
    quote_cell.alignment = Alignment(vertical='top', wrap_text=True)
    ws.row_dimensions[quote_row + 1].height = 60

    # 公式 URL
    url_row = quote_row + 3
    ws.cell(row=url_row, column=1, value='IT導入補助金 2026 通常枠 公募要領 PDF').font = label_font
    ws.cell(
        row=url_row, column=2,
        value='https://it-shien.smrj.go.jp/pdf/it2026_koubo_tsujyo.pdf',
    )
    ws.cell(row=url_row + 1, column=1, value='IT導入補助金 2026 交付申請マニュアル PDF').font = label_font
    ws.cell(
        row=url_row + 1, column=2,
        value='https://it-shien.smrj.go.jp/pdf/it2026_manual_kofu.pdf',
    )
    ws.cell(row=url_row + 2, column=1, value='通勤手当 非課税限度額').font = label_font
    ws.cell(
        row=url_row + 2, column=2,
        value='https://www.nta.go.jp/taxes/shiraberu/taxanswer/gensen/2585.htm（国税庁 No.2585）',
    )

    # データソース一覧（emp.source_file から重複排除）
    source_files: list[str] = []
    seen: set[str] = set()
    for emp in employees:
        src = (emp.source_file or '').strip()
        if src and src not in seen:
            seen.add(src)
            source_files.append(src)

    src_row = url_row + 4
    ws.cell(
        row=src_row, column=1,
        value=f'参照した賃金台帳ファイル（{len(source_files)}件）',
    ).font = label_font
    if not source_files:
        ws.cell(row=src_row + 1, column=2, value='（ファイル情報なし）').font = note_font
    else:
        for i, src in enumerate(source_files):
            ws.cell(row=src_row + 1 + i, column=2, value=src).font = note_font

    # 列幅
    ws.column_dimensions['A'].width = 24
    ws.column_dimensions['B'].width = 90


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

    # 公式「賃金状況報告シート（補助率引上げ・加点措置①用）」の判定数式は
    #   対象 = IF(AND(時間換算給与 < R7改定後, 時間換算給与 > 0), "対象", "対象外")
    # で、下限は「> 0」のみ。R6改定前は参考表示（改定前列）で判定には用いない。
    # かつて R6 を下限に加えていたが、それだと R6 未満に計算された従業員を取りこぼし、
    # 公式シートより厳しく「対象外」と誤判定して補助率2/3を逃すため、公式に合わせる。
    # （公募要領 2026 p.18 補助率2/3条件・p.26 加点項目14／賃金状況報告シート① 数式で確認）
    mw_r7 = result.min_wage_r7

    # ── 加点措置①（公式名: 補助率引上げ・加点措置① ／ 公募要領 加点項目14・補助率2/3トリガー）──
    # 令和6年10月〜令和7年9月の暦月のうち、R7改定後最低賃金未満の従業員が
    # 全従業員の30%以上である月が3か月以上 → 対象。
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
            # 公式シート①準拠: R7改定後未満なら対象（hourly>0は上の continue で担保済み）。
            # R6改定前の下限は設けない（公式数式は < R7改定後 のみ）。
            is_under_r7 = hourly < mw_r7

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

    # fill_bonus_sheet_2 が同じ月をシートに反映できるよう保存
    result.latest_month_idx = latest_month_idx

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
    latest_month_idx: int | None = None,
) -> Path:
    """
    加点措置②用シートに従業員データを入力

    加点措置②のシート構成:
      2つの賃金計算期間を横に並べて入力
      データ行は17行目から

    latest_month_idx を省略した場合、judge_bonus_points が
    動的決定して result.latest_month_idx に保存した月を使う。
    画面判定（最新データがある月）とExcel出力を一致させる目的。
    フォールバックとして 11（12月）を使う。
    """
    wb = openpyxl.load_workbook(str(template_path))
    ws = wb[wb.sheetnames[0]]

    if latest_month_idx is None:
        latest_month_idx = (
            result.latest_month_idx
            if result.latest_month_idx is not None else 11
        )

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
