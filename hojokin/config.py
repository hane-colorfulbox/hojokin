# -*- coding: utf-8 -*-
"""設定管理"""
from __future__ import annotations

import os
from pathlib import Path
from dataclasses import dataclass, field

# ── 環境変数から読み込み ──
CLAUDE_API_KEY = os.getenv('CLAUDE_API_KEY', '')
CLAUDE_MODEL = os.getenv('CLAUDE_MODEL', 'claude-sonnet-4-6')
GOOGLE_CREDENTIALS_PATH = os.getenv('GOOGLE_CREDENTIALS_PATH', 'credentials/service_account.json')
MANAGEMENT_SHEET_ID = os.getenv('MANAGEMENT_SHEET_ID', '')

# ── 標準パス ──
BASE_DIR = Path(os.getenv('HOJOKIN_BASE_DIR', '.'))

# ── 定数 ──
STANDARD_ANNUAL_HOURS = 2080  # 標準年間労働時間 (40h/週 × 52週)

# ── 最低賃金マスタ (R6年度) ──
MIN_WAGE_MAP = {
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


def detect_prefecture(address: str) -> str | None:
    """住所から都道府県を判定"""
    for pref in MIN_WAGE_MAP:
        if pref in address:
            return pref
    return None


def get_min_wage(address: str) -> tuple[str, int] | None:
    """住所から最低賃金を取得。(都道府県名, 金額) or None"""
    pref = detect_prefecture(address)
    if pref and pref in MIN_WAGE_MAP:
        return pref, MIN_WAGE_MAP[pref]
    return None


# ── テンプレート行マッピング (2026 通常枠) ──
@dataclass
class TemplateMapping:
    """テンプレートの行番号マッピング"""

    # 転記シート: ヒアリング行 → 転記行
    hearing_to_tenki: list[tuple[int, int, bool]] = field(default_factory=list)

    # 申請内容シート: フィールド名 → 行番号 (C列)
    shinsei: dict[str, int] = field(default_factory=dict)

    # 給与計算シート
    kyuyo_sheet_name: str = ''
    kyuyo: dict[str, tuple[int, int]] = field(default_factory=dict)  # field → (row, col)

    # 申請内容のクリア範囲
    shinsei_clear_range: tuple[int, int] = (5, 250)

    # 転記テキスト行の範囲
    tenki_text_range: tuple[int, int] = (16, 26)


# 2026 通常枠テンプレート
MAPPING_2026_TSUJO = TemplateMapping(
    hearing_to_tenki=[
        # (ヒアリング行, 転記行, 電話番号変換)
        (6,  8,  False),   # 企業名
        (8,  10, False),   # 企業名フリガナ
        (10, 12, False),   # 店舗事業所数
        (12, 14, False),   # 事業者URL
        (14, 16, False),   # 事業内容
        (20, 27, True),    # 代表電話番号
        (22, 29, False),   # 担当者氏名
        (24, 31, False),   # 担当者氏名フリガナ
        (26, 33, False),   # 担当者メールアドレス
        (28, 35, True),    # 担当者電話番号
        (30, 37, True),    # 担当者携帯番号
        (32, 39, False),   # 正規雇用
        (34, 41, False),   # 契約社員
        (36, 43, False),   # パートアルバイト
        (38, 45, False),   # 派遣社員
        (40, 47, False),   # その他
        (42, 49, False),   # 過去に補助金
        (59, 58, False),   # えるぼし
        (62, 61, False),   # くるみん
        (64, 63, False),   # SECURITY ACTION ID
        (69, 66, False),   # 正規雇用(前期)
        (70, 67, False),   # 契約社員(前期)
        (71, 68, False),   # パート(前期)
        (72, 69, False),   # 年間平均労働時間
        (75, 72, False),   # 自社の強み
        (76, 73, False),   # 自社の弱み
        (77, 74, False),   # IT投資年間金額
        (78, 75, False),   # IT投資プロセス
        (81, 78, False),   # 事業所内最低賃金
        (83, 80, False),   # 賃金引上げ表明
        (85, 82, False),   # 賃金引上げ幅
        (87, 84, False),   # 従業員代表者
        (88, 85, False),   # 給与担当者
        (89, 86, False),   # 事業所内最低賃金者
    ],
    shinsei={
        'headquarters_address': 39,
        'industry_code': 40,
        'industry_text': 41,
        'established_date': 42,
        'capital': 43,
        'business_description': 47,
        'fiscal_month': 48,
        'officer_count': 55,
        'rep_title': 56,
        'rep_name': 57,
        'rep_kana': 58,
        'officer_1_title': 60,
        'officer_1_name': 61,
        'officer_1_kana': 62,
        'officer_2_title': 63,
        'officer_2_name': 64,
        'officer_2_kana': 65,
        'officer_3_title': 66,
        'officer_3_name': 67,
        'officer_3_kana': 68,
        'officer_4_title': 69,
        'officer_4_name': 70,
        'officer_4_kana': 71,
        'officer_5_title': 72,
        'officer_5_name': 73,
        'officer_5_kana': 74,
        'past_subsidies': 80,
        'eruboshi': 85,
        'kurumin': 86,
        'tool_name': 84,
        'business_types': 105,
        'officer_count_prev': 114,
        'management_intent': 127,
        'security_status': 128,
        'future_goals': 129,
        'it_investment_status': 130,
        'it_utilization_status': 131,
        'min_wage': 163,
        'wage_raise_declaration': 172,
        'wage_raise_amount': 173,
        'wage_raise_method': 174,
        'wage_raise_date': 175,
    },
    kyuyo_sheet_name='給与支給総額計算',
    kyuyo={
        'revenue':          (10, 2),  # B10: 売上高
        'gross_profit':     (11, 2),  # B11: 粗利益
        'operating_profit': (12, 2),  # B12: 営業利益
        'ordinary_profit':  (13, 2),  # B13: 経常利益
        'depreciation':     (16, 5),  # E16: 減価償却費
        'salary':           (6,  5),  # E6:  従業員給与
        'officer_comp':     (5,  5),  # E5:  役員報酬
    },
    shinsei_clear_range=(5, 250),
    tenki_text_range=(16, 26),
)

# 2026 インボイス枠（通常枠と行番号が異なる箇所を上書き）
# 現時点では通常枠と同じマッピングを仮置き。テンプレート確認後に調整
MAPPING_2026_INVOICE = TemplateMapping(
    hearing_to_tenki=MAPPING_2026_TSUJO.hearing_to_tenki,
    shinsei=MAPPING_2026_TSUJO.shinsei.copy(),
    kyuyo_sheet_name='給与支給総額計算',
    kyuyo=MAPPING_2026_TSUJO.kyuyo.copy(),
    shinsei_clear_range=(5, 200),
    tenki_text_range=(16, 26),
)


def get_mapping(template_type: str) -> TemplateMapping:
    """テンプレートタイプからマッピングを取得"""
    mappings = {
        '通常枠_2026': MAPPING_2026_TSUJO,
        'インボイス枠_2026': MAPPING_2026_INVOICE,
    }
    if template_type not in mappings:
        raise ValueError(f'未対応のテンプレートタイプ: {template_type}。対応: {list(mappings.keys())}')
    return mappings[template_type]
