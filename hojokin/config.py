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

# ── 最低賃金マスタ (R7年度改定後) ──
MIN_WAGE_MAP = {
    '北海道': 1075, '青森県': 1029, '岩手県': 1031, '宮城県': 1038,
    '秋田県': 1031, '山形県': 1032, '福島県': 1033, '東京都': 1226,
    '茨城県': 1074, '栃木県': 1068, '群馬県': 1063, '埼玉県': 1141,
    '千葉県': 1140, '神奈川県': 1225, '新潟県': 1050, '富山県': 1062,
    '石川県': 1054, '福井県': 1053, '山梨県': 1052, '長野県': 1061,
    '岐阜県': 1065, '静岡県': 1097, '愛知県': 1140, '京都府': 1122,
    '大阪府': 1177, '三重県': 1087, '滋賀県': 1080, '兵庫県': 1116,
    '奈良県': 1051, '和歌山県': 1045, '鳥取県': 1030, '島根県': 1033,
    '岡山県': 1047, '広島県': 1085, '山口県': 1043, '徳島県': 1046,
    '香川県': 1036, '愛媛県': 1033, '高知県': 1023, '福岡県': 1057,
    '佐賀県': 1030, '長崎県': 1031, '大分県': 1035, '熊本県': 1034,
    '宮崎県': 1023, '鹿児島県': 1026, '沖縄県': 1023,
}


def detect_prefecture(address: str) -> str | None:
    """住所から都道府県を判定"""
    for pref in MIN_WAGE_MAP:
        if pref in address:
            return pref

    # 都道府県名なしの場合、市区町村名から推定
    CITY_TO_PREF = {
        '北九州市': '福岡県', '福岡市': '福岡県', '久留米市': '福岡県',
        '札幌市': '北海道', '仙台市': '宮城県', '新潟市': '新潟県',
        'さいたま市': '埼玉県', '千葉市': '千葉県', '川崎市': '神奈川県',
        '横浜市': '神奈川県', '相模原市': '神奈川県',
        '名古屋市': '愛知県', '浜松市': '静岡県', '静岡市': '静岡県',
        '京都市': '京都府', '大阪市': '大阪府', '堺市': '大阪府',
        '神戸市': '兵庫県', '姫路市': '兵庫県',
        '岡山市': '岡山県', '広島市': '広島県',
        '高松市': '香川県', '松山市': '愛媛県',
        '熊本市': '熊本県', '鹿児島市': '鹿児島県', '那覇市': '沖縄県',
    }
    for city, pref in CITY_TO_PREF.items():
        if city in address:
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
    tenki_text_range: tuple[int, int] = (15, 26)


# ── 2026 通常枠テンプレート ──
# テンプレート原本: ツール/【原本_法人】企業名_通常枠_法人2026.xlsx
# ヒアリングシート: ツール/ヒアリングシート2026_通常枠法人.xlsx
MAPPING_2026_TSUJO = TemplateMapping(
    hearing_to_tenki=[
        # (ヒアリング行, 転記行, 電話番号変換)
        (6,  8,  False),   # 企業名
        (8,  10, False),   # 企業名フリガナ
        (10, 12, False),   # 店舗事業所数
        (12, 14, False),   # 事業者URL
        (15, 16, False),   # 主な事業内容
        (16, 17, False),   # 強み
        (17, 18, False),   # 業務上の課題
        (18, 19, False),   # 人事・組織の課題
        (19, 20, False),   # ロス時間
        (20, 21, False),   # 活用したい主要機能
        (21, 22, False),   # 期待する効果（数値）
        (22, 23, False),   # 浮いた時間の活用
        (23, 24, False),   # 3年後の売上目標
        (24, 25, False),   # 取引先属性
        (26, 27, True),    # 代表電話番号
        (28, 29, False),   # 担当者氏名
        (30, 31, False),   # 担当者氏名フリガナ
        (32, 33, False),   # 担当者メールアドレス
        (34, 35, True),    # 担当者電話番号
        (36, 37, True),    # 担当者携帯番号
        (38, 39, False),   # 正規雇用
        (40, 41, False),   # 契約社員
        (42, 43, False),   # パートアルバイト
        (44, 45, False),   # 派遣社員
        (46, 47, False),   # その他
        (48, 49, False),   # 過去に補助金
        (50, 51, False),   # 申請年度
        (51, 52, False),   # 申請枠
        (52, 53, False),   # 申請回
        (53, 54, False),   # 申請ツール名
        (57, 58, False),   # えるぼし
        (60, 61, False),   # くるみん
        (62, 63, False),   # SECURITY ACTION ID
        (65, 66, False),   # 正規雇用(前期)
        (66, 67, False),   # 契約社員(前期)
        (67, 68, False),   # パート(前期)
        (68, 69, False),   # 年間平均労働時間
        (71, 72, False),   # 自社の強み（選択肢）
        (72, 73, False),   # 自社の弱み（選択肢）
        (73, 74, False),   # IT投資年間金額
        (74, 75, False),   # IT投資プロセス
        (77, 78, False),   # 事業所内最低賃金
        (79, 80, False),   # 賃金引上げ表明
        (82, 82, False),   # 賃金引上げ幅
        (84, 84, False),   # 従業員代表者
        (85, 85, False),   # 給与担当者
        (86, 86, False),   # 事業所内最低賃金者
    ],
    shinsei={
        # 基本情報入力 (行49〜)
        'headquarters_address': 54,
        'industry_code': 55,
        'industry_text': 56,
        'established_date': 57,
        'capital': 58,
        'business_description': 73,
        'fiscal_month': 74,
        'tool_name': 71,

        # 担当者・代表者 (行76〜)
        'officer_count': 81,
        'rep_title': 82,
        'rep_name': 83,
        'rep_kana': 84,

        # 役員 (最大10名) ※(1)〜(7)はズレなし、(8)〜(10)は新規
        'officer_1_title': 87,
        'officer_1_name': 88,
        'officer_1_kana': 89,
        'officer_2_title': 90,
        'officer_2_name': 91,
        'officer_2_kana': 92,
        'officer_3_title': 93,
        'officer_3_name': 94,
        'officer_3_kana': 95,
        'officer_4_title': 96,
        'officer_4_name': 97,
        'officer_4_kana': 98,
        'officer_5_title': 99,
        'officer_5_name': 100,
        'officer_5_kana': 101,
        'officer_6_title': 102,
        'officer_6_name': 103,
        'officer_6_kana': 104,
        'officer_7_title': 105,
        'officer_7_name': 106,
        'officer_7_kana': 107,
        'officer_8_title': 108,
        'officer_8_name': 109,
        'officer_8_kana': 110,
        'officer_9_title': 111,
        'officer_9_name': 112,
        'officer_9_kana': 113,
        'officer_10_title': 114,
        'officer_10_name': 115,
        'officer_10_kana': 116,

        # その他 ※+9
        'past_subsidies': 122,
        'eruboshi': 123,
        'kurumin': 124,
        'business_types': 144,

        # 財務情報入力 ※+9
        'officer_count_prev': 154,
        'fin_revenue': 156,
        'fin_gross_profit': 157,
        'fin_operating_profit': 158,
        'fin_ordinary_profit': 159,
        'fin_depreciation': 160,
        'fin_personnel': 161,
        'fin_capital': 162,

        # 経営状況 ※+9
        'management_intent': 168,
        'strength': 170,
        'weakness': 172,
        'it_investment_status': 173,
        'it_investment_amount': 174,
        'it_utilization_status': 175,
        'it_investment_process': 177,
        'security_status': 178,
        'improvement_process': 179,
        'expected_effect_dept': 180,
        'expected_effect': 181,
        'future_goals': 182,

        # 計画数値入力 ※+9
        'min_wage': 213,
        'min_wage_hourly': 214,
        'employee_count_fte': 215,
        'wage_raise_declaration': 222,
        'wage_raise_amount': 223,
        'wage_raise_method': 224,
        'wage_raise_date': 225,
    },
    kyuyo_sheet_name='生産性指標給与支給総額計算',
    kyuyo={
        'revenue':          (10, 2),  # B10: 売上高
        'gross_profit':     (11, 2),  # B11: 粗利益
        'operating_profit': (12, 2),  # B12: 営業利益
        'ordinary_profit':  (13, 2),  # B13: 経常利益
        'depreciation':     (21, 5),  # E21: 減価償却費
        'salary':           (5,  5),  # E5:  給料手当
        'misc_wages':       (6,  5),  # E6:  雑給
        'bonus':            (7,  5),  # E7:  賞与手当
        'officer_comp':     (4,  5),  # E4:  役員報酬（D4ラベル参照）
        'travel_expense':   (9,  5),  # E9:  旅費交通費
    },
    shinsei_clear_range=(5, 270),
    tenki_text_range=(15, 26),
)

# ── 2026 インボイス枠 ──
# テンプレート原本: ツール/【原本_法人】企業名_インボイス枠_法人2026.xlsx
# ヒアリングシート: ツール/ヒアリングシート2026_インボイス法人.xlsx
MAPPING_2026_INVOICE = TemplateMapping(
    hearing_to_tenki=[
        # (ヒアリング行, 転記行, 電話番号変換)
        (8,  8,  False),   # 企業名
        (10, 10, False),   # 企業名フリガナ
        (12, 12, False),   # 店舗事業所数
        (14, 14, False),   # 事業者URL
        (17, 16, False),   # 主な事業内容
        (18, 17, False),   # 強み
        (19, 18, False),   # 時間がかかっている業務
        (20, 19, False),   # 月間何時間
        (21, 20, False),   # どの機能で楽にしたいか
        (22, 21, False),   # 何%削減
        (23, 22, False),   # 浮いた時間の活用
        (24, 23, False),   # 3年後の売上目標
        (25, 24, False),   # 取引先属性
        (28, 26, True),    # 代表電話番号
        (30, 28, False),   # 担当者氏名
        (32, 30, False),   # 担当者氏名フリガナ
        (34, 32, False),   # 担当者メールアドレス
        (36, 34, True),    # 担当者電話番号
        (38, 36, True),    # 担当者携帯番号
        (40, 38, False),   # 正規雇用
        (42, 40, False),   # 契約社員
        (44, 42, False),   # パートアルバイト
        (46, 44, False),   # 派遣社員
        (48, 46, False),   # その他
        (50, 48, False),   # 過去に補助金
        (52, 50, False),   # 申請年度
        (53, 51, False),   # 申請枠
        (54, 52, False),   # 申請回
        (58, 56, False),   # えるぼし
        (61, 59, False),   # くるみん
        (63, 61, False),   # SECURITY ACTION ID
        (66, 64, False),   # 正規雇用(前期)
        (67, 65, False),   # 契約社員(前期)
        (68, 66, False),   # パート(前期)
        (69, 67, False),   # 年間平均労働時間
        (72, 70, False),   # 事業所内最低賃金
        (74, 72, False),   # 賃金引上げ表明
        (76, 74, False),   # 賃金引上げ幅
        (79, 76, False),   # 従業員代表者
        (80, 77, False),   # 給与担当者
        (81, 78, False),   # 事業所内最低賃金者
        (84, 81, False),   # インボイス登録状況
        (86, 83, False),   # インボイス登録予定
    ],
    shinsei={
        # 基本情報入力 (行39〜) ※役員枠より前なのでズレなし
        'headquarters_address': 45,
        'industry_code': 46,
        'industry_text': 47,
        'established_date': 48,
        'capital': 49,
        'business_description': 63,
        'fiscal_month': 64,
        'tool_name': 61,

        # 担当者・代表者 ※ズレなし
        'officer_count': 71,
        'rep_title': 72,
        'rep_name': 73,
        'rep_kana': 74,

        # 役員 (最大10名) ※(1)〜(7)はズレなし、(8)〜(10)は新規
        'officer_1_title': 77,
        'officer_1_name': 78,
        'officer_1_kana': 79,
        'officer_2_title': 80,
        'officer_2_name': 81,
        'officer_2_kana': 82,
        'officer_3_title': 83,
        'officer_3_name': 84,
        'officer_3_kana': 85,
        'officer_4_title': 86,
        'officer_4_name': 87,
        'officer_4_kana': 88,
        'officer_5_title': 89,
        'officer_5_name': 90,
        'officer_5_kana': 91,
        'officer_6_title': 92,
        'officer_6_name': 93,
        'officer_6_kana': 94,
        'officer_7_title': 95,
        'officer_7_name': 96,
        'officer_7_kana': 97,
        'officer_8_title': 98,
        'officer_8_name': 99,
        'officer_8_kana': 100,
        'officer_9_title': 101,
        'officer_9_name': 102,
        'officer_9_kana': 103,
        'officer_10_title': 104,
        'officer_10_name': 105,
        'officer_10_kana': 106,

        # その他 ※+9
        'past_subsidies': 112,
        'eruboshi': 116,
        'kurumin': 117,
        'business_types': 137,

        # 財務情報入力 ※+9
        'officer_count_prev': 146,
        'fin_revenue': 148,
        'fin_gross_profit': 149,
        'fin_operating_profit': 150,
        'fin_ordinary_profit': 151,
        'fin_depreciation': 152,
        'fin_personnel': 153,
        'fin_capital': 154,

        # 経営状況 ※+9
        'management_intent': 160,
        'security_status': 161,
        'future_goals': 162,
        'it_investment_status': 163,
        'it_utilization_scope': 164,
        'invoice_related_work': 165,

        # 計画数値入力 ※+9
        # 給与支給総額の計画値（C200:C204）
        'employee_count_fte': 200,    # 従業員数（FTE換算）
        'wage_total_base': 201,       # 直近決算期の給与支給総額
        'wage_total_y1': 202,         # 1年目計画（2025/4〜2026/3）
        'wage_total_y2': 203,         # 2年目計画（2026/4〜2027/3）
        'wage_total_y3': 204,         # 3年目計画（2027/4〜2028/3）

        'min_wage': 198,
        'min_wage_hourly': 199,
        'wage_raise_declaration': 207,
        'wage_raise_amount': 208,
        'wage_raise_method': 209,
        'wage_raise_date': 210,
    },
    kyuyo_sheet_name='給与支給総額計算',
    kyuyo={
        'revenue':          (10, 2),  # B10: 売上高
        'gross_profit':     (11, 2),  # B11: 粗利益
        'operating_profit': (12, 2),  # B12: 営業利益
        'ordinary_profit':  (13, 2),  # B13: 経常利益
        'depreciation':     (16, 5),  # E16: 減価償却費
        'salary':           (5,  5),  # E5:  給料手当
        'misc_wages':       (6,  5),  # E6:  雑給
        'bonus':            (7,  5),  # E7:  賞与手当
        'travel_expense':   (9,  5),  # E9:  旅費交通費
    },
    shinsei_clear_range=(5, 270),
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
