# -*- coding: utf-8 -*-
"""通常枠ヒアリングシートの選択肢（番号→本文）マスターと番号入力のデコード。

通常枠法人ヒアリングシートの「自社の強み(C71)」「自社の弱み(C72)」
「IT投資プロセス(C74)」は複数選択可。従来は単一選択プルダウンで本文が
そのまま入っていたが、配布マスター（Google Sheets）でプルダウンを廃し
「該当番号を記入（複数可、例 1,4,7）」に変更された。番号のまま下流に流すと:
  - normalize_value がカンマ区切りを連結整数に破壊する（"1,4,7"→147）
  - 転記シート→申請書の手動コピー / AI 事業内容生成で意味が失われる
そこで読み取り時に「N 本文」形式へ復元する。番号はプルダウンシート
（A列=強み / B列=弱み / C列=IT投資プロセス）の並び＝実シート番号に準拠。

対象は通常枠のみ。インボイス法人/個人にはこれらの番号選択欄が無い
（自由記述の「他社に負けない貴社の『強み』」だけ）。ラベルを「自社の強み」
「自社の弱み」＋(IT投資∧プロセス) に厳密一致させ、値が番号だけのときのみ
変換するため、自由記述の強みや他枠に誤発火しない。
"""
from __future__ import annotations

import re

# ── プルダウンシート A/B/C 列（実シート番号に準拠、本文のみ保持し番号は付与） ──
_STRENGTH = [
    '独自性・独創性', '営業力', '商圏・立地', '製品サービスの質',
    '商品・サービスの情報発信力', '顧客情報の収集・管理', '人材力', '技術力',
    '充実した設備力', 'ビジネスモデル', '特許などの知財', '社内チームワーク',
    '協力会社等との外部連携力', '伝統や長い社歴', '新製品や新サービスなどの開発力',
]
_WEAKNESS = [
    '競合他社との差別化が図れていない', '人材不足', '商圏・立地', '製品サービスの質',
    '商品・サービスの情報発信不足', '顧客情報の不足',
    '在庫管理・工程管理等、業務管理がうまく把握できていない', '社員の高齢化や退職',
    '人が育たない', '設備の陳腐化', '運転資金不足', '設備投資資金不足',
]
_IT_PROCESS = [
    '販売や店頭といったフロント業務の強化',
    '顧客のニーズや流行等を捉え、新規顧客獲得や新規市場開拓を行った',
    '事前の準備工程（施策、テスト、設計や計画立案、など）を強化',
    '生産管理・在庫管理・物流管理など、商品の動きの可視化・効率化',
    '案件管理・工程管理・進捗管理といった業務管理の可視化・効率化',
    '営業（現場）の業務効率化を図った',
    '人員配置の最適化を行った',
    '会計業務や清算業務の正確性・効率化を図った',
    '勤務時間の短縮・労働時間の適正化を図った',
    '単純な事務作業を自動化し、人手や時間の無駄を削った',
    '取引先や社内での情報共有を強化した',
]


def _to_map(items: list[str]) -> dict[int, str]:
    """本文リスト → {番号: 'N 本文'}（番号は1始まり）"""
    return {i: f'{i} {t}' for i, t in enumerate(items, start=1)}


STRENGTH_CHOICES = _to_map(_STRENGTH)
WEAKNESS_CHOICES = _to_map(_WEAKNESS)
IT_INVESTMENT_PROCESS_CHOICES = _to_map(_IT_PROCESS)

# 複数選択の区切り: 半角/全角カンマ・読点・スラッシュ・空白
_SEP = r'[,，、/／\s]+'
_JOIN = '、'


def _master_for_label(label) -> dict[int, str] | None:
    """ラベルから対象マスターを厳密に判定。自由記述の強み等には当てない。"""
    if label is None:
        return None
    s = str(label)
    if '自社の強み' in s:
        return STRENGTH_CHOICES
    if '自社の弱み' in s:
        return WEAKNESS_CHOICES
    if 'IT投資' in s and 'プロセス' in s:
        return IT_INVESTMENT_PROCESS_CHOICES
    return None


def decode_choice_field(label, raw_value):
    """選択肢フィールドのセル値を「N 本文」形式へ復元する。

    - 対象外ラベル → None（呼び出し側で normalize_value を使う）
    - 値が番号（＋区切り）だけ → 復元した文字列を返す（複数は「、」連結）
    - 既に本文を含む値（旧プルダウンのフルテキスト等） → None（非破壊、従来処理に委ねる）
    未知番号はマスターに無ければ番号のまま温存（人手確認用）。
    """
    master = _master_for_label(label)
    if master is None or raw_value is None:
        return None

    # Excel が数値として保持した単一回答（例: "1" → int 1）
    if isinstance(raw_value, (int, float)):
        n = int(raw_value)
        return master.get(n, str(n))

    s = str(raw_value).strip()
    if not s:
        return None

    # 番号＋区切りだけか（全角数字も Python の \d/int が解釈する）
    stripped = re.sub(_SEP, '', s)
    if not (stripped and stripped.isdigit()):
        return None  # 本文入り等 → 従来処理へ委ねる

    nums = [int(tok) for tok in re.findall(r'\d+', s)]
    parts = [master.get(n, str(n)) for n in nums]
    return _JOIN.join(parts) if parts else None


def apply_hearing_choice_overrides(ai_judgment, hearing_data) -> None:
    """通常枠: 顧客がヒアリングで選んだ 弱み/IT投資プロセス（複数可）を申請書へ反映。

    hearing_data の値は read_hearing_sheet で既に「N 本文」へデコード済み。
    顧客が選択している（非空）ときだけ AI 生成値を上書きし、空なら AI 値を温存する。
    ※呼び出し側で通常枠のときのみ実行する前提（インボイス/個人には該当ラベルが無い）。
    """
    if not hearing_data or ai_judgment is None:
        return
    for item in hearing_data.values():
        label = str(item.get('label', ''))
        value = item.get('value')
        if value is None or not str(value).strip():
            continue
        if '自社の弱み' in label:
            ai_judgment.weakness = str(value)
        elif 'IT投資' in label and 'プロセス' in label:
            ai_judgment.it_investment_process = str(value)
