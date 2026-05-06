# -*- coding: utf-8 -*-
"""賃金台帳抽出結果の自動品質検証

extract_wage_ledger / 決定論パーサーの出力に対して、抽出ミスのサインを検知し
警告メッセージとして返す。

検証項目:
1. 人数妥当性 — 前期従業員数（ヒアリングシート）と賃金台帳抽出人数の乖離
2. 月別カバレッジ — 各従業員の non-null 月数が極端に少ないものを検出
3. 値分布の異常 — 月別最大値が中央値の N 倍超（年間合計が混入しているサイン）

各関数は警告文字列（不整合あり）または空文字列（整合 / 判定不能）を返す。
pipeline 側で run_all_validations() を呼び、status.message に追記する。

API 呼出ゼロ。決定論的に動く（同じ入力 → 同じ出力）。
"""
from __future__ import annotations

import logging
from statistics import median

logger = logging.getLogger(__name__)

# 月別値が中央値の何倍を超えたら「年間合計混入」を疑うか。
# 健全な従業員は給与変動があっても max/median <= 2 程度。
# 3 倍超は明らかに異常（年間合計値 ≒ 中央値 × 12 のため、12 倍前後で出ることが多い）。
VALUE_OUTLIER_RATIO_THRESHOLD = 3.0

# 人数乖離の許容範囲（前期従業員数 ± この割合）
EMPLOYEE_COUNT_TOLERANCE = 0.30

# 月別データが「極端に少ない」と判定する non-null 月数の上限
MIN_MONTHLY_COVERAGE = 2


def _find_hearing_value(hearing_data: dict | None, *required_keywords: str) -> int | None:
    """ヒアリングシートからラベルに全キーワードを含む行の値（数値）を探す。

    テンプレート（通常枠/インボイス/個人）で行番号が異なるため、ラベルベースで検索。
    数値以外（文字列など）は None を返す。
    """
    if not hearing_data:
        return None
    for entry in hearing_data.values():
        if not isinstance(entry, dict):
            continue
        label = entry.get('label') or ''
        if all(kw in label for kw in required_keywords):
            v = entry.get('value')
            if isinstance(v, (int, float)):
                return int(v)
    return None


def check_employee_count_mismatch(
    hearing_data: dict | None,
    ledger_employees: list[dict] | None,
) -> str:
    """賃金台帳抽出人数と前期従業員数の乖離を検出。

    前期従業員数 = 正規雇用(前期) + 契約社員(前期) + パート(前期)（役員除く）
    賃金台帳人数 = ledger_employees のうち役員以外
    """
    if not hearing_data or not ledger_employees:
        return ''

    ledger_count = sum(
        1 for e in ledger_employees
        if '役員' not in (e.get('employment_type') or '')
    )
    if ledger_count == 0:
        return ''

    seishain = _find_hearing_value(hearing_data, '正規雇用', '前期') or 0
    keiyaku = _find_hearing_value(hearing_data, '契約社員', '前期') or 0
    part = _find_hearing_value(hearing_data, 'パート', '前期') or 0
    expected = seishain + keiyaku + part
    if expected <= 0:
        return ''

    diff_ratio = abs(ledger_count - expected) / expected
    if diff_ratio <= EMPLOYEE_COUNT_TOLERANCE:
        return ''

    direction = '少ない' if ledger_count < expected else '多い'
    return (
        f' ⚠ 賃金台帳の抽出人数({ledger_count}名)と前期従業員数({expected}名)に'
        f'{diff_ratio*100:.0f}%の乖離があります（賃金台帳が{direction}）。'
        f'抽出漏れ・退職者の混入・雇用区分誤認などを確認してください'
    )


def check_monthly_coverage(ledger_employees: list[dict] | None) -> str:
    """月別データのカバレッジが極端に低い従業員を検出。

    monthly_wages の non-null かつ > 0 の月数が MIN_MONTHLY_COVERAGE 以下の従業員が
    複数いれば警告（中途入退社の可能性もあるが、複数いるなら抽出漏れの方が疑わしい）。
    """
    if not ledger_employees:
        return ''

    insufficient = []
    for e in ledger_employees:
        wages = e.get('monthly_wages') or []
        if len(wages) != 12:
            continue
        non_null = sum(1 for w in wages if w is not None and isinstance(w, (int, float)) and w > 0)
        if non_null <= MIN_MONTHLY_COVERAGE:
            name = e.get('name') or '不明'
            insufficient.append(f'{name}({non_null}ヶ月)')

    if len(insufficient) < 2:
        return ''

    sample = '、'.join(insufficient[:3])
    suffix = '...' if len(insufficient) > 3 else ''
    return (
        f' ⚠ 月別データが極端に少ない従業員が{len(insufficient)}名います: '
        f'{sample}{suffix}。中途入退社でなければ抽出漏れの可能性があります'
    )


def check_value_distribution(ledger_employees: list[dict] | None) -> str:
    """月別値の異常分布を検出（年間合計が月別セルに混入しているサイン）。

    各従業員の monthly_wages について max / median > VALUE_OUTLIER_RATIO_THRESHOLD なら警告。
    年間合計が紛れ込むと max ≒ median × 12 になるため確実に検出できる。
    """
    if not ledger_employees:
        return ''

    suspicious = []
    for e in ledger_employees:
        wages = e.get('monthly_wages') or []
        non_null = [
            w for w in wages
            if w is not None and isinstance(w, (int, float)) and w > 0
        ]
        if len(non_null) < 3:
            # サンプルが少ないと中央値の信頼性が低いのでスキップ
            continue
        med = median(non_null)
        mx = max(non_null)
        if med > 0 and mx / med > VALUE_OUTLIER_RATIO_THRESHOLD:
            name = e.get('name') or '不明'
            suspicious.append(f'{name}(max/median={mx/med:.1f}倍)')

    if not suspicious:
        return ''

    sample = '、'.join(suspicious[:3])
    suffix = '...' if len(suspicious) > 3 else ''
    return (
        f' ⚠ 月別値の分布が異常な従業員が{len(suspicious)}名います: '
        f'{sample}{suffix}。年間合計が月別セルに混入している可能性があります'
    )


def run_all_validations(
    hearing_data: dict | None,
    ledger_employees: list[dict] | None,
) -> list[str]:
    """全検証を実行し、警告文字列のリストを返す（空は除外）。"""
    warnings = [
        check_employee_count_mismatch(hearing_data, ledger_employees),
        check_monthly_coverage(ledger_employees),
        check_value_distribution(ledger_employees),
    ]
    return [w for w in warnings if w]
