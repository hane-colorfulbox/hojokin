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


def check_bonus_omission(
    ledger_employees: list[dict] | None,
    financial,
) -> str:
    """賞与シート未参照を検出。

    健全な賃金台帳: 月別合計 ≒ PL の (給料+雑給+賞与)（賞与込みで集計されている）
    賞与未参照: 月別合計 ≒ PL の (給料+雑給) のみ（賞与分が抽出から抜けている）

    本番で観測された23%差ケース: 賞与が別タブシートにあり、Haiku が拾えなかった。
    PL の bonus と賃金台帳合計の関係から、この状態を機械的に検出する。
    """
    if not ledger_employees or not financial:
        return ''

    # 賃金台帳合計（役員除外、月別値の総和）
    ledger_total = 0
    for e in ledger_employees:
        if '役員' in (e.get('employment_type') or ''):
            continue
        wages = e.get('monthly_wages') or []
        ledger_total += sum(
            w for w in wages if isinstance(w, (int, float)) and w > 0
        )
    if ledger_total <= 0:
        return ''

    salary = (getattr(financial, 'salary', None) or 0) + (
        getattr(financial, 'misc_wages', None) or 0
    )
    bonus = getattr(financial, 'bonus', None) or 0

    if salary <= 0 or bonus <= 0:
        return ''
    # 賞与が小さすぎる場合は検出意義薄
    if bonus / salary < 0.05:
        return ''

    # 二仮説比較: 賃金台帳合計が「salary-only」と「salary+bonus」のどちらに近いか。
    # salary+bonus の方が近い、または同等なら賞与込みの健全パターン → 警告なし。
    # salary-only に明確に近い場合のみ「賞与未参照」と判定する（Codex 指摘の false positive 対策）。
    diff_to_salary_only = abs(ledger_total - salary)
    diff_to_full = abs(ledger_total - (salary + bonus))
    if diff_to_salary_only >= diff_to_full:
        return ''  # 賞与込み集計の方が近い = 健全

    # salary-only により近いとして、それでも誤差大ければ別問題（人数異常など）
    if diff_to_salary_only / salary > 0.10:
        return ''

    return (
        f' ⚠ 賞与未参照の可能性: 賃金台帳合計({ledger_total:,}円)が損益計算書の'
        f'給料手当+雑給({salary:,}円)とほぼ一致し、PL に計上されている'
        f'賞与{bonus:,}円が抽出に含まれていません。'
        f'賃金台帳PDFに賞与シートが別タブで存在する場合、Haiku/Document AI 経路で'
        f'拾えていない可能性があります'
    )


def check_similar_name_duplicates(ledger_employees: list[dict] | None) -> str:
    """OCR 誤読による別人扱いが疑われる類似氏名ペアを検出。

    判定: 「姓が同じ」かつ「名が編集距離1以内」の従業員ペア。
    実観測例: 「吉田 壽」と「吉田 靖」、「大嶋 晃輔」と「大崎 晃輔」
    （これらは _normalize_name_key の異体字辞書で救えるものは既に統合済み。
    残るのはより微妙な誤読ペア）

    統合は危険（実は別人の可能性もある）ので、警告のみ返して人間判断に委ねる。
    """
    if not ledger_employees or len(ledger_employees) < 2:
        return ''

    def _surname_given(name: str) -> tuple[str, str]:
        """姓名を空白で分割。空白が無ければ最初の2文字を姓・残りを名とする"""
        import unicodedata
        n = unicodedata.normalize('NFKC', name or '').strip()
        for sep in (' ', '　', '\t'):
            if sep in n:
                a, b = n.split(sep, 1)
                return a.strip(), b.strip()
        # 空白なし → 先頭2文字を姓と仮定
        if len(n) >= 3:
            return n[:2], n[2:]
        return n, ''

    def _edit_distance_le_1(a: str, b: str) -> bool:
        """編集距離が 1 以下か（最大長差1で済むかも判定）"""
        if a == b:
            return True
        if abs(len(a) - len(b)) > 1:
            return False
        if len(a) == len(b):
            diffs = sum(1 for x, y in zip(a, b) if x != y)
            return diffs <= 1
        # 長さ差1: 長い方から1文字消して一致するか
        long_s, short_s = (a, b) if len(a) > len(b) else (b, a)
        for i in range(len(long_s)):
            if long_s[:i] + long_s[i+1:] == short_s:
                return True
        return False

    pairs = []
    parsed = [(_surname_given(e.get('name') or ''), e.get('name') or '') for e in ledger_employees]
    n = len(parsed)
    for i in range(n):
        (sa, ga), na = parsed[i]
        if not sa or not ga:
            continue
        for j in range(i + 1, n):
            (sb, gb), nb = parsed[j]
            if not sb or not gb:
                continue
            if sa == sb and _edit_distance_le_1(ga, gb) and ga != gb:
                pairs.append(f'「{na}」と「{nb}」')

    if not pairs:
        return ''
    sample = '、'.join(pairs[:3])
    suffix = '...' if len(pairs) > 3 else ''
    return (
        f' ⚠ 類似氏名ペア{len(pairs)}件あり: {sample}{suffix}。'
        f'OCR誤読により同一人物が別人扱いされている可能性があります（壽⇔靖等）'
    )


def check_employment_type_missing(ledger_employees: list[dict] | None) -> str:
    """雇用区分（employment_type）が空欄の従業員を検出。

    プロンプトで「明示されていない場合は『正社員』を既定値」と指示しているが、
    指示違反した出力を検出する。
    """
    if not ledger_employees:
        return ''
    missing = [
        e.get('name') or '(無名)'
        for e in ledger_employees
        if not (e.get('employment_type') or '').strip()
    ]
    if not missing:
        return ''
    if len(missing) >= max(2, len(ledger_employees) // 3):
        sample = '、'.join(missing[:3])
        suffix = '...' if len(missing) > 3 else ''
        return (
            f' ⚠ 雇用区分が空欄の従業員が{len(missing)}名います: {sample}{suffix}。'
            f'抽出時に雇用形態列を読み取れていない可能性があります'
        )
    return ''


def run_all_validations(
    hearing_data: dict | None,
    ledger_employees: list[dict] | None,
    financial=None,
) -> list[str]:
    """全検証を実行し、警告文字列のリストを返す（空は除外）。

    Args:
        hearing_data: ヒアリングシート読込結果（{行番号: {label, value}}）
        ledger_employees: 賃金台帳から抽出された従業員リスト
        financial: PL 抽出結果（FinancialData）。賞与未参照検出に使用
    """
    warnings = [
        check_employee_count_mismatch(hearing_data, ledger_employees),
        check_monthly_coverage(ledger_employees),
        check_value_distribution(ledger_employees),
        check_bonus_omission(ledger_employees, financial),
        check_similar_name_duplicates(ledger_employees),
        check_employment_type_missing(ledger_employees),
    ]
    return [w for w in warnings if w]
