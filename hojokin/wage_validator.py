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


def _emp_get(emp, key, default=None):
    """dict / WageEmployee（dataclass）両対応のアクセサ。

    pipeline からは WageEmployee オブジェクトのリストが渡されるが、
    テスト・サブエージェント検証では dict のリストが渡されるため、両方を吸収する。
    """
    if emp is None:
        return default
    if isinstance(emp, dict):
        return emp.get(key, default)
    return getattr(emp, key, default)

# 月別値が中央値の何倍を超えたら「年間合計混入」を疑うか。
# 健全な従業員は給与変動があっても max/median <= 2 程度。
# **賞与月は月給の 3〜5 倍**になるのが正常なので、閾値は 5 倍より上に取る。
# 年間合計混入は max/median ≒ 12 になるため、5 倍超で検出すれば誤検出を避けつつ
# 年合計混入を確実に拾える。
VALUE_OUTLIER_RATIO_THRESHOLD = 5.0

# 人数乖離の許容範囲（前期従業員数 ± この割合）
EMPLOYEE_COUNT_TOLERANCE = 0.30

# 月別データが「極端に少ない」と判定する non-null 月数の上限
MIN_MONTHLY_COVERAGE = 2

# 抽出人数が PL 推定人数の何割を下回ったら「明らかに不足」と判定するか。
# 0.5 = 半数未満で警告（抽出処理が大量に取りこぼしているサイン）。
# 半数までは「中途入退社・パートタイム比率の高さ」等で許容範囲とする。
EXTRACTION_SIZE_VS_PL_TOLERANCE_RATIO = 0.5

# PL 人件費から推定する「最低 1 人当たり年収」（円）。これより小さい1人あたり給与は
# 現実的に存在しないため、この値で割って最低限の従業員数を推定する。
# 200万円 = パート想定の年収下限
EXTRACTION_SIZE_VS_PL_MIN_PER_PERSON = 2_000_000


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
        if '役員' not in (_emp_get(e, 'employment_type') or '')
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
        wages = _emp_get(e, 'monthly_wages') or []
        if len(wages) != 12:
            continue
        non_null = sum(1 for w in wages if w is not None and isinstance(w, (int, float)) and w > 0)
        if non_null <= MIN_MONTHLY_COVERAGE:
            name = _emp_get(e, 'name') or '不明'
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
        wages = _emp_get(e, 'monthly_wages') or []
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
            name = _emp_get(e, 'name') or '不明'
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
        if '役員' in (_emp_get(e, 'employment_type') or ''):
            continue
        wages = _emp_get(e, 'monthly_wages') or []
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
    parsed = [(_surname_given(_emp_get(e, 'name') or ''), _emp_get(e, 'name') or '') for e in ledger_employees]
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


def check_extraction_size_vs_pl(
    ledger_employees: list[dict] | None,
    financial,
) -> str:
    """PL の人件費規模に対して抽出人数が極端に少ない場合を検出。

    hearing_data が無い時でも「明らかに抽出失敗」を捉えるためのフォールバック。
    判定: PL の人件費（給料+雑給+賞与）から推定される人数(規模) に対して、
    抽出人数が著しく少ないなら警告。

    推定: 1人あたり年収の現実的下限 200万円（パート想定）として、
    PL人件費 / 200万 を超えていない人数なら「明らかに不足」。
    """
    if not ledger_employees or not financial:
        return ''
    salary = (
        (getattr(financial, 'salary', None) or 0)
        + (getattr(financial, 'misc_wages', None) or 0)
        + (getattr(financial, 'bonus', None) or 0)
    )
    if salary <= 0:
        return ''
    ledger_count = sum(
        1 for e in ledger_employees
        if '役員' not in (_emp_get(e, 'employment_type') or '')
    )
    # PL人件費 / EXTRACTION_SIZE_VS_PL_MIN_PER_PERSON = 最低限の従業員数の目安
    min_expected = max(1, int(salary / EXTRACTION_SIZE_VS_PL_MIN_PER_PERSON))
    # 半数まで許容: 「中途入退社が複数 / パート比率の高さ」等で正規分布から外れる
    # 範囲を見越して、min_expected の半分（浮動小数）を下回って初めて警告する。
    # 例: min_expected=5 なら threshold=2.5 → 抽出 ledger_count=2 で警告（< 2.5）、3 で許容（>= 2.5）
    threshold = min_expected * EXTRACTION_SIZE_VS_PL_TOLERANCE_RATIO
    if ledger_count >= threshold:
        return ''
    return (
        f' ⚠ 抽出従業員数({ledger_count}名)が PL人件費規模({salary:,}円)に対して'
        f'明らかに少ないです（最低{min_expected}名期待・許容下限{threshold:.1f}名）。'
        f'抽出処理が大量に取りこぼしている可能性があります（OCR失敗・大型PDF処理失敗等）'
    )


def check_employment_type_missing(ledger_employees: list[dict] | None) -> str:
    """雇用区分（employment_type）が空欄または「(推定)」付きの従業員を検出。

    補完前: 空文字列。空文字列のままなら抽出失敗を示唆。
    補完後: 「正社員(推定)」のような provenance 付き値。これも警告対象として
    人間チェックを促す（wage_reader.py で「(推定)」を付ける運用と整合）。
    """
    if not ledger_employees:
        return ''
    inferred = []
    empty = []
    for e in ledger_employees:
        et = (_emp_get(e, 'employment_type') or '').strip()
        name = _emp_get(e, 'name') or '(無名)'
        if not et:
            empty.append(name)
        elif '(推定)' in et or '(推測)' in et:
            inferred.append(name)
    flagged = empty + inferred
    if not flagged:
        return ''
    if len(flagged) >= max(2, len(ledger_employees) // 3):
        sample = '、'.join(flagged[:3])
        suffix = '...' if len(flagged) > 3 else ''
        kind = '空欄' if not inferred else '空欄/推定'
        return (
            f' ⚠ 雇用区分が{kind}の従業員が{len(flagged)}名います: {sample}{suffix}。'
            f'抽出時に雇用形態列を読み取れていない可能性があり、人事担当の確認が必要です'
        )
    return ''


# ── セル単位整合性チェック ─────────────────────────────────────
# PDFテキストから決定論的に取得した「物理列構造」と AI 出力の monthly_wages を
# セル単位で突合し、月給漏れ・月給誤配置・賞与漏れを検知する。
#
# 既存の PL 突合系チェックは「年間合計のマクロ整合」しか見えない（差 4-5% 程度
# だと素通り）ため、セル単位の漏れは検出できない。本チェックで補う。

# 賞与候補と見なすしきい値: 月給平均の何倍以上を「賞与込みの月」と推定するか。
# 賞与は通常 1〜3 ヶ月分のため、1.5 倍超で十分。低めに取って取りこぼしを減らす。
BONUS_DETECT_RATIO = 1.5

# 「基本給とほぼ同額」の判定許容率（賞与漏れ検知に使う）。
# 諸手当の月変動を許容して 5%。
BONUS_OMISSION_TOLERANCE = 0.05

# 「役員/正社員の定額連続性」検出: monthly_wages の (非null) 月の変動係数 (std/mean)
# がこの値より小さければ「定額」と判定する
DEFAULT_FLAT_CV_THRESHOLD = 0.05


def _name_match_key(name: str) -> str:
    """姓名の空白・記号差を吸収して比較用キーを返す。"""
    import unicodedata
    if not name:
        return ''
    s = unicodedata.normalize('NFKC', name).strip()
    s = re.sub(r'[\s\-_·・]+', '', s)
    return s


# `re` を遅延 import するため、関数内で必要時に呼ぶ:
import re  # noqa: E402


def check_cell_level_consistency(
    ledger_employees: list[dict] | None,
    pdf_layout,
) -> list[str]:
    """PDFレイアウト情報と AI 抽出結果のセル単位整合性をチェックする。

    Args:
        ledger_employees: AI 抽出された従業員リスト（dict or WageEmployee）
        pdf_layout: `wage_pdf_layout_parser.parse_wage_ledger_layout` の戻り値
            （PdfEmployee のリスト）。None または空なら検証スキップ。

    Returns:
        警告メッセージのリスト。「種類別・従業員別」で複数行に分けて返す。

    検証4種:
        C1: 月給漏れ — PDF に X月分列があるのに AI 出力で該当月 null
        C2: 月給誤配置 — PDF に X月分列がないのに AI 出力で該当月に値あり
            （賞与の支給月は例外）
        C3: 賞与漏れ — PDF に賞与記載があるのに、AI 出力の対応月セルに
            賞与額が反映されていない
        C4: 定額連続性の途切れ — 役員・正社員で「他月は同額なのに特定月だけ空白」
            （C1のフォールバック）
    """
    if not ledger_employees or not pdf_layout:
        return []

    # PDFレイアウト側を氏名キーで引けるようにする
    pdf_by_key = {}
    for pe in pdf_layout:
        key = _name_match_key(getattr(pe, 'name', '') or '')
        if key:
            pdf_by_key[key] = pe

    if not pdf_by_key:
        return []

    c1_missing: list[str] = []   # 月給漏れ
    c2_misplaced: list[str] = [] # 月給誤配置
    c3_bonus_lost: list[str] = []  # 賞与漏れ
    c4_flat_gap: list[str] = []  # 定額連続性の途切れ

    for emp in ledger_employees:
        name = _emp_get(emp, 'name') or ''
        key = _name_match_key(name)
        pe = pdf_by_key.get(key)
        if pe is None:
            continue  # PDF側に該当氏名が無い場合は別チェックの管轄

        wages = _emp_get(emp, 'monthly_wages') or []
        if len(wages) != 12:
            continue  # 形式異常は別チェックで検出
        emp_type = (_emp_get(emp, 'employment_type') or '')

        source_months = list(getattr(pe, 'source_months', []) or [])
        bonus_pays = list(getattr(pe, 'bonus_pays', []) or [])
        bonus_months = {mon for (mon, _amt) in bonus_pays}
        has_bonus = bool(getattr(pe, 'has_bonus_section', False))

        # C1: 月給漏れ
        # 「PDFに○月分列があるのに AI が null」のみ警告。
        # PDFテキスト側の月別値が 0/空欄なら「実際に支給なし」なので除外する
        # （長期欠勤など）。
        taxable = getattr(pe, 'monthly_taxable_totals', {}) or {}
        basic = getattr(pe, 'monthly_basic_pay', {}) or {}
        for mon in source_months:
            ai_val = wages[mon - 1] if 1 <= mon <= 12 else None
            if ai_val is not None and ai_val > 0:
                continue
            # PDF側に金額が確認できる月だけを「漏れ」と判定する
            pdf_val = taxable.get(mon, 0) or basic.get(mon, 0)
            if pdf_val > 0:
                c1_missing.append(
                    f'  - {name}（{emp_type}）: {mon}月セルが空白。'
                    f'PDFに「{mon}月分/月度」列があり、'
                    f'基本給/総支給額(課税)={pdf_val:,}円が読み取れます'
                )

        # C2: 月給誤配置
        # PDFに○月分列が無い × AI 出力に値あり × 賞与支給月でもない
        for mon in range(1, 13):
            ai_val = wages[mon - 1]
            if ai_val is None or ai_val <= 0:
                continue
            if mon in source_months:
                continue
            if mon in bonus_months:
                continue  # 賞与の支給月は許容
            c2_misplaced.append(
                f'  - {name}（{emp_type}）: {mon}月セルに {ai_val:,}円。'
                f'PDFには「{mon}月分/月度」列も賞与支給日({mon}月)もありません'
            )

        # C3: 賞与漏れ
        # bonus_pays に (mon, amount) があるのに、AI 出力の該当月セルが
        # 「基本給とほぼ同額」(=賞与が加算されていない) なら警告
        for (mon, amount) in bonus_pays:
            if not (1 <= mon <= 12):
                continue
            ai_val = wages[mon - 1]
            if ai_val is None:
                # 該当月が null = 賞与どころか月給自体が空（C1 で別途検出）
                continue
            base_value = basic.get(mon, 0)
            if base_value <= 0:
                # 基本給が PDF から読めなかった月は基準値が立てられない → 警告控えめに
                if amount and ai_val < amount * 0.5:
                    # 賞与額の半分未満 = 明らかに加算されていない
                    c3_bonus_lost.append(
                        f'  - {name}（{emp_type}）: {mon}月の賞与（PDF推定 {amount:,}円）'
                        f'が AI 出力に反映されていません（AI 出力 = {ai_val:,}円）'
                    )
                continue
            # 基本給と AI 値の差が許容範囲内なら賞与は加算されていない
            diff_ratio = abs(ai_val - base_value) / base_value
            if diff_ratio <= BONUS_OMISSION_TOLERANCE:
                amount_str = f'{amount:,}円' if amount else '（金額不明）'
                c3_bonus_lost.append(
                    f'  - {name}（{emp_type}）: {mon}月の賞与（PDF 推定 {amount_str}）'
                    f'が AI 出力に反映されていません'
                    f'（基本給 {base_value:,}円 ≒ AI 出力 {ai_val:,}円）'
                )

        # C4: 定額連続性の途切れ（C1のフォールバック）
        # PDF パースで source_months が取れなかったケース向け。
        # 役員・正社員で AI 出力の有効月が定額（変動係数 < しきい値）かつ、
        # 「途中の月だけ空白」の場合に警告。
        if not source_months and emp_type and ('役員' in emp_type or '正社員' in emp_type):
            valid_indices = [i for i, w in enumerate(wages) if w is not None and w > 0]
            if len(valid_indices) >= 6:
                vals = [wages[i] for i in valid_indices]
                mean = sum(vals) / len(vals)
                if mean > 0:
                    var = sum((v - mean) ** 2 for v in vals) / len(vals)
                    std = var ** 0.5
                    cv = std / mean
                    if cv < DEFAULT_FLAT_CV_THRESHOLD:
                        # 連続性チェック: 最初と最後の有効月の間に空白があれば異常
                        first, last = valid_indices[0], valid_indices[-1]
                        gaps = [
                            i + 1 for i in range(first, last + 1)
                            if wages[i] is None or wages[i] <= 0
                        ]
                        if gaps:
                            c4_flat_gap.append(
                                f'  - {name}（{emp_type}）: '
                                f'{first + 1}月〜{last + 1}月の在籍中に '
                                f'{gaps}月セルが空白。'
                                f'他月は {int(mean):,}円前後で定額のため、漏れの可能性'
                            )

    # 結果整形
    results: list[str] = []
    if c1_missing:
        results.append(
            ' ⚠ セル単位整合性: 月給漏れの可能性 '
            f'{len(c1_missing)}件 — PDF原本の該当月を確認してください\n'
            + '\n'.join(c1_missing)
        )
    if c3_bonus_lost:
        results.append(
            ' ⚠ セル単位整合性: 賞与漏れの可能性 '
            f'{len(c3_bonus_lost)}件 — PDF賞与ページの該当月を確認してください\n'
            + '\n'.join(c3_bonus_lost)
        )
    if c2_misplaced:
        results.append(
            ' ⚠ セル単位整合性: 月給誤配置の可能性 '
            f'{len(c2_misplaced)}件 — 月の入れ替えがないか確認してください\n'
            + '\n'.join(c2_misplaced)
        )
    if c4_flat_gap and not c1_missing:
        # C1 が出ているなら C4 はノイズになるので抑制
        results.append(
            ' ⚠ セル単位整合性: 定額連続性の途切れ '
            f'{len(c4_flat_gap)}件 — PDF未パース時のフォールバック検知\n'
            + '\n'.join(c4_flat_gap)
        )
    return results


def run_all_validations(
    hearing_data: dict | None,
    ledger_employees: list[dict] | None,
    financial=None,
    pdf_layout=None,
) -> list[str]:
    """全検証を実行し、警告文字列のリストを返す（空は除外）。

    Args:
        hearing_data: ヒアリングシート読込結果（{行番号: {label, value}}）
        ledger_employees: 賃金台帳から抽出された従業員リスト
        financial: PL 抽出結果（FinancialData）。賞与未参照検出に使用
        pdf_layout: `wage_pdf_layout_parser.parse_wage_ledger_layout` の戻り値。
            渡された場合、セル単位整合性チェック (C1〜C4) も実行する。
    """
    warnings = [
        check_employee_count_mismatch(hearing_data, ledger_employees),
        check_monthly_coverage(ledger_employees),
        check_value_distribution(ledger_employees),
        check_bonus_omission(ledger_employees, financial),
        check_similar_name_duplicates(ledger_employees),
        check_employment_type_missing(ledger_employees),
        check_extraction_size_vs_pl(ledger_employees, financial),
    ]
    if pdf_layout:
        warnings.extend(check_cell_level_consistency(ledger_employees, pdf_layout))
    return [w for w in warnings if w]
