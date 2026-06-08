# -*- coding: utf-8 -*-
"""賃金台帳PDFテキストのレイアウト構造を決定論的にパースする。

`pdf_text_extractor` 経由（pdfplumber / Document AI）で取得した
タブ区切りテキストを入力に、各従業員ごとに以下を抽出する。

- name: 氏名
- source_months: PDFに「○月分」「○月度」列として物理的に存在する月（1〜12）
- bonus_pays: 賞与の (支給月, 金額) リスト
- monthly_taxable_totals: 「総支給額(課税)」行から月別に取得した値（参照用）

これは AI 抽出（Sonnet）の結果を検証するための「正解側」として使われる。
検証ロジック本体は `wage_validator.py` 側に置く。

API 呼出ゼロ。決定論的に動く（同じ入力 → 同じ出力）。
"""
from __future__ import annotations

import logging
import re
import unicodedata
from dataclasses import dataclass, field
from typing import Iterable

logger = logging.getLogger(__name__)


# ── 表記揺れの吸収 ─────────────────────────────────────────────
# 「⽉」「月」両方を「月」に正規化。タブ・空白は単一空白に圧縮。
_MONTH_LABEL_RE = re.compile(r'(\d{1,2})\s*月\s*(分|度)')
_BONUS_LABEL_RE = re.compile(
    r'(賞与\s*\d+\s*回|夏季賞与|冬季賞与|期末賞与|決算賞与|特別賞与|ボーナス|賞\s*与\s*額)'
)
_JP_DATE_RE = re.compile(
    r'(?:令和|平成|昭和)\s*\d+\s*年\s*(\d{1,2})\s*月\s*\d{1,2}\s*日'
)
_WESTERN_DATE_RE = re.compile(r'(?:\d{4})[-/年]\s*(\d{1,2})[-/月]\s*\d{1,2}')

# 「総支給額(課税)」行を見つけるための正規表現。表記揺れに対応:
#   「総支給額(課税)」「総⽀給額(課税)」「総 支 給 額 ( 課 税 )」など
_TAXABLE_TOTAL_LABELS = ('総支給額(課税)', '課税支給合計', '課税支給額')

# 「賞与額」行
_BONUS_AMOUNT_LABELS = ('賞与額', '賞 与 額')

# 氏名行のパターン
#   PT1: 「氏 名\t...\tカナ」と「役 職\t...\t漢字氏名(No.000001)」が別行
#   PT2: 「氏名: 田中 太郎」のような一行型
_NAME_NOSUFFIX_RE = re.compile(r'氏\s*名[:：]\s*([^\s\t]+(?:\s+[^\s\t]+)?)')
_NAME_PAREN_NO_RE = re.compile(r'([^\s\t]+\s+[^\s\t]+)\s*\(\s*No\.?\s*\d+\s*\)')


def _normalize(s: str) -> str:
    """全角・半角・部首字を正規化、タブ/連続空白を単一空白に。"""
    if s is None:
        return ''
    n = unicodedata.normalize('NFKC', s)
    # 部首字「⽉」(U+2F49) → 「月」(U+6708) は NFKC で吸収される
    n = re.sub(r'[\t　]+', ' ', n)
    n = re.sub(r'  +', ' ', n)
    return n


@dataclass
class PdfEmployee:
    """PDFテキストから抽出した1従業員のレイアウト情報。"""
    name: str
    source_months: list[int] = field(default_factory=list)
    """PDFに『○月分/○月度』列として物理的に存在する月（1〜12）。"""

    bonus_pays: list[tuple[int, int | None]] = field(default_factory=list)
    """賞与の (支給月, 金額) リスト。金額が読めなかった場合は (月, None)。"""

    monthly_taxable_totals: dict[int, int] = field(default_factory=dict)
    """『総支給額(課税)』行から取得した月別値（参照用、月→円）。"""

    monthly_basic_pay: dict[int, int] = field(default_factory=dict)
    """『基本給』行から取得した月別値（賞与額の推定が困難な時の参照用）。"""

    source_pages: list[int] = field(default_factory=list)
    """この従業員レコードを構成した PDFページ番号（賞与が別ページにある等のため複数）。"""

    has_bonus_section: bool = False
    """賞与セクション（賞与1回/賞与2回/夏季賞与等）が PDF 上に存在するか。"""


# ── ページ分割 ────────────────────────────────────────────────
_PAGE_MARKER_RE = re.compile(r'^=====\s*page\s+(\d+)\s*=====')


def _split_pages(pdf_text: str) -> list[tuple[int, str]]:
    """`===== page N =====` マーカーでテキストをページ分割。

    マーカーが無い（DocAI 経路）場合は1ページ扱いで全文を返す。
    """
    pages: list[tuple[int, str]] = []
    current_page: int | None = None
    current_lines: list[str] = []
    for line in pdf_text.splitlines():
        m = _PAGE_MARKER_RE.match(line)
        if m:
            if current_page is not None:
                pages.append((current_page, '\n'.join(current_lines)))
            current_page = int(m.group(1))
            current_lines = []
        else:
            current_lines.append(line)
    if current_page is not None:
        pages.append((current_page, '\n'.join(current_lines)))
    elif current_lines:
        pages.append((1, '\n'.join(current_lines)))
    return pages


# ── ページごとのパース ────────────────────────────────────────
def _extract_name_from_page(page_text: str) -> str | None:
    """ページから氏名を抽出する。

    レイアウト1（実案件B系）: 「役 職\t社員\t\t山田 太郎(No.000002)」
    レイアウト2（給与計算ソフト系）: 「氏名: 田中 太郎」
    """
    for raw_line in page_text.splitlines():
        line = _normalize(raw_line)
        # PT2: 一行型
        m = _NAME_NOSUFFIX_RE.search(line)
        if m:
            name = m.group(1).strip()
            if name and not _looks_like_address(name):
                return name
        # PT1: 役職行に (No.XXX) 付き
        m = _NAME_PAREN_NO_RE.search(line)
        if m:
            name = m.group(1).strip()
            if name and not _looks_like_address(name):
                return name
    return None


def _looks_like_address(s: str) -> bool:
    """住所の誤抽出ガード。"""
    return bool(re.search(r'[市県区都道町村]', s))


def _extract_month_columns(page_text: str) -> tuple[list[int], int | None]:
    """ページの月分ヘッダから「物理的に存在する月」のリストと、その行Indexを返す。

    Returns:
        (months, header_line_index)
        months: 重複なし・出現順を保ったリスト（例: [3,4,5,6,7,8,9,10,11,12,1]）
        header_line_index: ヘッダ行のIndex（None なら見つからず）
    """
    lines = page_text.splitlines()
    for i, raw in enumerate(lines):
        line = _normalize(raw)
        matches = _MONTH_LABEL_RE.findall(line)
        if len(matches) >= 3:  # 3列以上の「○月分」「○月度」があればヘッダ行と判断
            months: list[int] = []
            for m_str, _suffix in matches:
                m = int(m_str)
                if 1 <= m <= 12 and m not in months:
                    months.append(m)
            return months, i
    return [], None


def _extract_bonus_layout(
    page_text: str,
) -> tuple[bool, list[tuple[int, int | None]]]:
    """賞与セクションを検出して (has_section, [(支給月, 金額)]) を返す。

    検出ロジック:
      1. ヘッダ行に「賞与1回」「賞与2回」「夏季賞与」「冬季賞与」「期末賞与」等が
         少なくとも1個ある → 賞与セクション存在
      2. ヘッダ直下の行に支給日があれば → 支給月リストを取得
      3. 「賞与額」「総支給額(課税)」行から金額列を引き当て
    """
    lines = page_text.splitlines()
    header_idx = None
    header_normalized = ''
    for i, raw in enumerate(lines):
        line = _normalize(raw)
        # 「賞 与 額」だけの行は集計行（金額行）。ヘッダ行ではない。
        # ヘッダ行は「賞与1回」「夏季賞与」等の **区分名** を含む。
        if _BONUS_LABEL_RE.search(line):
            # 「賞 与 額」のみで他の区分が無い行は除外
            non_amount = re.sub(r'賞\s*与\s*額', '', line)
            if _BONUS_LABEL_RE.search(non_amount):
                header_idx = i
                header_normalized = line
                break
    if header_idx is None:
        return False, []

    # 支給月の取得: ヘッダ直下の数行で日付を探す
    payment_months: list[int] = []
    for j in range(header_idx + 1, min(header_idx + 5, len(lines))):
        target = _normalize(lines[j])
        if not target.strip():
            continue
        m_dates = _JP_DATE_RE.findall(target) + _WESTERN_DATE_RE.findall(target)
        if m_dates:
            for m in m_dates:
                payment_months.append(int(m))
            break

    # 金額の取得: 「賞与額」または「総支給額(課税)」行から数値列を取得
    bonus_amounts: list[int] = []
    for j in range(header_idx + 1, min(header_idx + 40, len(lines))):
        line = _normalize(lines[j])
        if any(lbl in line for lbl in _BONUS_AMOUNT_LABELS) or any(
            lbl in line for lbl in _TAXABLE_TOTAL_LABELS
        ):
            nums = _extract_numbers(line)
            if nums:
                # 末尾の「合計」「総合計」列は除外する。
                # ヒューリスティック: 末尾の値が前項目の和に近ければ合計列とみなして除外
                if len(nums) >= 2 and abs(sum(nums[:-1]) - nums[-1]) <= 2:
                    nums = nums[:-1]
                bonus_amounts = nums
                break

    pays: list[tuple[int, int | None]] = []
    for k, mon in enumerate(payment_months):
        amount = bonus_amounts[k] if k < len(bonus_amounts) else None
        pays.append((mon, amount))
    # 支給日が取れず金額だけある場合は捨てる（月配置の根拠が立たないため）
    return True, pays


def _extract_numbers(line: str) -> list[int]:
    """カンマ区切り数値を整数リストとして抽出。

    `1,089,000` 等のカンマ区切り、または `100000` 等のプレーン数値。
    """
    raw = _normalize(line)
    # ラベル部分（先頭の「総支給額(課税)」等）を除外するため、
    # 数値らしき部分のみを取り出す
    nums: list[int] = []
    for m in re.finditer(r'[\d,]+', raw):
        s = m.group(0).replace(',', '')
        if s.isdigit() and len(s) >= 4:  # 4桁以上 = 金額とみなす（日付・コードを除外）
            nums.append(int(s))
    return nums


def _label_in_line(label: str, line: str) -> bool:
    """ラベル比較。空白差「基 本 給」「基本給」を吸収する。"""
    norm_line = re.sub(r'\s+', '', line)
    norm_label = re.sub(r'\s+', '', label)
    return norm_label in norm_line


def _cell_to_int(cell: str) -> int | None:
    """タブ区切りセルを整数に。空セル・非数値は None（その月は支給なし＝空欄）。"""
    s = re.sub(r'[\s,]', '', _normalize(cell))
    if s.isdigit() and len(s) >= 4:  # 4桁以上=金額（日付・コードを除外）
        return int(s)
    return None


def _align_tabbed_row(
    raw_line: str, n_cols: int, label_candidates: Iterable[str],
) -> dict[int, int] | None:
    """タブ区切りの金額行を「列位置を保ったまま」月Indexに割り付ける。

    `_extract_numbers` は空セルを潰して詰めてしまい、中途月のある行（パート・
    途中入退社）で月ズレを起こす。table-aware TSV はタブで列位置を保持しているので、
    タブ分割し、ラベルセル（「課税支給合計」等）より後ろのセルを空セルも含めて
    位置対応させる（落とし穴①対策・§1.1.0）。先頭に空セルや見出しセルが付く
    レイアウト（例: 「\t総支給額(課税)\t…」）にも対応するため、ラベルセルの位置を探す。

    Returns:
        {0始まりの列Index: 金額} or 整列できなければ None（呼出し側でフォールバック）。
    """
    if '\t' not in raw_line:
        return None
    cells = raw_line.split('\t')
    # ラベルを含むセルの位置を探し、その後ろを値セルとする
    label_idx = None
    for i, c in enumerate(cells):
        if any(_label_in_line(lbl, c) for lbl in label_candidates):
            label_idx = i
            break
    if label_idx is None:
        return None
    vals = [_cell_to_int(c) for c in cells[label_idx + 1:]]
    # 末尾の合計列を判定して除外（本体の和に一致するとき）
    if len(vals) == n_cols + 1 and vals[-1] is not None:
        body = [v for v in vals[:-1] if v is not None]
        if body and abs(sum(body) - vals[-1]) <= 2:
            vals = vals[:-1]
    if len(vals) < n_cols:
        return None
    return {idx: v for idx, v in enumerate(vals[:n_cols]) if v is not None}


def _extract_monthly_amounts(
    page_text: str,
    month_columns: list[int],
    header_line_index: int,
    label_candidates: Iterable[str],
) -> dict[int, int]:
    """月分ヘッダの直下から、指定ラベル行の値を月別 dict にする。

    密な行（全月に値）は従来どおり数値抽出で位置対応（既存挙動を維持）。
    数値が列数に満たない＝空セルのある中途月行のみ、タブ整列で列位置を復元する
    （空セルを潰して詰める事故＝月ズレを防ぐ。§1.1.0）。
    """
    if not month_columns or header_line_index is None:
        return {}
    lines = page_text.splitlines()
    n_cols = len(month_columns)
    for j in range(header_line_index + 1, min(header_line_index + 80, len(lines))):
        raw = lines[j]
        line = _normalize(raw)
        if not any(_label_in_line(lbl, line) for lbl in label_candidates):
            continue
        nums = _extract_numbers(line)
        # 末尾が合計列なら除外
        if len(nums) == n_cols + 1 and abs(sum(nums[:-1]) - nums[-1]) <= 2:
            nums = nums[:-1]
        # 密な行（ちょうど列数）は位置対応で確定（既存レイアウトの挙動を維持）
        if len(nums) == n_cols:
            return {mon: nums[idx] for idx, mon in enumerate(month_columns)}
        # 空セルで数値が欠ける中途月行 → タブ整列で列位置を復元
        aligned = _align_tabbed_row(raw, n_cols, label_candidates)
        if aligned is not None:
            return {month_columns[idx]: v for idx, v in aligned.items()}
        # 最終手段: 数値が列数以上あれば先頭から割当（旧フォールバック）
        if len(nums) >= n_cols:
            return {mon: nums[idx] for idx, mon in enumerate(month_columns)}
        return {}
    return {}


def _parse_employee_page(
    page_number: int, page_text: str,
) -> PdfEmployee | None:
    """1ページ分のテキストから PdfEmployee を構築。

    氏名が取れない、月分ヘッダも賞与ヘッダもないページは無視。
    """
    name = _extract_name_from_page(page_text)
    if not name:
        return None

    months, header_idx = _extract_month_columns(page_text)
    has_bonus, bonus_pays = _extract_bonus_layout(page_text)

    emp = PdfEmployee(name=name, source_pages=[page_number])
    emp.source_months = months
    emp.has_bonus_section = has_bonus
    emp.bonus_pays = bonus_pays

    # 月別の課税支給合計・基本給を取得
    if months and header_idx is not None:
        emp.monthly_taxable_totals = _extract_monthly_amounts(
            page_text, months, header_idx, _TAXABLE_TOTAL_LABELS,
        )
        emp.monthly_basic_pay = _extract_monthly_amounts(
            page_text, months, header_idx, ('基本給',),
        )
    return emp


# ── 同一氏名のマージ（賞与ページ別など） ────────────────────────
def _normalize_name_key(name: str) -> str:
    """マージ用のキー。空白を全部詰める。"""
    return re.sub(r'\s+', '', _normalize(name))


def _merge_employees(emps: list[PdfEmployee]) -> list[PdfEmployee]:
    """同一氏名（空白除去で一致）のページを統合する。

    月給ページと賞与ページが分離しているレイアウト（実案件B系）に必要。
    """
    merged: dict[str, PdfEmployee] = {}
    order: list[str] = []
    for e in emps:
        key = _normalize_name_key(e.name)
        if not key:
            continue
        if key not in merged:
            merged[key] = e
            order.append(key)
        else:
            base = merged[key]
            # 月給情報: 既存に source_months があれば優先、なければ新規分を採用
            for m in e.source_months:
                if m not in base.source_months:
                    base.source_months.append(m)
            # 賞与: 統合
            if e.has_bonus_section:
                base.has_bonus_section = True
                base.bonus_pays.extend(e.bonus_pays)
            # 月別金額: 後勝ち（賞与ページの「総支給額(課税)」で月給を上書きしないよう
            # 元ページに値があれば残す）
            for m, v in e.monthly_taxable_totals.items():
                base.monthly_taxable_totals.setdefault(m, v)
            for m, v in e.monthly_basic_pay.items():
                base.monthly_basic_pay.setdefault(m, v)
            base.source_pages.extend(e.source_pages)
    return [merged[k] for k in order]


# ── 公開関数 ───────────────────────────────────────────────────
def parse_wage_ledger_layout(pdf_text: str) -> list[PdfEmployee]:
    """PDFテキストから従業員別のレイアウト情報を抽出する。

    Args:
        pdf_text: pdf_text_extractor 経由で取得したタブ区切りテキスト。
            `===== page N =====` マーカー有り想定だが、無くても1ページ扱いで処理する。

    Returns:
        PdfEmployee のリスト（PDFの出現順）。同一氏名のページは統合済み。
        氏名が取れないページは無視される。
    """
    if not pdf_text or not pdf_text.strip():
        return []

    pages = _split_pages(pdf_text)
    per_page: list[PdfEmployee] = []
    for page_no, page_text in pages:
        emp = _parse_employee_page(page_no, page_text)
        if emp:
            per_page.append(emp)

    return _merge_employees(per_page)


def parse_wage_ledger_layout_from_pdf(pdf_bytes: bytes) -> list[PdfEmployee]:
    """PDFバイトから直接ページ別テキストを取得してレイアウト解析する。

    本関数は `pdf_text_extractor.extract_pdf_as_text_with_source` には依存しない。
    後者は Document AI 経由のテキストを返す経路があり、その場合ページマーカー無し・
    タブ区切りなしのフラットなテキストになるため、本パーサーが氏名・月分列を
    抽出できない（テーブル構造に依存しているため）。

    代わりに `pdf_text_extractor.extract_pdf_text_table_aware` を呼び、
    pdfplumber.extract_tables() ベースでテーブル構造を保ったまま取得する。
    画像PDF（テキスト層なし）の場合は空リスト。その場合は既存のマクロチェック
    （PL 突合・人数妥当性）に検証を委ねる。

    Args:
        pdf_bytes: PDFのバイト列

    Returns:
        PdfEmployee のリスト。テキスト抽出に失敗した場合は空リスト。
    """
    if not pdf_bytes:
        return []

    from .pdf_text_extractor import extract_pdf_text_table_aware
    text = extract_pdf_text_table_aware(pdf_bytes)
    if not text or not text.strip():
        return []
    return parse_wage_ledger_layout(text)


def summarize_layout(employees: list[PdfEmployee]) -> str:
    """ログ・デバッグ用のサマリ文字列。"""
    lines = [f'PDF レイアウト解析: {len(employees)}名']
    for e in employees:
        lines.append(
            f'  - {e.name}: months={e.source_months}, '
            f'has_bonus={e.has_bonus_section}, '
            f'bonus_pays={e.bonus_pays}, '
            f'pages={e.source_pages}'
        )
    return '\n'.join(lines)
