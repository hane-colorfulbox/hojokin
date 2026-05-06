# -*- coding: utf-8 -*-
"""
Claude APIによるPDFデータ抽出
- StubExtractor: APIキーなしで動作するスタブ（プレースホルダ値を返す）
- ClaudeExtractor: 実際のAPI呼出し（APIキー必要）
"""
from __future__ import annotations

import json
import logging
import time
from abc import ABC, abstractmethod
from typing import Callable, Optional

from .models import (
    CompanyInfo, FinancialData, TaxCertificate,
    Employee, MonthlyWageData, EstimateData, AIJudgment,
)

logger = logging.getLogger(__name__)


# ── リトライ設定 ──
# 初回+1回リトライ = 最大2回試行、バックオフは 2s
# timeout エラーは即 fail（再試行しても 300 秒 × N 回浪費するだけのため）
MAX_API_ATTEMPTS = 2
API_BACKOFF_SECONDS = [2]

# リトライ対象のHTTPステータス（APIStatusError系の一時的失敗）
# 422: "context reduction is suggested" 等のflakyエラー
# 429: rate_limit
# 500/502/503/504: サーバ側の一時障害
# 529: overloaded_error
RETRYABLE_STATUS_CODES = {422, 429, 500, 502, 503, 504, 529}

# 残高切れ判定用の文字列（400 invalid_request_error の message に含まれる）
CREDIT_BALANCE_MARKER = 'credit balance is too low'

# 賃金台帳PDFの事前分割閾値（ページ数 / バイト数）
# どちらかを超えたら 1 PDF を半分に分割して送信する。
# 経験則:
#   - 15ページを超えると Sonnet が出力打ち切り (max_tokens) する確率が上がる
#   - 4MB を超えると API timeout (300〜480秒) に近づく
WAGE_LEDGER_SPLIT_PAGE_THRESHOLD = 15
WAGE_LEDGER_SPLIT_BYTES_THRESHOLD = 4 * 1024 * 1024

# 進捗コールバックの型: (attempt, max_attempts, wait_seconds, error_summary) -> None
RetryCallback = Callable[[int, int, float, str], None]


class APICreditExhaustedError(RuntimeError):
    """API残高切れ（400 credit_balance_too_low）を表す専用例外"""
    pass


# ============================================================
# 賃金台帳PDF分割・マージ ヘルパ
# ============================================================

def _pdf_should_be_split(pdf_bytes: bytes) -> bool:
    """PDF が事前分割の閾値を超えているか判定。

    - WAGE_LEDGER_SPLIT_BYTES_THRESHOLD バイト超 → True
    - WAGE_LEDGER_SPLIT_PAGE_THRESHOLD ページ超 → True
    PyMuPDF を使う。失敗時はサイズだけで判定。
    """
    if len(pdf_bytes) > WAGE_LEDGER_SPLIT_BYTES_THRESHOLD:
        return True
    try:
        import fitz  # type: ignore
        doc = fitz.open(stream=pdf_bytes, filetype='pdf')
        n_pages = len(doc)
        doc.close()
        return n_pages > WAGE_LEDGER_SPLIT_PAGE_THRESHOLD
    except Exception as e:
        logger.warning(f'PDF ページ数判定失敗: {e}')
        return False


def _split_pdfs_in_half(
    pdf_files: list[tuple[str, bytes]],
) -> tuple[list[tuple[str, bytes]], list[tuple[str, bytes]]]:
    """PDFリストを前半・後半に分割。

    - 各 PDF を PyMuPDF で前半・後半に分割
    - 2 ページ以下の PDF は分割せず part1 にだけ含める（情報量保持）
    - 分割失敗した PDF は元のまま part1 に入れる
    Returns: (part1_files, part2_files)
    """
    import fitz  # type: ignore

    part1: list[tuple[str, bytes]] = []
    part2: list[tuple[str, bytes]] = []
    for fname, pdf_bytes in pdf_files:
        try:
            doc = fitz.open(stream=pdf_bytes, filetype='pdf')
            n_pages = len(doc)
            if n_pages <= 2:
                part1.append((fname, pdf_bytes))
                doc.close()
                continue
            half = n_pages // 2
            doc1 = fitz.open()
            doc1.insert_pdf(doc, from_page=0, to_page=half - 1)
            buf1 = doc1.tobytes()
            doc1.close()
            doc2 = fitz.open()
            doc2.insert_pdf(doc, from_page=half, to_page=n_pages - 1)
            buf2 = doc2.tobytes()
            doc2.close()
            doc.close()
            part1.append((f'{fname}#part1of2', buf1))
            part2.append((f'{fname}#part2of2', buf2))
            logger.info(
                f'PDF分割: {fname} ({n_pages}P, {len(pdf_bytes)/1_000_000:.2f}MB) → '
                f'{half}P + {n_pages-half}P'
            )
        except Exception as e:
            logger.warning(f'PDF分割失敗 ({fname}): {e} → 元のまま part1 に入れる')
            part1.append((fname, pdf_bytes))
    return part1, part2


def _is_pl_extraction_suspicious(d: dict) -> tuple[bool, str]:
    """PL抽出結果が「製造原価/完成工事原価の見落とし疑い」かを複合条件で検知。

    Codex指摘: 単一閾値は会社規模差で誤判定する。
    複数条件の OR で「明らかに小さすぎる」ケースを拾う。

    Returns: (suspicious, reason)
    """
    if not isinstance(d, dict):
        return False, ''
    salary = d.get('salary') or 0
    misc = d.get('misc_wages') or 0
    bonus = d.get('bonus') or 0
    revenue = d.get('revenue') or 0
    operating = d.get('operating_profit') or 0
    gross = d.get('gross_profit') or 0
    cost_of_sales = d.get('cost_of_sales') or 0

    total_personnel = salary + misc + bonus

    # 条件1: 売上があるのに人件費合計が0
    if total_personnel == 0 and revenue > 0:
        return True, '人件費合計0で売上あり'

    # 条件2: 給料手当が極端に小さく売上が大きい中小企業（建設業/製造業の典型）
    if salary < 1_000_000 and revenue > 10_000_000:
        return True, f'給料手当{salary:,}で売上{revenue:,}（製造原価/工事原価の見落とし疑い）'

    # 条件3: 給料/売上比率が異常に低い + 営業利益プラス
    if revenue > 0 and (salary / revenue) < 0.01 and operating > 0:
        ratio = salary / revenue * 100
        return True, f'給料/売上比率{ratio:.2f}%で営業利益プラス'

    # 条件4: 売上原価が大きいのに人件費合計が小さい（労務費が原価に流れている疑い）
    if cost_of_sales > 5_000_000 and total_personnel > 0 and total_personnel < cost_of_sales * 0.05:
        return True, (
            f'売上原価{cost_of_sales:,}に対して人件費{total_personnel:,}が極端に小さい'
            f'（原価部に労務費がある可能性）'
        )

    # 条件5: 売上総利益が大きいのに人件費が3M未満で営業利益プラス
    if total_personnel < 3_000_000 and operating > 0 and gross > 10_000_000:
        return True, f'売上総利益{gross:,}・営業利益プラスなのに人件費{total_personnel:,}'

    # 条件6: 売上原価/人件費 比 > 10倍 = 建設業/製造業で原価部に労務費が大量にある典型
    # （後藤造園の実本番ケース: 売上原価160M / 人件費12.7M = 12.6倍 で検知される）
    # 医療法人など製造原価のない業種は売上原価そのものが小さいので発動しない
    if cost_of_sales > 0 and total_personnel > 0:
        ratio = cost_of_sales / total_personnel
        if ratio > 10:
            return True, (
                f'売上原価/人件費 = {ratio:.1f}倍 '
                f'（売上原価{cost_of_sales:,} / 人件費{total_personnel:,}、'
                f'建設業/製造業の原価部労務費未合算疑い）'
            )

    return False, ''


def _merge_pl_with_cost_report(pl: dict, cost: dict) -> dict:
    """販管費抽出結果と原価フォーカス抽出結果をマージする（人件費系のみ加算）"""
    if not isinstance(pl, dict):
        pl = {}
    if not isinstance(cost, dict):
        return pl
    merged = dict(pl)
    for key in ('salary', 'misc_wages', 'bonus', 'legal_welfare', 'welfare', 'depreciation'):
        pl_val = pl.get(key) or 0
        cost_val = cost.get(key) or 0
        if cost_val > 0:
            merged[key] = pl_val + cost_val
    return merged


def _merge_wage_employees_by_month(
    chunks: list[list[dict]],
) -> list[dict]:
    """複数チャンクの従業員データを NFKC 正規化キーで統合し、**月単位で補完マージ**する。

    Codex 指摘 (2026-05): 同名見つけたら捨てる方式だと、同一人物の月データが
    別チャンクに分かれた場合に欠落する（例: 山田太郎が前半PDFに 1-6月、後半PDFに 7-12月のデータ）。

    同月衝突ポリシー:
        - 片方null → もう片方を採用（補完）
        - 両方値あり・差小 → そのまま
        - 両方値あり・差大 → max を採用（部分入力・欠損で値が小さくなるミスを救う）
    """
    import re as _re
    import unicodedata as _ud

    def _key(name: str) -> str:
        return _re.sub(r'[\s　]+', '', _ud.normalize('NFKC', str(name)))

    def _merge_arr(a: list, b: list, threshold: float = 1.0) -> list:
        """12要素配列を月単位補完マージ"""
        a = a or [None] * 12
        b = b or [None] * 12
        merged = []
        for i in range(12):
            va = a[i] if i < len(a) else None
            vb = b[i] if i < len(b) else None
            if va is None:
                merged.append(vb)
            elif vb is None:
                merged.append(va)
            else:
                try:
                    fa, fb = float(va), float(vb)
                    if abs(fa - fb) < threshold:
                        merged.append(va)
                    else:
                        merged.append(max(fa, fb))
                except (TypeError, ValueError):
                    merged.append(va)
        return merged

    merged: dict[str, dict] = {}
    order: list[str] = []
    for chunk in chunks:
        for emp in chunk or []:
            name = emp.get('name', '')
            if not name:
                continue
            key = _key(name)
            if not key:
                continue
            if key not in merged:
                # 新規 — 配列フィールドはコピーで持つ（in-place変更を防ぐ）
                merged[key] = {
                    'name': emp.get('name', ''),
                    'employment_type': emp.get('employment_type', ''),
                    'monthly_wages': list(emp.get('monthly_wages') or [None] * 12),
                    'monthly_hours': list(emp.get('monthly_hours') or [None] * 12),
                    'monthly_work_days': list(emp.get('monthly_work_days') or [None] * 12),
                }
                order.append(key)
                continue
            ex = merged[key]
            ex['monthly_wages'] = _merge_arr(
                ex.get('monthly_wages'), emp.get('monthly_wages'), threshold=1.0,
            )
            ex['monthly_hours'] = _merge_arr(
                ex.get('monthly_hours'), emp.get('monthly_hours'), threshold=0.1,
            )
            ex['monthly_work_days'] = _merge_arr(
                ex.get('monthly_work_days'), emp.get('monthly_work_days'), threshold=0.5,
            )
            # employment_type は空でない方を優先
            if not ex.get('employment_type') and emp.get('employment_type'):
                ex['employment_type'] = emp['employment_type']
    return [merged[k] for k in order]


# ── プロンプトテンプレート ──

PROMPT_REGISTRY = """**出力は ```json コードブロック1個のみ。前置き禁止。単一 dict で返すこと。**

この履歴事項全部証明書の画像から、以下の情報をJSON形式で抽出してください。
読み取れない項目はnullにしてください。

重要ルール:
- 履歴事項には役員の就任・退任・重任の履歴が記録されています。同一人物が複数回登場する場合は、最新の役職のみを採用してください。
- 下線が引かれた（抹消された）情報は過去のものなので無視してください。
- 代表者はofficersには含めないでください（representative_name/representative_titleに記載）。
- 退任済みの役員は含めないでください。

```json
{
  "name": "法人名（株式会社等含む）",
  "name_kana": "法人名フリガナ（カタカナ）",
  "address": "本店所在地",
  "postal_code": "郵便番号（わかれば）",
  "established_date": "設立年月日 yyyy-mm-dd形式",
  "capital": 資本金（円、整数）,
  "representative_name": "代表者氏名",
  "representative_title": "代表者役職",
  "officers": [
    {"title": "役職", "name": "氏名", "kana": "フリガナ（推定でOK）"}
  ],
  "business_purposes": ["目的1", "目的2"]
}
```"""

PROMPT_PL = """**出力は ```json コードブロック1個のみ。前置き・解説・複数年度の配列禁止。単一の dict で返すこと。**

この決算書類（損益計算書・販管費内訳書・製造原価報告書・完成工事原価報告書・工事原価報告書・売上原価内訳書 など全体）から、以下をJSON形式で抽出してください。
該当項目がない場合はnullにしてください。金額は円単位の整数で。
複数年度の決算書がある場合は **直近期1期分のみ** を dict で返してください（配列にしないこと）。

【必須・契約: 原価部の人件費は販管費と必ず合算する】
PDF/画像の中に以下のいずれかが含まれる場合、**販管費部分の同種項目と必ず合算した値** を返してください:
  - 製造原価報告書 / 製造原価明細書
  - 完成工事原価報告書 / 工事原価報告書（建設業）
  - 売上原価内訳書 / 売上原価報告書
  - 役務原価報告書（サービス業）

合算ルール（**販管費の値 + 原価部の値**）:
  - salary  ← 販管費「給料手当」 + 原価部「賃金」「給料」「労務費」
  - misc_wages ← 販管費「雑給」 + 原価部「雑給」
  - bonus ← 販管費「賞与」 + 原価部「賞与」
  - legal_welfare ← 販管費「法定福利費」 + 原価部「法定福利費」
  - welfare ← 販管費「福利厚生費」 + 原価部「福利厚生費」
  - depreciation ← 販管費「減価償却費」 + 原価部「減価償却費」

例（建設業、製造原価報告書がある決算書）:
  販管費 給料手当 200,000 + 完成工事原価 賃金 13,900,000 → salary: 14,100,000
  販管費 法定福利費 100,000 + 工事原価 法定福利費 2,500,000 → legal_welfare: 2,600,000

**販管費の値だけで返すのは抽出ミスです。原価部があれば必ず加算してください。**

個人事業主の「所得税の青色申告決算書」または「収支内訳書」の場合:
- revenue = 売上（収入）金額
- gross_profit = 売上（収入）金額 - 売上原価
- operating_profit / ordinary_profit = 所得金額（青色申告特別控除前）
- salary = 給料賃金
- 役員報酬・賞与・経常利益といった法人特有の項目は null
- 専従者給与がある場合は misc_wages に計上

```json
{
  "fiscal_year_start": "事業年度開始日 yyyy-mm-dd",
  "fiscal_year_end": "事業年度終了日 yyyy-mm-dd",
  "revenue": 売上高,
  "cost_of_sales": 売上原価,
  "gross_profit": 売上総利益,
  "operating_profit": 営業利益（損失ならマイナス）,
  "ordinary_profit": 経常利益（損失ならマイナス）,
  "net_profit": 当期純利益（損失ならマイナス）,
  "salary": 給料手当,
  "misc_wages": 雑給,
  "bonus": 賞与,
  "officer_compensation": 役員報酬,
  "legal_welfare": 法定福利費,
  "welfare": 福利厚生費,
  "depreciation": 減価償却費,
  "travel_expense": 旅費交通費
}
```"""

PROMPT_PL_COST_REPORT_FOCUS = """**出力は ```json コードブロック1個のみ。前置き禁止。単一 dict で返すこと。**

この決算書類のうち、**製造原価報告書 / 完成工事原価報告書 / 工事原価報告書 / 売上原価内訳書 / 役務原価報告書** に
記載されている人件費・減価償却費 **のみ** を抽出してください。販管費部分は無視してください。

各項目は **原価部の値のみ** を返してください（販管費との合算は呼出側で行います）:
  - salary: 原価部の「給料手当」「賃金」「労務費」「給料」の合計
  - misc_wages: 原価部の「雑給」
  - bonus: 原価部の「賞与」
  - legal_welfare: 原価部の「法定福利費」
  - welfare: 原価部の「福利厚生費」
  - depreciation: 原価部の「減価償却費」

該当項目がなければ 0 を返してください（null ではなく 0）。
原価報告書が含まれていない（販管費のみの決算書）場合は、すべて 0 を返してください。

```json
{
  "salary": 0,
  "misc_wages": 0,
  "bonus": 0,
  "legal_welfare": 0,
  "welfare": 0,
  "depreciation": 0
}
```"""


# ============================================================
# Phase 1: 構造分解 PL 抽出（3呼出方式）
# ============================================================
# 1回の視覚抽出に「ページ選別・年度選択・販管費/原価合算・検算」を背負わせると
# Sonnet が揺らいで抽出ミスする問題への抜本対策。
# タスクを「基本PL」「販管費部の人件費」「原価部の人件費」の3つに分解し、
# 各プロンプトを1枚絵レベルに単純化することでブレを抑える。
# 結果はコード側で機械的に合算（二重計上リスクなし）。

PROMPT_PL_PAGE_INVENTORY = """**出力は ```json コードブロック1個のみ。前置き禁止。単一 dict。**

これから提示する決算書類画像（複数ページ）について、各ページを分類してインベントリを作成してください。
判別ポイントは「**何が記載されているか**」と「**どの事業年度の書類か**」です。

各ページは以下のいずれか、または複数のラベルに該当します:
  - "pl_basic": 損益計算書本表（売上高・売上原価・営業利益等のサマリ）
  - "pl_section": 販管費及び一般管理費の内訳明細
  - "cost_section": 製造原価報告書 / 完成工事原価報告書 / 工事原価報告書 / 売上原価内訳書
  - "balance_sheet": 貸借対照表
  - "cash_flow": キャッシュフロー計算書
  - "stockholders_equity": 株主資本等変動計算書
  - "notes": 個別注記表 / 注記事項
  - "other": その他（表紙・目次・別添資料など）

事業年度は「年度ラベル」として記載してください（例: "R6"="2024年度", "R7"="2025年度"）。
直近期（最新の事業年度）は is_latest=true としてください。

```json
{
  "pages": [
    {"page": 1, "labels": ["pl_basic"], "fiscal_year_label": "R6", "is_latest": true},
    {"page": 2, "labels": ["pl_section"], "fiscal_year_label": "R6", "is_latest": true},
    {"page": 3, "labels": ["cost_section"], "fiscal_year_label": "R6", "is_latest": true},
    {"page": 4, "labels": ["balance_sheet"], "fiscal_year_label": "R6", "is_latest": true},
    {"page": 5, "labels": ["pl_basic"], "fiscal_year_label": "R5", "is_latest": false}
  ],
  "latest_fiscal_year_label": "R6",
  "latest_fiscal_year_period": "2024-04 to 2025-03"
}
```

ページ番号は 1 から始まる連番。labels は配列（複数該当可）。
fiscal_year_label が判別できないページは null。
すべてのページを必ず含めてください。"""


PROMPT_PL_BASIC = """**出力は ```json コードブロック1個のみ。前置き禁止。単一 dict。配列禁止。**

この決算書類の **損益計算書本表（一番表紙のPL）** から、財務サマリ項目だけを抽出してください。
販管費明細・製造原価報告書・貸借対照表は**無視**してください。

複数年度がある場合は **直近期1期分のみ** を返してください。
個人事業主の場合: revenue=売上(収入)金額、operating_profit=所得金額、salary系はnull。

```json
{
  "fiscal_year_start": "yyyy-mm-dd",
  "fiscal_year_end": "yyyy-mm-dd",
  "revenue": 売上高,
  "cost_of_sales": 売上原価,
  "gross_profit": 売上総利益,
  "operating_profit": 営業利益（損失ならマイナス）,
  "ordinary_profit": 経常利益（損失ならマイナス）,
  "net_profit": 当期純利益（損失ならマイナス）
}
```

該当項目がない/読み取り不能なら null。"""


PROMPT_PL_PL_SECTION = """**出力は ```json コードブロック1個のみ。前置き禁止。単一 dict。配列禁止。**

この決算書類の **販売費及び一般管理費の内訳** から、人件費系・減価償却費の項目を抽出してください。
**製造原価報告書 / 完成工事原価報告書 / 工事原価報告書 は無視** してください（別途抽出するため）。

販管費「のみ」の値を返してください。複数年度なら直近期1期のみ。

```json
{
  "salary": 販管費の「給料手当」（円、整数）,
  "misc_wages": 販管費の「雑給」,
  "bonus": 販管費の「賞与」,
  "officer_compensation": 販管費の「役員報酬」,
  "legal_welfare": 販管費の「法定福利費」,
  "welfare": 販管費の「福利厚生費」,
  "depreciation": 販管費の「減価償却費」,
  "travel_expense": 販管費の「旅費交通費」
}
```

該当項目がなければ 0 を返してください（null ではなく 0、合算時の不確実性を避けるため）。
販管費明細自体が無い場合は、全て 0 を返してください。"""


PROMPT_PL_COST_SECTION = """**出力は ```json コードブロック1個のみ。前置き禁止。単一 dict。配列禁止。**

この決算書類のうち、**原価部 = 製造原価報告書 / 完成工事原価報告書 / 工事原価報告書 / 売上原価内訳書 / 役務原価報告書** に
記載されている人件費系・減価償却費の項目だけを抽出してください。
**販売費及び一般管理費は完全に無視** してください（別途抽出するため）。

原価部「のみ」の値を返してください。複数年度なら直近期1期のみ。

```json
{
  "salary": 原価部の「給料手当 + 賃金 + 労務費 + 給料」の合計,
  "misc_wages": 原価部の「雑給」,
  "bonus": 原価部の「賞与」,
  "legal_welfare": 原価部の「法定福利費」,
  "welfare": 原価部の「福利厚生費」,
  "depreciation": 原価部の「減価償却費」,
  "travel_expense": 原価部の「旅費交通費」
}
```

**該当項目がなければ 0 を返してください（null ではなく 0）。**
**原価報告書が含まれていない（販管費のみの決算書）場合は、すべて 0 を返してください。**

二重計上を避けるため、販管費部の値は絶対に含めないでください。"""


PROMPT_TAX = """**出力は ```json コードブロック1個のみ。前置き禁止。単一 dict で返すこと（配列禁止）。**

この納税証明書の画像から、以下をJSON形式で抽出してください。

```json
{
  "tax_type": "証明書の種類（その1、その2等）",
  "tax_amount": 納税額（円、整数）,
  "fiscal_year": "事業年度"
}
```"""

PROMPT_WAGES = """この給与支給控除一覧表の画像から、従業員ごとのデータをJSON配列で抽出してください。
全従業員を漏れなく抽出してください。

```json
[
  {
    "name": "氏名",
    "department": "所属（例: 総本店）",
    "employee_id": "社員番号",
    "employment_type": "正社員 または パート・アルバイト",
    "working_days": 出勤日数,
    "scheduled_hours": 所定労働時間,
    "base_salary": 基本給,
    "taxable_total": 課税支給合計,
    "total_pay": 支給合計,
    "deductions": 控除合計,
    "net_pay": 差引支給額
  }
]
```

判定ヒント:
- 社員番号100xxx台 → 正社員、200xxx台 → パート・アルバイト
- 所属欄に「正社員」「アルバイト」の記載があればそれを使う"""

PROMPT_ESTIMATE = """**出力は ```json コードブロック1個のみ。前置き禁止。単一 dict で返すこと。**

この見積書の画像から、以下をJSON形式で抽出してください。

```json
{
  "vendor_name": "発行元の会社名",
  "tool_name": "ツール/サービス名",
  "items": [
    {"name": "項目名", "amount": 金額}
  ],
  "total_amount": 合計金額（税抜）,
  "tax_amount": 消費税額
}
```"""

PROMPT_AI_JUDGMENT = """以下の会社情報に基づいて、補助金申請に必要な判断項目を埋めてください。

会社情報:
- 会社名: {company_name}
- 事業内容: {business_purposes}
- 所在地: {address}
- 営業利益: {operating_profit}円
- ツール名: {tool_name}

ヒアリングシート回答:
- 主な事業内容: {main_business}
- 強み: {strength}
- 課題（時間がかかっている業務）: {challenge}
- 月間所要時間: {monthly_hours}
- ツールで楽にしたいこと: {tool_usage}
- 削減見込み: {reduction}
- 浮いた時間の活用: {freed_time}
- 3年後の売上目標: {sales_target}
- IT投資実績: {it_investment_answer}
- IT投資金額: {it_investment_amount}
- IT投資プロセス: {it_investment_process}

※ヒアリングシートの回答が空欄の項目がある場合は、履歴事項の事業目的や決算書の情報から合理的に推定してください。

以下をJSON形式で回答してください。
重要: ヒアリングシートの回答を最優先で参照し、矛盾しないようにしてください。

```json
{{
  "industry_code": "日本標準産業分類の細分類コード（4桁）",
  "industry_text": "大分類 X xxx / 中分類 xx xxx / 小分類 xxx xxx / 細分類 xxxx xxx",
  "business_description": "事業内容の説明文。250-255文字。会社の現状・課題・ツール導入による解決策・期待効果を含む",
  "management_intent": "営業利益がプラスなら '事業の拡大に積極的'、マイナスなら '事業の維持に注力'",
  "future_goals": "営業利益がプラスなら '事業の拡大'、マイナスなら '利益の確保'",
  "security_status": "パソコンやサーバなどには、IDやパスワードを設け情報セキュリティ管理を行っている",
  "business_types": "履歴事項の目的から該当する日本標準産業分類の大分類をカンマ区切りで",
  "it_investment_status": "ヒアリングのIT投資実績が「はい」なら過去にIT投資を行ったことがある旨を記載。「いいえ」なら今までIT投資を行っていなかった",
  "it_utilization_status": "ヒアリングのIT投資実績に基づき適切に選択",
  "it_utilization_scope": "ITツールの導入により電子化する事務の範囲（例: '会計', '受発注', '決済' 等から該当するものをカンマ区切りで）",
  "invoice_related_work": "ITツールの導入によりインボイス対応に資する業務（例: '請求書の発行・受領', '仕入税額控除の計算' 等）",
  "weakness": "下記の選択肢から、会社の状況に最も合う番号と本文を 'N 本文' 形式で1つだけ返す（複数候補がある場合は最重要を1つ）",
  "it_investment_process": "下記の選択肢から、過去にIT投資を行ったプロセスを 'N 本文' 形式で1つだけ返す。過去IT投資なし/未回答なら空文字列",
  "improvement_process": "下記の選択肢から、補助金で最も改善したい業務プロセスを 'N 本文' 形式で1つだけ返す",
  "expected_effect_dept": "下記の選択肢から、強化したい部門を 'N 本文' 形式で1つだけ返す",
  "expected_effect": "下記の選択肢から、期待する効果を 'N 本文' 形式で1つだけ返す"
}}
```

【選択肢一覧（番号は厳密に対応させること）】
■ weakness（弱み）
 1 競合他社との差別化が図れていない
 2 人材不足
 3 商圏・立地
 4 製品サービスの質
 5 商品・サービスの情報発信不足
 6 顧客情報の不足
 7 在庫管理・工程管理等、業務管理がうまく把握できていない
 8 社員の高齢化や退職
 9 人が育たない
 10 設備の陳腐化
 11 運転資金不足
 12 設備投資資金不足

■ it_investment_process / improvement_process（IT投資プロセス・改善したいプロセス、共通選択肢）
 1 販売や店頭といったフロント業務の強化
 2 顧客のニーズや流行等を捉え、新規顧客獲得や新規市場開拓を行った
 3 事前の準備工程（施策、テスト、設計や計画立案、など）を強化
 4 生産管理・在庫管理・物流管理など、商品の動きの可視化・効率化
 5 案件管理・工程管理・進捗管理といった業務管理の可視化・効率化
 6 営業（現場）の業務効率化を図った
 7 会計業務や清算業務の正確性・効率化を図った
 8 人員配置の最適化を行った
 9 勤務時間の短縮・労働時間の適正化を図った
 10 単純な事務作業を自動化し、人手や時間の無駄を削った
 11 取引先や社内での情報共有を強化した

■ expected_effect_dept（強化したい部門）
 1 営業、店頭等顧客と接する部門・業務
 2 現場（実施政策に関わる業務/制作現場管理・工程管理・品質管理等）
 3 現場（準備に関わる業務/企画・開発・設計など）
 4 仕入れ・受発注・在庫管理・物流管理等
 5 総務・法務
 6 人事（勤怠管理・賃金管理等）
 7 会計・経理
 8 情報システム部門

■ expected_effect（期待効果）
 1 新規市場開拓・新規顧客獲得による売り上げの向上・拡大
 2 原価コストの圧縮
 3 勤務時間の短縮もしくは適正化
 4 会計の正確性
 5 ニーズに合った製品やサービスの提供
 6 製品やサービスの質の向上
 7 社内の情報が共有化されて、風通しの良い環境
 8 経営状況の正確な把握

判定ヒント: ヒアリングの「ツールで楽にしたいこと」「課題」「IT投資金額」を最優先参照。営業赤字なら weakness は経営/資金系、人材課題があれば 2/8/9 など。"""


# ============================================================
# Prompt Caching 用に PROMPT_WAGE_LEDGER を固定部 / 動的部 で分離
# - STATIC: 顧客間で共通の指示・出力形式・重要な注意。cache_control 対象。
# - TAIL: fiscal_period_section と tsv_data の挿入。顧客ごとに変動するため cache 対象外。
# 順序を STATIC → TAIL にすることで「期間フィルタ → 賃金台帳データ」が末尾に並び、
# 元の PROMPT_WAGE_LEDGER と意味的に等価（AIの判断に影響しない範囲の再配置）。
# 既存の PROMPT_WAGE_LEDGER は USE_PROMPT_CACHING=false 時のフォールバックとして維持。
# ============================================================
PROMPT_WAGE_LEDGER_STATIC = """**最優先指示: 出力は ```json コードブロック1個のみ。**
**前置き・解説・分析過程・後書き・思考の説明は厳禁。**
**応答の最初の文字は ``` で始め、最後の文字は ``` で終わること。**
（前置きで出力トークンを使い切ると応答が途中で打ち切られ、抽出失敗します）

以下は賃金台帳のExcelをTSV形式に変換したテキストです。
各従業員の月別給与データ・労働時間データを抽出し、JSON形式で返してください。

【抽出ルール】
1. 全シート・全テーブルを横断して、登場するすべての従業員を抽出してください。
2. monthly_wages: 月別の課税支給合計（または支給合計、税込支給額、給与+賞与の合算）。
   - **★ 配列の Index は必ず西暦/和暦カレンダーの月で固定**: Index 0=1月, Index 1=2月, ..., Index 11=12月
   - **★ 台帳の列の物理的順序とは無関係**: 例えば「R6.12月, R7.1月, R7.2月, ..., R7.11月」の順で並んでいても、
     R7.1月→Index 0, R7.2月→Index 1, ..., R7.11月→Index 10, **R6.12月→Index 11** に格納する
   - **★ 「13ヶ月レイアウト」（前年12月 + 当年1〜11月）**: 賃金台帳によくあるパターン。
     先頭が「12月」と書かれていて、それが前年度のものなら Index 11 に格納（前年度はそれ自体を採用、後述の期間フィルタで除外可能）
   - **★ 24ヶ月以上ある台帳**: 同じ月（例: R6.5月とR7.5月）は重複しないため、対象期間（後述の【期間フィルタ】参照）の月だけ抽出する
   - **★ 「年合計」「合計」「累計」列は決して monthly_wages の月別セルに混入させない**。年合計が1月や12月セルに紛れ込むと、後段の検証で異常検知される。年合計列は無視する
   - データがない月は null
3. monthly_hours: 月別の **総労働時間 / 実労働時間 / 所定労働時間**。Index は monthly_wages と同じ規則（Index 0=1月固定）。
   - **重要**: 「残業時間」「所定時間外」「時間外労働」は労働時間ではありません。混同しないでください。
   - 賃金台帳に労働時間の欄が存在しない（労働日数のみ等）場合は **必ず null** にしてください。推測値は禁止。
4. monthly_work_days: 月別の **労働日数 / 出勤日数 / 勤務日数**。Index は monthly_wages と同じ規則。
   - 「有給休暇日数」は労働日数ではないので含めないこと。
5. employment_type: 雇用形態（正社員・パート・アルバイト・役員等）。元の表記をそのまま入れてください。役員は「役員」を含む表記に。
   - 雇用形態が台帳に明示されていない場合は「正社員」を既定値とする（null/空文字列にせず文字列を必ず埋める）
6. **【名前抽出の厳密ルール】**
   - 賃金台帳の各行は通常、以下の列構成：フリガナ / 氏名 / 性別 / 生年月日 / 住所 / 給与・労働時間...
   - 氏名は『氏名』『社員名』『名前』など明示的なラベルが付いた欄 **からのみ** 抽出。決して隣接する「住所」「市町村」欄から文字を引っ張らない
   - 抽出後、名前に「市」「県」「区」「都」「道」「町」「村」など行政地名が含まれていないか確認。含まれていたら、その行の住所欄を参照して誤抽出部分を削除する
   - 住所欄が参照できず確実に分離できない場合は、その従業員レコード全体を除外する（name=null）
   - **★ 「氏名」「フリガナ」「氏」「名」「社員名」だけが入った空行はヘッダー残骸であり従業員ではない。これらは抽出対象から外す**
   - **★ OCR 由来の異体字・誤読**: 「壽⇔靖」「栁⇔櫛」「嶋⇔崎⇔島」「月⇔目」「日⇔Π」など類似字形の混乱がよくある。
     同一台帳内で類似氏名（最後の1文字だけ違う等）が複数出現したら、別人ではなく **同一人物の OCR 誤読の可能性が高い** ため統合を検討
7. 給与と賞与が別行・別シートに分かれている場合は、同月分を**合算**してください。
   - **★ 重要: 賞与一覧シート / 賞与明細シート / 期末手当シート / ボーナス一覧 等の名称で別タブに賞与情報がある場合、必ず対応する従業員の monthly_wages にその月の賞与額を加算してください。**
   - 賞与は通常「夏季賞与（6月・7月）」「冬季賞与（12月）」「期末賞与（決算期月）」として支給される。該当月の monthly_wages = 月給 + 賞与 で合算する
   - 「賞与」「ボーナス」「期末手当」「特別手当」を含むシート名・列名・行名は **絶対に見落とさない**こと
   - **★ 月別カラムに「6月（賞与）」「12月（賞与）」のような独立列がある場合**: 該当月の通常給与とは別物として、同月の monthly_wages に加算する
   - **★ 月度コードが「YYYYMM」形式（例: 202521, 202522）で月部分が 13 以上の場合**: これは「賞与1（21）」「賞与2（22）」を意味する給与ソフト独自表記。該当年の夏季賞与（202521）と冬季賞与（202522）として、それぞれ 7月（Index 6）と 12月（Index 11）の monthly_wages に加算する
   - 賞与シートが見つからない場合は、PL の「賞与手当」科目との突合で漏れを検出するため、可能な限り全シートを探索する
8. 月の判定は以下のいずれかを使用:
   - 列ヘッダ「1月」「5月」等のプレーン表記
   - 「R6.5月」「R7.4月」「令和6年5月」等の和暦付き表記（年は無視して月だけ使用）
   - 「2024年5月」「2024/05」「202405」等の西暦表記
   - 「対象年月」「給与年月」列の値
   - **OCR 誤読対策**: 「月」が「目」「Π」に化けていても文脈で月と判定する（数字 + 月/目/Π = 月度ヘッダ）
9. **【数値の正規化】**
   - **千区切りと小数点の混乱**: OCR が「90,000」を「90.000」と誤読することがある。
     金額は通常 **整数（円単位）**で記録される。`90.000` のような小数点が出てきても、**金額として整数に変換**（90000 円）して扱う
   - 「3桁ごとに区切られた数値」は千区切りとして整数化（カンマ・ピリオド区別なし）
   - 「90,000」「90.000」「90000」「90,000円」「￥90,000」はすべて 90000 と解釈
10. **【ページ境界・チャンク分割の処理】**
    - OCR テキスト中に「`----- pdf chunk boundary -----`」「`===== page N =====`」「`--- table N ---`」のようなマーカーがある場合、これは PDF 分割やページ区切りのため。
      **同じ従業員のデータがマーカーで分断されている可能性がある**。マーカーの前後で同名の従業員が出てきたら同一人物として統合する
    - マーカー自体は従業員データではないため、抽出対象から除外する

【出力形式（厳密に従ってください）】
```json
[
  {
    "name": "従業員名",
    "employment_type": "正社員",
    "monthly_wages": [430000, 316000, null, null, null, null, null, null, null, null, null, null],
    "monthly_hours": [160, 160, null, null, null, null, null, null, null, null, null, null],
    "monthly_work_days": [20, 21, null, null, null, null, null, null, null, null, null, null]
  }
]
```
↑ 上の例では Index 0=1月(430000円), Index 1=2月(316000円), 残り全 null。
   仮に「12月のみ給与あり」なら `[null,null,null,null,null,null,null,null,null,null,null,500000]` となる。

【重要な注意】
- monthly_wages / monthly_hours / monthly_work_days は **必ず12要素** の配列にしてください。データがない月は null。
- 金額は **円単位の整数**。コンマや「円」記号は付けないでください。
- 労働時間が記載されていない賃金台帳では monthly_hours は全て null で構いません（後段で労働日数×8hで補完します）。
- 役員報酬は役員として抽出してください（employment_type に「役員」を含める）。
- 名前のフリガナや空欄行は無視してください。
- **登場するすべての従業員を漏れなく出力してください。N名いれば N要素のJSON配列にする。最初の数名で打ち切ったり、代表者だけ抽出することは禁止です。**
- JSON以外のコメント・説明文は一切含めないでください（先頭・末尾とも純粋なJSON配列のみ）。
"""

PROMPT_WAGE_LEDGER_TAIL = """{fiscal_period_section}

【賃金台帳データ】
{tsv_data}
"""


PROMPT_WAGE_LEDGER = """**最優先指示: 出力は ```json コードブロック1個のみ。**
**前置き・解説・分析過程・後書き・思考の説明は厳禁。**
**応答の最初の文字は ``` で始め、最後の文字は ``` で終わること。**
（前置きで出力トークンを使い切ると応答が途中で打ち切られ、抽出失敗します）

以下は賃金台帳のExcelをTSV形式に変換したテキストです。
各従業員の月別給与データ・労働時間データを抽出し、JSON形式で返してください。

【抽出ルール】
1. 全シート・全テーブルを横断して、登場するすべての従業員を抽出してください。
2. monthly_wages: 月別の課税支給合計（または支給合計、税込支給額、給与+賞与の合算）。
   - **★ 配列の Index は必ず西暦/和暦カレンダーの月で固定**: Index 0=1月, Index 1=2月, ..., Index 11=12月
   - **★ 台帳の列の物理的順序とは無関係**: 例えば「R6.12月, R7.1月, R7.2月, ..., R7.11月」の順で並んでいても、
     R7.1月→Index 0, R7.2月→Index 1, ..., R7.11月→Index 10, **R6.12月→Index 11** に格納する
   - **★ 「13ヶ月レイアウト」（前年12月 + 当年1〜11月）**: 賃金台帳によくあるパターン。先頭の12月を Index 11 に格納
   - **★ 24ヶ月以上ある台帳**: 同じ月（例: R6.5月とR7.5月）は重複しないため、対象期間（{fiscal_period_section}参照）の月だけ抽出する
   - **★ 「年合計」「合計」「累計」列は決して monthly_wages の月別セルに混入させない**。年合計列は無視する
   - データがない月は null
3. monthly_hours: 月別の **総労働時間 / 実労働時間 / 所定労働時間**。Index は monthly_wages と同じ規則（Index 0=1月固定）。
   - **重要**: 「残業時間」「所定時間外」「時間外労働」は労働時間ではありません。混同しないでください。
   - 賃金台帳に労働時間の欄が存在しない（労働日数のみ等）場合は **必ず null** にしてください。推測値は禁止。
4. monthly_work_days: 月別の **労働日数 / 出勤日数 / 勤務日数**。Index は monthly_wages と同じ規則。
   - 「有給休暇日数」は労働日数ではないので含めないこと。
5. employment_type: 雇用形態（正社員・パート・アルバイト・役員等）。元の表記をそのまま入れてください。役員は「役員」を含む表記に。
   - 雇用形態が台帳に明示されていない場合は「正社員」を既定値とする
6. **【名前抽出の厳密ルール】**
   - 賃金台帳の各行は通常、以下の列構成：フリガナ / 氏名 / 性別 / 生年月日 / 住所 / 給与・労働時間...
   - 氏名は『氏名』『社員名』『名前』など明示的なラベルが付いた欄 **からのみ** 抽出。決して隣接する「住所」「市町村」欄から文字を引っ張らない
   - 抽出後、名前に「市」「県」「区」「都」「道」「町」「村」など行政地名が含まれていないか確認。含まれていたら、その行の住所欄を参照して誤抽出部分を削除する
   - 住所欄が参照できず確実に分離できない場合は、その従業員レコード全体を除外する（name=null）
   - **★ 「氏名」「フリガナ」「氏」「名」「社員名」だけが入った空行はヘッダー残骸であり従業員ではない**
   - **★ OCR 異体字・誤読**: 「壽⇔靖」「栁⇔櫛」「嶋⇔崎⇔島」「月⇔目」「日⇔Π」の類似字形混乱に注意。
     類似氏名（最後の1文字だけ違う等）が複数出現したら、同一人物の OCR 誤読として統合を検討
7. 給与と賞与が別行・別シートに分かれている場合は、同月分を**合算**してください。
   - **★ 重要: 賞与一覧シート / 賞与明細シート / 期末手当シート / ボーナス一覧 等の名称で別タブに賞与情報がある場合、必ず対応する従業員の monthly_wages にその月の賞与額を加算してください。**
   - 賞与は通常「夏季賞与（6月・7月）」「冬季賞与（12月）」「期末賞与（決算期月）」として支給される。該当月の monthly_wages = 月給 + 賞与 で合算する
   - 「賞与」「ボーナス」「期末手当」「特別手当」を含むシート名・列名・行名は **絶対に見落とさない**こと
   - **★ 月別カラムに「6月（賞与）」「12月（賞与）」のような独立列がある場合**: 通常給与とは別物として同月の monthly_wages に加算
   - **★ 月度コードが「YYYYMM」形式で月部分が 13 以上**（例: 202521, 202522）: 賞与の給与ソフト独自表記。202521→7月、202522→12月にそれぞれ加算
   - 賞与シートが見つからない場合は、PL の「賞与手当」科目との突合で漏れを検出するため、可能な限り全シートを探索する
8. 月の判定は以下のいずれかを使用:
   - 列ヘッダ「1月」「5月」等のプレーン表記
   - 「R6.5月」「R7.4月」「令和6年5月」等の和暦付き表記（年は無視して月だけ使用）
   - 「2024年5月」「2024/05」「202405」等の西暦表記
   - 「対象年月」「給与年月」列の値
   - **OCR 誤読対策**: 「月」が「目」「Π」に化けていても文脈で月と判定
9. **【数値の正規化】**
   - 千区切りと小数点の混乱（OCR が「90,000」を「90.000」と誤読する）に注意。金額は **整数（円単位）** に正規化
   - 「3桁ごとに区切られた数値」は千区切りとして整数化（カンマ・ピリオド区別なし）
10. **【ページ境界・チャンク分割の処理】**
    - 「`----- pdf chunk boundary -----`」「`===== page N =====`」「`--- table N ---`」のようなマーカーは PDF 分割・ページ区切り。マーカーの前後で同名の従業員が出てきたら統合する。マーカー自体は抽出対象外

{fiscal_period_section}

【出力形式（厳密に従ってください）】
```json
[
  {{
    "name": "従業員名",
    "employment_type": "正社員",
    "monthly_wages": [430000, 316000, null, null, null, null, null, null, null, null, null, null],
    "monthly_hours": [160, 160, null, null, null, null, null, null, null, null, null, null],
    "monthly_work_days": [20, 21, null, null, null, null, null, null, null, null, null, null]
  }}
]
```
↑ 上の例では Index 0=1月(430000円), Index 1=2月(316000円), 残り全 null。
   仮に「12月のみ給与あり」なら `[null,null,null,null,null,null,null,null,null,null,null,500000]` となる。

【重要な注意】
- monthly_wages / monthly_hours / monthly_work_days は **必ず12要素** の配列にしてください。データがない月は null。
- 金額は **円単位の整数**。コンマや「円」記号は付けないでください。
- 労働時間が記載されていない賃金台帳では monthly_hours は全て null で構いません（後段で労働日数×8hで補完します）。
- 役員報酬は役員として抽出してください（employment_type に「役員」を含める）。
- 名前のフリガナや空欄行は無視してください。
- **登場するすべての従業員を漏れなく出力してください。N名いれば N要素のJSON配列にする。最初の数名で打ち切ったり、代表者だけ抽出することは禁止です。**
- JSON以外のコメント・説明文は一切含めないでください（先頭・末尾とも純粋なJSON配列のみ）。

【賃金台帳データ】
{tsv_data}
"""

PROMPT_WAGE_LEDGER_FISCAL_FILTER = """【前事業年度フィルタ】
納税証明書から判定された前事業年度の決算期は **{fiscal_period}** です。
賃金台帳に複数年度のデータが含まれている場合は、この期間に該当する12ヶ月分のデータのみ抽出してください。
それ以外の月のデータは monthly_wages / monthly_hours に含めないでください（該当月のセルはあっても null）。

ただし、賃金台帳が既に前事業年度の12ヶ月分のみで構成されている場合（例: 「R6.5月」〜「R7.4月」の12列のみ）は、すべてのデータを抽出してください。
"""

PROMPT_WAGE_LEDGER_NO_FILTER = """【期間フィルタ】
納税証明書からの決算期情報は提供されていません。賃金台帳に登場するすべての月のデータを抽出してください。
複数年度に渡る場合は、各従業員について **直近12ヶ月** のデータを優先してください。
"""

PROMPT_WAGE_LEDGER_PDF_NOTE = """【添付PDFについて】
このメッセージには賃金台帳のPDFが添付されています。Excel(TSV)が提供されていない場合はPDFのみから、両方ある場合は両方を統合して抽出してください。
PDFが複数ページにわたる場合も漏れなく全ページを参照し、表中の全従業員を抽出してください。
PDF内の数値はカンマ・円記号・空白を取り除き、純粋な整数に正規化してください。
"""


class BaseExtractor(ABC):
    """データ抽出の基底クラス"""

    @abstractmethod
    def extract_registry(self, images: list[bytes]) -> CompanyInfo:
        """履歴事項全部証明書から会社情報を抽出"""
        ...

    @abstractmethod
    def extract_pl(self, images: list[bytes]) -> FinancialData:
        """損益計算書から財務データを抽出"""
        ...

    @abstractmethod
    def extract_tax(self, images: list[bytes]) -> TaxCertificate:
        """納税証明書からデータを抽出"""
        ...

    @abstractmethod
    def extract_wages(self, images: list[bytes], year_month: str) -> MonthlyWageData:
        """給与支給控除一覧から従業員データを抽出"""
        ...

    @abstractmethod
    def extract_estimate(self, images: list[bytes]) -> EstimateData:
        """見積書からデータを抽出"""
        ...

    @abstractmethod
    def generate_ai_judgment(self, company: CompanyInfo, financial: FinancialData,
                              tool_name: str, hearing_data: dict | None = None) -> AIJudgment:
        """AI判断項目を生成"""
        ...

    @abstractmethod
    def extract_wage_ledger(
        self,
        tsv_data: str,
        fiscal_period_hint: str | None = None,
        pdf_files: list[tuple[str, bytes]] | None = None,
    ) -> list[dict]:
        """賃金台帳のTSV/PDFから従業員データを抽出。

        Args:
            tsv_data: 全シートをTSV形式で結合したテキスト（PDFのみのときは空文字でも可）
            fiscal_period_hint: 前事業年度の決算期（例: 'R6.5-R7.4' または '2024-05〜2025-04'）
            pdf_files: (ファイル名, PDF bytes) のリスト。Excelとの混在もOK

        Returns:
            従業員データのリスト。各要素は {name, employment_type, monthly_wages[12], monthly_hours[12]}
        """
        ...


class StubExtractor(BaseExtractor):
    """
    APIキーなしで動作するスタブ。
    全フィールドにプレースホルダ値 '[要API: xxx]' を設定。
    """

    STUB_MARKER = '[要API]'

    def extract_registry(self, images: list[bytes]) -> CompanyInfo:
        logger.warning(f'{self.STUB_MARKER} 履歴事項の読取にはClaude APIが必要です')
        return CompanyInfo(
            name=f'{self.STUB_MARKER} 法人名',
            name_kana=f'{self.STUB_MARKER} フリガナ',
            address=f'{self.STUB_MARKER} 所在地',
            established_date=None,
            capital=0,
            representative_name=f'{self.STUB_MARKER} 代表者名',
            representative_title='代表取締役',
            officers=[],
            business_purposes=[],
        )

    def extract_pl(self, images: list[bytes]) -> FinancialData:
        logger.warning(f'{self.STUB_MARKER} 損益計算書の読取にはClaude APIが必要です')
        return FinancialData()

    def extract_tax(self, images: list[bytes]) -> TaxCertificate:
        logger.warning(f'{self.STUB_MARKER} 納税証明書の読取にはClaude APIが必要です')
        return TaxCertificate()

    def extract_wages(self, images: list[bytes], year_month: str) -> MonthlyWageData:
        logger.warning(f'{self.STUB_MARKER} 給与データの読取にはClaude APIが必要です')
        return MonthlyWageData(year_month=year_month)

    def extract_estimate(self, images: list[bytes]) -> EstimateData:
        logger.warning(f'{self.STUB_MARKER} 見積書の読取にはClaude APIが必要です')
        return EstimateData()

    def generate_ai_judgment(self, company, financial, tool_name, hearing_data=None) -> AIJudgment:
        logger.warning(f'{self.STUB_MARKER} AI判断にはClaude APIが必要です')

        # 営業利益の符号だけで判定できる部分はスタブでも埋める
        is_profitable = financial.operating_profit > 0 if financial.operating_profit else False
        return AIJudgment(
            industry_code=f'{self.STUB_MARKER}',
            industry_text=f'{self.STUB_MARKER}',
            business_description=f'{self.STUB_MARKER} 事業内容（250-255文字）',
            management_intent=(
                '■事業の拡大に積極的\n□事業の維持に注力\n□事業の売却・整備・廃業を考えている\n□特に意識したことは無い'
                if is_profitable else
                '□事業の拡大に積極的\n■事業の維持に注力\n□事業の売却・整備・廃業を考えている\n□特に意識したことは無い'
            ),
            future_goals=(
                '■事業の拡大\n□利益の確保' if is_profitable else '□事業の拡大\n■利益の確保'
            ),
            security_status=(
                '□緊急時の対応マニュアルや手順を定め、定期的に訓練を行っている\n'
                '■パソコンやサーバなどには、IDやパスワードを設け情報セキュリティ管理を行っている\n'
                '□セキュリティ対策は講じていないため、対策を講じていく\n'
                '□セキュリティ対策を講じておらず、今後もその予定はない'
            ),
            business_types=f'{self.STUB_MARKER}',
            it_investment_status='■今までIT投資を行っていなかった',
            it_utilization_status='■ITツールを導入しておらず、今回が初めてである',
        )

    def extract_wage_ledger(
        self,
        tsv_data: str,
        fiscal_period_hint: str | None = None,
        pdf_files: list[tuple[str, bytes]] | None = None,
    ) -> list[dict]:
        logger.warning(f'{self.STUB_MARKER} 賃金台帳のAI抽出にはClaude APIが必要です')
        return []


class ClaudeExtractor(BaseExtractor):
    """Claude API による実データ抽出"""

    def __init__(
        self,
        api_key: str,
        model: str = 'claude-sonnet-4-6',
        retry_callback: Optional[RetryCallback] = None,
        timeout: float = 480.0,
    ):
        try:
            import anthropic
        except ImportError:
            raise ImportError('anthropic パッケージが必要です: pip install anthropic')

        # PDF含む長時間レスポンスでクライアント側がぶら下がるのを防ぐ。
        # 480秒応答なしで例外 → _messages_create_with_retry が指数バックオフで再試行。
        # （事前分割で 15P/4MB 超は分割送信するため、1チャンクあたりは確実にこの範囲内に収まる）
        self.client = anthropic.Anthropic(api_key=api_key, timeout=timeout)
        self.model = model
        self.retry_callback = retry_callback
        # 抽出関数 (_extract_pl_pl_section 等) が API エラーで失敗した際の「失敗理由」を記録。
        # caller名 → ユーザー向けメッセージ（_format_api_error の戻り値）を蓄積し、
        # 上位の信頼度判定（_extract_pl_structured）が confidence.reason に転記する。
        # 結果として pipeline._build_confidence_warnings → UI「📋 確認キュー」まで届く。
        self._extraction_errors: dict[str, str] = {}
        logger.info(f'Claude API 初期化完了 (model={model}, timeout={timeout}s)')

    def _format_api_error(self, e: Exception) -> str:
        """API例外をユーザー向けの分かりやすい日本語メッセージに変換する。

        UI の「📋 確認キュー」にそのまま表示されるため、
        「何が起きたか」+「どう対処するか」を1文で含める。
        """
        import anthropic

        if isinstance(e, APICreditExhaustedError):
            return 'API残高切れ — APIキー管理担当者にチャージを依頼してください'
        if isinstance(e, anthropic.APITimeoutError):
            return 'APIタイムアウト (480秒) — リクエストが大きすぎる可能性があります'
        if isinstance(e, anthropic.APIConnectionError):
            return 'ネットワーク接続エラー — しばらく待って再実行してください'
        if isinstance(e, anthropic.APIStatusError):
            status = getattr(e, 'status_code', None)
            if status == 529:
                return (
                    'Anthropic API 過負荷 (529 Overloaded) — 一時的な状態です。'
                    '10〜30分待ってから再実行してください'
                )
            if status == 429:
                return (
                    'Anthropic API レート制限 (429) — '
                    'リクエストが多すぎます。少し時間を置いて再実行してください'
                )
            if status in (500, 502, 503, 504):
                return f'Anthropic API サーバー障害 ({status}) — 一時的な状態です。再実行してください'
            if status == 422:
                return f'API入力検証エラー (422) — {e}'
            return f'Anthropic API エラー ({status})'
        return f'抽出エラー ({type(e).__name__}): {e}'

    def _messages_create_with_retry(self, *, caller: str, stats: str, **kwargs):
        """messages.create をバックオフ付きで呼び出す。

        - 422/429/5xx/529/connection エラーは最大1回まで再試行
        - **timeout（480秒経過）は即失敗**（再試行しても同条件で再度 timeout するだけのため）
        - 400 credit_balance_too_low は APICreditExhaustedError に変換して即失敗
        - その他の 400/401/403/404/413 は即失敗
        - 再試行時は retry_callback(attempt, max_attempts, wait, err_summary) を呼ぶ
        """
        import anthropic

        last_error: Optional[Exception] = None
        for attempt in range(1, MAX_API_ATTEMPTS + 1):
            try:
                return self.client.messages.create(**kwargs)

            except anthropic.BadRequestError as e:
                # 400: 残高切れだけは専用例外に、それ以外は即失敗（リトライしても無駄）
                if CREDIT_BALANCE_MARKER in str(e).lower():
                    logger.error(f'[API残高切れ] caller={caller} {stats}')
                    raise APICreditExhaustedError(
                        'APIの残高が不足しています。APIキー管理担当者にチャージを依頼してください。'
                    ) from e
                logger.error(f'[API失敗/400] caller={caller} {stats} error={e}')
                raise

            except (anthropic.AuthenticationError,
                    anthropic.PermissionDeniedError,
                    anthropic.NotFoundError) as e:
                # 401/403/404: 設定ミス系、リトライ無意味
                logger.error(f'[API失敗/非リトライ] caller={caller} {stats} error={type(e).__name__}: {e}')
                raise

            except anthropic.APIStatusError as e:
                # 422/429/5xx/529 などステータスコード付きエラー
                status = getattr(e, 'status_code', None)
                if status in RETRYABLE_STATUS_CODES and attempt < MAX_API_ATTEMPTS:
                    last_error = e
                    wait = API_BACKOFF_SECONDS[attempt - 1]
                    err_summary = f'{status} {type(e).__name__}'
                    logger.warning(
                        f'[API再試行] caller={caller} {attempt}/{MAX_API_ATTEMPTS} '
                        f'wait={wait}s error={err_summary}: {e}'
                    )
                    if self.retry_callback:
                        try:
                            self.retry_callback(attempt, MAX_API_ATTEMPTS, wait, err_summary)
                        except Exception as cb_err:
                            logger.warning(f'retry_callback実行失敗: {cb_err}')
                    time.sleep(wait)
                    continue
                logger.error(f'[API失敗/確定] caller={caller} {stats} status={status} error={e}')
                raise

            # ⚠ 重要: APITimeoutError は APIConnectionError のサブクラス。
            # この2つの except 節の並び順は **絶対に入れ替えないこと**。
            # 入れ替えると Timeout が Connection 扱いでリトライされ、
            # 300 秒 × N 回 = 数分〜十数分のハングが再発する。
            except anthropic.APITimeoutError as e:
                # timeout は即失敗（リトライしても 300 秒 × N 回浪費するだけ）
                logger.error(
                    f'[API失敗/timeout即失敗] caller={caller} {stats} '
                    f'error={type(e).__name__}: {e}'
                )
                raise

            except anthropic.APIConnectionError as e:
                # ネットワーク瞬断は1回だけリトライ
                if attempt < MAX_API_ATTEMPTS:
                    last_error = e
                    wait = API_BACKOFF_SECONDS[attempt - 1]
                    err_summary = type(e).__name__
                    logger.warning(
                        f'[API再試行] caller={caller} {attempt}/{MAX_API_ATTEMPTS} '
                        f'wait={wait}s error={err_summary}: {e}'
                    )
                    if self.retry_callback:
                        try:
                            self.retry_callback(attempt, MAX_API_ATTEMPTS, wait, err_summary)
                        except Exception as cb_err:
                            logger.warning(f'retry_callback実行失敗: {cb_err}')
                    time.sleep(wait)
                    continue
                logger.error(f'[API失敗/確定] caller={caller} {stats} error={e}')
                raise

        # ループを抜けた = リトライ全敗（通常到達しない。安全網）
        if last_error:
            raise last_error
        raise RuntimeError('API呼出しリトライが想定外に終了しました')

    def _call_api(self, images: list[bytes], prompt: str, max_tokens: int = 4096) -> str:
        """画像+プロンプトでAPIを呼び出し、テキストを返す"""
        import base64
        import traceback
        content = []

        raw_sizes = []
        b64_sizes = []
        for img in images:
            b64 = base64.standard_b64encode(img).decode('ascii')
            raw_sizes.append(len(img))
            b64_sizes.append(len(b64))
            content.append({
                'type': 'image',
                'source': {'type': 'base64', 'media_type': 'image/png', 'data': b64}
            })

        content.append({'type': 'text', 'text': prompt})

        # 送信直前のペイロード統計（422/413/529 の原因切り分け用）
        n = len(images)
        raw_mb = sum(raw_sizes) / 1_000_000
        b64_mb = sum(b64_sizes) / 1_000_000
        raw_max = max(raw_sizes) / 1_000_000 if raw_sizes else 0
        prompt_chars = len(prompt)
        caller = traceback.extract_stack()[-2].name  # extract_tax 等、どのメソッドからの呼び出しか
        stats = (
            f'images={n}枚 raw合計={raw_mb:.2f}MB raw最大={raw_max:.2f}MB '
            f'base64合計={b64_mb:.2f}MB prompt={prompt_chars}chars max_tokens={max_tokens}'
        )
        logger.warning(f'[API送信] caller={caller} {stats}')

        response = self._messages_create_with_retry(
            caller=caller,
            stats=stats,
            model=self.model,
            max_tokens=max_tokens,
            messages=[{'role': 'user', 'content': content}],
        )

        text = response.content[0].text
        logger.warning(
            f'[API成功] caller={caller} '
            f'応答={len(text)}chars '
            f'tokens={response.usage.input_tokens}in+{response.usage.output_tokens}out'
        )
        return text

    def _parse_json(self, text: str) -> dict | list:
        """API応答からJSONを抽出・パース"""
        # ```json ... ``` ブロックがあれば中身を取り出す
        if '```json' in text:
            start = text.index('```json') + 7
            end = text.index('```', start)
            text = text[start:end].strip()
        elif '```' in text:
            start = text.index('```') + 3
            end = text.index('```', start)
            text = text[start:end].strip()

        return json.loads(text)

    def _ensure_dict(self, data, caller: str) -> dict:
        """dict 期待の API 応答が list で返ってきた時のフォールバック。

        Sonnet が「複数年度の決算書を返したい」等の理由で JSON 配列を返すケースが
        実本番で観測された。プロンプト指定は dict だが、Sonnet が独自解釈で list 化
        することがあるため、安全側で list の最初の要素 (dict) を採用する。
        どちらでもなければ空 dict を返して呼出側でデフォルト値で埋める。
        """
        if isinstance(data, dict):
            return data
        if isinstance(data, list) and data and isinstance(data[0], dict):
            logger.warning(
                f'[{caller}] AI 応答が JSON 配列で返却されました '
                f'(要素数={len(data)}) → 先頭要素を採用します'
            )
            return data[0]
        logger.error(
            f'[{caller}] AI 応答が dict でも有効な list でもありません: '
            f'type={type(data).__name__}'
        )
        return {}

    def extract_registry(self, images: list[bytes]) -> CompanyInfo:
        text = self._call_api(images, PROMPT_REGISTRY)
        data = self._ensure_dict(self._parse_json(text), 'extract_registry')

        # 役員リスト（同一人物の重複を排除）
        officers = []
        seen_names = set()
        rep_name = data.get('representative_name', '')
        for o in data.get('officers', []):
            name = o.get('name', '').strip()
            if not name or name in seen_names or name == rep_name:
                continue
            seen_names.add(name)
            officers.append({
                'title': o.get('title', ''),
                'name': name,
                'kana': o.get('kana', ''),
            })

        from datetime import datetime
        est = None
        if data.get('established_date'):
            try:
                est = datetime.strptime(data['established_date'], '%Y-%m-%d')
            except ValueError:
                pass

        return CompanyInfo(
            name=data.get('name') or '',
            name_kana=data.get('name_kana') or '',
            address=data.get('address') or '',
            postal_code=data.get('postal_code') or '',
            established_date=est,
            capital=data.get('capital', 0) or 0,
            representative_name=data.get('representative_name') or '',
            representative_title=data.get('representative_title') or '',
            officers=officers,
            business_purposes=data.get('business_purposes') or [],
        )

    def extract_pl(self, images: list[bytes]) -> FinancialData:
        """PL抽出。構造分解フロー（3呼出方式）と従来フロー（1呼出方式）を切替可能。

        環境変数 USE_STRUCTURED_PL_EXTRACTION で切替（デフォルト: true=新方式）。
        新方式は「基本PL」「販管費部」「原価部」の3呼出に分けて、コード側で機械的に合算。
        Codex 指摘の「1回の視覚抽出に複数タスクを背負わせる」問題への抜本対策。
        """
        from .config import USE_STRUCTURED_PL_EXTRACTION

        if USE_STRUCTURED_PL_EXTRACTION:
            return self._extract_pl_structured(images)

        # ── 従来フロー（保険として残す。USE_STRUCTURED_PL_EXTRACTION=false で有効） ──
        text = self._call_api(images, PROMPT_PL)
        d = self._ensure_dict(self._parse_json(text), 'extract_pl')

        # 異常検知: 製造原価/完成工事原価の見落とし疑いがあれば、原価フォーカスで再抽出して合算
        # （建設業・製造業の決算書で、販管費しか抽出しない事象が本番で多発したため）
        suspicious, reason = _is_pl_extraction_suspicious(d)
        if suspicious and images:
            logger.warning(
                f'[extract_pl] PL抽出に異常検知: {reason} → 原価報告書フォーカスで再抽出'
            )
            cost_data = self._extract_pl_cost_report_only(images)
            d = _merge_pl_with_cost_report(d, cost_data)
            logger.warning(
                f'[extract_pl] 原価合算後: salary={d.get("salary", 0):,} '
                f'misc_wages={d.get("misc_wages", 0):,} bonus={d.get("bonus", 0):,}'
            )

        # 決算月を事業年度終了日から推定
        fiscal_month = ''
        if d.get('fiscal_year_end'):
            month = d['fiscal_year_end'].split('-')[1] if '-' in d['fiscal_year_end'] else ''
            month_names = {'01': '1月', '02': '2月', '03': '3月', '04': '4月',
                          '05': '5月', '06': '6月', '07': '7月', '08': '8月',
                          '09': '9月', '10': '10月', '11': '11月', '12': '12月'}
            fiscal_month = month_names.get(month, '')

        return FinancialData(
            fiscal_year_start=d.get('fiscal_year_start', ''),
            fiscal_year_end=d.get('fiscal_year_end', ''),
            fiscal_month=fiscal_month,
            revenue=d.get('revenue', 0) or 0,
            cost_of_sales=d.get('cost_of_sales', 0) or 0,
            gross_profit=d.get('gross_profit', 0) or 0,
            operating_profit=d.get('operating_profit', 0) or 0,
            ordinary_profit=d.get('ordinary_profit', 0) or 0,
            net_profit=d.get('net_profit', 0) or 0,
            salary=d.get('salary', 0) or 0,
            misc_wages=d.get('misc_wages', 0) or 0,
            bonus=d.get('bonus', 0) or 0,
            officer_compensation=d.get('officer_compensation', 0) or 0,
            legal_welfare=d.get('legal_welfare', 0) or 0,
            welfare=d.get('welfare', 0) or 0,
            depreciation=d.get('depreciation', 0) or 0,
            travel_expense=d.get('travel_expense', 0) or 0,
        )

    def extract_tax(self, images: list[bytes]) -> TaxCertificate:
        text = self._call_api(images, PROMPT_TAX)
        d = self._ensure_dict(self._parse_json(text), 'extract_tax')
        return TaxCertificate(
            tax_type=d.get('tax_type', ''),
            tax_amount=d.get('tax_amount', 0) or 0,
            fiscal_year=d.get('fiscal_year', ''),
        )

    def extract_wages(self, images: list[bytes], year_month: str) -> MonthlyWageData:
        text = self._call_api(images, PROMPT_WAGES, max_tokens=8192)
        data = self._parse_json(text)

        employees = []
        for e in data:
            employees.append(Employee(
                name=e.get('name', ''),
                department=e.get('department', ''),
                employee_id=e.get('employee_id', ''),
                employment_type=e.get('employment_type', ''),
                working_days=e.get('working_days', 0) or 0,
                scheduled_hours=e.get('scheduled_hours', 0) or 0,
                base_salary=e.get('base_salary', 0) or 0,
                taxable_total=e.get('taxable_total', 0) or 0,
                total_pay=e.get('total_pay', 0) or 0,
                deductions=e.get('deductions', 0) or 0,
                net_pay=e.get('net_pay', 0) or 0,
            ))

        return MonthlyWageData(year_month=year_month, employees=employees)

    def extract_estimate(self, images: list[bytes]) -> EstimateData:
        text = self._call_api(images, PROMPT_ESTIMATE)
        d = self._ensure_dict(self._parse_json(text), 'extract_estimate')

        items = [{'name': i.get('name', ''), 'amount': i.get('amount', 0)}
                 for i in d.get('items', []) if isinstance(i, dict)]

        return EstimateData(
            vendor_name=d.get('vendor_name', ''),
            tool_name=d.get('tool_name', ''),
            items=items,
            total_amount=d.get('total_amount', 0) or 0,
            tax_amount=d.get('tax_amount', 0) or 0,
        )

    def generate_ai_judgment(self, company, financial, tool_name, hearing_data=None) -> AIJudgment:
        # ヒアリングデータから各種情報を取得
        hearing_fields = {
            'it_investment_answer': '不明',
            'it_investment_amount': '不明',
            'it_investment_process': '不明',
            'main_business': '',
            'strength': '',
            'challenge': '',
            'monthly_hours': '',
            'tool_usage': '',
            'reduction': '',
            'freed_time': '',
            'sales_target': '',
        }
        if hearing_data:
            FIELD_KEYWORDS = {
                'main_business': ['主な事業内容'],
                'strength': ['強み'],
                'challenge': ['時間がかかっている'],
                'monthly_hours': ['月間何時間'],
                'tool_usage': ['どの機能'],
                'reduction': ['何％', '何時間'],
                'freed_time': ['浮いた時間'],
                'sales_target': ['売上目標'],
            }
            for row_num, item in hearing_data.items():
                label = str(item.get('label', ''))
                value = item.get('value')
                if 'IT投資' in label and '金額' in label:
                    hearing_fields['it_investment_answer'] = 'はい' if value else 'いいえ'
                    hearing_fields['it_investment_amount'] = str(value) if value else 'なし'
                elif 'IT投資' in label and 'プロセス' in label:
                    hearing_fields['it_investment_process'] = str(value) if value else 'なし'
                else:
                    for field_key, keywords in FIELD_KEYWORDS.items():
                        if any(kw in label for kw in keywords):
                            hearing_fields[field_key] = str(value) if value else ''
                            break

        prompt = PROMPT_AI_JUDGMENT.format(
            company_name=company.name,
            business_purposes=', '.join(company.business_purposes),
            address=company.address,
            operating_profit=financial.operating_profit,
            tool_name=tool_name,
            **hearing_fields,
        )

        # AI判断はテキストのみ（画像なし）
        stats = f'images=0枚 prompt={len(prompt)}chars max_tokens=4096'
        logger.warning(f'[API送信] caller=generate_ai_judgment {stats}')
        response = self._messages_create_with_retry(
            caller='generate_ai_judgment',
            stats=stats,
            model=self.model,
            max_tokens=4096,
            messages=[{'role': 'user', 'content': prompt}],
        )
        text = response.content[0].text
        logger.warning(
            f'[API成功] caller=generate_ai_judgment '
            f'応答={len(text)}chars '
            f'tokens={response.usage.input_tokens}in+{response.usage.output_tokens}out'
        )
        d = self._parse_json(text)

        # 最低賃金はconfig.pyから取得
        from .config import get_min_wage
        mw = get_min_wage(company.address)
        min_wage_text = f'{mw[0]}/{mw[1]}円' if mw else d.get('min_wage', '')

        return AIJudgment(
            industry_code=d.get('industry_code', ''),
            industry_text=d.get('industry_text', ''),
            business_description=d.get('business_description', ''),
            management_intent=d.get('management_intent', ''),
            future_goals=d.get('future_goals', ''),
            security_status=d.get('security_status', ''),
            business_types=d.get('business_types', ''),
            min_wage_text=min_wage_text,
            it_investment_status=d.get('it_investment_status', ''),
            it_utilization_status=d.get('it_utilization_status', ''),
            it_utilization_scope=d.get('it_utilization_scope', ''),
            invoice_related_work=d.get('invoice_related_work', ''),
            weakness=d.get('weakness', ''),
            it_investment_process=d.get('it_investment_process', ''),
            improvement_process=d.get('improvement_process', ''),
            expected_effect_dept=d.get('expected_effect_dept', ''),
            expected_effect=d.get('expected_effect', ''),
        )

    def extract_wage_ledger(
        self,
        tsv_data: str,
        fiscal_period_hint: str | None = None,
        pdf_files: list[tuple[str, bytes]] | None = None,
        _retry_depth: int = 0,
    ) -> list[dict]:
        """賃金台帳のTSV/PDFから従業員データを抽出。

        Args:
            _retry_depth: 内部用。0=初回、1=分割中（再々分割しないためのカウンタ）

        事前分割（API呼出 +1回）:
            初回のみ、PDF が WAGE_LEDGER_SPLIT_PAGE_THRESHOLD ページ超または
            WAGE_LEDGER_SPLIT_BYTES_THRESHOLD バイト超なら、最初から半分に分割して送信。
            timeout (480秒) 前に確実に処理できる規模に抑える。

        事後分割（API呼出 +1回）:
            事前分割しなかった場合、stop_reason=='max_tokens' を検出した時に
            分割再抽出する。誤発動防止のため発動条件は max_tokens のみ。

        API呼出上限:
            事前分割した場合は事後分割を発動しない（既に分割済みのため）→ 最大 2回
            事前分割しなかった場合は事後分割を最大 1回発動 → 最大 2回
            → 1案件あたり API呼出は最大 2回に制限される（コスト爆増防止）

        Phase 2 — PDFテキスト前処理（USE_PDF_TEXT_PREPROCESSING=true 時）:
            画像PDFを送る前に pdfplumber/PyMuPDF でテキスト化してtsv_dataに統合し、
            PDF送信そのものをスキップする。画像トークンが消えるため、賃金台帳のような
            大量PDF案件で総コストを 1/5〜1/10 に圧縮できる。
            テキスト化失敗時は自動的に画像経路にフォールバック（既存挙動維持）。
        """
        # ── Phase 2/3: PDFテキスト前処理 (pdfplumber) / OCR (Document AI) ──
        # どちらかのフラグONかつ初回のみ。pdf_text_extractor 側で経路判定する。
        from .config import USE_PDF_TEXT_PREPROCESSING, USE_DOCUMENT_AI_OCR
        if (USE_PDF_TEXT_PREPROCESSING or USE_DOCUMENT_AI_OCR) and pdf_files and _retry_depth == 0:
            try:
                from .pdf_text_extractor import extract_pdf_as_text
                blocks = []
                for fname, pdf_bytes in pdf_files:
                    text = extract_pdf_as_text(pdf_bytes)
                    blocks.append(f'### {fname}（PDFテキスト化済み）\n{text}')
                pdf_as_text = '\n\n'.join(blocks)
                if tsv_data and tsv_data.strip():
                    tsv_data = (
                        tsv_data
                        + '\n\n=== 賃金台帳PDF（テキスト化済み・原本順） ===\n'
                        + pdf_as_text
                    )
                else:
                    tsv_data = '=== 賃金台帳PDF（テキスト化済み・原本順） ===\n' + pdf_as_text
                logger.warning(
                    f'[extract_wage_ledger] PDF→テキスト化成功: '
                    f'{len(pdf_files)}件 → 統合TSV={len(tsv_data)}chars, 画像送信スキップ'
                )
                pdf_files = None  # 以降の経路は PDF なし扱い
            except Exception as e:
                logger.warning(
                    f'[extract_wage_ledger] PDF→テキスト化失敗、画像経路にフォールバック: {e}'
                )
                # pdf_files はそのまま、既存経路へ

        # ── 事前分割判定（初回のみ） ──
        if _retry_depth == 0 and pdf_files:
            should_pre_split = any(
                _pdf_should_be_split(pdf_bytes)
                for _, pdf_bytes in pdf_files
            )
            if should_pre_split:
                logger.warning(
                    f'[extract_wage_ledger] 事前分割発動: '
                    f'PDF {len(pdf_files)}件中に閾値超 '
                    f'(>{WAGE_LEDGER_SPLIT_PAGE_THRESHOLD}P or '
                    f'>{WAGE_LEDGER_SPLIT_BYTES_THRESHOLD/1_000_000:.0f}MB) があります'
                )
                return self._extract_with_pre_split(
                    tsv_data, fiscal_period_hint, pdf_files,
                )

        if fiscal_period_hint:
            fiscal_section = PROMPT_WAGE_LEDGER_FISCAL_FILTER.format(
                fiscal_period=fiscal_period_hint
            )
        else:
            fiscal_section = PROMPT_WAGE_LEDGER_NO_FILTER

        tsv_for_prompt = tsv_data if tsv_data else '(Excel(TSV)データなし、添付PDFのみを参照してください)'

        # 出力JSONは従業員数に比例して大きくなるため最大16Kトークン
        max_tokens = 16384

        # PDF document ブロック（prompt caching 有無で配置位置が変わるため先に作っておく）
        pdf_blocks: list = []
        pdf_total_bytes = 0
        if pdf_files:
            import base64
            for fname, pdf_bytes in pdf_files:
                pdf_total_bytes += len(pdf_bytes)
                b64 = base64.standard_b64encode(pdf_bytes).decode('ascii')
                pdf_blocks.append({
                    'type': 'document',
                    'source': {
                        'type': 'base64',
                        'media_type': 'application/pdf',
                        'data': b64,
                    },
                    'title': fname,
                })

        # メッセージ content を構築。
        # USE_PROMPT_CACHING=true: 固定指示(cache対象) → 動的指示 → PDF の順
        # USE_PROMPT_CACHING=false: PDF → プロンプト全体（1テキスト）の旧構成
        from .config import USE_PROMPT_CACHING
        content: list = []
        if USE_PROMPT_CACHING:
            tail_text = PROMPT_WAGE_LEDGER_TAIL.format(
                tsv_data=tsv_for_prompt,
                fiscal_period_section=fiscal_section,
            )
            if pdf_blocks:
                tail_text += '\n\n' + PROMPT_WAGE_LEDGER_PDF_NOTE
            content.append({
                'type': 'text',
                'text': PROMPT_WAGE_LEDGER_STATIC,
                'cache_control': {'type': 'ephemeral'},
            })
            content.append({'type': 'text', 'text': tail_text})
            content.extend(pdf_blocks)
            prompt_for_log = PROMPT_WAGE_LEDGER_STATIC + tail_text
        else:
            prompt = PROMPT_WAGE_LEDGER.format(
                tsv_data=tsv_for_prompt,
                fiscal_period_section=fiscal_section,
            )
            if pdf_blocks:
                prompt = prompt + '\n\n' + PROMPT_WAGE_LEDGER_PDF_NOTE
            content.extend(pdf_blocks)
            content.append({'type': 'text', 'text': prompt})
            prompt_for_log = prompt

        # Haiku モード: OCR テキストを Haiku 4.5 で抽出（コスト 1/3）
        from .config import USE_OCR_HAIKU_EXTRACTION, HAIKU_MODEL
        effective_model = HAIKU_MODEL if USE_OCR_HAIKU_EXTRACTION else self.model

        pdf_count = len(pdf_files) if pdf_files else 0
        stats = (
            f'pdfs={pdf_count}件 pdf合計={pdf_total_bytes/1_000_000:.2f}MB '
            f'prompt={len(prompt_for_log)}chars max_tokens={max_tokens} '
            f'retry={_retry_depth} cache={"on" if USE_PROMPT_CACHING else "off"} '
            f'model={effective_model}'
        )
        logger.warning(f'[API送信] caller=extract_wage_ledger {stats}')
        response = self._messages_create_with_retry(
            caller='extract_wage_ledger',
            stats=stats,
            model=effective_model,
            max_tokens=max_tokens,
            messages=[{'role': 'user', 'content': content}],
        )
        text = response.content[0].text
        output_tokens = response.usage.output_tokens
        cache_create = getattr(response.usage, 'cache_creation_input_tokens', 0) or 0
        cache_read = getattr(response.usage, 'cache_read_input_tokens', 0) or 0
        stop_reason = getattr(response, 'stop_reason', None)
        logger.warning(
            f'[API成功] caller=extract_wage_ledger '
            f'応答={len(text)}chars '
            f'tokens={response.usage.input_tokens}in+{output_tokens}out '
            f'cache={cache_create}create/{cache_read}read '
            f'stop_reason={stop_reason}'
        )

        # 打ち切り検出: stop_reason==max_tokens が最も信頼できるシグナル
        # （Anthropic API が「出力をmax_tokensで切った」と明示してくる）
        truncated = (stop_reason == 'max_tokens')

        try:
            data = self._parse_json(text)
        except json.JSONDecodeError as e:
            logger.error(f'[extract_wage_ledger] JSON解析失敗: {e}, 応答先頭500文字: {text[:500]}')
            # 打ち切り + PDF あり + 初回 → 分割再抽出 (JSON が途中で切れて parse 失敗の可能性大)
            if truncated and pdf_files and _retry_depth == 0:
                logger.warning('[extract_wage_ledger] JSON parse失敗 + max_tokens打ち切り → PDF分割再抽出')
                return self._retry_extract_with_split_pdf(
                    tsv_data, fiscal_period_hint, pdf_files, partial_data=[]
                )
            return []

        if not isinstance(data, list):
            logger.error(f'[extract_wage_ledger] 応答がリストではありません: type={type(data).__name__}')
            return []

        # 打ち切り検出 + PDF あり + 初回 → 分割再抽出してマージ
        # (JSON parse は通ったが、配列途中で max_tokens 到達した可能性 = 後半の従業員が欠けている)
        if truncated and pdf_files and _retry_depth == 0:
            logger.warning(
                f'[extract_wage_ledger] max_tokens打ち切り検出 (出力{output_tokens}tok使用) '
                f'→ {len(data)}名の取得後、PDF分割再抽出で残りを補完'
            )
            return self._retry_extract_with_split_pdf(
                tsv_data, fiscal_period_hint, pdf_files, partial_data=data
            )

        return data

    def _extract_pl_structured(self, images: list[bytes]) -> FinancialData:
        """構造分解 PL 抽出: 3つの専門呼出 → コード側で合算 + 信頼度付与（Phase 1+2+3）。

        - basic: 売上・利益基本（PL本表のみフォーカス）
        - pl_section: 販管費の人件費系（販管費明細のみフォーカス）
        - cost_section: 原価部の人件費系（製造原価/完成工事原価/工事原価のみフォーカス）

        各プロンプトを1枚絵レベルに単純化することで Sonnet の揺らぎを抑える。
        二重計上を避けるため、各プロンプトが排他的範囲を抽出し、コード側で機械的に加算。

        Phase 2: 各フィールドに FieldConfidence (level / source_component / reason) を付与。
        異常検知に引っかかったフィールドは level='low' として扱い、
        申請書転記側で「空欄+警告」として処理する。

        Phase 3: PDF が3ページ超なら、まず軽量プロンプトで全ページを分類（インベントリ化）し、
        各専門呼出には関連ページだけ送信。複数年度PDFで前年度のページが混入する問題を解消。
        """
        from .models import FieldConfidence
        from .config import USE_PL_PAGE_INVENTORY
        if not images:
            return FinancialData()

        # 前回の抽出失敗ログをクリア（同一インスタンスで複数案件処理する場合の混入防止）
        self._extraction_errors.clear()

        # ── Phase 3: ページインベントリ化（4ページ以上で発動） ──
        # 3ページ以下は分類のメリットが薄いのでスキップ（コスト最適化）
        inventory: dict | None = None
        if USE_PL_PAGE_INVENTORY and len(images) >= 4:
            inventory = self._inventory_pdf_pages(images)

        basic_imgs = images
        pl_imgs = images
        cost_imgs = images
        if inventory:
            basic_imgs = self._select_pages_by_labels(images, inventory, ['pl_basic'], latest_only=True)
            pl_imgs = self._select_pages_by_labels(images, inventory, ['pl_section'], latest_only=True)
            cost_imgs = self._select_pages_by_labels(images, inventory, ['cost_section'], latest_only=True)
            # 該当ページがない場合は全ページにフォールバック（インベントリ判定ミスを救う）
            if not basic_imgs: basic_imgs = images
            if not pl_imgs: pl_imgs = images
            if not cost_imgs: cost_imgs = images
            logger.warning(
                f'[extract_pl_structured] ページ絞込: '
                f'basic={len(basic_imgs)}/{len(images)}枚, '
                f'pl={len(pl_imgs)}/{len(images)}枚, '
                f'cost={len(cost_imgs)}/{len(images)}枚'
            )

        # 並列で 3 呼出（API リクエストはネットワーク I/O bound なので concurrent.futures で並列化可能だが、
        # 現状は 同期で順次実行。コスト試算で 3呼出 ≈ ¥3-9/案件）
        basic = self._extract_pl_basic_section(basic_imgs)
        pl_part = self._extract_pl_pl_section(pl_imgs)
        cost_part = self._extract_pl_cost_section(cost_imgs)

        # 抽出失敗判定（空dict = API例外 or JSON parse失敗）
        basic_failed = not basic
        pl_failed = not pl_part
        cost_failed = not cost_part

        # 決算月を事業年度終了日から推定
        fiscal_month = ''
        end = basic.get('fiscal_year_end') or ''
        if end and '-' in end:
            month = end.split('-')[1]
            month_names = {'01': '1月', '02': '2月', '03': '3月', '04': '4月',
                           '05': '5月', '06': '6月', '07': '7月', '08': '8月',
                           '09': '9月', '10': '10月', '11': '11月', '12': '12月'}
            fiscal_month = month_names.get(month, '')

        # 人件費系: 販管費 + 原価部 を機械的に加算（二重計上なし）
        def _sum(key: str) -> int:
            v_pl = pl_part.get(key) or 0
            v_cost = cost_part.get(key) or 0
            try:
                return int(float(v_pl) + float(v_cost))
            except (TypeError, ValueError):
                return 0

        # 信頼度メタ: 各フィールドの source_component と level を判定
        confidence: dict = {}

        # _extract_pl_*_section の except 節で記録された失敗理由を参照（529/timeout 等を識別）
        basic_err = self._extraction_errors.get('basic', '')
        pl_err = self._extraction_errors.get('pl_section', '')
        cost_err = self._extraction_errors.get('cost_section', '')

        def _basic_conf(key: str) -> FieldConfidence:
            if basic_failed:
                reason = basic_err or '損益計算書本表の抽出失敗'
                return FieldConfidence(level='low', source_component='basic', reason=reason)
            if basic.get(key) is None:
                return FieldConfidence(level='low', source_component='basic', reason=f'{key} が null')
            return FieldConfidence(level='high', source_component='basic')

        def _personnel_conf(key: str) -> FieldConfidence:
            """人件費系（販管費 + 原価部）の信頼度判定"""
            if pl_failed and cost_failed:
                # 両方失敗 → 個別の理由を結合（同じ 529 でも両方表示することで「API 障害」が明確に）
                parts = []
                if pl_err:
                    parts.append(f'販管費: {pl_err}')
                if cost_err:
                    parts.append(f'原価部: {cost_err}')
                reason = ' / '.join(parts) if parts else '販管費・原価部とも抽出失敗'
                return FieldConfidence(level='low', source_component='unknown', reason=reason)
            if pl_failed:
                reason = f'{pl_err}（原価部のみで判定）' if pl_err else '販管費抽出失敗、原価部のみ'
                return FieldConfidence(level='medium', source_component='cost', reason=reason)
            if cost_failed:
                reason = f'{cost_err}（販管費のみで判定）' if cost_err else '原価部抽出失敗、販管費のみ'
                return FieldConfidence(level='medium', source_component='PL', reason=reason)
            v_pl = pl_part.get(key) or 0
            v_cost = cost_part.get(key) or 0
            if v_pl > 0 and v_cost > 0:
                src = 'PL+cost'
            elif v_pl > 0:
                src = 'PL'
            elif v_cost > 0:
                src = 'cost'
            else:
                src = 'unknown'
            return FieldConfidence(level='high', source_component=src)

        for key in ('fiscal_year_start', 'fiscal_year_end', 'revenue', 'cost_of_sales',
                    'gross_profit', 'operating_profit', 'ordinary_profit', 'net_profit'):
            confidence[key] = _basic_conf(key)
        for key in ('salary', 'misc_wages', 'bonus', 'legal_welfare', 'welfare',
                    'depreciation', 'travel_expense'):
            confidence[key] = _personnel_conf(key)
        # 役員報酬は販管費のみ
        if pl_failed:
            confidence['officer_compensation'] = FieldConfidence(
                level='low', source_component='PL',
                reason=pl_err or '販管費抽出失敗')
        else:
            confidence['officer_compensation'] = FieldConfidence(
                level='high', source_component='PL')

        # 整合性チェック: 売上原価/(販管費人件費+原価部人件費) > 50倍 など極端な値を low に降格
        total_personnel = _sum('salary') + _sum('misc_wages') + _sum('bonus')
        cost_of_sales = int(basic.get('cost_of_sales') or 0)
        if cost_of_sales > 5_000_000 and total_personnel > 0:
            ratio = cost_of_sales / total_personnel
            if ratio > 50:
                # 50倍超 = 原価部抽出も含めても人件費が極端に小さい = 何か根本的におかしい
                for k in ('salary', 'misc_wages', 'bonus'):
                    confidence[k] = FieldConfidence(
                        level='low', source_component=confidence[k].source_component,
                        reason=f'売上原価/人件費={ratio:.1f}倍で異常（人件費抽出全体に問題）')

        result = FinancialData(
            fiscal_year_start=basic.get('fiscal_year_start') or '',
            fiscal_year_end=end,
            fiscal_month=fiscal_month,
            revenue=int(basic.get('revenue') or 0),
            cost_of_sales=cost_of_sales,
            gross_profit=int(basic.get('gross_profit') or 0),
            operating_profit=int(basic.get('operating_profit') or 0),
            ordinary_profit=int(basic.get('ordinary_profit') or 0),
            net_profit=int(basic.get('net_profit') or 0),
            salary=_sum('salary'),
            misc_wages=_sum('misc_wages'),
            bonus=_sum('bonus'),
            officer_compensation=int(pl_part.get('officer_compensation') or 0),
            legal_welfare=_sum('legal_welfare'),
            welfare=_sum('welfare'),
            depreciation=_sum('depreciation'),
            travel_expense=_sum('travel_expense'),
            confidence=confidence,
        )

        # 低信頼項目をログに集約
        low_fields = [k for k, c in confidence.items() if c.level == 'low']
        if low_fields:
            logger.warning(
                f'[extract_pl_structured] 低信頼フィールド {len(low_fields)}件: '
                f'{", ".join(low_fields[:5])}'
                f'{"..." if len(low_fields) > 5 else ""}'
            )
        logger.warning(
            f'[extract_pl_structured] 統合結果: '
            f'salary={result.salary:,} (販管費{int(pl_part.get("salary") or 0):,} + '
            f'原価{int(cost_part.get("salary") or 0):,}) '
            f'misc_wages={result.misc_wages:,} bonus={result.bonus:,}'
        )
        return result

    def _inventory_pdf_pages(self, images: list[bytes]) -> dict | None:
        """PDF全ページを軽量プロンプトで分類（Phase 3）。

        Returns: {'pages': [{page, labels, fiscal_year_label, is_latest}, ...],
                  'latest_fiscal_year_label': str, 'latest_fiscal_year_period': str}
                 または None（失敗時）
        失敗時は呼出側で全ページ送信にフォールバック。
        ただし APICreditExhaustedError は即 raise（残高切れで何やっても無駄）。
        """
        try:
            text = self._call_api(images, PROMPT_PL_PAGE_INVENTORY, max_tokens=4096)
            d = self._ensure_dict(self._parse_json(text), 'pl_page_inventory')
            pages = d.get('pages')
            if not isinstance(pages, list) or not pages:
                logger.warning('[pl_page_inventory] pages が空 → インベントリスキップ')
                return None
            latest_year = d.get('latest_fiscal_year_label', '?')
            page_count = len(pages)
            label_summary: dict[str, int] = {}
            for p in pages:
                for lab in (p.get('labels') or []):
                    label_summary[lab] = label_summary.get(lab, 0) + 1
            logger.warning(
                f'[pl_page_inventory] {page_count}ページ分類完了: 直近年度={latest_year}, '
                f'内訳={dict(sorted(label_summary.items(), key=lambda x: -x[1]))}'
            )
            return d
        except APICreditExhaustedError:
            raise  # 残高切れは即停止（他の API 呼出を試行しない）
        except Exception as e:
            logger.warning(f'[pl_page_inventory] 失敗: {e} → 全ページ送信にフォールバック')
            return None

    def _select_pages_by_labels(
        self,
        images: list[bytes],
        inventory: dict,
        labels: list[str],
        latest_only: bool = True,
    ) -> list[bytes]:
        """インベントリ結果から指定ラベルのページだけを抽出。

        Args:
            images: 全PDFページ画像
            inventory: _inventory_pdf_pages の結果
            labels: 抽出したいラベル（例: ['pl_basic', 'pl_section']）
            latest_only: True なら is_latest=True のページだけ
        """
        pages_meta = inventory.get('pages') or []
        selected_indices: set[int] = set()
        for meta in pages_meta:
            if not isinstance(meta, dict):
                continue
            page_num = meta.get('page')
            if not isinstance(page_num, int) or page_num < 1 or page_num > len(images):
                continue
            page_labels = meta.get('labels') or []
            if not any(lab in page_labels for lab in labels):
                continue
            if latest_only and not meta.get('is_latest', True):
                continue
            selected_indices.add(page_num - 1)  # 1-indexed → 0-indexed
        # 順序維持
        return [images[i] for i in sorted(selected_indices) if 0 <= i < len(images)]

    def _extract_pl_basic_section(self, images: list[bytes]) -> dict:
        """構造分解PL: 損益計算書本表（売上・利益サマリ）のみ抽出。

        APICreditExhaustedError は再 raise（残高切れで他の呼出も無駄なため）。
        """
        try:
            text = self._call_api(images, PROMPT_PL_BASIC)
            d = self._ensure_dict(self._parse_json(text), 'extract_pl_basic')
            logger.warning(
                f'[extract_pl_basic] 売上={d.get("revenue", 0):,} '
                f'営業利益={d.get("operating_profit", 0):,} '
                f'経常利益={d.get("ordinary_profit", 0):,}'
            )
            return d
        except APICreditExhaustedError:
            raise
        except Exception as e:
            self._extraction_errors['basic'] = self._format_api_error(e)
            logger.warning(f'[extract_pl_basic] 失敗: {e}')
            return {}

    def _extract_pl_pl_section(self, images: list[bytes]) -> dict:
        """構造分解PL: 販管費の人件費系のみ抽出（原価部は無視）"""
        try:
            text = self._call_api(images, PROMPT_PL_PL_SECTION)
            d = self._ensure_dict(self._parse_json(text), 'extract_pl_pl_section')
            logger.warning(
                f'[extract_pl_pl_section] 販管費 salary={d.get("salary", 0):,} '
                f'misc_wages={d.get("misc_wages", 0):,} bonus={d.get("bonus", 0):,} '
                f'officer={d.get("officer_compensation", 0):,}'
            )
            return d
        except APICreditExhaustedError:
            raise
        except Exception as e:
            self._extraction_errors['pl_section'] = self._format_api_error(e)
            logger.warning(f'[extract_pl_pl_section] 失敗: {e}')
            return {}

    def _extract_pl_cost_section(self, images: list[bytes]) -> dict:
        """構造分解PL: 原価部（製造原価/完成工事原価/工事原価/売上原価/役務原価）の人件費系のみ抽出"""
        try:
            text = self._call_api(images, PROMPT_PL_COST_SECTION)
            d = self._ensure_dict(self._parse_json(text), 'extract_pl_cost_section')
            logger.warning(
                f'[extract_pl_cost_section] 原価部 salary={d.get("salary", 0):,} '
                f'misc_wages={d.get("misc_wages", 0):,} bonus={d.get("bonus", 0):,}'
            )
            return d
        except APICreditExhaustedError:
            raise
        except Exception as e:
            self._extraction_errors['cost_section'] = self._format_api_error(e)
            logger.warning(f'[extract_pl_cost_section] 失敗: {e}')
            return {}

    def _extract_pl_cost_report_only(self, images: list[bytes]) -> dict:
        """原価報告書（製造原価/完成工事原価/工事原価/売上原価/役務原価）のみフォーカスで再抽出。

        販管費は無視し、原価部の人件費・減価償却費だけ取る。
        異常検知 (_is_pl_extraction_suspicious) で True になった時だけ呼ばれるため、
        全案件のコストを増やさず、必要な案件だけ +1回 API。
        APICreditExhaustedError は再 raise（残高切れで他の呼出も無駄なため）。
        """
        try:
            text = self._call_api(images, PROMPT_PL_COST_REPORT_FOCUS)
            d = self._ensure_dict(self._parse_json(text), 'extract_pl_cost_report')
            logger.warning(
                f'[extract_pl_cost_report] 原価部抽出: '
                f'salary={d.get("salary", 0):,} misc_wages={d.get("misc_wages", 0):,} '
                f'bonus={d.get("bonus", 0):,} depreciation={d.get("depreciation", 0):,}'
            )
            return d
        except APICreditExhaustedError:
            raise
        except Exception as e:
            logger.warning(f'[extract_pl_cost_report] 失敗: {e} → 原価部=0で進める')
            return {}

    def _extract_with_pre_split(
        self,
        tsv_data: str,
        fiscal_period_hint: str | None,
        pdf_files: list[tuple[str, bytes]],
    ) -> list[dict]:
        """大型PDFを最初から半分に分割して送信し、月単位補完マージで統合。

        各チャンクは _retry_depth=1 で呼ばれるため、事後分割は発動しない。
        → 1案件あたり API 呼出は **最大 2回** （前半 + 後半）に制限される。
        """
        split_part1, split_part2 = _split_pdfs_in_half(pdf_files)

        chunks: list[list[dict]] = []
        for label, part_pdfs, with_tsv in [
            ('part1', split_part1, True),   # TSV は最初の chunk のみ送信
            ('part2', split_part2, False),
        ]:
            if not part_pdfs:
                continue
            try:
                partial = self.extract_wage_ledger(
                    tsv_data=tsv_data if with_tsv else '',
                    fiscal_period_hint=fiscal_period_hint,
                    pdf_files=part_pdfs,
                    _retry_depth=1,  # 事後分割を禁止（無限再帰防止 + コスト上限）
                )
                logger.info(f'事前分割 {label}: {len(partial)}名')
                chunks.append(partial)
            except Exception as e:
                logger.warning(f'事前分割 {label} 失敗（{type(e).__name__}: {e}）→ skip')

        merged = _merge_wage_employees_by_month(chunks)
        logger.warning(
            f'[extract_wage_ledger] 事前分割マージ完了: '
            f'チャンク{len(chunks)}個 → 統合後{len(merged)}名'
        )
        return merged

    def _retry_extract_with_split_pdf(
        self,
        tsv_data: str,
        fiscal_period_hint: str | None,
        pdf_files: list[tuple[str, bytes]],
        partial_data: list[dict],
    ) -> list[dict]:
        """事後分割: max_tokens 打ち切り検出時のみ呼ばれる。

        - PDF を PyMuPDF で前半・後半に分割（事前分割と同じヘルパ使用）
        - 各分割を _retry_depth=1 で再 API 呼出
        - partial_data + retry_results を月単位補完マージで統合
          （同一人物の月データが別チャンクに分かれた場合の欠落を防ぐ）
        - 失敗時は partial_data をそのまま返す
        """
        split_part1, split_part2 = _split_pdfs_in_half(pdf_files)

        retry_chunks: list[list[dict]] = []
        for label, split_pdfs in [('part1', split_part1), ('part2', split_part2)]:
            if not split_pdfs:
                continue
            try:
                partial = self.extract_wage_ledger(
                    tsv_data='',  # TSV は初回呼出で送信済み、分割では PDF のみ
                    fiscal_period_hint=fiscal_period_hint,
                    pdf_files=split_pdfs,
                    _retry_depth=1,
                )
                logger.info(f'分割再抽出 {label}: {len(partial)}名')
                retry_chunks.append(partial)
            except Exception as e:
                logger.warning(f'分割再抽出 {label} 失敗: {e}')

        # 月単位補完マージ: partial_data + retry_chunks を統合
        all_chunks = ([partial_data] if partial_data else []) + retry_chunks
        merged = _merge_wage_employees_by_month(all_chunks)
        logger.warning(
            f'[extract_wage_ledger] 事後分割マージ完了: '
            f'初回{len(partial_data or [])}名 + 再抽出{sum(len(c) for c in retry_chunks)}名 '
            f'→ 月単位補完マージ後{len(merged)}名'
        )
        return merged


def create_extractor(
    api_key: str = '',
    retry_callback: Optional[RetryCallback] = None,
    model: str | None = None,
) -> BaseExtractor:
    """APIキーの有無に応じて適切なExtractorを返す。

    model を省略した場合は config.CLAUDE_MODEL（環境変数 CLAUDE_MODEL で
    上書き可能）を使う。Cloud Secrets / .env でモデル切替できるようにする。
    """
    if api_key:
        from .config import CLAUDE_MODEL
        selected_model = model or CLAUDE_MODEL
        logger.info(f'Claude API Extractor を使用 (model={selected_model})')
        return ClaudeExtractor(
            api_key,
            model=selected_model,
            retry_callback=retry_callback,
        )
    else:
        logger.warning('APIキー未設定 → StubExtractor を使用（PDF読取不可）')
        return StubExtractor()
