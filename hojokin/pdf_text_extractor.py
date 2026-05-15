# -*- coding: utf-8 -*-
"""賃金台帳PDFのテキスト前処理（コスト削減）

画像PDFをClaudeに送ると 1ページ ~2000トークン (画像入力単価) 消費する。
テキスト化できれば 1ページ数百トークン (テキスト入力単価) で済むため、
賃金台帳のような大きなPDFで月額コストを大幅に削減できる。

優先順:
  1. Document AI Enterprise OCR (USE_DOCUMENT_AI_OCR=true) — スキャンPDFも OCR で読める
  2. pdfplumber.extract_tables — 表構造を保ったTSV風出力（テキストPDF向け）
  3. pdfplumber.extract_text — 表抽出失敗時のテキストフォールバック
  4. PyMuPDF page.get_text — pdfplumber 完全失敗時の最終フォールバック

呼び出し側は ValueError をキャッチして画像経路にフォールバックする想定。
"""
from __future__ import annotations

import io
import logging
from typing import Optional

logger = logging.getLogger(__name__)


def _is_meaningful_text(text: str) -> bool:
    """ページマーカー / テーブル区切り以外の本文が含まれているか。

    `===== page N =====` や `--- table N ---` の行のみ抽出された場合 (= スキャン画像PDF)
    これらを除外した上で、何らかの本文が残っているかをチェックする。
    """
    for line in text.splitlines():
        s = line.strip()
        if not s:
            continue
        if s.startswith('===== page') or s.startswith('--- table'):
            continue
        return True
    return False


def _table_to_tsv(table: list[list[Optional[str]]]) -> str:
    """pdfplumber の table（list of rows）をTSV文字列に変換。空セルは空文字に。"""
    lines = []
    for row in table:
        cells = []
        for cell in row:
            if cell is None:
                cells.append('')
            else:
                cells.append(str(cell).replace('\n', ' ').replace('\t', ' ').strip())
        lines.append('\t'.join(cells))
    return '\n'.join(lines)


def _extract_with_pdfplumber(pdf_bytes: bytes) -> str:
    """pdfplumber で PDF を表構造保ったテキストに変換。"""
    import pdfplumber

    parts: list[str] = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for i, page in enumerate(pdf.pages, 1):
            parts.append(f'===== page {i} =====')
            tables = page.extract_tables()
            if tables:
                for ti, table in enumerate(tables, 1):
                    parts.append(f'--- table {ti} ---')
                    parts.append(_table_to_tsv(table))
            else:
                text = page.extract_text() or ''
                if text.strip():
                    parts.append(text)
    return '\n'.join(parts)


def _extract_with_pymupdf(pdf_bytes: bytes) -> str:
    """PyMuPDF で PDF からテキストを抽出（pdfplumber 失敗時のフォールバック）。"""
    import fitz

    parts: list[str] = []
    doc = fitz.open(stream=pdf_bytes, filetype='pdf')
    try:
        for i, page in enumerate(doc, 1):
            parts.append(f'===== page {i} =====')
            text = page.get_text() or ''
            if text.strip():
                parts.append(text)
    finally:
        doc.close()
    return '\n'.join(parts)


def _try_document_ai(pdf_bytes: bytes) -> Optional[str]:
    """Document AI を試す。設定不足/失敗時は None を返して次経路に。"""
    from .config import (
        USE_DOCUMENT_AI_OCR, DOCUMENT_AI_PROJECT_ID,
        DOCUMENT_AI_LOCATION, DOCUMENT_AI_PROCESSOR_ID,
    )
    if not USE_DOCUMENT_AI_OCR:
        return None
    if not DOCUMENT_AI_PROJECT_ID or not DOCUMENT_AI_PROCESSOR_ID:
        logger.warning(
            '[pdf_text] USE_DOCUMENT_AI_OCR=true だが PROJECT_ID/PROCESSOR_ID が未設定 → スキップ'
        )
        return None
    try:
        from .document_ai_ocr import extract_pdf_via_document_ai
        text = extract_pdf_via_document_ai(
            pdf_bytes=pdf_bytes,
            project_id=DOCUMENT_AI_PROJECT_ID,
            location=DOCUMENT_AI_LOCATION,
            processor_id=DOCUMENT_AI_PROCESSOR_ID,
        )
        return text if text.strip() else None
    except Exception as e:
        logger.warning(f'[pdf_text] Document AI 失敗: {e} → 次経路にフォールバック')
        return None


def get_pdf_pages_text(pdf_path) -> list[str]:
    """PDF をページ別テキストのリストとして返す（ページ番号特定用）。

    1始まりインデックスではなく 0始まりのリストを返す。i 番目の要素は (i+1) ページ目。
    決算書からの抽出値をテキストで逆引きしてページ番号を機械的に特定する用途に使う。
    AI抽出値の検証（PDF テキストに値が含まれているか）にも兼用できる。

    画像PDF（テキスト層なし）の場合は空文字のリストが返るのでフォールバック側で
    「ページ特定不可」と扱うこと。

    Returns:
        各ページのテキスト（取れない場合は空文字）。PDF 自体が読めない場合は空リスト。
    """
    from pathlib import Path
    pdf_path = Path(pdf_path)
    # pdfplumber 優先（表構造を保ったまま読めるので数値検索が確実）
    try:
        import pdfplumber
        with pdfplumber.open(str(pdf_path)) as pdf:
            return [(page.extract_text() or '') for page in pdf.pages]
    except Exception:
        pass
    # PyMuPDF フォールバック
    try:
        import fitz
        doc = fitz.open(str(pdf_path))
        try:
            return [(page.get_text() or '') for page in doc]
        finally:
            doc.close()
    except Exception:
        return []


def find_value_pages(pages_text: list[str], value: int | float) -> list[int]:
    """整数値が含まれるページ番号（1始まり）のリストを返す。

    カンマ区切り（'80,153,961'）と無区切り（'80153961'）の両方を検索する。
    マイナス値は絶対値で検索（PDFで「△30,694,465」「(30,694,465)」のような
    表記揺れがあるため、符号付き完全一致だと取れない）。

    Returns:
        値が見つかったページ番号のリスト（1始まり、昇順）。1件も無ければ空リスト。
    """
    if not value:
        return []
    abs_value = int(abs(value))
    if abs_value == 0:
        return []
    formatted = f'{abs_value:,}'  # 'カンマ区切り'
    raw = str(abs_value)
    pages = []
    for i, text in enumerate(pages_text, 1):
        if not text:
            continue
        # カンマ区切り版（'80,153,961'）または無区切り版（'80153961'）を含むか
        # 念のためカンマ・空白を除去した正規化テキストでも検索
        normalized_text = text.replace(' ', '').replace(',', '').replace('，', '')
        if formatted in text or raw in normalized_text:
            pages.append(i)
    return pages


def extract_pdf_as_text(pdf_bytes: bytes) -> str:
    """PDFのバイト列を表構造保ったテキストに変換（後方互換ラッパ）。

    Returns:
        テキスト。空文字は返さない（その場合は ValueError を投げる）。

    Raises:
        ValueError: 全経路（Document AI / pdfplumber / PyMuPDF）が失敗した場合。
            呼び出し側はこれをキャッチして画像経路にフォールバックすべき。
    """
    text, _source = extract_pdf_as_text_with_source(pdf_bytes)
    return text


def extract_pdf_as_text_with_source(pdf_bytes: bytes) -> tuple[str, str]:
    """PDFのバイト列をテキスト化し、どの経路で成功したかも返す。

    Returns:
        (text, source) のタプル。source は以下のいずれか:
          'document_ai'  : Document AI で OCR 成功
          'pdfplumber'   : pdfplumber で抽出成功（テキストPDF）
          'pymupdf'      : PyMuPDF フォールバックで抽出成功

    Raises:
        ValueError: 全経路で失敗した場合。
    """
    if not pdf_bytes:
        raise ValueError('PDFバイト列が空')

    # 1: Document AI（フラグONかつ設定OK時のみ）
    text = _try_document_ai(pdf_bytes)
    if text:
        logger.info(f'[pdf_text] Document AI 成功: {len(text)}chars')
        return text, 'document_ai'

    # 2: pdfplumber（テキストPDF向け）
    try:
        text = _extract_with_pdfplumber(pdf_bytes)
        if _is_meaningful_text(text):
            logger.info(f'[pdf_text] pdfplumber 成功: {len(text)}chars')
            return text, 'pdfplumber'
        logger.warning('[pdf_text] pdfplumber で本文ゼロ → PyMuPDFにフォールバック')
    except Exception as e:
        logger.warning(f'[pdf_text] pdfplumber 失敗: {e} → PyMuPDFにフォールバック')

    # 3: PyMuPDF（最終フォールバック）
    try:
        text = _extract_with_pymupdf(pdf_bytes)
        if _is_meaningful_text(text):
            logger.info(f'[pdf_text] PyMuPDF 成功: {len(text)}chars')
            return text, 'pymupdf'
        raise ValueError('全経路で本文テキスト無し → スキャン画像PDF (要OCR)')
    except Exception as e:
        raise ValueError(f'PDFテキスト化に完全失敗: {e}') from e
