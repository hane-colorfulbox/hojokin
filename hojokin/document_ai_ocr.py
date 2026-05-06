# -*- coding: utf-8 -*-
"""Google Document AI (Enterprise Document OCR) で賃金台帳PDFをテキスト化

スキャン画像PDF (テキストレイヤー無し) でも OCR でテキスト化できる。
pdfplumber/PyMuPDF が空文字を返す場合のフォールバック先として使う。

Cost: $0.0015/ページ (Enterprise OCR、2026年公式価格)
Region: us (Enterprise OCR の安定対応リージョン)

呼び出し側は ValueError をキャッチして次のフォールバック (画像経路など) に流す想定。
"""
from __future__ import annotations

import logging
import os
from pathlib import Path
from typing import Optional

logger = logging.getLogger(__name__)

# Document AI 同期API のサイズ上限（公式 20MB、安全マージン込で 18MB）
DOCUMENT_AI_MAX_BYTES = 18 * 1024 * 1024
DOCUMENT_AI_MAX_PAGES = 30  # 同期処理上限（Enterprise OCR）


def _get_credentials():
    """Service Account 認証情報を構築。GOOGLE_SERVICE_ACCOUNT_JSON を優先参照。"""
    from google.oauth2 import service_account

    sa_path = os.getenv('GOOGLE_SERVICE_ACCOUNT_JSON', 'credentials/service_account.json')
    if not Path(sa_path).exists():
        raise FileNotFoundError(
            f'Service Account JSON が見つかりません: {sa_path}\n'
            f'.env の GOOGLE_SERVICE_ACCOUNT_JSON を確認してください'
        )
    return service_account.Credentials.from_service_account_file(sa_path)


def extract_pdf_via_document_ai(
    pdf_bytes: bytes,
    project_id: str,
    location: str,
    processor_id: str,
) -> str:
    """Document AI (Enterprise Document OCR) で PDF をテキスト化して返す。

    Args:
        pdf_bytes: PDFのバイト列
        project_id: GCP プロジェクトID（例 'hojokin-automation'）
        location: Processor のリージョン（例 'us'）
        processor_id: Processor のID（例 'a9bacc9ec399367b'）

    Returns:
        OCR で抽出されたテキスト全文（ページ区切りは含まない、Document AI が自然な
        順序で結合したもの）。

    Raises:
        ValueError: 設定不足、サイズ超過、API失敗、空テキストのいずれか。
            呼び出し側はこれをキャッチしてフォールバックすべき。
    """
    if not pdf_bytes:
        raise ValueError('PDFバイト列が空')
    if not project_id or not processor_id:
        raise ValueError(
            f'Document AI 設定不足 (project_id={project_id!r}, processor_id={processor_id!r}). '
            f'.env の DOCUMENT_AI_PROJECT_ID / DOCUMENT_AI_PROCESSOR_ID を確認'
        )
    if len(pdf_bytes) > DOCUMENT_AI_MAX_BYTES:
        raise ValueError(
            f'PDFサイズ {len(pdf_bytes)/1_000_000:.1f}MB が同期API上限 '
            f'{DOCUMENT_AI_MAX_BYTES/1_000_000:.0f}MB を超えています'
        )

    from google.cloud import documentai
    from google.api_core.client_options import ClientOptions
    from google.api_core import exceptions as gax_exceptions

    credentials = _get_credentials()
    opts = ClientOptions(api_endpoint=f'{location}-documentai.googleapis.com')
    client = documentai.DocumentProcessorServiceClient(
        credentials=credentials, client_options=opts
    )
    name = client.processor_path(project_id, location, processor_id)

    raw_document = documentai.RawDocument(
        content=pdf_bytes, mime_type='application/pdf'
    )
    # imageless_mode=True で同期API のページ上限を 15→30 に拡張。
    # 画像レイヤーの返却を省く代わりにテキストOCR のみ取得（賃金台帳の用途には十分）。
    request = documentai.ProcessRequest(
        name=name,
        raw_document=raw_document,
        imageless_mode=True,
    )

    try:
        result = client.process_document(request=request)
    except gax_exceptions.InvalidArgument as e:
        raise ValueError(f'Document AI InvalidArgument: {e}') from e
    except gax_exceptions.PermissionDenied as e:
        raise ValueError(f'Document AI PermissionDenied: {e}') from e
    except Exception as e:
        raise ValueError(f'Document AI 呼出失敗: {type(e).__name__}: {e}') from e

    text = result.document.text or ''
    pages = len(result.document.pages) if result.document.pages else 0
    if not text.strip():
        raise ValueError(f'Document AI が空テキスト返却 (pages={pages})')

    logger.warning(
        f'[document_ai] OCR成功: pages={pages}, text={len(text)}chars '
        f'(概算コスト: ${pages * 0.0015:.4f})'
    )
    return text
