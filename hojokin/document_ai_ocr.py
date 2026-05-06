# -*- coding: utf-8 -*-
"""Google Document AI (Enterprise Document OCR) で賃金台帳PDFをテキスト化

スキャン画像PDF (テキストレイヤー無し) でも OCR でテキスト化できる。
pdfplumber/PyMuPDF が空文字を返す場合のフォールバック先として使う。

Cost: $0.0015/ページ (Enterprise OCR、2026年公式価格)
Region: us (Enterprise OCR の安定対応リージョン)

呼び出し側は ValueError をキャッチして次のフォールバック (画像経路など) に流す想定。
"""
from __future__ import annotations

import json
import logging
import os
from pathlib import Path
from typing import Optional

logger = logging.getLogger(__name__)

# Document AI 同期API のサイズ上限（公式 20MB、安全マージン込で 18MB）
DOCUMENT_AI_MAX_BYTES = 18 * 1024 * 1024
DOCUMENT_AI_MAX_PAGES = 30  # 同期処理上限（Enterprise OCR）
# 再帰分割の保険上限（これを超えるとバッチAPI 利用を案内するエラー）
DOCUMENT_AI_RECURSION_PAGE_CEILING = 480
# 分割チャンク間の境界マーカー（下流プロンプトに「ここでチャンクが切れた」と知らせる）
CHUNK_BOUNDARY_MARKER = '\n\n----- pdf chunk boundary -----\n\n'


def _get_credentials():
    """Service Account 認証情報を構築。

    優先順:
    1. 環境変数 GOOGLE_SERVICE_ACCOUNT_JSON_CONTENT に JSON 文字列が直接入っていればそれを使う。
       Streamlit Cloud / Cloud Run など、Service Account JSON ファイルを配置できない環境向け。
       app.py 側で st.secrets['gcp_service_account'] (TOML テーブル) を JSON 文字列化して
       この環境変数に橋渡しする想定。
    2. 環境変数 GOOGLE_SERVICE_ACCOUNT_JSON にファイルパスが指定されていてファイルが存在すれば、
       それを使う。ローカル開発向け（デフォルト credentials/service_account.json）。
    """
    from google.oauth2 import service_account

    sa_content = os.getenv('GOOGLE_SERVICE_ACCOUNT_JSON_CONTENT', '').strip()
    if sa_content:
        try:
            info = json.loads(sa_content)
        except json.JSONDecodeError as e:
            raise ValueError(
                f'GOOGLE_SERVICE_ACCOUNT_JSON_CONTENT が不正な JSON: {e}'
            ) from e
        return service_account.Credentials.from_service_account_info(info)

    sa_path = os.getenv('GOOGLE_SERVICE_ACCOUNT_JSON', 'credentials/service_account.json')
    if not Path(sa_path).exists():
        raise FileNotFoundError(
            f'Service Account JSON が見つかりません: {sa_path}\n'
            f'ローカル: .env の GOOGLE_SERVICE_ACCOUNT_JSON を確認してください\n'
            f'本番(Streamlit Cloud): Secrets に [gcp_service_account] セクション、'
            f'または GOOGLE_SERVICE_ACCOUNT_JSON_CONTENT を設定してください'
        )
    return service_account.Credentials.from_service_account_file(sa_path)


def _count_pdf_pages(pdf_bytes: bytes) -> int:
    """PyMuPDF で PDF のページ数を取得（失敗時は 0）"""
    try:
        import fitz  # type: ignore
        doc = fitz.open(stream=pdf_bytes, filetype='pdf')
        n = len(doc)
        doc.close()
        return n
    except Exception:
        return 0


def _split_pdf_in_half(pdf_bytes: bytes) -> tuple[bytes, bytes]:
    """PyMuPDF で PDF を前半・後半に分割。

    garbage=4, deflate=True で再シリアライズ時の肥大化を防ぐ
    （元PDFが圧縮済の場合でも、再書き出しで未圧縮になるのを抑制）。

    Returns: (前半 bytes, 後半 bytes)
    """
    import fitz  # type: ignore
    doc = fitz.open(stream=pdf_bytes, filetype='pdf')
    n = len(doc)
    half = n // 2
    doc1 = fitz.open()
    doc1.insert_pdf(doc, from_page=0, to_page=half - 1)
    buf1 = doc1.tobytes(garbage=4, deflate=True)
    doc1.close()
    doc2 = fitz.open()
    doc2.insert_pdf(doc, from_page=half, to_page=n - 1)
    buf2 = doc2.tobytes(garbage=4, deflate=True)
    doc2.close()
    doc.close()
    return buf1, buf2


def extract_pdf_via_document_ai(
    pdf_bytes: bytes,
    project_id: str,
    location: str,
    processor_id: str,
) -> str:
    """Document AI (Enterprise Document OCR) で PDF をテキスト化して返す。

    30ページ超の PDF は同期API 上限を超えるため、再帰的に半分ずつ分割して OCR し、
    結果を連結する。サイズ超過は分割しても解消しないため即エラー。

    Args:
        pdf_bytes: PDFのバイト列
        project_id: GCP プロジェクトID（例 'hojokin-automation'）
        location: Processor のリージョン（例 'us'）
        processor_id: Processor のID（例 'a9bacc9ec399367b'）

    Returns:
        OCR で抽出されたテキスト全文。分割した場合はチャンクをページ順に連結。

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

    # ── ページ分割で解消できない極大PDFは即エラー（同期APIでは非効率） ──
    pages = _count_pdf_pages(pdf_bytes)
    if pages > DOCUMENT_AI_RECURSION_PAGE_CEILING:
        raise ValueError(
            f'PDF {pages} ページが再帰分割上限 {DOCUMENT_AI_RECURSION_PAGE_CEILING} を超過。'
            f'同期API では非効率なため、Document AI バッチAPI（非同期）を検討してください'
        )

    # ── ページ数 or サイズが上限超過 → ページ分割を試す ──
    # ページ単位で分割すれば多くの場合サイズも縮む（PyMuPDF の garbage+deflate 効果あり）。
    # 単ページPDFがサイズ超過しているケースのみ救えない（その時だけエラー）。
    over_pages = pages > DOCUMENT_AI_MAX_PAGES
    over_bytes = len(pdf_bytes) > DOCUMENT_AI_MAX_BYTES
    if over_pages or over_bytes:
        if pages < 2:
            raise ValueError(
                f'PDFサイズ {len(pdf_bytes)/1_000_000:.1f}MB が同期API上限 '
                f'{DOCUMENT_AI_MAX_BYTES/1_000_000:.0f}MB を超過、かつ単ページのため分割不可'
            )
        reason = (
            f'{pages}ページ > 上限{DOCUMENT_AI_MAX_PAGES}'
            if over_pages else
            f'サイズ{len(pdf_bytes)/1_000_000:.1f}MB > 上限{DOCUMENT_AI_MAX_BYTES/1_000_000:.0f}MB'
        )
        logger.warning(f'[document_ai] {reason} → 半分ずつ分割してOCR')
        try:
            buf1, buf2 = _split_pdf_in_half(pdf_bytes)
        except Exception as e:
            raise ValueError(f'PDF分割失敗: {type(e).__name__}: {e}') from e
        text1 = extract_pdf_via_document_ai(buf1, project_id, location, processor_id)
        text2 = extract_pdf_via_document_ai(buf2, project_id, location, processor_id)
        return text1 + CHUNK_BOUNDARY_MARKER + text2

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
    pages_processed = len(result.document.pages) if result.document.pages else 0
    if not text.strip():
        raise ValueError(f'Document AI が空テキスト返却 (pages={pages_processed})')

    logger.warning(
        f'[document_ai] OCR成功: pages={pages_processed}, text={len(text)}chars '
        f'(概算コスト: ${pages_processed * 0.0015:.4f})'
    )
    return text
