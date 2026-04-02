# -*- coding: utf-8 -*-
"""PDF → 画像変換"""
from __future__ import annotations

import io
import base64
import logging
from pathlib import Path

logger = logging.getLogger(__name__)


MAX_IMAGE_SIZE = 4_500_000  # API制限5MBに余裕を持たせる


def pdf_to_images(pdf_path: Path, dpi: int = 150) -> list[bytes]:
    """PDFの各ページをPNG画像のbytesリストに変換"""
    try:
        import fitz  # PyMuPDF
    except ImportError:
        raise ImportError(
            'PyMuPDF がインストールされていません。\n'
            'pip install PyMuPDF でインストールしてください。'
        )

    doc = fitz.open(str(pdf_path))
    images = []

    for page_num in range(len(doc)):
        page = doc[page_num]
        # dpi指定でラスタライズ
        zoom = dpi / 72
        mat = fitz.Matrix(zoom, zoom)
        pix = page.get_pixmap(matrix=mat)
        img_bytes = pix.tobytes('png')

        # 5MB制限対策: 超えたらDPIを下げてリトライ
        retry_dpi = dpi
        while len(img_bytes) > MAX_IMAGE_SIZE and retry_dpi > 72:
            retry_dpi -= 25
            zoom = retry_dpi / 72
            mat = fitz.Matrix(zoom, zoom)
            pix = page.get_pixmap(matrix=mat)
            img_bytes = pix.tobytes('png')
            logger.debug(f'{pdf_path.name} p{page_num + 1}: DPI{retry_dpi}でリサイズ ({len(img_bytes):,} bytes)')

        images.append(img_bytes)
        logger.debug(f'{pdf_path.name} ページ{page_num + 1}: {len(img_bytes):,} bytes')

    doc.close()
    logger.info(f'{pdf_path.name}: {len(images)}ページ変換完了')
    return images


def images_to_base64(images: list[bytes]) -> list[str]:
    """画像bytesリストをbase64文字列リストに変換（API送信用）"""
    return [base64.standard_b64encode(img).decode('ascii') for img in images]


def get_pdf_page_count(pdf_path: Path) -> int:
    """PDFのページ数を取得"""
    try:
        import fitz
        doc = fitz.open(str(pdf_path))
        count = len(doc)
        doc.close()
        return count
    except ImportError:
        logger.warning('PyMuPDF未インストール。ページ数不明。')
        return -1
