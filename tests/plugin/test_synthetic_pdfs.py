"""実機確認用の合成 PDF が期待どおりに作れるか（テキスト版は文字が読め、画像版は文字の層が無い）。"""
import json
import tempfile
from pathlib import Path

import fitz

from _common import Checker

import make_synthetic_pdfs as mk


def main() -> int:
    c = Checker('合成 PDF（実機確認用）')
    with tempfile.TemporaryDirectory() as tmp:
        paths = mk.make(Path(tmp))
        with fitz.open(paths['text_pdf']) as doc:
            text = ''.join(p.get_text() for p in doc)
            c.check('テキスト版は3名×1ページ', doc.page_count == len(mk.EMPLOYEES))
        c.check('テキスト版から氏名と金額が読める', '甲野花子' in text and '210,000' in text, text[:200])
        with fitz.open(paths['image_pdf']) as doc:
            c.check('画像版は指定ページ数', doc.page_count == mk.IMAGE_PAGES)
            c.check('画像版には文字の層が無い', not ''.join(p.get_text() for p in doc).strip())
        expected = json.loads((Path(tmp) / '期待値.json').read_text(encoding='utf-8'))
        c.check('期待値に3名分の年額がある', [e['annual_taxable'] for e in expected['employees']]
                == [(b + cm) * 12 for _n, _k, b, cm, _d, _h in mk.EMPLOYEES])
    return c.finish()


if __name__ == '__main__':
    raise SystemExit(main())
