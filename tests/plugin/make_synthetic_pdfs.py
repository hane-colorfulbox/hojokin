"""実機確認（チャット・Cowork・Claude Code）用の合成の賃金台帳 PDF を作る。実在の人・会社は使わない。

  - 合成賃金台帳_テキスト版.pdf : 架空の会社・3名×12か月（1人1ページ・文字を選択できる PDF）
  - 合成賃金台帳_画像版.pdf     : 上の1〜2ページ目を画像にしたもの（文字の層が無い＝画像 PDF の確認用）
  - 期待値.json                 : 変換結果の答え合わせ用（従業員ごとの月額・通勤手当・年額）

使い方: python -B tests/plugin/make_synthetic_pdfs.py <出力フォルダ>
依存: PyMuPDF（fitz）。日本語は PyMuPDF 内蔵の CJK フォントを PDF に埋め込んで描く
（埋め込まないと、日本語フォントを持たない表示環境＝チャットの PDF 画像化などで文字が消える）。
"""
import json
import sys
from pathlib import Path

import fitz

for _stream in (sys.stdout, sys.stderr):
    if hasattr(_stream, 'reconfigure'):
        _stream.reconfigure(encoding='utf-8')

COMPANY = '株式会社サンプル商会（架空）'
PERIOD = '令和7年1月〜令和7年12月（12月決算）'
MONTHS = [f'{m}月' for m in range(1, 13)]
# (氏名, 雇用形態, 基本給, 通勤手当, 出勤日数, 労働時間)。通勤手当は課税支給額に含める
EMPLOYEES = [
    ('甲野花子', '正社員', 200000, 10000, 20, 160),
    ('乙野太郎', 'パート', 88000, 0, 16, 80),
    ('丙野次郎', '正社員', 250000, 12000, 21, 168),
]
FONT = 'jpfont'                    # ページに埋め込むフォントの名前（中身は PyMuPDF 内蔵の CJK フォント）
FONT_BUFFER = fitz.Font('cjk').buffer
PAGE_W, PAGE_H = 842, 595          # A4 横
LEFT, TOP, ROW_H = 30, 110, 26
LABEL_W, CELL_W = 90, 55
IMAGE_DPI = 110
IMAGE_PAGES = 2
JPEG_QUALITY = 80                  # 画像版を小さく保つ（添付の上限 30MB を気にしなくて済む大きさ）


def _row_values(base, commute, days, hours):
    taxable = base + commute
    return [('基本給', base), ('通勤手当', commute), ('課税支給額', taxable), ('出勤日数', days), ('労働時間', hours)]


def _draw_page(page, emp) -> None:
    name, kind, base, commute, days, hours = emp
    page.insert_font(fontname=FONT, fontbuffer=FONT_BUFFER)
    page.insert_text((LEFT, 40), f'賃金台帳　{COMPANY}', fontname=FONT, fontsize=14)
    page.insert_text((LEFT, 62), f'対象期間：{PERIOD}', fontname=FONT, fontsize=10)
    page.insert_text((LEFT, 84), f'氏名：{name}　雇用形態：{kind}', fontname=FONT, fontsize=11)
    heads = ['項目'] + MONTHS + ['合計']
    for i, h in enumerate(heads):
        x = LEFT + (0 if i == 0 else LABEL_W + (i - 1) * CELL_W)
        page.insert_text((x + 2, TOP), h, fontname=FONT, fontsize=9)
    for r, (label, value) in enumerate(_row_values(base, commute, days, hours), 1):
        y = TOP + r * ROW_H
        page.insert_text((LEFT + 2, y), label, fontname=FONT, fontsize=9)
        for m in range(12):
            page.insert_text((LEFT + LABEL_W + m * CELL_W + 2, y), f'{value:,}', fontname=FONT, fontsize=8)
        page.insert_text((LEFT + LABEL_W + 12 * CELL_W + 2, y), f'{value * 12:,}', fontname=FONT, fontsize=8)
    right = LEFT + LABEL_W + 13 * CELL_W
    for r in range(0, 7):
        y = TOP - 16 + r * ROW_H
        page.draw_line((LEFT, y), (right, y))


def make(out_dir: Path) -> dict:
    out_dir.mkdir(parents=True, exist_ok=True)
    text_pdf = out_dir / '合成賃金台帳_テキスト版.pdf'
    doc = fitz.open()
    for emp in EMPLOYEES:
        _draw_page(doc.new_page(width=PAGE_W, height=PAGE_H), emp)
    doc.subset_fonts()                      # 使っている文字だけを埋め込む（フォント全体だと数MBになる）
    doc.save(text_pdf, garbage=3, deflate=True)
    image_pdf = out_dir / '合成賃金台帳_画像版.pdf'
    img_doc = fitz.open()
    for i in range(IMAGE_PAGES):
        pix = doc[i].get_pixmap(dpi=IMAGE_DPI)
        page = img_doc.new_page(width=PAGE_W, height=PAGE_H)
        page.insert_image(page.rect, stream=pix.tobytes('jpeg', jpg_quality=JPEG_QUALITY))
    img_doc.save(image_pdf)
    doc.close()
    img_doc.close()
    expected = {
        'company': COMPANY, 'period': PERIOD, 'fiscal_month': 12,
        'employees': [{'name': n, 'type': k, 'monthly_taxable': b + c, 'monthly_commute': c,
                       'annual_taxable': (b + c) * 12, 'annual_commute': c * 12, 'hours': h}
                      for n, k, b, c, _d, h in EMPLOYEES],
        'image_pdf_pages': IMAGE_PAGES,
    }
    (out_dir / '期待値.json').write_text(json.dumps(expected, ensure_ascii=False, indent=1), encoding='utf-8')
    return {'text_pdf': text_pdf, 'image_pdf': image_pdf}


def main() -> int:
    if len(sys.argv) != 2:
        print(__doc__)
        return 2
    paths = make(Path(sys.argv[1]))
    for p in paths.values():
        print(f'作成: {p}')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
