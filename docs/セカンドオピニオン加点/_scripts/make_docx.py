"""Word（.docx）生成 — 営業ヒアリングシート(05)の編集可能版。

使い方:
    python make_docx.py <05_面談ヒアリングシート.md>                汎用版（設問は hearing_form.FORM）
    python make_docx.py --from-md <05_..._会社別.md> [-o 出力先]     会社別版（設問は md をパース）
入力 md と同じ場所（-o 指定時はそこ）に `<basename>.docx` を出力する。

設計:
- □ は Word のチェックボックス コンテンツコントロール（w14:checkbox）にする。Word 上で
  クリックすれば ☒ に切り替わる。w14 は mc:Ignorable 宣言済みのため、対応しないビューアでも
  ☐ の文字として読める（＝壊れない）。
- 記入欄は下線付きの空白ラン。クリックしてそのまま打てて、印刷しても記入線として機能する。
- フォントは游ゴシック（eastAsia も明示。指定しないと日本語が MS 明朝系にフォールバックする）。
- レイアウトの狙いは PDF 版（md_to_pdf.py）と同じ見え方。見出しは薄い塗り＋左の色帯。
"""
import sys
import argparse
import pathlib

sys.stdout.reconfigure(encoding="utf-8")

import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import qn
from docx.shared import Mm, Pt, RGBColor

from hearing_form import FORM
from md_form_parser import parse_md

# ---- レイアウト定数 ----
FONT_JA = "游ゴシック"
FONT_BOX = "MS Gothic"           # ☐/☒ が確実に出る等幅フォント
PAGE_MARGIN_TB, PAGE_MARGIN_LR = Mm(14), Mm(13)
SIZE_TITLE, SIZE_BODY, SIZE_NOTE, SIZE_SECTION, SIZE_SUB = 15, 10.5, 9, 12, 10.5
COLOR_NOTE = RGBColor(0x55, 0x55, 0x55)
COLOR_SECTION = RGBColor(0x1F, 0x3A, 0x52)
FILL_SECTION = "EEF1F4"
BAR_SECTION = "5B6B7A"
FIELD_CHAR = "　"            # 記入欄に敷く全角スペース
FIELD_PT_PER_CHAR = 15           # md_form_parser の幅(pt)を全角スペース数に戻す係数
FIELD_MIN_CHARS, FIELD_MAX_CHARS = 2, 22
SPACE_BEFORE_SECTION, SPACE_AFTER_SECTION = Pt(10), Pt(4)
INSTR_DOCX = (
    "□ をクリックすると ☒ に切り替わります（Word で開いた場合）。"
    "記入欄は下線部にそのまま入力してください。"
)

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml"

CHECKBOX_XML = f"""<w:sdt xmlns:w="{W_NS}" xmlns:w14="{W14_NS}">
  <w:sdtPr>
    <w:rPr><w:rFonts w:ascii="{FONT_BOX}" w:hAnsi="{FONT_BOX}" w:eastAsia="{FONT_BOX}"/></w:rPr>
    <w:id w:val="%(id)d"/>
    <w14:checkbox>
      <w14:checked w14:val="0"/>
      <w14:checkedState w14:val="2612" w14:font="{FONT_BOX}"/>
      <w14:uncheckedState w14:val="2610" w14:font="{FONT_BOX}"/>
    </w14:checkbox>
  </w:sdtPr>
  <w:sdtContent>
    <w:r><w:rPr><w:rFonts w:ascii="{FONT_BOX}" w:hAnsi="{FONT_BOX}" w:eastAsia="{FONT_BOX}"/>
      <w:sz w:val="%(sz)d"/><w:szCs w:val="%(sz)d"/></w:rPr><w:t>☐</w:t></w:r>
  </w:sdtContent>
</w:sdt>"""


def _set_ja_font(run, name=FONT_JA, size=SIZE_BODY, bold=False, color=None, underline=False):
    run.font.name = name
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.underline = underline
    if color is not None:
        run.font.color.rgb = color
    run._element.rPr.rFonts.set(qn("w:eastAsia"), name)


def _shade(paragraph, fill, bar=None):
    """段落に背景色（と左の色帯）を付ける。"""
    pPr = paragraph._p.get_or_add_pPr()
    pPr.append(parse_xml(
        f'<w:shd xmlns:w="{W_NS}" w:val="clear" w:color="auto" w:fill="{fill}"/>'))
    if bar:
        pPr.append(parse_xml(
            f'<w:pBdr xmlns:w="{W_NS}">'
            f'<w:left w:val="single" w:sz="24" w:space="4" w:color="{bar}"/></w:pBdr>'))


class Doc:
    def __init__(self, form=None):
        self.form = form if form is not None else FORM
        self.d = docx.Document()
        self.sdt_id = 1000
        self._setup()

    def _setup(self):
        sec = self.d.sections[0]
        sec.page_width, sec.page_height = Mm(210), Mm(297)
        sec.top_margin = sec.bottom_margin = PAGE_MARGIN_TB
        sec.left_margin = sec.right_margin = PAGE_MARGIN_LR
        style = self.d.styles["Normal"]
        style.font.name = FONT_JA
        style.font.size = Pt(SIZE_BODY)
        style.element.rPr.rFonts.set(qn("w:eastAsia"), FONT_JA)
        pf = style.paragraph_format
        pf.space_before, pf.space_after = Pt(0), Pt(3)
        pf.line_spacing = 1.15

    def _para(self, align=WD_ALIGN_PARAGRAPH.LEFT):
        p = self.d.add_paragraph()
        p.alignment = align
        return p

    def title(self, text):
        p = self._para()
        _set_ja_font(p.add_run(text), size=SIZE_TITLE, bold=True)

    def note(self, text):
        p = self._para()
        _set_ja_font(p.add_run(text), size=SIZE_NOTE, color=COLOR_NOTE)

    def section(self, text):
        p = self._para()
        p.paragraph_format.space_before = SPACE_BEFORE_SECTION
        p.paragraph_format.space_after = SPACE_AFTER_SECTION
        _set_ja_font(p.add_run(text), size=SIZE_SECTION, bold=True, color=COLOR_SECTION)
        _shade(p, FILL_SECTION, BAR_SECTION)

    def subsection(self, text):
        p = self._para()
        p.paragraph_format.space_before = Pt(6)
        _set_ja_font(p.add_run(text), size=SIZE_SUB, bold=True, color=COLOR_SECTION)

    def _checkbox(self, paragraph):
        self.sdt_id += 1
        paragraph._p.append(parse_xml(
            CHECKBOX_XML % {"id": self.sdt_id, "sz": int(SIZE_BODY * 2)}))

    def row(self, segments):
        p = self._para()
        for seg in segments:
            kind = seg[0]
            if kind == "t":
                _set_ja_font(p.add_run(seg[1]))
            elif kind == "c":
                self._checkbox(p)
                _set_ja_font(p.add_run(seg[1] + " "))
            elif kind == "f":
                n = min(FIELD_MAX_CHARS, max(FIELD_MIN_CHARS, round(seg[1] / FIELD_PT_PER_CHAR)))
                _set_ja_font(p.add_run(FIELD_CHAR * n), underline=True)
                _set_ja_font(p.add_run(" "))

    def build(self, with_instruction=True):
        emitted_instruction = not with_instruction
        for block in self.form:
            kind = block[0]
            if kind == "title":
                self.title(block[1])
            elif kind == "note":
                self.note(block[1])
            elif kind == "section":
                self.section(block[1])
            elif kind == "subsection":
                self.subsection(block[1])
            elif kind == "row":
                self.row(block[1])
            elif kind == "space":
                self.d.add_paragraph()
            if not emitted_instruction and kind == "title":
                self.note(INSTR_DOCX)
                emitted_instruction = True
        return self

    def save(self, path):
        self.d.save(str(path))


def main():
    ap = argparse.ArgumentParser(description="Word版ヒアリングシートを作る")
    ap.add_argument("md", help="入力 .md（出力名の決定に使う。--from-md 指定時は設問の取得元にもなる）")
    ap.add_argument("--from-md", action="store_true",
                    help="設問を md からパースする（会社別シート用）。既定は hearing_form.FORM")
    ap.add_argument("-o", "--outdir", default=None, help="出力先フォルダ（既定: md と同じ場所）")
    args = ap.parse_args()

    md = pathlib.Path(args.md)
    form = parse_md(md)["blocks"] if args.from_md else None
    outdir = pathlib.Path(args.outdir) if args.outdir else md.parent
    outdir.mkdir(parents=True, exist_ok=True)
    out = outdir / (md.stem + ".docx")
    Doc(form).build().save(out)
    print(f"OK: {out.name}  ({out.stat().st_size:,} bytes)")


if __name__ == "__main__":
    main()
