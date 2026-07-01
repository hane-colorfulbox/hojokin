"""入力可能（フォーム）PDF 生成 — 営業ヒアリングシート(05)の iPad/PC タップ入力版。

使い方:
    python make_fillable_pdf.py <05_面談ヒアリングシート.md>
入力 md と同じ場所に `<basename>_入力用.pdf` を出力する（md は出力名の決定だけに使う）。

設計:
- reportlab の AcroForm（checkbox / textfield）で、タップ/クリックで入力できる PDF を作る。
- 日本語は reportlab 組込みの CID フォント HeiseiKakuGo-W5（外部TTF不要）。
- フォーム要素は下の FORM 構造で**明示的に保持**する（md の完全パースは行わない）。
  ⚠️ 同期注意：`05_面談ヒアリングシート.md` の設問・選択肢を変えたら、この FORM も合わせて直すこと
  （片方だけ直すと紙版と入力版がズレる）。紙用の静的PDFは `md_to_pdf.py` 側。
"""
import sys
import pathlib

sys.stdout.reconfigure(encoding="utf-8")

from reportlab.lib.pagesizes import A4
from reportlab.lib.colors import black, white, Color
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.acroform import PDFFromString
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.pdfgen import canvas

from hearing_form import FORM  # 設問データ（PDF/xlsx/md 共通の単一の真実）

# 入力欄に打った日本語をビューア側で描画させる（reportlab の AcroForm テキスト欄は
# 標準14フォントしか指定できないため、フォントは Helvetica にして NeedAppearances=true を立てる）。
FIELD_FONT = "Helvetica"

FONT = "HeiseiKakuGo-W5"
pdfmetrics.registerFont(UnicodeCIDFont(FONT))
SIZE = 10
PAGE_W, PAGE_H = A4
L_MARGIN, R_MARGIN, T_MARGIN, B_MARGIN = 40, 40, 52, 40
RIGHT = PAGE_W - R_MARGIN
WRAP_INDENT = L_MARGIN + 12
BOX = 11
LINE_IN = 16     # 行内で折り返したときの送り
ROW_GAP = 20     # 行（設問）の送り
SEC_BEFORE = 12  # セクション見出し前の余白
GAP = 9          # チェック肢どうしの間隔
GREY = Color(0.4, 0.4, 0.4)
BAR_BG = Color(0.93, 0.95, 0.96)
BAR_AC = Color(0.36, 0.42, 0.48)
FIELD_BORDER = Color(0.7, 0.72, 0.75)

class Form:
    def __init__(self, path):
        self.c = canvas.Canvas(str(path), pagesize=A4)
        self.y = PAGE_H - T_MARGIN
        self.x = L_MARGIN
        self.n = 0

    def _name(self, kind):
        self.n += 1
        return f"{kind}{self.n}"

    def _page_break_if_needed(self, need):
        if self.y - need < B_MARGIN:
            self.c.showPage()
            self.y = PAGE_H - T_MARGIN

    def newline(self, dy):
        self.y -= dy
        self._page_break_if_needed(0)

    def width(self, text):
        return pdfmetrics.stringWidth(text, FONT, SIZE)

    def title(self, text):
        self.c.setFillColor(black)
        self.c.setFont(FONT, 15)
        self.c.drawString(L_MARGIN, self.y, text)
        self.newline(ROW_GAP)

    def note(self, text):
        self.c.setFillColor(GREY)
        self.c.setFont(FONT, 9)
        self.c.drawString(L_MARGIN, self.y, text)
        self.c.setFillColor(black)
        self.newline(LINE_IN)

    def section(self, text):
        self.newline(SEC_BEFORE)
        self._page_break_if_needed(24)
        self.c.setFillColor(BAR_BG)
        self.c.rect(L_MARGIN, self.y - 4, RIGHT - L_MARGIN, 16, fill=1, stroke=0)
        self.c.setFillColor(BAR_AC)
        self.c.rect(L_MARGIN, self.y - 4, 4, 16, fill=1, stroke=0)
        self.c.setFillColor(black)
        self.c.setFont(FONT, 11.5)
        self.c.drawString(L_MARGIN + 10, self.y, text)
        self.newline(ROW_GAP)

    def _wrap(self, need):
        if self.x + need > RIGHT and self.x > L_MARGIN:
            self.newline(LINE_IN)
            self.x = WRAP_INDENT

    def row(self, segments):
        self._page_break_if_needed(ROW_GAP)
        self.x = L_MARGIN
        self.c.setFont(FONT, SIZE)
        for seg in segments:
            kind = seg[0]
            if kind == "t":
                text = seg[1]
                w = self.width(text)
                self._wrap(w)
                self.c.setFillColor(black)
                self.c.setFont(FONT, SIZE)
                self.c.drawString(self.x, self.y, text)
                self.x += w + 2
            elif kind == "c":
                label = seg[1]
                lw = self.width(label)
                total = BOX + 3 + lw + GAP
                self._wrap(total)
                self.c.acroForm.checkbox(
                    name=self._name("cb"), x=self.x, y=self.y - 1.5, size=BOX,
                    buttonStyle="check", borderWidth=0.8, borderColor=black,
                    fillColor=white, textColor=black, forceBorder=True,
                    fieldFlags="",  # reportlab既定の'required'を解除（任意チェックのため）
                )
                self.c.setFillColor(black)
                self.c.setFont(FONT, SIZE)
                self.c.drawString(self.x + BOX + 3, self.y, label)
                self.x += total
            elif kind == "f":
                w = seg[1]
                self._wrap(w)
                self.c.acroForm.textfield(
                    name=self._name("tf"), x=self.x, y=self.y - 2.5, width=w, height=14,
                    borderWidth=0.7, borderColor=FIELD_BORDER, fillColor=white,
                    fontName=FIELD_FONT, fontSize=9, forceBorder=True,
                )
                self.x += w + 4
        self.newline(ROW_GAP)

    def build(self):
        for block in FORM:
            kind = block[0]
            if kind == "title":
                self.title(block[1])
            elif kind == "note":
                self.note(block[1])
            elif kind == "section":
                self.section(block[1])
            elif kind == "row":
                self.row(block[1])
            elif kind == "space":
                self.newline(block[1])
        # 入力した日本語をビューアに描画させる
        self.c.acroForm.extras["NeedAppearances"] = PDFFromString("true")
        self.c.save()


def main():
    if len(sys.argv) < 2:
        print("input .md を指定してください（出力名の決定に使います）")
        sys.exit(1)
    md = pathlib.Path(sys.argv[1])
    out = md.with_name(md.stem + "_入力用.pdf")
    Form(out).build()
    print(f"OK: {out.name}  ({out.stat().st_size:,} bytes)")


if __name__ == "__main__":
    main()
