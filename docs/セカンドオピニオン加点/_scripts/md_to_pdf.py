"""配布用 Markdown → PDF 変換（営業ヒアリングシート用）。

使い方:
    python md_to_pdf.py <input1.md> [<input2.md> ...]
各 .md と同じ場所に同名 .pdf を出力する。

方針:
- markdown ライブラリで HTML 化（extra / sane_lists / nl2br）。
- 章（## 見出し）ごとに <section class="keep"> で包み、CSS break-inside:avoid で
  「見出しと選択肢が改ページで割れない」ようにする。
- 日本語フォントは Windows 標準の Yu Gothic。Chrome ヘッドレスの --print-to-pdf で印刷。
"""
import sys
import shutil
import subprocess
import tempfile
import pathlib

sys.stdout.reconfigure(encoding="utf-8")

import markdown

CHROME = r"C:\Program Files\Google\Chrome\Application\chrome.exe"

CSS = """
@page { size: A4; margin: 14mm 13mm; }
* { box-sizing: border-box; }
body {
  font-family: "Yu Gothic", "Meiryo", "MS PGothic", sans-serif;
  font-size: 10.5pt; line-height: 1.55; color: #1a1a1a; margin: 0;
}
h1 { font-size: 15pt; margin: 0 0 6px; }
h2 {
  font-size: 12pt; margin: 14px 0 6px; padding: 3px 8px;
  background: #eef1f4; border-left: 4px solid #5b6b7a; break-after: avoid;
}
p { margin: 5px 0; }
ul { list-style: none; padding-left: 0; margin: 3px 0 6px; }
li { margin: 2px 0; line-height: 1.5; }
hr { display: none; }
section.keep { break-inside: avoid; }
section.header { margin-bottom: 4px; }
section.header p:first-of-type { color: #555; font-size: 9.5pt; }
strong { font-weight: 700; }
"""


def md_to_html_body(md_text: str) -> str:
    lines = md_text.replace("\r\n", "\n").split("\n")
    header, sections, cur = [], [], None
    for ln in lines:
        if ln.startswith("## "):
            if cur is not None:
                sections.append(cur)
            cur = [ln]
        elif cur is None:
            header.append(ln)
        else:
            cur.append(ln)
    if cur is not None:
        sections.append(cur)

    md = markdown.Markdown(extensions=["extra", "sane_lists", "nl2br"])

    def conv(chunk_lines):
        md.reset()
        return md.convert("\n".join(chunk_lines).strip())

    parts = [f'<section class="keep header">{conv(header)}</section>']
    for s in sections:
        parts.append(f'<section class="keep">{conv(s)}</section>')
    return "\n".join(parts)


def make_pdf(md_path: pathlib.Path) -> pathlib.Path:
    body = md_to_html_body(md_path.read_text(encoding="utf-8"))
    html = (
        "<!doctype html><html lang='ja'><head><meta charset='utf-8'>"
        f"<style>{CSS}</style></head><body>{body}</body></html>"
    )
    pdf_path = md_path.with_suffix(".pdf")
    with tempfile.NamedTemporaryFile(
        "w", suffix=".html", delete=False, encoding="utf-8"
    ) as f:
        f.write(html)
        html_path = pathlib.Path(f.name)
    # Chrome は日本語パスへの --print-to-pdf 出力に失敗することがあるため、
    # ASCII の一時パスへ出力してから Python で最終パスへ移動する。
    tmp_pdf = pathlib.Path(tempfile.gettempdir()) / f"_md2pdf_{abs(hash(md_path.name))}.pdf"
    tmp_pdf.unlink(missing_ok=True)
    with tempfile.TemporaryDirectory() as udir:
        res = subprocess.run(
            [
                CHROME,
                "--headless=new",
                "--disable-gpu",
                "--no-sandbox",
                "--no-pdf-header-footer",
                "--run-all-compositor-stages-before-draw",
                "--virtual-time-budget=4000",
                f"--user-data-dir={udir}",
                f"--print-to-pdf={tmp_pdf}",
                html_path.as_uri(),
            ],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
        )
    html_path.unlink(missing_ok=True)
    if not tmp_pdf.exists():
        raise RuntimeError(
            f"Chrome が PDF を生成しませんでした (rc={res.returncode})\n"
            f"--- stderr ---\n{res.stderr}\n--- stdout ---\n{res.stdout}"
        )
    pdf_path.unlink(missing_ok=True)
    shutil.move(str(tmp_pdf), str(pdf_path))
    return pdf_path


def main():
    if len(sys.argv) < 2:
        print("input .md を指定してください")
        sys.exit(1)
    for arg in sys.argv[1:]:
        p = pathlib.Path(arg)
        out = make_pdf(p)
        print(f"OK: {out.name}  ({out.stat().st_size:,} bytes)")


if __name__ == "__main__":
    main()
