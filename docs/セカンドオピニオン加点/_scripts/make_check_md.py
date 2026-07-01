"""チェック可能 Markdown 生成 — 営業ヒアリングシート(05)のオンライン商談・ドキュメント版。

使い方:
    python make_check_md.py <05_面談ヒアリングシート.md>
入力 md と同じ場所に `<basename>_オンライン版.md` を出力する（入力 md は出力名の決定だけに使う）。
設問データは hearing_form.FORM を参照する。

設計（Google ドキュメント互換）:
- 選択肢は GitHub 互換の `- [ ]` チェックリスト。Google ドキュメントに貼り付け、対象行を選んで
  「箇条書き ＞ チェックリスト」を適用すると、クリックでチェックできるリストになる
  （画面共有しながらの読み上げ＋チェックに向く）。
- 記入欄は下線（全角アンダースコア）で示し、その場で入力できる。
- PDF にしたいときは生成した .md を `md_to_pdf.py` に渡す。

⚠️ 同期注意：設問・選択肢は hearing_form.FORM が単一の真実。05.md / FORM を変えたら本スクリプトを
   再実行して .md を作り直すこと（make_fillable_pdf.py / make_check_xlsx.py も同様）。
"""
import sys
import pathlib

sys.stdout.reconfigure(encoding="utf-8")

from hearing_form import FORM, split_check_row, freetext_items

TITLE_TEXT = "面談ヒアリングシート（オンライン商談・入力用）"
INSTR_TEXT = (
    "オンライン商談で画面共有しながら、あてはまる項目にチェックを入れてください（複数選択可）。"
    "下線部（＿＿＿）は直接入力できます。"
    "_Google ドキュメントでは、選択肢の行を選んで「箇条書き ＞ チェックリスト」にすると、"
    "クリックでチェックできます。_"
)
BLANK = "＿＿＿＿＿＿＿＿＿＿"


def fill_line(label, checkbox):
    """記入欄つきの行。ラベル末尾のコロン重複を避ける。"""
    box = "[ ] " if checkbox else ""
    sep = "" if label.endswith(("：", ":")) else "："
    return f"- {box}{label}{sep}{BLANK}"


def render_check_row(segments, out):
    prompt, options = split_check_row(segments)
    if prompt:
        out.append(f"**{prompt}**")
    for kind, label in options:
        if kind == "check":
            out.append(f"- [ ] {label}")
        elif "その他" in label:  # 選択もできる＋内容も書ける
            out.append(fill_line(label, checkbox=True))
        else:  # 氏名・内容などの記入項目
            out.append(fill_line(label, checkbox=False))


def render_freetext_row(segments, out):
    items = freetext_items(segments)
    if items and items[0][0][:1] in "①②③":  # 候補日は1行にまとめる
        label = " ".join(lbl for lbl, _ in items if lbl)
        out.append(fill_line(label, checkbox=False))
        return
    for label, has_input in items:
        if has_input:
            out.append(fill_line(label, checkbox=False))
        else:
            out.append(label)  # ※注記などラベルのみの行


def join_md(lines):
    """連続する箇条書き同士以外は、ブロック間に必ず空行を入れる（markdown のリスト認識のため）。"""
    final = []
    for line in lines:
        if final and not (final[-1].startswith("- ") and line.startswith("- ")):
            final.append("")
        final.append(line)
    return "\n".join(final).strip() + "\n"


def build_md():
    out = [f"# {TITLE_TEXT}", INSTR_TEXT]
    for block in FORM:
        kind = block[0]
        if kind == "title":
            continue  # 既定タイトルは上で出力済み
        if kind == "note":
            out.append(block[1])
        elif kind == "section":
            out.append(f"## {block[1]}")
        elif kind == "row":
            segments = block[1]
            if any(s[0] == "c" for s in segments):
                render_check_row(segments, out)
            elif any(s[0] == "f" for s in segments):
                render_freetext_row(segments, out)
            else:
                out.append(f"**{segments[0][1]}**")
    return join_md(out)


def main():
    if len(sys.argv) < 2:
        print("input .md を指定してください（出力名の決定に使います）")
        sys.exit(1)
    md = pathlib.Path(sys.argv[1])
    out = md.with_name(md.stem + "_オンライン版.md")
    out.write_text(build_md(), encoding="utf-8")
    print(f"OK: {out.name}  ({out.stat().st_size:,} bytes)")


if __name__ == "__main__":
    main()
