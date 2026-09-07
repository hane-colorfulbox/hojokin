"""会社別 面談ヒアリングシート(05) の配布物を全形式まとめて生成する。

使い方:
    python build_company_sheets.py                    6社すべて
    python build_company_sheets.py --only ピスケス     名前で絞り込み
    python build_company_sheets.py --list             会社とフォルダ名の対応だけ表示

出力: `配布物テンプレ/05_会社別/_配布/<会社名（ツール名）>/` に5形式
    - <シート名>.pdf                  印刷用（md_to_pdf。紙に刷って手書き）
    - <シート名>_入力用.pdf           入力できるPDF（AcroForm。iPad/PCでタップ入力）
    - <シート名>.docx                 Word（□はクリックでチェック・記入欄は下線）
    - <シート名>_オンライン入力用.xlsx  スプレッドシート（未/済のドロップダウン。Googleへ変換して使う）
    - <シート名>.md                   正本のコピー（Drive へ上げると Google ドキュメントに変換される）

このフォルダ構成は Drive の「セカンドオピニオン直下の会社フォルダ」と1対1で対応させている
（Drive へ上げるときに迷わないため）。設問の単一の真実は各社の Markdown。
Markdown を直したら本スクリプトを回し直し、Drive 側も差し替える。
"""
import sys
import shutil
import argparse
import json
import pathlib

sys.stdout.reconfigure(encoding="utf-8")

import md_to_pdf
import make_docx
import make_fillable_pdf
import make_check_xlsx
from md_form_parser import parse_md, verify

SHEET_DIR = pathlib.Path(__file__).resolve().parent.parent / "配布物テンプレ" / "05_会社別"
OUT_ROOT = SHEET_DIR / "_配布"
# 一部の会社は会社フォルダ名をリポジトリに置かず、ローカルの JSON で与える
# （{"配車頭": "会社名（配車頭）"} の形。無い会社は生成・アップロードを飛ばす）
LOCAL_FOLDERS_FILE = pathlib.Path.home() / ".hojokin_pii" / "company_folders.json"

# md のファイル名末尾（ツール名） → Drive/ローカルの会社フォルダ名
# 単位は「会社」。1社が複数ツールを持つ場合があり、送客されるまでツールが確定しないため
# （2026-09-02 会議決定）。フォルダ名は 会社名（ツール名） で、営業がどちらでも探せる形にする。
COMPANY_FOLDERS = {
    "Scale人事評価": "カラフルボックス株式会社（Scale人事評価）",
    "ピスケスアポ": "株式会社ピスケス（ピスケスアポ）",
    "irohana": "いろはな株式会社（irohana）",
    "配車頭": None,  # 会社フォルダ名は LOCAL_FOLDERS_FILE
    "ABILI_ClipLine": "ClipLine株式会社（ABILI Clip）",
    "クラフトバンクオフィス": "クラフトバンク株式会社（クラフトバンクオフィス）",
}
STEM_PREFIX = "05_面談ヒアリングシート_"


def company_sheets():
    """(md パス, 会社フォルダ名) の並びを返す。"""
    local = _local_folders()
    out = []
    for md in sorted(SHEET_DIR.glob(f"{STEM_PREFIX}*.md")):
        key = md.stem[len(STEM_PREFIX):]
        if key not in COMPANY_FOLDERS:
            print(f"⚠ COMPANY_FOLDERS に未登録: {md.name}（会社フォルダ名を追記してください）")
            continue
        folder = COMPANY_FOLDERS[key] or local.get(key)
        if not folder:
            print(f"⚠ {md.name} の会社フォルダ名が {LOCAL_FOLDERS_FILE} にありません（飛ばします）")
            continue
        out.append((md, folder))
    return out


def _local_folders():
    if not LOCAL_FOLDERS_FILE.exists():
        return {}
    try:
        data = json.loads(LOCAL_FOLDERS_FILE.read_text(encoding="utf-8-sig"))
    except (OSError, ValueError) as e:
        sys.exit(f"✗ {LOCAL_FOLDERS_FILE} を読めません: {e}")
    if not isinstance(data, dict):
        sys.exit(f"✗ {LOCAL_FOLDERS_FILE} は {{ツール名: 会社名（ツール名）}} の JSON にしてください")
    return data


def build_one(md, folder):
    outdir = OUT_ROOT / folder
    outdir.mkdir(parents=True, exist_ok=True)
    made = []

    # 1) 印刷用PDF（md_to_pdf は md と同じ場所に出すため、生成後に移動する）
    pdf = md_to_pdf.make_pdf(md)
    dest = outdir / pdf.name
    dest.unlink(missing_ok=True)
    shutil.move(str(pdf), str(dest))
    made.append(dest)

    parsed = parse_md(md)
    form = parsed["blocks"]

    # 2) 入力できるPDF
    out = outdir / (md.stem + "_入力用.pdf")
    make_fillable_pdf.Form(out, form).build()
    made.append(out)

    # 3) Word
    out = outdir / (md.stem + ".docx")
    make_docx.Doc(form).build().save(out)
    made.append(out)

    # 4) スプレッドシート（xlsx）
    out = outdir / (md.stem + "_オンライン入力用.xlsx")
    title = parsed["title"].replace("（営業記入用）", "（オンライン商談・入力用）")
    sheet = make_check_xlsx.Sheet(form, title)
    sheet.build()
    sheet.save(out)
    made.append(out)

    # 5) md 本体のコピー（Drive へ上げると Google ドキュメントに変換され、ブラウザで読める形になる）
    out = outdir / md.name
    shutil.copyfile(md, out)
    made.append(out)

    return made


def main():
    ap = argparse.ArgumentParser(description="会社別ヒアリングシートを全形式生成する")
    ap.add_argument("--only", default=None, help="会社名・ツール名の部分一致で絞り込む")
    ap.add_argument("--list", action="store_true", help="会社とフォルダ名の対応を表示して終了")
    args = ap.parse_args()

    targets = company_sheets()
    if args.only:
        targets = [(md, f) for md, f in targets if args.only in md.stem or args.only in f]
    if args.list:
        for md, folder in targets:
            print(f"{md.name}  →  {folder}")
        return

    if not targets:
        print("対象が見つかりませんでした")
        sys.exit(1)

    # 記法崩れの検知（md → 設問データ の取りこぼしが無いことを先に確認する）
    ng = [md.name for md, _ in targets if not verify(md)[0]]
    if ng:
        print("記法が崩れている md があります（生成を中止）: " + ", ".join(ng))
        print("  `python md_form_parser.py <md>` で不一致箇所を確認してください")
        sys.exit(1)

    for md, folder in targets:
        made = build_one(md, folder)
        total = sum(p.stat().st_size for p in made)
        print(f"OK  {folder}")
        for p in made:
            print(f"      {p.name}  ({p.stat().st_size:,} bytes)")
        print(f"      計 {len(made)}ファイル / {total:,} bytes")
    print(f"\n完了: {len(targets)}社 → {OUT_ROOT}")


if __name__ == "__main__":
    main()
