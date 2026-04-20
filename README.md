# 補助金書類 自動化プロジェクト

IT導入補助金の申請書類をClaude API + openpyxlで半自動作成するツール群。
案件管理（Google Sheets）・資料受領（Google Drive/Gmail）・書類生成（Python）・Streamlit UIを含む。

---

## はじめに読むもの

| 目的 | ドキュメント |
|------|------------|
| 業務全体の流れを知りたい | [docs/運用マニュアル.md](docs/運用マニュアル.md) |
| 書類作成ツールの使い方 | [docs/マニュアル_書類作成.md](docs/マニュアル_書類作成.md) |
| システム設計・構成 | [docs/設計_API自動化.md](docs/設計_API自動化.md) |
| GAS（案件管理表など）の仕様 | [gas/README.md](gas/README.md) |
| 個別案件の調査・修正ログ | [docs/案件メモ/](docs/案件メモ/) |

---

## ディレクトリ構成

```
補助金/
├── README.md                      このファイル（入口）
├── app.py                         Streamlit UI（書類作成）
├── run.py                         CLI実行
├── transfer*.py                   ヒアリング転記（AI版・非AI版）
├── wage_calc_kyo.py               給与計算（案件別）
├── hojokin/                       コアモジュール
│   ├── pipeline.py                処理パイプライン + FileDetector
│   ├── ai_extractor.py            Claude APIでPDF読取
│   ├── hearing_reader.py          ヒアリングシート読取
│   ├── wage_reader.py             賃金台帳読取 + 加点判定
│   ├── wage_calculator.py         給与計算
│   ├── template_filler.py         Excelテンプレート書き込み
│   ├── google_drive.py            Drive連携
│   ├── google_sheets.py           Sheets連携
│   └── config.py                  定数・最低賃金表
├── gas/                           Google Apps Script（案件管理表等）
├── docs/                          ドキュメント
│   ├── 運用マニュアル.md
│   ├── マニュアル_書類作成.md
│   ├── 設計_API自動化.md
│   └── 案件メモ/                  案件ごとの調査・修正ログ
├── テンプレート原本.xlsx           書類テンプレート（触らない）
├── 1.交付申請_過去採択なし/        交付申請用資料
├── 補助金加点/                    加点措置関連
├── output/                        生成ファイル出力先（gitignore）
└── credentials/                   API認証情報（gitignore）
```

---

## 開発ルール

- 詳細: [../CLAUDE.md](../CLAUDE.md)（カラフルボックス全体）と [../../CLAUDE.md](../../CLAUDE.md)（共通）
- Excel操作は `openpyxl`（xlwings禁止）
- APIキーは `.env` の `ANTHROPIC_API_KEY` から読む。コード直書き禁止
- 出力ファイル名は `{会社名}_{テンプレート種別}_AI版.xlsx`
- 外部書き込み（Drive更新・Sheets更新）は必ず `--dry-run` で確認してから本番実行

---

## 新しくMDを書くとき

| 種類 | 置き場所 | 命名 |
|------|---------|------|
| 業務フロー・使い方マニュアル | `docs/` | `マニュアル_〇〇.md` |
| 設計・技術仕様 | `docs/` | `設計_〇〇.md` |
| 個別案件の調査・修正ログ | `docs/案件メモ/` | `YYYY-MM-DD_担当者_会社名.md` |
| 特定サブシステムの説明 | そのサブフォルダ直下 | `README.md` |

ルート直下にはこのREADME以外のMDを置かない。
