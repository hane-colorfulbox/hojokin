# 補助金プロジェクト - Claude向けメモ

このファイルは Claude Code が自動で読み込む、プロジェクト横断のコンテキストファイル。

## プロジェクト概要

IT導入補助金の申請書類を、ヒアリングシート・PDF資料・Excel賃金台帳から AI で自動生成するツール。
本番運用は Streamlit Cloud で動作。

## ディレクトリ構成

```
補助金/
├── app.py                  Streamlit エントリポイント
├── run.py                  CLI エントリポイント
├── hojokin/                コア処理パッケージ
│   ├── ai_extractor.py    Claude API 呼び出し（PDF/CSV/Excel 抽出）
│   ├── config.py          設定・テンプレートマッピング
│   ├── pipeline.py        申請書作成・給与計算のオーケストレーション
│   ├── template_filler.py Excel 書込み
│   ├── wage_reader.py     賃金台帳の決定論パーサー
│   ├── pdf_reader.py      PDF テキスト抽出
│   ├── google_drive.py    Drive API クライアント
│   ├── google_sheets.py   Sheets API クライアント
│   └── ...
├── gas/                    Google Apps Script（管理表自動化）
├── docs/                   運用マニュアル・設計ドキュメント
├── ツール/                 申請テンプレート・ヒアリングシート原本
└── credentials/            Google サービスアカウント鍵（gitignore）
```

## 主要設定

- 環境変数: `.env` 参照（`.env.example` を雛形として）
- 主な変数: `CLAUDE_API_KEY`, `GOOGLE_CREDENTIALS_PATH`, `MANAGEMENT_SHEET_ID`, `USE_AI_WAGE_EXTRACTION`
- Google Sheet/Drive の各種 ID は `.env` で管理。コードにハードコードしない。

## 開発時の注意

### コード規約
- マジックナンバーは定数化、上部にまとめる
- `sys.stdout.reconfigure(encoding='utf-8')` をスクリプト先頭に
- 日本語パスは `pathlib.Path.iterdir()` で解決
- Excel に書き込む文字列の先頭に `=` を入れない（数式と誤認）

### 個人情報・機密情報の取り扱い（重要）
- **顧客企業名・顧客従業員氏名・年収・連絡先など、特定可能な個人情報や顧客情報をリポジトリにコミットしない**
- 作業ログや調査メモは、抽象化した形（「顧客企業A」「担当者」等）で記述する
- 顧客実データ（ヒアリングシート、賃金台帳、給与明細等）は `.gitignore` 配下に置く
- Google Sheet ID・Drive フォルダ ID・Webhook URL 等は環境変数 / Apps Script Properties に逃がす

### 本番デプロイ
- Streamlit Cloud と GitHub の `main` ブランチが連携。push で自動再デプロイ。
- ロールバックは `git revert` または Streamlit Cloud secrets での機能フラグ切替
- `USE_AI_WAGE_EXTRACTION=false` で AI 抽出を旧経路（決定論パーサー）に戻せる

## テスト方針（暗黙ルール）

- **テスト実行時に Anthropic API は絶対に呼ばない**（API課金ゼロが原則）。
  `_debug/test_*.py` は `StubExtractor` / `MagicMock` / Sonnet サブエージェント代用で動くこと
- AI モデルは **Sonnet 4.6** を仮定（コードのデフォルト `claude-sonnet-4-6` 通り）。
  Haiku 等への切替は別途検討。現状は **Sonnet 統一方針**
- 大規模変更後は `_debug/` 配下のテストを全件回し、回帰なしを確認してから push

## ロールバック・運用フラグ

| フラグ | デフォルト | 役割 |
|---|---|---|
| `USE_AI_WAGE_EXTRACTION` | `true` | 賃金台帳の AI 抽出 ON/OFF。`false` で決定論パーサーのみ |

## 関連ドキュメント

- `docs/マニュアル_書類作成.md` — Streamlit/Claude Code を使った書類作成手順
- `docs/運用マニュアル.md` — 案件管理から書類作成までの運用フロー
- `docs/設計_API自動化.md` — 完全自動化バージョンの設計案
- `docs/案件メモ/` — 個別案件の調査ログ（テンプレートのみ。実案件メモは gitignore）
