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

## 業務体制・運用ルール

### 連絡窓口
- **賃上げ申請関連の指示・依頼は坂平さん（窓口）経由で受ける**
- 1次振り返りMTG（2026-05-14）で混乱回避のため一本化決定
- 山田さん・村上さんから個別に依頼が来た場合も、坂平さんに集約してから動く

### 賃金台帳の回収方針
- 顧客への依頼は **原則 Excel / CSV 形式**（フォーマット統一によりAPI節約・自動化精度向上）
- PDF しか出ない場合は **ローカル（Claude Code）で Excel に変換**してから自動化に投入する運用
- 変換用テンプレート（必要項目：氏名 / 月 / 課税支給総額 / 12ヶ月合計 / 勤務時間）を羽根側で作成・共有

### 定例報告
- 補助金TM MTG: **毎月25日**（社内）。ツール改善・運用変更の進捗を報告
- 次回2次申請スケジュール（暫定）：
  - 推奨送客期限 6/1（月）
  - 疎明期限 6/3
  - 資料提出期限 6/8
  - 申請締切 6/15（月）

### 残タスクの参照先
- 2次申請に向けた改善タスクは [`docs/TODO_2次申請改善.md`](docs/TODO_2次申請改善.md) に集約

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
- **`docs/補助金_実務知識ベース.md`** — **R215/R216 や賃金台帳の取り扱いに迷ったらまずここを参照**。公募要領の定義（給与支給総額・役員除外・賞与の扱い）、賃金台帳の構造解釈（段1/段2/役員ブロック）、Codex 調査結果のサマリを集約。鮮度管理あり。

## R215/R216 や賃金台帳に迷ったら

- まず `docs/補助金_実務知識ベース.md` を読む（公募要領と Codex 調査結果が集約済み）
- 公募要領の WebFetch / Codex のフル再調査は **本資料で判断できないとき** のみ
- 案件特有の事情（複数集計行の意味、賞与シートの有無など）は最終的にユーザーが PDF 原本で確認する運用
