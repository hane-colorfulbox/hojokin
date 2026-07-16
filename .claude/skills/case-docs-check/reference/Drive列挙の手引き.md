# Drive 列挙の手引き（case-docs-check 経路A）

Google Drive コネクタ（MCP）で案件フォルダを列挙・取得する手順。
**ツール名は環境ごとに異なる**（サーバー名のプレフィックスが uuid 等になる）ため、
本スキルはツールを**名前ではなく機能で指定**する。自分の環境で以下の機能を持つ
Drive コネクタツールを探して使うこと:

| 機能 | ツール名の典型（末尾） | 用途 |
|---|---|---|
| 構造化クエリ検索 | `search_files` | フォルダ内列挙・顧客名検索 |
| メタデータ取得 | `get_file_metadata` | 親フォルダ・mimeType の確認 |
| バイナリ取得 | `download_file_content` | base64 でDL（中身チェック用） |
| テキスト表現取得 | `read_file_content` | PDF/xlsx の概観（正確な検査はDL+Read推奨） |

> 🔴 このスキルで使ってよいのは**読み取り系のみ**。`create_file` / `copy_file` 等の
> 書込系ツールがコネクタにあっても**使用禁止**（SKILL.md §0.3）。

## 1. folderId の確定

- **URLをもらった場合**: `https://drive.google.com/drive/folders/<ID>` または `...?id=<ID>` から
  `<ID>` を抜く（`[-\w]{25,}` 程度の英数ハイフン列）。
- **顧客名をもらった場合**: `search_files` に `title contains '<顧客名>' and mimeType contains 'folder'`
  相当のクエリを投げる。案件フォルダの命名は `NNN.[顧客企業名]_[申請枠]`
  （例: `012.株式会社サンプル建設_通常枠`）なので、ヒットのうちこの形のものを採る。
  - 複数ヒットで絞れない場合のみ、チェック開始**前**にユーザーへ1回だけ確認（§0.2）。
  - フォルダ名の末尾 `_通常枠` / `_インボイス` から申請枠を自動推定する。

## 2. フォルダ内の再帰列挙

`search_files` の構造化クエリで親を指定する:

```
parentId = '<folderId>'
```

- 返ってきた各アイテムの `mimeType` が `application/vnd.google-apps.folder` のものはサブフォルダ。
  **サブフォルダ名を NFC 正規化して `申請時使用` に一致したら、その配下には降りない**
  （税理士納品の要約版PDF置き場。R216 算定を壊すためツールも見に行かない運用）。
  それ以外（`1.交付申請` `2.実績報告` 等）は同じクエリで再帰する。
- ページネーション: 結果が多い場合は `pageToken`（`next_page_token`）で最後まで取り切る。
  途中で打ち切ると「不足」の誤報告になるので必ず全件取得。
- Google Drive のファイル名は NFD（濁点分離）になりがち。**照合は常に NFC 正規化後**に行う。

### クエリが通らない環境のフォールバック

1. ツールスキーマに親フォルダ指定のパラメータがあればそれを使う。
2. `title contains '<フォルダ名>'` でファイル候補を検索し、`get_file_metadata` の親情報で
   対象フォルダ配下かを確認して絞る。
3. それでも列挙できない場合は、**チェックを開始せずに**「コネクタでフォルダ列挙ができない環境のため、
   案件フォルダをローカルにダウンロードして `--local` で再実行してほしい」と案内して終了する
   （途中停止ではなく、開始前の pre-flight で判定する）。

## 3. manifest.json の作成

列挙結果を `{WORKDIR}/manifest.json` に保存する（check_docs.py `--manifest` の入力）:

```json
{"files": [
  {"name": "履歴事項全部証明書.pdf", "id": "<fileId>", "mimeType": "application/pdf",
   "parent_path": "", "size": 123456, "modifiedTime": "2026-06-01T00:00:00Z"}
]}
```

- `parent_path` は案件フォルダからの相対（サブフォルダ名を `/` 区切り）。
- `mimeType` は必ず入れる（Googleネイティブ形式・ショートカットの判定に使う）。
  - `application/vnd.google-apps.shortcut` → check_docs.py が「実体解決不能」の要確認に落とす。
  - `application/vnd.google-apps.spreadsheet` → 拡張子なしでも `.xlsx` 扱いで分類される。

## 4. 中身チェック対象のダウンロード

check_docs.py（1回目）の JSON にある `content_check_targets` のファイルだけを
`download_file_content` で取得する（全件DLしない）:

1. 返却は **base64 文字列**（JSON `{content, id, mimeType, title}` の `content`）。
   - 応答が大きい場合、Claude Code はツール結果を**自動でローカルファイルに退避**し、そのパスを
     示す（実測: 82KB の PDF → 11万文字の JSON が `tool-results/...txt` に保存された）。
     **base64 をコンテキストで読もうとしない**こと。退避ファイルから Python で直接復号する:
   ```
   python -c "import base64,json,pathlib; d=json.loads(pathlib.Path(r'<退避ファイル>').read_text(encoding='utf-8')); pathlib.Path(r'<WORKDIR>/files/'+d['title']).write_bytes(base64.b64decode(d['content']))"
   ```
   - 復号後に先頭バイトを確認する（PDF なら `%PDF`、xlsx なら `PK`）。化けていたら形式を疑う。
2. Googleネイティブ形式は `exportMimeType` を指定する
   （スプレッドシート → `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`）。
3. **サイズ上限**: 巨大ファイル（目安 20MB 超、フルスキャン決算書等）は DL を試みて失敗・切断したら
   粘らず「サイズ超過につき中身チェック省略・要目視」として報告に落とす。
4. 復号先は必ず `{WORKDIR}/files/`（セッション scratchpad 配下）。**リポジトリ・案件フォルダには置かない**。
5. DL 完了後、check_docs.py を `--files-dir {WORKDIR}/files` 付きで再実行すると、
   Excel 構造チェック（テンプレ規格・ヒアリング様式判定）まで機械判定される。

`read_file_content`（テキスト表現）は概観・当たり付けには使えるが、証明日・期末日・様式の
正確な確認は **DL した実ファイルを Read（PDF はページ指定可）で見る**のが正
（テキスト表現は形式が保証されず、レイアウト情報が落ちる）。
