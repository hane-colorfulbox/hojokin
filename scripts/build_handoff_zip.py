"""補助金業務の引き継ぎコンテキストパックを ZIP にまとめて配布物を作る。

GitHub アカウントを使わない担当者へ、引き継ぎパック（引き継ぎ/）と、そこから
参照している配布可能なドキュメントを1つの ZIP で渡すためのスクリプト。
各自の PC の Claude Code が、解凍したフォルダをローカルファイルとして読める。

実行方法:
    python scripts/build_handoff_zip.py

出力:
    _dist/hojokin-handoff.zip            最新版エイリアス
    _dist/hojokin-handoff-YYYYMMDD.zip   日付付き
    _dist/handoff_version.json           バージョン情報

配布フロー:
    1. このスクリプトを実行して _dist/ に ZIP を生成
    2. _dist/hojokin-handoff.zip を Google Drive の配布フォルダにアップロード
    3. 担当者が解凍 → そのフォルダで Claude Code を開き 引き継ぎ/00_はじめに.md から読む

重要（PII ガード）:
    含めるものは下の INCLUDE 許可リストで「明示的に挙げたものだけ」。
    顧客実データや個人情報を含む gitignore 配下（docs/補助金_実務知識ベース.md,
    docs/案件メモ/, docs/TODO_2次申請改善.md, docs/セカンドオピニオン加点/ の社外秘分,
    _debug/, credentials/, output/, *資料/ 等）は許可リストに入れない＝絶対に同梱しない。
    新しく参照先ドキュメントを増やすときは、それが配布可能（PII 無し）か確認のうえ
    INCLUDE に追記する。
    セカンドオピニオン加点は配布物テンプレの汎用版のみ明示同梱し、自社製品特化版
    (09/11/12/13)・社長意向/設計/議事/参考_Scale・案件別・レンダー物(*.pdf/*.xlsx)は
    同梱しない。
"""
import json
import shutil
import sys
import zipfile
from datetime import date
from pathlib import Path

sys.stdout.reconfigure(encoding='utf-8')

ROOT = Path(__file__).resolve().parent.parent
DIST_DIR = ROOT / '_dist'

# 配布する中身（許可リスト）。ディレクトリは再帰的に含める。
# 引き継ぎパックが相対リンク（../docs/..., ../CLAUDE.md, ../.claude/skills/...）で
# 参照する配布可能ドキュメントを、リポジトリ相対のパス構造を保ったまま同梱する。
# スキル(wagebook-convert / case-docs-check)も同梱する（v2 一本化）。受け取り手が補助金
# フォルダ直下に解凍すると `<補助金>/.claude/skills/...` = project-scoped スキルとして
# 自動有効になり、別DL不要になる。同梱前に build_skill_zip.check_template_sync()（テンプレ
# xlsx byte 一致）と check_docscheck_sync.check_docs_sync()（分類定数の AST 突合）を
# 必ず通す（main 冒頭）。
INCLUDE_DIRS = [
    '引き継ぎ',
    '.claude/skills/wagebook-convert',
    '.claude/skills/case-docs-check',
    '.claude/skills/katen-judge',
]
# 注: CLAUDE.md は社内担当者の氏名を含む開発設定ファイルのため、配布スナップショットには
# 同梱しない（引き継ぎパックは CLAUDE.md にリンクしない自己完結構成）。GitHub 経由の閲覧者は
# リポジトリ内で参照できる。
INCLUDE_FILES = [
    'README.md',
    'docs/運用マニュアル.md',
    'docs/マニュアル_書類作成.md',
    'docs/設計_API自動化.md',
    'docs/警告一覧.md',
    'docs/弊社成果物と書式.md',
    # セカンドオピニオン加点の面談準備テンプレ（汎用版のみ・PII なし）。
    # 自社製品特化版(09/11/12/13)・社長意向/設計/議事/参考_Scale・案件別・*.pdf/*.xlsx は
    # 明示列挙しない＝同梱しない（1ファイルずつ挙げる許可リストなので構造的に混入しない）。
    'docs/セカンドオピニオン加点/配布物テンプレ/00_予約フォーム仕様.md',
    'docs/セカンドオピニオン加点/配布物テンプレ/05_面談ヒアリングシート.md',
    'docs/セカンドオピニオン加点/配布物テンプレ/05_面談ヒアリングシート_オンライン版.md',
    'docs/セカンドオピニオン加点/配布物テンプレ/05_別紙_お客様向け事前案内.md',
    'docs/セカンドオピニオン加点/配布物テンプレ/05_別紙_営業向けカンペ.md',
    # SO 実務プロンプトは 07b（確認メール）→ 06（予約フォーム入力文）の2本立て
    # （2026-07-21 決定。07・06b は使わない＝配布から除外。ファイル自体はリポジトリに残置）
    'docs/セカンドオピニオン加点/配布物テンプレ/06_入力整理プロンプト.md',
    'docs/セカンドオピニオン加点/配布物テンプレ/07b_お客様確認メール生成プロンプト.md',
    'docs/セカンドオピニオン加点/配布物テンプレ/08_面談回答例_汎用中立版.md',
    'docs/セカンドオピニオン加点/配布物テンプレ/10_経営者向け_一問一答の回答例_汎用.md',
    'docs/セカンドオピニオン加点/配布物テンプレ/14_記入見本_予約フォーム入力_製造業例.md',
    'docs/セカンドオピニオン加点/配布物テンプレ/15_記入見本_経営者一問一答_製造業例.md',
    'docs/セカンドオピニオン加点/_sample/05_記入サンプル_ダミー工務店.md',
    # 加点判定ローカル完結（katen-judge）の判定エンジンと様式。ビルド毎にリポジトリから
    # コピーするため、ツール本番と常に同一コード（同期ドリフトが構造的に起きない）。
    'hojokin/__init__.py',
    'hojokin/config.py',
    'hojokin/models.py',
    'hojokin/wage_calculator.py',
    'hojokin/wage_reader.py',
    'hojokin/xlsx_surgical.py',
    'hojokin/bonus_wage_ledger_writer.py',
    'ツール/加点判定用賃金台帳テンプレート.xlsx',
    '補助金加点/加点措置①用.xlsx',
    '補助金加点/加点措置②用.xlsx',
    '補助金加点/補助率引き上げ・加点措置①用.xlsx',
    'docs/チェックリスト_加点判定.md',
    # 業務全体の俯瞰マップ（新任オンボーディング用・脱PII）
    'docs/補助金_業務フロー全体マップ.md',
    # 受け取り手が「引き継ぎを最新に更新して」で使う更新スクリプト（ZIPに同梱＝初回展開後は自走）。
    'scripts/update_handoff.py',
]

# ZIP のルートに特別配置するファイル（リポジトリ内ソース → ZIP 内の配置名）。
# 配布用CLAUDE.md は受け取り手の「補助金フォルダ直下」に CLAUDE.md として置き、
# Claude Code に起動時自動ロードさせる索引にする（引き継ぎ/ 配下としては含めない）。
ROOT_PLACEMENTS = {'引き継ぎ/配布用CLAUDE.md': 'CLAUDE.md'}

# 念のための除外パターン（許可ディレクトリ配下にうっかり入った生成物・機密を弾く）。
EXCLUDE_PARTS = {'__pycache__', '_transcripts', '案件別'}
EXCLUDE_SUFFIXES = {'.pyc'}
# 再帰収集から除く固有ファイル名（ROOT_PLACEMENTS でルート配置するものは二重に入れない）。
EXCLUDE_NAMES = {'配布用CLAUDE.md'}
# 配布対象外のファイル名キーワード（顧客資料の典型語）。許可リスト運用の二重安全網。
EXCLUDE_NAME_KEYWORDS = ('賃金台帳_', '決算', '給与明細', 'ヒアリング結果', '社長意向', '議事_', '参考_')


def _excluded(path: Path) -> bool:
    if EXCLUDE_PARTS & set(path.parts):
        return True
    if path.suffix in EXCLUDE_SUFFIXES:
        return True
    name = path.name
    if name in EXCLUDE_NAMES:
        return True
    return any(kw in name for kw in EXCLUDE_NAME_KEYWORDS)


def _collect() -> list[tuple[Path, str]]:
    """(ソースパス, ZIP内の配置名) の一覧を返す。"""
    items: list[tuple[Path, str]] = []
    for d in INCLUDE_DIRS:
        base = ROOT / d
        if not base.exists():
            print(f'⚠ 許可ディレクトリが見つかりません（スキップ）: {d}', file=sys.stderr)
            continue
        for f in sorted(base.rglob('*')):
            if f.is_file() and not _excluded(f.relative_to(ROOT)):
                items.append((f, f.relative_to(ROOT).as_posix()))
    for f in INCLUDE_FILES:
        p = ROOT / f
        if p.exists():
            items.append((p, Path(f).as_posix()))
        else:
            print(f'⚠ 許可ファイルが見つかりません（スキップ）: {f}', file=sys.stderr)
    for src_rel, arcname in ROOT_PLACEMENTS.items():
        p = ROOT / src_rel
        if p.exists():
            items.append((p, arcname))
        else:
            print(f'⚠ ルート配置ファイルが見つかりません（スキップ）: {src_rel}', file=sys.stderr)
    return items


def main() -> int:
    handoff_dir = ROOT / '引き継ぎ'
    if not handoff_dir.exists():
        print(f'❌ 引き継ぎパックが見つかりません: {handoff_dir}', file=sys.stderr)
        return 1

    # スキル同梱の前提: スキル同梱 xlsx == ツール原本 xlsx（byte一致）。崩れていたら中止。
    from build_skill_zip import check_template_sync
    if not check_template_sync():
        print('❌ スキルテンプレ同期NG: 中止（スキル同梱の前提が崩れています）', file=sys.stderr)
        return 1

    # case-docs-check の前提: スキル同梱 check_docs.py の分類・必須判定定数 ==
    # hojokin/pipeline.py（FileDetector）・app.py の正典（AST 突合）。崩れていたら中止。
    from check_docscheck_sync import check_docs_sync
    if not check_docs_sync():
        print('❌ case-docs-check 定数同期NG: 中止（原本と スキル側 check_docs.py を同時に直すこと）',
              file=sys.stderr)
        return 1

    items = _collect()
    if not items:
        print('❌ 同梱対象のファイルが0件です', file=sys.stderr)
        return 1

    DIST_DIR.mkdir(exist_ok=True)
    today = date.today().isoformat()
    zip_dated = DIST_DIR / f'hojokin-handoff-{today}.zip'
    zip_latest = DIST_DIR / 'hojokin-handoff.zip'

    included: list[str] = []
    with zipfile.ZipFile(zip_dated, 'w', zipfile.ZIP_DEFLATED) as zf:
        for src, arcname in items:
            zf.write(src, arcname)
            included.append(arcname)
        version = {
            'name': 'hojokin-handoff',
            'built_on': today,
            'zip': zip_dated.name,
            'file_count': len(included),
            'files': included,
        }
        version_json = json.dumps(version, ensure_ascii=False, indent=2)
        # 受け取り手の update_handoff.py が差分適用に使う manifest を ZIP 内にも同梱する。
        zf.writestr('handoff_version.json', version_json)

    shutil.copy2(zip_dated, zip_latest)

    (DIST_DIR / 'handoff_version.json').write_text(version_json, encoding='utf-8')

    size_kb = zip_dated.stat().st_size / 1024
    print('✅ 引き継ぎパック ZIP 生成完了')
    print(f'   日付付き  : {zip_dated.relative_to(ROOT)} ({size_kb:.1f} KB)')
    print(f'   最新版    : {zip_latest.relative_to(ROOT)}')
    print(f'   ファイル数: {len(included)}')
    print()
    print('同梱ファイル:')
    for name in included:
        print(f'   - {name}')
    print()
    print('次の手順: _dist/hojokin-handoff.zip を Drive 配布フォルダにアップロード → 担当者へ共有')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
