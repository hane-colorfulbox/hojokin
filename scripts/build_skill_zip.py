"""wagebook-convert Skill を ZIP にパッケージ化して配布物を作る。

実行方法:
    python scripts/build_skill_zip.py

出力:
    _dist/wagebook-convert.zip           最新版エイリアス
    _dist/wagebook-convert-YYYYMMDD.zip  日付付き
    _dist/version.json                   バージョン情報

配布フロー:
    1. このスクリプトを実行して _dist/ に ZIP を生成
    2. _dist/wagebook-convert.zip を Google Drive の配布フォルダにアップロード
    3. Streamlit アプリの「Skill インストール案内」内の「最新更新日」を更新（手動）
    4. git push で Streamlit Cloud を再デプロイ

担当者の初回セットアップ:
    1. Streamlit アプリ上の Drive リンクから ZIP をダウンロード
    2. ~/.claude/skills/ 配下に展開（既存があれば上書き）
    3. Claude Code を再起動して /wagebook-convert が表示されることを確認
"""
import hashlib
import json
import shutil
import sys
import zipfile
from datetime import date
from pathlib import Path

sys.stdout.reconfigure(encoding='utf-8')

ROOT = Path(__file__).resolve().parent.parent
SKILL_DIR = ROOT / '.claude' / 'skills' / 'wagebook-convert'
DIST_DIR = ROOT / '_dist'

# スキル同梱テンプレと、ツールが「賃金台帳の作成」タスクで使う原本テンプレ。
# 両者が byte 単位で一致していないと、CC スキル出力とツール出力でフォーマットが
# ズレ、後段の申請書作成（決定論パーサーの列読取）で不具合が出る。ビルド時に強制する。
SKILL_TEMPLATE = SKILL_DIR / 'templates' / '賃金台帳テンプレート.xlsx'
TOOL_TEMPLATE = ROOT / 'ツール' / '賃金台帳テンプレート.xlsx'


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _check_template_sync() -> bool:
    """スキル同梱テンプレ == ツール原本テンプレ を検証。不一致ならビルド中止。"""
    if not SKILL_TEMPLATE.exists():
        print(f'❌ スキル同梱テンプレが見つかりません: {SKILL_TEMPLATE}', file=sys.stderr)
        return False
    if not TOOL_TEMPLATE.exists():
        print(f'⚠ ツール原本テンプレが見つかりません（同期チェックをスキップ）: {TOOL_TEMPLATE}',
              file=sys.stderr)
        return True
    skill_hash = _sha256(SKILL_TEMPLATE)
    tool_hash = _sha256(TOOL_TEMPLATE)
    if skill_hash != tool_hash:
        print(
            '❌ テンプレ不一致: スキル同梱テンプレとツール原本テンプレが異なります。\n'
            f'   skill: {SKILL_TEMPLATE} ({skill_hash[:16]})\n'
            f'   tool : {TOOL_TEMPLATE} ({tool_hash[:16]})\n'
            '   → どちらかを最新に揃えてから再ビルドしてください'
            '（CC スキル出力とツール出力のフォーマット一致を担保するため）。',
            file=sys.stderr,
        )
        return False
    print(f'✅ テンプレ同期OK（skill == tool, sha256={skill_hash[:16]}）')
    return True


def check_template_sync() -> bool:
    """公開ラッパ。build_handoff_zip がスキル同梱前にテンプレ byte 一致を確認するために流用する。

    スキル同梱 xlsx（`.claude/skills/wagebook-convert/templates/賃金台帳テンプレート.xlsx`）と
    ツール原本（`ツール/賃金台帳テンプレート.xlsx`）が一致していないと、CC スキル出力とツール
    出力のフォーマットがズレて後段パーサーが壊れる。同梱経路（handoff）でも必ず通す。
    """
    return _check_template_sync()


def main() -> int:
    if not SKILL_DIR.exists():
        print(f'❌ Skill ディレクトリが見つかりません: {SKILL_DIR}', file=sys.stderr)
        return 1

    if not _check_template_sync():
        return 1

    DIST_DIR.mkdir(exist_ok=True)

    today = date.today().isoformat()
    zip_dated = DIST_DIR / f'wagebook-convert-{today}.zip'
    zip_latest = DIST_DIR / 'wagebook-convert.zip'

    files_included: list[str] = []
    with zipfile.ZipFile(zip_dated, 'w', zipfile.ZIP_DEFLATED) as zf:
        for file in sorted(SKILL_DIR.rglob('*')):
            if file.is_file():
                # __pycache__ / .pyc は配布物に含めない
                # （review/verify_wagebook.py の実行時に生成されるキャッシュ）
                if '__pycache__' in file.parts or file.suffix == '.pyc':
                    continue
                arcname = Path('wagebook-convert') / file.relative_to(SKILL_DIR)
                zf.write(file, arcname.as_posix())
                files_included.append(arcname.as_posix())

    shutil.copy2(zip_dated, zip_latest)

    version = {
        'name': 'wagebook-convert',
        'built_on': today,
        'zip': zip_dated.name,
        'file_count': len(files_included),
        'files': files_included,
    }
    (DIST_DIR / 'version.json').write_text(
        json.dumps(version, ensure_ascii=False, indent=2),
        encoding='utf-8',
    )

    size_kb = zip_dated.stat().st_size / 1024
    print('✅ ZIP 生成完了')
    print(f'   日付付き  : {zip_dated.relative_to(ROOT)} ({size_kb:.1f} KB)')
    print(f'   最新版    : {zip_latest.relative_to(ROOT)}')
    print(f'   ファイル数: {len(files_included)}')
    print(f'   version  : {(DIST_DIR / "version.json").relative_to(ROOT)}')
    print()
    print('次の手順:')
    print('  1. _dist/wagebook-convert.zip を Drive 配布フォルダにアップロード')
    print('  2. app.py の WAGEBOOK_SKILL_VERSION を今日の日付に更新')
    print('  3. git commit & push で Streamlit Cloud を再デプロイ')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
