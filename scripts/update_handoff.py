"""引き継ぎコンテキストパックを最新版へ差分適用で更新する（受け取り手のPCで動く）。

使い方（受け取り手の Claude Code が実行）:
    python scripts/update_handoff.py <入力>

<入力> は次のいずれか:
  - Google Drive コネクタで download_file_content した base64 を保存した .txt / .b64
    （※ base64 文字列だけを保存すること。JSON等で包まない）
  - 手動DLした hojokin-handoff.zip そのもの

やること:
  - 入力を復号・検証（ZIPとして壊れていないか＝DL切れの検知）
  - 補助金フォルダ直下（このスクリプトの2階層上）に差分適用
    - 引き継ぎ/ と .claude/skills/wagebook-convert/ は丸ごと入替（消えたファイルも反映）
    - docs 等の許可ファイルは更新、前回 manifest から消えたものは削除
    - あなた自身の CLAUDE.md は保護（配布版マーカーが無ければ上書きせず CLAUDE.handoff.md に退避）
  - 反映後は Claude Code を再起動

初回インストールはこのスクリプトではなく、ZIP を手動で解凍して補助金フォルダ直下に置く。
"""
import base64
import json
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path

sys.stdout.reconfigure(encoding='utf-8')

ROOT = Path(__file__).resolve().parent.parent  # <補助金> フォルダ直下
MANIFEST_NAME = 'handoff_version.json'
LOCAL_VERSION = ROOT / '.handoff_version'
CLAUDE_MD = 'CLAUDE.md'
CLAUDE_MARKER = '<!-- handoff-distributed-claude-md -->'
# 丸ごと入れ替える（stale ファイルも消す）管理ディレクトリ。
MANAGED_DIRS = ('引き継ぎ', '.claude/skills/wagebook-convert')
ZIP_MAGIC = b'PK\x03\x04'


def _load_zip_bytes(input_path: Path) -> bytes:
    """入力（zip そのもの or base64テキスト）から ZIP の生バイトを得る。"""
    raw = input_path.read_bytes()
    if raw[:4] == ZIP_MAGIC:
        return raw
    text = raw.decode('utf-8', errors='ignore')
    try:
        decoded = base64.b64decode(text, validate=False)  # 空白/改行は破棄される
    except Exception as exc:  # noqa: BLE001
        raise SystemExit(f'❌ 入力を base64 として復号できません: {exc}\n'
                         '   → DLが不完全の可能性。手動でZIPをDLし直してください。')
    if decoded[:4] != ZIP_MAGIC:
        raise SystemExit('❌ 入力がZIPではありません（base64復号後もPK署名なし）。\n'
                         '   → DLが途中で切れている可能性。手動DLで再取得してください。')
    return decoded


def _validate_zip(zip_path: Path) -> None:
    if not zipfile.is_zipfile(zip_path):
        raise SystemExit('❌ ZIPとして開けません（DL不完全/破損）。手動DLで再試行してください。')
    with zipfile.ZipFile(zip_path) as zf:
        bad = zf.testzip()
        if bad is not None:
            raise SystemExit(f'❌ ZIP内 {bad} がCRC不一致（破損/切り詰め）。手動DLで再試行を。')


def _read_manifest(extract_root: Path) -> tuple[str, set[str]]:
    """展開済みディレクトリから (built_on, 管理ファイル集合) を読む。"""
    mf = extract_root / MANIFEST_NAME
    if mf.exists():
        data = json.loads(mf.read_text(encoding='utf-8'))
        return data.get('built_on', '?'), set(data.get('files', []))
    files = {p.relative_to(extract_root).as_posix()
             for p in extract_root.rglob('*') if p.is_file() and p.name != MANIFEST_NAME}
    return '?', files


def _prev_version() -> dict:
    if LOCAL_VERSION.exists():
        return json.loads(LOCAL_VERSION.read_text(encoding='utf-8'))
    return {}


def _under_managed(rel: str) -> bool:
    return any(rel == d or rel.startswith(d + '/') for d in MANAGED_DIRS)


def _apply_claude_md(src: Path) -> str:
    """CLAUDE.md をマーカー判定で安全に反映。結果メッセージを返す。"""
    dst = ROOT / CLAUDE_MD
    if not dst.exists():
        shutil.copy2(src, dst)
        return '設置（新規）'
    if CLAUDE_MARKER in dst.read_text(encoding='utf-8', errors='ignore'):
        shutil.copy2(src, dst)
        return '上書き（配布版を最新に更新）'
    shutil.copy2(src, ROOT / 'CLAUDE.handoff.md')
    return ('保護（あなた独自のCLAUDE.mdを検出）→ CLAUDE.handoff.md に配置。'
            '案内ブロックを自分のCLAUDE.mdに追記 or `@CLAUDE.handoff.md` で取り込んでください')


def _apply(extract_root: Path, new_files: set[str], old_files: set[str]) -> tuple[int, int, str]:
    updated = 0
    for d in MANAGED_DIRS:  # 1) 管理ディレクトリは丸ごと入替（stale も消える）
        src_dir = extract_root / d
        if not src_dir.exists():
            continue
        dst_dir = ROOT / d
        if dst_dir.exists():
            shutil.rmtree(dst_dir)
        shutil.copytree(src_dir, dst_dir)
        updated += sum(1 for p in dst_dir.rglob('*') if p.is_file())

    claude_msg = 'CLAUDE.md 同梱なし'  # 2) CLAUDE.md（マーカー保護）
    if (extract_root / CLAUDE_MD).exists():
        claude_msg = _apply_claude_md(extract_root / CLAUDE_MD)

    for rel in sorted(new_files):  # 3) その他の許可ファイル（docs, scripts, README 等）
        if rel == CLAUDE_MD or _under_managed(rel):
            continue
        src = extract_root / rel
        if not src.exists():
            continue
        dst = ROOT / rel
        dst.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(src, dst)
        updated += 1

    deleted = 0  # 4) manifest から消えたファイルを削除（管理ディレクトリ外・CLAUDE.md 以外）
    for rel in sorted(old_files - new_files):
        if rel == CLAUDE_MD or _under_managed(rel):
            continue
        target = ROOT / rel
        if target.is_file():
            target.unlink()
            deleted += 1
    return updated, deleted, claude_msg


def main(argv: list[str]) -> int:
    if len(argv) < 2:
        print('使い方: python scripts/update_handoff.py <base64ファイル or hojokin-handoff.zip>',
              file=sys.stderr)
        return 2
    if not (ROOT / '引き継ぎ').exists():
        print(f'❌ ここは補助金フォルダ直下ではないようです（引き継ぎ/ が無い）: {ROOT}\n'
              '   初回はZIPを手動で解凍して補助金フォルダ直下に置いてください。', file=sys.stderr)
        return 1

    input_path = Path(argv[1]).expanduser().resolve()
    if not input_path.exists():
        print(f'❌ 入力が見つかりません: {input_path}', file=sys.stderr)
        return 1

    zip_bytes = _load_zip_bytes(input_path)

    with tempfile.TemporaryDirectory(prefix='handoff_update_') as tmp:
        tmp_dir = Path(tmp)
        zip_path = tmp_dir / 'handoff.zip'
        zip_path.write_bytes(zip_bytes)
        _validate_zip(zip_path)

        extract_root = tmp_dir / 'unzipped'
        extract_root.mkdir()
        with zipfile.ZipFile(zip_path) as zf:
            zf.extractall(extract_root)

        built_on, new_files = _read_manifest(extract_root)
        prev = _prev_version()
        if prev.get('built_on') == built_on and built_on != '?' and '--force' not in argv:
            print(f'✅ 既に最新です（built_on={built_on}）。差分なし。')
            return 0

        updated, deleted, claude_msg = _apply(extract_root, new_files, set(prev.get('files', [])))
        LOCAL_VERSION.write_text(json.dumps(
            {'built_on': built_on, 'files': sorted(new_files)},
            ensure_ascii=False, indent=2), encoding='utf-8')

    print('✅ 引き継ぎパックを更新しました')
    print(f'   版(built_on): {built_on}')
    print(f'   更新ファイル: {updated}')
    print(f'   削除ファイル: {deleted}')
    print(f'   CLAUDE.md   : {claude_msg}')
    print('   → 反映のため Claude Code を再起動してください。')
    return 0


if __name__ == '__main__':
    raise SystemExit(main(sys.argv))
