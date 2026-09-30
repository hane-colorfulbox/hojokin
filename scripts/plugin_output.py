"""配布プラグインの出力（マニフェスト・プラグイン用リポジトリの直下ファイル・書き出し・ZIP）。"""
import hashlib
import json
import zipfile
from pathlib import Path

import distribution_manifest as dm

ZIP_FIXED_TIME = (2026, 1, 1, 0, 0, 0)   # ZIP を同じ中身なら同じバイト列にするための固定時刻
PRESERVED_TOP = {'.git'}                 # プラグイン用リポジトリで消さないもの
HOOK_PATH = '.githooks/pre-commit'


def _json(obj) -> bytes:
    return (json.dumps(obj, ensure_ascii=False, indent=2) + '\n').encode('utf-8')


def plugin_json(version: str) -> bytes:
    return _json({
        'name': dm.PLUGIN_NAME,
        'version': version,
        'description': dm.PLUGIN_DESCRIPTION,
        'author': {'name': dm.ORG_NAME},
    })


def marketplace_json() -> bytes:
    return _json({
        'name': dm.MARKETPLACE_NAME,
        'owner': {'name': dm.ORG_NAME},
        'metadata': {'description': dm.MARKETPLACE_DESCRIPTION},
        'plugins': [{
            'name': dm.PLUGIN_NAME,
            'source': f'./plugins/{dm.PLUGIN_NAME}',
            'description': dm.PLUGIN_DESCRIPTION,
        }],
    })


README_TEMPLATE = """# hojokin-plugins（生成物・手で直さない）

カラフルボックス社内向けの補助金業務プラグイン `{name}` の配布用リポジトリです。
組織設定の「プラグインとスキル」がこのリポジトリと同期し、メンバーの claude.ai（チャット・Cowork）と
Claude Code に届きます。

- 中身は開発リポジトリ（補助金ツール）の `scripts/build_plugin.py` が自動生成します。ここを手で直さないでください
  （次のビルドで上書きされます）。直したい内容は開発リポジトリ側で直してからビルドします。
- 現在の版: `{version}`（元コミット `{sha}`）。ファイルごとの sha256 は `BUILD_INFO.json` にあります。
- 既定ブランチに push すると、自動同期でメンバー全員に届きます。push の前に、差分に顧客名・従業員名・金額が
  無いことを確かめてください（`.githooks/pre-commit` が PII スキャンを実行します）。

## 更新の手順

1. 開発リポジトリで変更をコミットする
2. `python scripts/build_plugin.py --out <このリポジトリ>` を実行する（全ゲートに通らなければ何も書き換わらない）
3. このリポジトリで差分を確かめてコミットし、既定ブランチに push する

## 問題が起きたときの戻し方

1. すぐ止める: 組織設定 > プラグインとスキル で、このプラグインの配り方を「利用不可」にする
   （インストール済みの人にも止まる）
2. 直す: このリポジトリで `git revert <問題のコミット>` をして push する（開発リポジトリ側の原因コミットも revert する）
3. 配り方を元に戻す（「インストール可能」）

「Delete source」（同期元の削除）は使わないこと。同期したプラグインが組織から消え、元に戻せない。
"""

HOOK_SCRIPT = """#!/bin/sh
# プラグイン用リポジトリの pre-commit: ステージした追加行を PII スキャンする。
# 開発リポジトリの場所は HOJOKIN_DEV_REPO で変えられる。
DEV="${HOJOKIN_DEV_REPO:-$HOME/projects/カラフルボックス/補助金}"
if [ ! -f "$DEV/scripts/plugin_pii_scan.py" ]; then
  echo "❌ PII スキャナが見つかりません: $DEV/scripts/plugin_pii_scan.py（HOJOKIN_DEV_REPO を設定）" >&2
  exit 1
fi
python "$DEV/scripts/plugin_pii_scan.py" --staged || exit 1
"""


def repo_root_files(version: str, sha: str, plugin_files: dict[str, bytes]) -> dict[str, bytes]:
    root = {
        '.claude-plugin/marketplace.json': marketplace_json(),
        'README.md': README_TEMPLATE.format(name=dm.PLUGIN_NAME, version=version, sha=sha[:12]).encode(),
        '.gitignore': b'__pycache__/\n*.pyc\n',
        HOOK_PATH: HOOK_SCRIPT.encode('utf-8'),
    }
    everything = {**plugin_files, **root}
    root['BUILD_INFO.json'] = _json({
        'plugin': dm.PLUGIN_NAME,
        'version': version,
        'source_repo': 'hane-colorfulbox/hojokin',
        'source_commit': sha,
        'files': {p: hashlib.sha256(b).hexdigest() for p, b in sorted(everything.items())},
    })
    return root


def write_tree(dest: Path, files: dict[str, bytes]) -> None:
    for rel, data in files.items():
        p = dest / rel
        p.parent.mkdir(parents=True, exist_ok=True)
        p.write_bytes(data)


def _existing_files(dest: Path) -> set[str]:
    out = set()
    for p in dest.rglob('*'):
        rel = p.relative_to(dest).as_posix()
        if p.is_file() and rel.split('/', 1)[0] not in PRESERVED_TOP:
            out.add(rel)
    return out


def sync_repo(dest: Path, files: dict[str, bytes]) -> None:
    """dest を files と完全に一致させる（.git は触らない）。dest は空か git の作業コピーであること。"""
    dest = dest.resolve()
    if dest.exists() and any(dest.iterdir()) and not (dest / '.git').exists():
        raise SystemExit(f'❌ 出力先が git の作業コピーではありません（空でもない）: {dest}')
    dest.mkdir(parents=True, exist_ok=True)
    existing = _existing_files(dest)
    removed = sorted(existing - set(files))
    for rel in removed:
        (dest / rel).unlink()
    changed = [rel for rel, data in files.items()
               if not (dest / rel).is_file() or (dest / rel).read_bytes() != data]
    write_tree(dest, {rel: files[rel] for rel in changed})
    for d in sorted({p.parent for p in (dest / r for r in removed)}, key=lambda x: -len(x.parts)):
        if d != dest and d.exists() and not any(d.iterdir()):
            d.rmdir()
    print(f'   追加・変更 {len(changed)} / 削除 {len(removed)} / 合計 {len(files)} ファイル')


def write_plugin_zip(path: Path, repo_files: dict[str, bytes], plugin_root: str, name: str) -> None:
    """プラグインのフォルダ（最上位に1つ）だけを ZIP にする（claude.ai の個人アップロード用）。"""
    path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as zf:
        for rel, data in sorted(repo_files.items()):
            if rel.startswith(plugin_root):
                info = zipfile.ZipInfo(f'{name}/{rel[len(plugin_root):]}', ZIP_FIXED_TIME)
                info.compress_type = zipfile.ZIP_DEFLATED
                info.external_attr = 0o644 << 16
                zf.writestr(info, data)
