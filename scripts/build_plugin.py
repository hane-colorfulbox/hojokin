"""補助金業務の組織配布プラグイン（hojokin）を、git の HEAD にコミット済みの内容だけから組み立てる。

出力先はプラグイン用リポジトリ（colorful-box/hojokin-plugins・private・生成物専用）。組織設定の
「プラグインとスキル」がこのリポジトリと同期し、メンバーのチャット・Cowork・Claude Code に届く。
中身は scripts/distribution_manifest.py の許可リストで決まる。ゲート（plugin_gates.py と PII スキャン、
同期チェック、claude plugin validate）が1つでも NG なら何も出力しない。

使い方:
    python scripts/build_plugin.py --check-only                 # 組み立てて検査だけ（出力しない）
    python scripts/build_plugin.py --out ../hojokin-plugins      # プラグイン用リポジトリを更新
    python scripts/build_plugin.py --zip _dist/hojokin-plugin.zip   # 個人アップロード用の ZIP
    python scripts/build_plugin.py --retire-zip _dist/hojokin-handoff.zip  # 旧パックの退役版
    （開発中の確認だけ）--from-worktree で未コミットの作業ツリーから組み立てる（出力は不可）

終了コード: 0=成功 / 1=ゲート NG / 2=使い方・環境エラー
"""
import argparse
import hashlib
import json
import sys
import tempfile
from pathlib import Path

for _stream in (sys.stdout, sys.stderr):
    if hasattr(_stream, 'reconfigure'):
        _stream.reconfigure(encoding='utf-8', line_buffering=True)   # 標準出力とエラーの順序を保つ

sys.path.insert(0, str(Path(__file__).resolve().parent))

import distribution_manifest as dm  # noqa: E402
import plugin_output  # noqa: E402
from plugin_source import GitSource, WorktreeSource, SourceError  # noqa: E402
from plugin_checks import run_all_gates  # noqa: E402

ROOT = Path(__file__).resolve().parent.parent
PLUGIN_ROOT = f'plugins/{dm.PLUGIN_NAME}/'


def _excluded(path: str) -> bool:
    parts = path.split('/')
    name = parts[-1]
    return bool(dm.EXCLUDE_PARTS & set(parts)) or Path(name).suffix in dm.EXCLUDE_SUFFIXES \
        or any(kw in name for kw in dm.EXCLUDE_NAME_KEYWORDS)


def assemble_skills(src) -> dict[str, bytes]:
    """プラグイン内の相対パス（skills/<名前>/...）→ 中身。_build.json はまだ付けない。"""
    files: dict[str, bytes] = {}
    wanted = [dm.SELFCHECK_SOURCE]
    for spec in dm.SKILLS:
        wanted.append(spec.skill_md)
        for d in ([spec.source_dir] if spec.source_dir else []) + list(spec.mirror_dirs):
            wanted.extend(src.list_dir(d))
        wanted.extend(spec.mirror_files)
    src.prefetch(wanted)   # git から1回でまとめて読む
    for spec in dm.SKILLS:
        base = f'skills/{spec.name}/'
        files[base + 'SKILL.md'] = src.read(spec.skill_md)
        if spec.source_dir:
            for p in src.list_dir(spec.source_dir):
                rel = p[len(spec.source_dir) + 1:]
                if rel != 'SKILL.md' and not _excluded(p):
                    files[base + rel] = src.read(p)
        for d in spec.mirror_dirs:
            for p in src.list_dir(d):
                if p not in spec.mirror_exclude and not _excluded(p):
                    files[base + p] = src.read(p)
        for p in spec.mirror_files:
            files[base + p] = src.read(p)
        files[base + dm.SELFCHECK_NAME] = src.read(dm.SELFCHECK_SOURCE)
    return files


def make_version(src) -> str:
    return f'{src.commit_date():%Y.%m.%d}-{src.commit_sha()[:12]}' + src.version_suffix


def attach_build_info(skill_files: dict[str, bytes], version: str, src) -> dict[str, bytes]:
    out = dict(skill_files)
    for spec in dm.SKILLS:
        base = f'skills/{spec.name}/'
        listed = {p[len(base):]: hashlib.sha256(b).hexdigest()
                  for p, b in sorted(skill_files.items()) if p.startswith(base)}
        info = {'plugin': dm.PLUGIN_NAME, 'skill': spec.name, 'version': version,
                'source_commit': src.commit_sha(), 'source_commit_date': src.commit_date().isoformat(),
                'files': listed}
        out[base + dm.BUILD_INFO_NAME] = (json.dumps(info, ensure_ascii=False, indent=1) + '\n').encode()
    return out


def build(src) -> tuple[dict[str, bytes], str]:
    """(リポジトリ相対パス → 中身, 版) を返す。ゲートは呼び出し側で通す。"""
    version = make_version(src)
    skills = attach_build_info(assemble_skills(src), version, src)
    plugin = {PLUGIN_ROOT + p: b for p, b in skills.items()}
    plugin[PLUGIN_ROOT + '.claude-plugin/plugin.json'] = plugin_output.plugin_json(version)
    repo = dict(plugin)
    repo.update(plugin_output.repo_root_files(version, src.commit_sha(), plugin))
    return repo, version


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description='補助金業務の組織配布プラグインを組み立てる')
    ap.add_argument('--out', type=Path, help='プラグイン用リポジトリ（作業コピー）の場所')
    ap.add_argument('--zip', type=Path, help='個人アップロード用のプラグイン ZIP の出力先')
    ap.add_argument('--retire-zip', type=Path, help='旧パック（hojokin-handoff.zip）の退役版の出力先')
    ap.add_argument('--check-only', action='store_true', help='組み立てと検査だけ行う')
    ap.add_argument('--from-worktree', action='store_true', help='（開発中の確認用）作業ツリーから組み立てる')
    ap.add_argument('--skip-cli', action='store_true', help='claude CLI での検証を省く（CLI が無い環境用）')
    args = ap.parse_args(argv)
    if not (args.out or args.zip or args.retire_zip or args.check_only):
        ap.error('--out / --zip / --retire-zip / --check-only のどれかを指定してください')
    if args.from_worktree and (args.out or args.zip or args.retire_zip):
        ap.error('--from-worktree は検査専用です（出力はコミット済みの HEAD からだけ作ります）')

    try:
        src = WorktreeSource(ROOT) if args.from_worktree else GitSource(ROOT)
        repo_files, version = build(src)
    except SourceError as e:
        print(f'❌ 組み立てに失敗: {e}', file=sys.stderr)
        return 1
    print(f'組み立て: 版 {version}（元コミット {src.commit_sha()[:12]}）・{len(repo_files)} ファイル')

    with tempfile.TemporaryDirectory(prefix='hojokin-plugin-') as tmp:
        staging = Path(tmp) / 'repo'
        plugin_output.write_tree(staging, repo_files)
        if not run_all_gates(src, repo_files, staging, PLUGIN_ROOT, skip_cli=args.skip_cli):
            print('❌ ゲート NG のため出力しません', file=sys.stderr)
            return 1
        if args.check_only:
            print('✅ 全ゲート合格（--check-only のため出力なし）')
            return 0
        if args.out:
            plugin_output.sync_repo(args.out, repo_files)
            print(f'✅ プラグイン用リポジトリを更新: {args.out}（コミット・push は別に行う）')
        if args.zip:
            plugin_output.write_plugin_zip(args.zip, repo_files, PLUGIN_ROOT, dm.PLUGIN_NAME)
            print(f'✅ 個人アップロード用 ZIP: {args.zip}')
        if args.retire_zip:
            import retire_pack
            retire_pack.write_retire_zip(args.retire_zip, src)
            print(f'✅ 旧パックの退役版 ZIP: {args.retire_zip}')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
