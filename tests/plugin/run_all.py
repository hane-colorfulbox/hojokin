"""配布プラグインのテストを一括で実行する（合成データのみ・API は呼ばない）。

1. build_plugin.build() で配布版を組み立て、リポジトリの外の一時フォルダに書き出す
   （ゲートは通さない。ゲートの検査は build_plugin.py 側の役目）
2. tests/plugin/test_*.py を1本ずつ別プロセスで実行する（作業フォルダは一時フォルダ）

使い方:
    python -B tests/plugin/run_all.py                  # HEAD から組み立てて実行
    python -B tests/plugin/run_all.py --from-worktree  # 未コミットの作業ツリーから（開発中の確認）
終了コード: 0=全件成功 / 1=失敗あり
"""
import argparse
import os
import subprocess
import sys
import tempfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
from _common import REPO, SKILLS_ENV, STAGE_ENV  # noqa: E402

import build_plugin  # noqa: E402
import plugin_output  # noqa: E402
from plugin_source import GitSource, WorktreeSource  # noqa: E402

NO_BILLING_ENV = {
    'CLAUDE_API_KEY': 'sk-ant-invalid-no-billing',
    'ANTHROPIC_API_KEY': 'sk-ant-invalid-no-billing',
    'ANTHROPIC_BASE_URL': 'http://127.0.0.1:9',
}


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument('--from-worktree', action='store_true')
    ap.add_argument('-k', help='名前にこの文字を含むテストだけ実行する')
    args = ap.parse_args()
    src = WorktreeSource(REPO) if args.from_worktree else GitSource(REPO)
    files, version = build_plugin.build(src)
    tests = sorted(p for p in HERE.glob('test_*.py') if not args.k or args.k in p.name)
    results = {}
    with tempfile.TemporaryDirectory(prefix='hojokin-plugin-test-') as tmp:
        stage = Path(tmp) / 'repo'
        plugin_output.write_tree(stage, files)
        env = dict(os.environ, **NO_BILLING_ENV, PYTHONIOENCODING='utf-8', PYTHONDONTWRITEBYTECODE='1')
        env[SKILLS_ENV] = str(stage / build_plugin.PLUGIN_ROOT / 'skills')
        env[STAGE_ENV] = str(stage)
        print(f'配布版 {version} を {stage} に組み立て（{len(files)} ファイル）\n')
        for t in tests:
            cp = subprocess.run([sys.executable, '-B', str(t)], cwd=tmp, env=env, capture_output=True,
                                text=True, encoding='utf-8', errors='replace')
            print(cp.stdout.rstrip())
            if cp.stderr.strip():
                print('[stderr]', cp.stderr.strip()[-1500:])
            results[t.name] = cp.returncode
            print()
    failed = [n for n, rc in results.items() if rc != 0]
    print('=' * 60)
    for name, rc in results.items():
        print(f'{"PASS" if rc == 0 else "FAIL"}  {name}')
    print(f'合計 {len(results)} 本 / 失敗 {len(failed)} 本')
    return 1 if failed else 0


if __name__ == '__main__':
    raise SystemExit(main())
