"""配布プラグインの全ゲートを順に実行して結果を表示する（build_plugin.py から呼ぶ）。

1. HEAD との整合: 配るファイルと同期チェックの入力に、未コミットの変更が無い
2. 同期: 賃金台帳テンプレ（スキル同梱＝ツール原本）、case-docs-check の定数、通常枠個人の原本
3. claude.ai の規則・リンク・古い語・共通ブロック・xlsx のメタデータ（plugin_gates.py）
4. PII スキャン（plugin_pii_scan.py。名簿が無い・少ないときも NG）
5. claude plugin validate --strict（マーケットプレイスとプラグイン）と、スキル4本の読み込み確認
"""
import re
import shutil
import subprocess
import sys
from pathlib import Path

import distribution_manifest as dm
import plugin_gates as gates
import plugin_pii_scan as pii

ROOT = Path(__file__).resolve().parent.parent
WAGEBOOK_TEMPLATE_PAIR = ('.claude/skills/wagebook-convert/templates/賃金台帳テンプレート.xlsx',
                          'ツール/賃金台帳テンプレート.xlsx')
GENPON_VERIFIER = 'scripts/verify_genpon_tsujo_kojin.py'
CLI_TIMEOUT_SEC = 120


def _report(title: str, problems: list[str], limit: int = 40) -> bool:
    if not problems:
        print(f'  ✅ {title}')
        return True
    print(f'  ❌ {title}（{len(problems)} 件）')
    for msg in problems[:limit]:
        print(f'     - {msg}')
    if len(problems) > limit:
        print(f'     … ほか {len(problems) - limit} 件')
    return False


def gate_head_consistency(src) -> list[str]:
    dirty = src.dirty_paths(src.read_paths() | set(dm.SYNC_GATE_INPUTS) | set(WAGEBOOK_TEMPLATE_PAIR))
    return [f'未コミットの変更があります: {p}' for p in dirty]


def gate_sync(src) -> list[str]:
    problems = []
    skill_tpl, tool_tpl = WAGEBOOK_TEMPLATE_PAIR
    for p in WAGEBOOK_TEMPLATE_PAIR:
        if not src.exists(p):
            problems.append(f'賃金台帳テンプレがありません: {p}')
    if not problems and src.read(skill_tpl) != src.read(tool_tpl):
        problems.append('スキル同梱の賃金台帳テンプレとツール原本が一致しません（両方を揃えてから配布）')
    sys.path.insert(0, str(ROOT / 'scripts'))
    import check_docscheck_sync
    if not check_docscheck_sync.check_docs_sync():
        problems.append('case-docs-check の分類定数が hojokin/pipeline.py・app.py と一致しません')
    cp = subprocess.run([sys.executable, '-B', str(ROOT / GENPON_VERIFIER)], cwd=ROOT, capture_output=True,
                        text=True, encoding='utf-8', errors='replace')
    if cp.returncode != 0:
        tail = (cp.stdout + cp.stderr).strip().splitlines()[-3:]
        problems.append(f'通常枠個人の原本の検証が NG（{GENPON_VERIFIER}）: {" / ".join(tail)}')
    return problems


def gate_content(plugin_files: dict[str, bytes]) -> list[tuple[str, list[str]]]:
    canon = 'skills/hojokin-guide/SKILL.md'
    all_blocks = dm.SHARED_BLOCKS
    required = {f'skills/{s.name}/SKILL.md': all_blocks for s in dm.SKILLS}
    for doc in ('00_はじめに.md', '06_運用ルールと体制.md'):
        required[f'skills/hojokin-guide/引き継ぎ/{doc}'] = ('customer-data-rules',)
    return [
        ('claude.ai の frontmatter 規則', gates.check_frontmatter(plugin_files)),
        ('プラグインの構成（bin/・件数・サイズ）', gates.check_layout(plugin_files, '')),
        ('相対リンク（スキルのフォルダ内で解決）', gates.check_links(plugin_files)),
        ('ZIP 配布時代の古い語', gates.check_stale_terms(plugin_files)),
        ('共通ブロックの一致', gates.check_shared_blocks(plugin_files, canon, required)),
        ('xlsx のメタデータ', gates.check_xlsx_metadata(plugin_files)),
    ]


def gate_pii(staging: Path) -> list[str]:
    try:
        hits = pii.run([str(staging)])
    except (FileNotFoundError, ValueError) as e:
        return [f'PII スキャンを実行できません: {e}']
    return [f'{h.path}:{h.line_no} [{h.rule}] {h.masked()}' for h in hits]


def _claude(*args: str) -> subprocess.CompletedProcess:
    exe = shutil.which('claude')   # Windows では claude.CMD などのラッパーを実体パスで呼ぶ必要がある
    if exe is None:
        raise OSError('claude コマンドが見つかりません')
    return subprocess.run([exe, *args], capture_output=True, text=True, encoding='utf-8',
                          errors='replace', timeout=CLI_TIMEOUT_SEC)


def gate_cli(staging: Path, plugin_dir: Path) -> list[str]:
    problems = []
    for target in (staging, plugin_dir):
        try:
            cp = _claude('plugin', 'validate', str(target), '--strict', '--json')
        except (OSError, subprocess.TimeoutExpired) as e:
            return [f'claude CLI を実行できません（--skip-cli で省略可）: {e}']
        if cp.returncode != 0:
            problems.append(f'claude plugin validate NG（{target.name}）: {(cp.stdout or cp.stderr).strip()[:600]}')
    cp = _claude('--plugin-dir', str(plugin_dir), 'plugin', 'details', dm.PLUGIN_NAME)
    out = cp.stdout + cp.stderr
    m = re.search(r'Skills\s*\((\d+)\)', out)
    want = len(dm.SKILLS)
    if cp.returncode != 0 or not m or int(m.group(1)) != want:
        problems.append(f'claude plugin details でスキル {want} 本が見えません: {out.strip()[:400]}')
    return problems


def run_all_gates(src, repo_files: dict[str, bytes], staging: Path, plugin_root: str,
                  skip_cli: bool = False) -> bool:
    plugin_files = {p[len(plugin_root):]: b for p, b in repo_files.items() if p.startswith(plugin_root)}
    print('ゲート:')
    ok = _report('HEAD との整合（未コミットの変更なし）', gate_head_consistency(src))
    ok &= _report('同期（テンプレ・定数・通常枠個人の原本）', gate_sync(src))
    for title, problems in gate_content(plugin_files):
        ok &= _report(title, problems)
    ok &= _report('PII スキャン（語は伏せ字）', gate_pii(staging))
    if skip_cli:
        print('  ⚠ claude CLI での検証は省略しました（--skip-cli）')
    else:
        ok &= _report('claude plugin validate / details', gate_cli(staging, staging / plugin_root))
    return ok
