#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""case-docs-check スキルの分類定数がツール正典と一致するかの build 前ゲート。

スキル同梱 `.claude/skills/case-docs-check/review/check_docs.py` は、配布先に
hojokin パッケージが無い前提で `hojokin/pipeline.py`（FileDetector）と `app.py`
の分類・必須判定定数を独立コピーしている。原本だけ直してスキル側を直し忘れると、
スキルの判定とツールの実挙動がズレたまま配布される。wagebook-convert の
「テンプレ xlsx sha256 一致ゲート」（build_skill_zip.check_template_sync）と
同じ構図で、ビルド時にここで突合して不一致なら中止する。

import ではなく AST でリテラルを抽出する理由:
- app.py は import すると Streamlit 実行が走る
- pipeline.py は anthropic 等の依存で headless 環境では import 自体が壊れうる
抽出に失敗した場合も沈黙せず NG を返す（誤 PASS を出さない）。

単体実行: python scripts/check_docscheck_sync.py  → exit 0=一致 / 1=不一致
"""
import ast
import sys
from pathlib import Path

sys.stdout.reconfigure(encoding='utf-8')
sys.stderr.reconfigure(encoding='utf-8')

ROOT = Path(__file__).resolve().parent.parent
PIPELINE_PY = ROOT / 'hojokin' / 'pipeline.py'
APP_PY = ROOT / 'app.py'
SKILL_CHECK_PY = ROOT / '.claude' / 'skills' / 'case-docs-check' / 'review' / 'check_docs.py'

# (正典ファイル, クラス名 or None, 正典側の変数名) → スキル側の変数名
SYNC_MAP = [
    (PIPELINE_PY, 'FileDetector', 'PATTERNS', 'PATTERNS'),
    (PIPELINE_PY, 'FileDetector', 'OUTPUT_FILE_MARKERS', 'OUTPUT_FILE_MARKERS'),
    (PIPELINE_PY, 'FileDetector', 'ALLOWED_EXTS', 'ALLOWED_EXTS'),
    (APP_PY, None, 'DRIVE_EXCLUDED_SUBFOLDERS', 'EXCLUDED_SUBFOLDERS'),
    (APP_PY, None, '_REGISTRY_CONFIRMED_KEYWORD', 'REGISTRY_CONFIRMED_KEYWORD'),
    (APP_PY, None, '_REQUIRED_CATS_BY_TASK', 'REQUIRED_CATS_BY_TASK'),
    (APP_PY, None, '_REQUIRED_CATS_APPLICATION_KOJIN', 'REQUIRED_CATS_APPLICATION_KOJIN'),
]


def _literals_from_body(body, names: set) -> dict:
    out = {}
    for node in body:
        if isinstance(node, ast.Assign):
            targets = [t.id for t in node.targets if isinstance(t, ast.Name)]
            value = node.value
        elif isinstance(node, ast.AnnAssign) and isinstance(node.target, ast.Name):
            targets = [node.target.id]
            value = node.value
        else:
            continue
        for name in targets:
            if name in names and value is not None:
                try:
                    out[name] = ast.literal_eval(value)
                except (ValueError, SyntaxError):
                    pass  # リテラルでない代入は「抽出失敗」として欠落させる
    return out


def _extract(path: Path, class_name, names: set) -> dict:
    tree = ast.parse(path.read_text(encoding='utf-8'))
    if class_name is None:
        return _literals_from_body(tree.body, names)
    for node in tree.body:
        if isinstance(node, ast.ClassDef) and node.name == class_name:
            return _literals_from_body(node.body, names)
    return {}


def _short(v, limit=160) -> str:
    s = repr(v)
    return s if len(s) <= limit else s[:limit] + '…'


def check_docs_sync() -> bool:
    """公開ラッパ。build_handoff_zip がスキル同梱前に定数一致を確認するために呼ぶ。"""
    for p in (PIPELINE_PY, APP_PY, SKILL_CHECK_PY):
        if not p.exists():
            print(f'❌ 突合対象が見つかりません: {p}', file=sys.stderr)
            return False

    src_cache = {}
    ok = True
    skill_names = {m[3] for m in SYNC_MAP}
    skill_vals = _extract(SKILL_CHECK_PY, None, skill_names)

    for path, class_name, src_name, skill_name in SYNC_MAP:
        key = (path, class_name)
        if key not in src_cache:
            src_cache[key] = _extract(path, class_name,
                                      {m[2] for m in SYNC_MAP
                                       if (m[0], m[1]) == key})
        src_vals = src_cache[key]
        where = f'{path.name}' + (f':{class_name}' if class_name else '')
        if src_name not in src_vals:
            print(f'❌ 正典から抽出できません: {where}.{src_name}', file=sys.stderr)
            ok = False
            continue
        if skill_name not in skill_vals:
            print(f'❌ スキル側から抽出できません: check_docs.py.{skill_name}',
                  file=sys.stderr)
            ok = False
            continue
        if src_vals[src_name] != skill_vals[skill_name]:
            print(f'❌ 定数不一致: {where}.{src_name} != check_docs.py.{skill_name}',
                  file=sys.stderr)
            print(f'   正典  : {_short(src_vals[src_name])}', file=sys.stderr)
            print(f'   スキル: {_short(skill_vals[skill_name])}', file=sys.stderr)
            ok = False

    if ok:
        print(f'✓ case-docs-check 定数同期OK（{len(SYNC_MAP)}項目）')
    return ok


def main() -> int:
    return 0 if check_docs_sync() else 1


if __name__ == '__main__':
    sys.exit(main())
