"""旧パック（ZIP 配布の hojokin-handoff.zip）を退役させるための最終版 ZIP を作る。

旧パックの利用者が「引き継ぎを最新に更新して」を実行すると、手元の scripts/update_handoff.py が
この ZIP を差分適用する。update_handoff.py の動き（ZIP に入っている管理フォルダは丸ごと入れ替え・
前回の一覧から消えたファイルは削除・配布版マーカー付きの CLAUDE.md だけ上書き）を使って、次の状態にする。
  - 旧スキル3本は「廃止・プラグインを使う」案内だけのスキルに置き換わる（自動では使われない）
  - 補助金フォルダ直下の配布版 CLAUDE.md は移行案内に置き換わる（本人が書いた CLAUDE.md は残し、
    移行案内は CLAUDE.handoff.md に置かれる）
  - 前回の更新の記録（.handoff_version）がある利用者だけ、前回配った docs/・hojokin/ などが削除される
    （手で展開しただけの利用者のものは残る。片付けは下の手順で本人が行う）
  - 引き継ぎ/ は ZIP に入れない（丸ごと入れ替えにすると、利用者がそこに置いたファイルまで消えるため）
build_plugin.py --retire-zip から呼ぶ。Drive の同じファイルの新しい版として上げるのは管理者本人。
手順の中身は hojokin-guide スキルの §0（legacy-migration ブロック）と揃えること。
"""
import json
import zipfile
from pathlib import Path

ZIP_FIXED_TIME = (2026, 1, 1, 0, 0, 0)
CLAUDE_MARKER = '<!-- handoff-distributed-claude-md -->'
OLD_SKILLS = {
    'wagebook-convert': '賃金台帳変換',
    'case-docs-check': '必要書類チェック',
    'katen-judge': '加点判定',
}

CLEANUP_STEPS = f"""1. claude.ai の「カスタマイズ」→「プラグイン」で、組織のプラグイン「hojokin」をインストールする
   （一度入れれば、以後は管理者が直すたびに自動で最新になる。チャット・Cowork・Claude Code のどれでも使える）
2. 旧パックを展開した補助金フォルダの `.claude/skills/` から、`wagebook-convert`・`case-docs-check`・`katen-judge` の
   3フォルダを削除する（自分で作ったスキルは残す）
3. 同じフォルダの `handoff_version.json`・`.handoff_version`・`CLAUDE.handoff.md` を削除する
4. 同じフォルダの `CLAUDE.md` の先頭行が `{CLAUDE_MARKER}` なら配布版なので削除する
   （そうでなければ自分の `CLAUDE.md` なので残し、補助金の案内を書き足していた部分だけを消す）
5. パソコン全体用の `~/.claude/skills/` に `wagebook-convert`・`case-docs-check`・`katen-judge` があれば削除する
   （同じ名前だとこちらが優先されるため、この廃止案内より古い版が使われてしまう）
6. 残っている `引き継ぎ/`・`docs/`・`hojokin/` などの古い資料は使わない（案件のファイルを置いていなければ削除してよい）
7. Claude Code を再起動する
"""

TOMBSTONE_SKILL = """---
name: {name}
description: （廃止）ZIP 配布時代の旧パックの{label}スキル。もう更新されない。組織のプラグイン hojokin に入っている {name} を使うこと。
disable-model-invocation: true
---

# このスキルは廃止されました

このフォルダの `{name}` は、ZIP で配っていた旧パックの版で、もう更新されません。
組織のプラグイン「hojokin」に入っている最新版を使ってください（Claude Code では `/hojokin:{name}`）。

Claude への指示: このスキルが呼ばれたら、作業はせずに、ユーザーへ次の手順を案内する（削除は本人が行う）。

{steps}"""

MIGRATION_CLAUDE_MD = f"""{CLAUDE_MARKER}
# 補助金業務（旧パックは廃止されました）

このフォルダの引き継ぎパック（hojokin-handoff.zip）は廃止し、組織のプラグイン「hojokin」で配るようになりました。
プラグインは管理者が直すたびに自動で最新になり、claude.ai のチャット・Cowork・Claude Code のどれでも使えます。

## Claude Code への指示

- 補助金の質問・作業では、組織プラグインの hojokin-guide スキル（`/hojokin:hojokin-guide`）を使う。
- このフォルダに残っている `引き継ぎ/`・`docs/`・`hojokin/` などの古い資料は参照しない。
- プラグインが入っていない場合は、下の手順1を案内する。削除などの片付けは本人が行う（Claude は案内だけ）。

## 片付けの手順

{CLEANUP_STEPS}"""


def retire_files() -> dict[str, bytes]:
    files = {f'.claude/skills/{name}/SKILL.md':
             TOMBSTONE_SKILL.format(name=name, label=label, steps=CLEANUP_STEPS).encode('utf-8')
             for name, label in OLD_SKILLS.items()}
    files['CLAUDE.md'] = MIGRATION_CLAUDE_MD.encode('utf-8')
    return files


def write_retire_zip(path: Path, src) -> None:
    files = retire_files()
    manifest = {
        'name': 'hojokin-handoff',
        'built_on': f'{src.commit_date():%Y-%m-%d}-retired',
        'retired': True,
        'file_count': len(files),
        'files': sorted(files),
    }
    path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as zf:
        for name, data in sorted(files.items()):
            info = zipfile.ZipInfo(name, ZIP_FIXED_TIME)
            info.compress_type = zipfile.ZIP_DEFLATED
            zf.writestr(info, data)
        info = zipfile.ZipInfo('handoff_version.json', ZIP_FIXED_TIME)
        info.compress_type = zipfile.ZIP_DEFLATED
        zf.writestr(info, json.dumps(manifest, ensure_ascii=False, indent=2))
