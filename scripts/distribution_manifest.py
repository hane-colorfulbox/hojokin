"""配布プラグイン（補助金業務）の中身を決める許可リストの単一の正本。

build_plugin.py がこれを読み、git の HEAD にコミット済みの内容だけからプラグインを組み立てる。
許可リストに挙げたものだけを配る（顧客実データや社外秘の資料は挙げない＝構造的に混入しない）。
新しく配るファイルを増やすときは、それが配布してよいもの（PII なし）か確かめてからここに足す。

スキルごとの置き方:
  - source_dir: そのフォルダの中身をスキル直下にそのまま置く（.claude/skills/<名前>/ の既存スキル）
  - mirror: リポジトリ相対のパスのまま、スキルのフォルダ内に再現する
    （例 docs/チェックリスト_加点判定.md → skills/katen-judge/docs/チェックリスト_加点判定.md）
    相対リンク（../docs/...）と、run_katen_judge.py の親フォルダ探索がそのまま効くようにするため。
チャット（claude.ai）ではスキルのフォルダしか実行環境にコピーされないので、依存するファイルは
必ず同じスキルの中に置く（スキルをまたぐリンクは張らず、文章で「○○スキルの…」と書く）。
"""
import string
import unicodedata
from dataclasses import dataclass, field

PLUGIN_NAME = 'hojokin'
MARKETPLACE_NAME = 'colorfulbox-hojokin'
ORG_NAME = 'カラフルボックス株式会社'
PLUGIN_DESCRIPTION = ('デジタル化・AI導入補助金（旧IT導入補助金）の申請支援業務の手引きとスキル一式'
                      '（業務ガイド・賃金台帳変換・加点判定・必要書類チェック）。社内専用。')
MARKETPLACE_DESCRIPTION = 'カラフルボックス社内向けの補助金業務プラグイン（社内専用・非公開）'

# プラグイン用リポジトリ（生成物専用・private）
PLUGIN_REPO = 'colorful-box/hojokin-plugins'

# 各スキルに共通で入れる文言ブロック（目印コメントで囲む）。正本は hojokin-guide の SKILL.md。
SHARED_BLOCKS = ('skill-paths', 'legacy-pack-check', 'chat-env', 'customer-data-rules')
SHARED_BLOCK_CANON = '引き継ぎ/配布用SKILL_hojokin-guide.md'

# スキルごとに生成して置くファイル（ソースはリポジトリ内）
SELFCHECK_SOURCE = 'scripts/plugin_selfcheck.py'
SELFCHECK_NAME = 'selfcheck.py'
BUILD_INFO_NAME = '_build.json'

ENGINE_FILES = (
    'hojokin/__init__.py',
    'hojokin/application_context.py',
    'hojokin/subsidy_rules.py',
    'hojokin/template_versions.py',
    'hojokin/config.py',
    'hojokin/models.py',
    'hojokin/wage_calculator.py',
    'hojokin/wage_reader.py',
    'hojokin/xlsx_surgical.py',
    'hojokin/bonus_wage_ledger_writer.py',
)

ROUND7_GUIDES = (
    'docs/第7次申請_変更点と使い方.md',
    'docs/第7次申請_確認待ち事項.md',
    'docs/第7次原本_改訂内容と再取込手順.md',
    'docs/第7次申請_ツール入力手順.md',
    'docs/第7次申請_公式資料で判断できることと確認の範囲.md',
)

GUIDE_DOCS = ROUND7_GUIDES + (
    'docs/運用マニュアル.md',
    'docs/マニュアル_書類作成.md',
    'docs/設計_API自動化.md',
    'docs/警告一覧.md',
    'docs/弊社成果物と書式.md',
    'docs/チェックリスト_加点判定.md',
    'docs/補助金_業務フロー全体マップ.md',
    'docs/補助金_制度リファレンス.md',
    'docs/制度一次情報/sources.json',
    # セカンドオピニオン加点の面談準備テンプレ（汎用版のみ・PII なし）。自社製品特化版・社長意向・
    # 設計・議事・案件別・レンダー物（*.pdf/*.xlsx）は挙げない＝配らない。
    # 実務プロンプトは 07b（確認メール）→ 06（予約フォーム入力文）の2本立て（07・06b は配らない）。
    'docs/セカンドオピニオン加点/配布物テンプレ/00_予約フォーム仕様.md',
    'docs/セカンドオピニオン加点/配布物テンプレ/05_面談ヒアリングシート.md',
    'docs/セカンドオピニオン加点/配布物テンプレ/05_面談ヒアリングシート_オンライン版.md',
    'docs/セカンドオピニオン加点/配布物テンプレ/05_別紙_お客様向け事前案内.md',
    'docs/セカンドオピニオン加点/配布物テンプレ/05_別紙_営業向けカンペ.md',
    'docs/セカンドオピニオン加点/配布物テンプレ/06_入力整理プロンプト.md',
    'docs/セカンドオピニオン加点/配布物テンプレ/07b_お客様確認メール生成プロンプト.md',
    'docs/セカンドオピニオン加点/配布物テンプレ/08_面談回答例_汎用中立版.md',
    'docs/セカンドオピニオン加点/配布物テンプレ/10_経営者向け_一問一答の回答例_汎用.md',
    'docs/セカンドオピニオン加点/配布物テンプレ/14_記入見本_予約フォーム入力_製造業例.md',
    'docs/セカンドオピニオン加点/配布物テンプレ/15_記入見本_経営者一問一答_製造業例.md',
    'docs/セカンドオピニオン加点/_sample/05_記入サンプル_ダミー工務店.md',
    'docs/セカンドオピニオン加点/_sample/07_記入サンプル_情報少なめ_ダミー工務店.md',
    # 通常枠×個人の申請書原本（Drive 正本に機械的修正を当てた配布版）。他3テンプレは Drive で配る。
    'ツール/【原本_個人】企業名_通常枠_個人2026.xlsx',
    'ツール/第7次以降/【原本_個人】企業名_通常枠_個人2026.xlsx',
)

KATEN_FILES = ENGINE_FILES + (
    'ツール/第7次以降/加点判定用賃金台帳テンプレート.xlsx',
    '補助金加点/第7次以降/加点措置①用.xlsx',
    '補助金加点/第7次以降/加点措置②用.xlsx',
    '補助金加点/第7次以降/補助率引き上げ・加点措置①用.xlsx',
    'ツール/加点判定用賃金台帳テンプレート.xlsx',
    '補助金加点/加点措置①用.xlsx',
    '補助金加点/加点措置②用.xlsx',
    '補助金加点/補助率引き上げ・加点措置①用.xlsx',
    'docs/チェックリスト_加点判定.md',
)


@dataclass(frozen=True)
class SkillSpec:
    name: str
    skill_md: str                              # SKILL.md の元（リポジトリ相対）
    source_dir: str | None = None              # 中身をスキル直下に置くフォルダ
    mirror_dirs: tuple[str, ...] = ()          # リポジトリ相対のまま再現するフォルダ
    mirror_files: tuple[str, ...] = ()         # リポジトリ相対のまま再現するファイル
    mirror_exclude: tuple[str, ...] = field(default_factory=tuple)
    source_extra: tuple[str, ...] = ()         # source_dir の中で、拡張子の許可外でも配るファイル（名前で挙げる）


# フォルダごと取り込む場所（source_dir・mirror_dirs）で配ってよい拡張子。これ以外のファイル
# （開発中に置いた PDF・成果物の xlsx など）は、source_extra で名前を挙げない限りビルドを止める。
FOLDER_SUFFIXES_OK = frozenset({'.md', '.py', '.json'})


SKILLS = (
    SkillSpec(
        name='hojokin-guide',
        skill_md=SHARED_BLOCK_CANON,
        mirror_dirs=('引き継ぎ',),
        mirror_files=GUIDE_DOCS,
        mirror_exclude=(SHARED_BLOCK_CANON,),
    ),
    SkillSpec(name='wagebook-convert', skill_md='.claude/skills/wagebook-convert/SKILL.md',
              source_dir='.claude/skills/wagebook-convert',
              mirror_files=ROUND7_GUIDES,
              source_extra=('.claude/skills/wagebook-convert/templates/賃金台帳テンプレート.xlsx',)),
    SkillSpec(name='katen-judge', skill_md='.claude/skills/katen-judge/SKILL.md',
              source_dir='.claude/skills/katen-judge', mirror_files=KATEN_FILES + ROUND7_GUIDES),
    SkillSpec(name='case-docs-check', skill_md='.claude/skills/case-docs-check/SKILL.md',
              source_dir='.claude/skills/case-docs-check', mirror_files=ROUND7_GUIDES),
)

# 同期ゲートが読むファイル（作業ツリーに未コミットの差分があるとビルドを止める）
SYNC_GATE_INPUTS = (
    'hojokin/pipeline.py',
    'app.py',
    'ツール/賃金台帳テンプレート.xlsx',
    'ツール/ヒアリングシート2026_通常枠個人.xlsx',
    'hojokin/config.py',
)

# 念のための除外（許可フォルダにうっかり入った生成物・機密を弾く二重の安全網）
EXCLUDE_PARTS = frozenset({'__pycache__', '_transcripts', '案件別', '.pytest_cache'})
EXCLUDE_SUFFIXES = frozenset({'.pyc', '.pyo', '.log', '.b64'})
EXCLUDE_NAME_KEYWORDS = ('賃金台帳_', '賃金台帳一覧', '決算', '給与明細', 'ヒアリング結果', '社長意向', '議事_',
                         '参考_', '_加点措置', '_AI版', '_給与支給総額')   # 顧客資料・成果物の典型的な名前

# プラグインの制約（claude.ai へのアップロード・組織同期）
ALLOWED_FRONTMATTER_KEYS = frozenset({'name', 'description', 'license', 'compatibility', 'metadata',
                                      'allowed-tools'})
SKILL_NAME_PATTERN = r'^[a-z0-9]+(-[a-z0-9]+)*$'
SKILL_NAME_MAX = 64
DESCRIPTION_MAX = 1024
MAX_FILES = 5000
MAX_TOTAL_BYTES = 200 * 1024 * 1024

# 配布するファイルのパスに使える文字。claude.ai は記号入りのパスを含む ZIP を
# 「Zip file contains path with invalid characters」で拒否する（2026-09-30 実測: 漢字・かなは可、
# 【】①②・ は不可）。確認できた範囲だけを許可し、それ以外の文字は置き換えるかゲートで止める。
PATH_ASCII_OK = frozenset(string.ascii_letters + string.digits + '._-/')
PATH_UNICODE_CATEGORIES_OK = frozenset({'Lo', 'Lm', 'Lu', 'Ll'})
# 置き換え（katen-judge の run_katen_judge.py の DIST_NAME_REPLACEMENTS と同じ内容にする）
PATH_CHAR_REPLACEMENTS = (('【', ''), ('】', '_'), ('①', '1'), ('②', '2'), ('・', '_'))


def dist_path(path: str) -> str:
    """リポジトリ内のパスを、配布物の中のパスにする（claude.ai が受け付けない記号を置き換える）。"""
    for old, new in PATH_CHAR_REPLACEMENTS:
        path = path.replace(old, new)
    return path


def path_chars_ok(path: str) -> bool:
    return all(ch in PATH_ASCII_OK or (not ch.isascii() and unicodedata.category(ch) in PATH_UNICODE_CATEGORIES_OK)
               for ch in path)


# 配布物に残っていてはいけない古い語（ZIP 配布時代の手順）。例外は build_plugin 側で扱う。
STALE_TERMS = ('hojokin-handoff.zip', 'update_handoff', '引き継ぎを最新に更新', '.claude/skills/',
               '{SKILL_DIR}', '{ROOT}/', 'wagebook-convert.zip')
