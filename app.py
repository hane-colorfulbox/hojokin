# -*- coding: utf-8 -*-
"""
補助金書類自動作成 Webアプリ
Streamlit で動作するシンプルなUI
"""
from __future__ import annotations

import sys
import os
import shutil
import tempfile
import logging
import re
import unicodedata
import hashlib
from pathlib import Path
from datetime import datetime


def _nfc_filename(name: str) -> str:
    """ファイル名を NFC 正規化する。

    macOS のファイルシステムやブラウザは日本語ファイル名を NFD（濁点・半濁点を分離）で
    送ってくることがあり、Linux サーバ上で NFC キーワードと比較したときに一致しない。
    保存時点で必ず NFC に揃えることで、Drive モード／アップロードモードの差異と
    Mac/Windows ローカルの差異を吸収する。
    """
    return unicodedata.normalize('NFC', name)


def _safe_company_name(name: str) -> str:
    """会社名をファイル名に使える形にサニタイズする。

    会社名は出力ファイル名（例: ○○株式会社_通常枠_2026_AI版.xlsx）に直接連結されるため、
    Windows で使えない文字（/ \\ : * ? " < > |）が含まれているとパス解釈が壊れて
    FileNotFoundError になる。実例: 会社名「テスト5/1」→ パスが「テスト5\1_...」と解釈され失敗。
    禁止文字を - に置換し、制御文字・先頭末尾の空白は除去する。
    """
    if not name:
        return ''
    # Windows 禁止文字 + 改行・タブ
    forbidden = '/\\:*?"<>|\r\n\t'
    safe = ''.join('-' if c in forbidden else c for c in name)
    # 制御文字（U+0000-U+001F, U+007F）の除去
    safe = ''.join(c for c in safe if c.isprintable())
    return safe.strip()

import streamlit as st

# .env読み込み
from dotenv import load_dotenv
load_dotenv()

# ── Streamlit Cloud Secrets を os.environ に橋渡し ──
# 【重要】hojokin.* の import より前に実行する必要がある。
# config.py は import された時点で os.getenv('USE_*') を評価するため、
# Secrets 展開がそれより後だとフラグがデフォルト値で固定される（前バグ）。
# 既に環境変数で設定済みのキーは上書きしない（ローカル .env / OSの環境変数を優先）。
# dict 値（gcp_service_account など TOML セクション）は別経路で読むため対象外。
try:
    if hasattr(st, 'secrets'):
        for _k, _v in st.secrets.items():
            if isinstance(_v, (str, int, float, bool)) and _k not in os.environ:
                os.environ[_k] = str(_v)
except Exception:
    # ローカル開発で secrets.toml が無い場合は無害（os.getenv は load_dotenv 経由で .env を読む）
    pass

# ── gcp_service_account TOML セクションを JSON 文字列として env に橋渡し ──
# document_ai_ocr.py など、ファイル配置できないクラウド環境で Service Account 認証を
# 必要とするモジュールは、この環境変数を読んで from_service_account_info() を呼ぶ。
# Drive 連携は credentials_dict 引数で直接渡す既存経路があるためここでは触らない。
try:
    if hasattr(st, 'secrets') and 'gcp_service_account' in st.secrets:
        if 'GOOGLE_SERVICE_ACCOUNT_JSON_CONTENT' not in os.environ:
            import json as _json
            os.environ['GOOGLE_SERVICE_ACCOUNT_JSON_CONTENT'] = _json.dumps(
                dict(st.secrets['gcp_service_account'])
            )
except Exception:
    pass

# パッケージパス追加
sys.path.insert(0, str(Path(__file__).parent))

from hojokin.ai_extractor import create_extractor
from hojokin.config import CLAUDE_API_KEY, detect_prefecture
from hojokin.pipeline import (
    FileDetector, run_application_transfer, run_wage_calculation,
    run_wage_ledger_conversion, run_bonus_wage_ledger_creation,
)
from hojokin.wage_reader import (
    judge_bonus_points, read_bonus_wage_ledger, is_bonus_wage_ledger,
    fill_bonus_sheet_1, fill_bonus_sheet_2,
)

# Drive連携（認証情報がある場合のみ）
logger = logging.getLogger(__name__)

_drive_client = None
_DRIVE_CREDS = os.getenv('GOOGLE_SERVICE_ACCOUNT_JSON', '')
_DRIVE_PARENT_ID = os.getenv('DRIVE_PARENT_FOLDER_ID', '')


def _has_streamlit_secret(key: str) -> bool:
    """st.secrets に key が存在するかを安全に判定する。

    ローカル開発で secrets.toml が無い場合、`'key' in st.secrets` が
    StreamlitSecretNotFoundError を出すため try で包む。Cloud では正常動作。
    """
    try:
        return hasattr(st, 'secrets') and key in st.secrets
    except Exception:
        return False


def _get_drive_client():
    global _drive_client
    if _drive_client is not None:
        return _drive_client

    from hojokin.drive_client import DriveClient

    # 方法1: ローカルのJSONファイル
    if _DRIVE_CREDS and Path(_DRIVE_CREDS).exists():
        _drive_client = DriveClient(credentials_path=_DRIVE_CREDS)
        return _drive_client

    # 方法2: Streamlit Secrets（Cloud用）
    # ローカルで secrets.toml が無い場合や、Cloud で形式不正な場合を吸収する。
    # 失敗時は return None でローカル認証経路へフォールバックさせる設計。
    # 観測性のため WARN 出力（ロガーは Streamlit Cloud のログから確認可能）。
    try:
        if _has_streamlit_secret('gcp_service_account'):
            _drive_client = DriveClient(
                credentials_dict=dict(st.secrets['gcp_service_account']),
            )
            return _drive_client
    except Exception as e:
        logger.warning(
            f'Streamlit Secrets経由のDrive認証に失敗: {e}',
            exc_info=True,
        )

    return None


# ── Drive 取得用キャッシュ関数（モジュールレベル定義 = cache_key が引数で確定する） ──
# Drive 案件フォルダ配下で、配下まで降りずに無視するサブフォルダ名。
# 「申請時使用」は税理士納品の要約版PDFを置くカラフル運用上のサブフォルダ。
# 要約版PDFは販管費の科目内訳が無く R216 算定が壊れるため、ツールは見に行かない。
# 将来「2.実績報告」「アーカイブ」等で同様の事象が出たらここに追加するだけで対応。
DRIVE_EXCLUDED_SUBFOLDERS: set[str] = {'申請時使用'}


@st.cache_data(ttl=60)
def _cached_list_drive_folders(folder_id: str) -> list[dict]:
    c = _get_drive_client()
    return c.list_folders(folder_id) if c else []


@st.cache_data(ttl=30)
def _cached_list_drive_files_recursive(folder_id: str) -> list[dict]:
    c = _get_drive_client()
    return c.list_files_recursive(
        folder_id, exclude_folder_names=DRIVE_EXCLUDED_SUBFOLDERS
    ) if c else []


# ── 定数 ──
TEMPLATE_OPTIONS = {
    '通常枠 2026（法人）': '通常枠_2026',
    'インボイス枠 2026（法人）': 'インボイス枠_2026',
    'インボイス枠 2026（個人）': 'インボイス枠_個人_2026',
}

TASK_OPTIONS = {
    # 推奨フロー順:
    #   ① 賃金台帳の作成 → ② 一人当たり給与支給総額 → ③ 申請書作成
    # 賃金台帳が標準テンプレ形式 + 人間チェック済の状態を作ってから申請書を作る。
    # 申請書作成タスクは賃金台帳を AI 再抽出しない（決定論パーサー一本）ため、
    # ②と同じ決定論パーサーで R215/R216 を算出する。中途入退社が多い案件は
    # 注意喚起を出すが自動転記は行い、②と数値が一致する設計。
    # 例外: 賃金台帳の記録期間が直近事業年度から完全に外れている疑いがあるケース
    # （全月在籍者0名）のみ hard stop して手動入力を促す。
    #
    # 旧「給与計算のみ(wage)」「両方(all)」は UI から除外（2026-06）。
    # 申請書作成タスクが「給与支給総額計算」「従業員別明細」「生産性指標」シートを
    # AI版.xlsx に内包する（pipeline.py の per_employee_wage シート統合）ため完全に冗長。
    # コードパス（task_type 'wage'/'all'）は CLI(run.py)用に残置。
    '賃金台帳の作成': 'wage_ledger_creation',
    '一人当たり給与支給総額': 'per_employee_wage',
    '申請書作成': 'application',
    # 賃上げ加点は2工程: ①加点判定用賃金台帳を作る → ②それを入力に加点判定。
    # 加点判定は「基本給÷所定労働時間＝時間換算給与」を暦月固定（R6/10〜R7/9＋申請直近月）で
    # 見るため、R215/R216用の標準賃金台帳とは別物の専用台帳が必要（標準台帳には基本給・
    # 所定労働時間・暦月固定列が無く、正社員が判定からこぼれる）。
    '加点判定用賃金台帳の作成': 'bonus_wage_ledger_creation',
    '加点判定': 'bonus',
}

# 加点項目の全体像（デジタル化・AI導入補助金2026 通常枠 公募要領 p.24-27、2026-06-01 確認）。
# 本ツールが賃金台帳から自動判定できるのは 14)・15) のみ。他は ITツール選定・外部認定・
# 賃上げ計画の表明など賃金台帳の外で確認する項目なので、手動チェック用に列挙する。
BONUS_ITEMS_REFERENCE_MD = """\
**デジタル化・AI導入補助金2026 通常枠の加点項目**（公募要領 p.24-26）

本ツールが賃金台帳から自動判定するのは **14)・15) の2項目のみ**です。残りは別途ご確認ください。

| # | 加点項目 | 判定 |
|---|---|---|
| 1) | 導入ITツールがクラウド製品 | 手動（ITツール選定）|
| 2) | サイバーセキュリティお助け隊サービスを選定 | 手動 |
| 3) | インボイス制度対応製品を選定 | 手動 |
| 4) | デジタル化セカンドオピニオンの取組み（第3回公募回〜）| 手動 |
| 5)-8) | 賃上げ計画（事業場内最低賃金 +30円/+50円・給与支給総額の年平均成長率 3%/3.5%・計画表明。補助金額と過去交付有無で要件が変わる）| 手動（申請書側で表明）|
| 9) | IT戦略ナビwith を申請前に実施 | 手動 |
| 10) | 健康経営優良法人2026 に認定 | 手動 |
| 11) | えるぼし／くるみん等の認定 | 手動 |
| 12) | 成長加速マッチングサービスで課題登録 | 手動 |
| 13) | 省力化ナビ活用 | 手動 |
| **14)** | 令和6年10月〜令和7年9月で「R7改定後最低賃金 未満」雇用の従業員が全従業員の30%以上の月が3か月以上（**補助率1/2→2/3 のトリガー** 兼 加点）| **自動（加点措置①）**|
| **15)** | 交付申請直近月の事業場内最低賃金 ≥ 令和7年7月＋63円 | **自動（加点措置②）**|

※ 減点項目あり（過去のIT/デジタル化補助金の交付決定、インボイス枠との機能重複申請、プロセス重複、賃上げ計画の未達歴 等。公募要領 p.27）。
"""

# ── ページ設定 ──
st.set_page_config(
    page_title='補助金書類自動作成',
    page_icon='📋',
    layout='wide',
)

# ── スタイル ──
st.markdown("""
<style>
    .main-header {
        font-size: 2.2rem;
        font-weight: bold;
        margin-bottom: 0.2rem;
        color: #1a1a2e;
    }
    .sub-header {
        color: #555;
        font-size: 1.1rem;
        margin-bottom: 1.5rem;
    }
    .step-number {
        display: inline-block;
        background: #0068c9;
        color: white;
        width: 2rem;
        height: 2rem;
        border-radius: 50%;
        text-align: center;
        line-height: 2rem;
        font-weight: bold;
        margin-right: 0.5rem;
    }
    .step-title {
        font-size: 1.3rem;
        font-weight: bold;
        color: #1a1a2e;
    }
    .file-card {
        background: #f8f9fa;
        border: 1px solid #dee2e6;
        border-radius: 0.5rem;
        padding: 1rem;
        margin: 0.3rem 0;
    }
    .file-required {
        border-left: 4px solid #ff4b4b;
    }
    .file-optional {
        border-left: 4px solid #21c354;
    }
    .badge-required {
        background: #ff4b4b;
        color: white;
        padding: 0.15rem 0.5rem;
        border-radius: 0.8rem;
        font-size: 0.75rem;
        font-weight: bold;
    }
    .badge-optional {
        background: #21c354;
        color: white;
        padding: 0.15rem 0.5rem;
        border-radius: 0.8rem;
        font-size: 0.75rem;
        font-weight: bold;
    }
    .keyword-tag {
        display: inline-block;
        background: #e8f0fe;
        color: #1967d2;
        padding: 0.1rem 0.5rem;
        border-radius: 0.3rem;
        font-size: 0.85rem;
        font-weight: bold;
        margin: 0.1rem;
    }
    .stFileUploader > div > div {
        padding: 2rem;
    }
    .how-it-works {
        background: #f0f7ff;
        border-radius: 0.5rem;
        padding: 1rem 1.5rem;
        margin: 1rem 0;
    }
</style>
""", unsafe_allow_html=True)


def _find_wage_ledger_template(base_dir: Path) -> Path | None:
    """賃金台帳テンプレート（ツール/賃金台帳テンプレート.xlsx）を解決する。

    探索順:
      1. base_dir / 'ツール' / '賃金台帳テンプレート.xlsx'
      2. base_dir / '賃金台帳テンプレート.xlsx'
      3. プロジェクトルート（このファイルの親） / 'ツール' / '賃金台帳テンプレート.xlsx'

    存在するファイルを最初にヒットした順で返す。見つからなければ None。
    """
    target_name = '賃金台帳テンプレート.xlsx'
    candidates = [
        base_dir / 'ツール' / target_name,
        base_dir / target_name,
        Path(__file__).parent / 'ツール' / target_name,
    ]
    for p in candidates:
        if p.exists() and not p.name.startswith('~$'):
            return p
    return None


def _find_bonus_wage_ledger_template(base_dir: Path) -> Path | None:
    """加点判定用賃金台帳テンプレート（ツール/加点判定用賃金台帳テンプレート.xlsx）を解決する。"""
    target_name = '加点判定用賃金台帳テンプレート.xlsx'
    candidates = [
        base_dir / 'ツール' / target_name,
        base_dir / target_name,
        Path(__file__).parent / 'ツール' / target_name,
    ]
    for p in candidates:
        if p.exists() and not p.name.startswith('~$'):
            return p
    return None


def _parse_app_month_input(s: str) -> tuple[int, int] | None:
    """交付申請月の入力（yyyy/mm 等）を (西暦年, 月) に解釈する。失敗時 None。"""
    if not s:
        return None
    m = re.match(r'^\s*(\d{4})[\s/\-.年]+(\d{1,2})', str(s))
    if m:
        year, month = int(m.group(1)), int(m.group(2))
        if 1 <= month <= 12:
            return (year, month)
    return None


def _parse_bonus_months_input(s: str) -> list[tuple[int, int]] | None:
    """賞与の支給月入力（yyyy/mm のカンマ区切り）を [(年, 月), ...] に解釈する。

    空入力や全トークン不正なら None（＝指定なし扱い）。一部だけ解釈できた場合は
    解釈できた分のみ返す。各トークンは _parse_app_month_input を再利用して解釈する。
    """
    if not s or not s.strip():
        return None
    months: list[tuple[int, int]] = []
    for tok in re.split(r'[,、，]', s):
        parsed = _parse_app_month_input(tok.strip())
        if parsed:
            months.append(parsed)
    return months or None


def find_template(base_dir: Path, template_type: str) -> Path | None:
    """テンプレートファイルを検索（v2を優先）

    ルートとツール/の両方から候補を集め、v2を優先して返す。
    """
    import unicodedata
    keywords = {
        '通常枠_2026': ['原本', '通常枠', '2026'],
        'インボイス枠_2026': ['原本', 'インボイス', '法人', '2026'],
        'インボイス枠_個人_2026': ['原本', 'インボイス', '個人', '2026'],
    }
    kws = keywords.get(template_type, [])
    candidates = []
    # ルートとツール/の両方から候補を集める
    search_dirs = [base_dir]
    tool_dir = base_dir / 'ツール'
    if tool_dir.exists():
        search_dirs.append(tool_dir)
    for d in search_dirs:
        for p in d.iterdir():
            name = unicodedata.normalize('NFC', p.name)
            if p.suffix == '.xlsx' and all(kw in name for kw in kws) and not name.startswith('~$'):
                candidates.append(p)
    if not candidates:
        return None
    # v2を優先（ファイル名に'v2'が含まれるものを優先）
    for c in candidates:
        if 'v2' in c.name:
            return c
    return candidates[0]


def save_uploaded_files(uploaded_files, target_dir: Path) -> list[str]:
    """アップロードファイルを一時ディレクトリに保存（ファイル名は NFC 統一）"""
    saved = []
    for f in uploaded_files:
        name = _nfc_filename(f.name)
        dest = target_dir / name
        dest.write_bytes(f.getvalue())
        saved.append(name)
    return saved


def run_processing(
    company_name: str,
    template_type: str,
    task_type: str,
    work_dir: Path,
    template_dir: Path,
    progress_callback=None,
    prefecture: str = '',
    application_ym: tuple[int, int] | None = None,
    fiscal_month_override: int | None = None,
    has_cost_report_hint: bool = False,
    selection_override: dict[str, list[Path]] | None = None,
    bonus_paid_months: list[tuple[int, int]] | None = None,
):
    """メイン処理を実行"""
    results = {}

    # Extractor作成（加点判定用賃金台帳の作成は PDF/Excel から AI 抽出する）。
    # 加点判定（bonus）は専用台帳を決定論で直読みするため AI 不要。
    extractor = None
    if task_type in ('application', 'wage', 'per_employee_wage', 'all',
                     'bonus_wage_ledger_creation'):
        def _on_api_retry(attempt: int, max_attempts: int, wait: float, err: str):
            # Anthropic APIの一時エラー（422/429/5xx/529/timeout等）時の再試行をユーザーに通知
            try:
                st.toast(
                    f'API一時エラー ({err}) — {wait}秒後に再試行します（試行 {attempt}/{max_attempts}）',
                    icon='⚠️',
                )
            except Exception:
                pass  # UI表示に失敗しても処理は継続

        extractor = create_extractor(CLAUDE_API_KEY, retry_callback=_on_api_retry)

    if task_type in ('application', 'all'):
        if progress_callback:
            progress_callback('申請書を作成中...')

        template_path = find_template(template_dir, template_type)
        if template_path is None:
            # work_dir内も探す
            template_path = find_template(work_dir, template_type)

        if template_path is None:
            results['application'] = {
                'status': 'エラー',
                'message': 'テンプレートファイルが見つかりません。原本Excelもアップロードしてください。',
            }
        else:
            output_path = work_dir / f'{company_name}_{template_type}_AI版.xlsx'
            status = run_application_transfer(
                resource_folder=work_dir,
                template_path=template_path,
                template_type=template_type,
                output_path=output_path,
                extractor=extractor,
                fiscal_month_override=fiscal_month_override,
                has_cost_report_hint=has_cost_report_hint,
                selection_override=selection_override,
            )
            # 追加出力ファイル（賃金台帳一覧等）を収集
            extra_files = {}
            if status.status == '完了':
                for fname in status.output_files:
                    fpath = work_dir / fname
                    if fpath.exists() and fpath != output_path:
                        extra_files[fname] = fpath

            results['application'] = {
                'status': status.status,
                'message': status.message,
                'output_path': output_path if status.status == '完了' else None,
                'empty_cells': status.empty_cells,
                # Phase 4: 低信頼項目の確認キュー（項目・値・根拠・警告理由）
                'confidence_warnings': getattr(status, 'confidence_warnings', []),
                'extra_files': extra_files,
                # 直近年度として選定された決算書ファイル（UI 明示用）
                'pl_selected_filename': getattr(status, 'pl_selected_filename', ''),
                'pl_selected_end': getattr(status, 'pl_selected_end', ''),
                'pl_selection_warnings': getattr(status, 'pl_selection_warnings', []),
                # all タスクで run_wage_calculation に渡してAPI重複を防ぐためのキャッシュ
                '_cached_financial': status.financial,
                '_cached_ledger_employees': status.ledger_employees,
            }

    if task_type in ('wage', 'all'):
        if progress_callback:
            progress_callback('給与支給総額を計算中...')

        # all タスクの場合、申請書作成タスクのAI抽出結果を再利用してAPI呼出スキップ
        cached_financial = None
        cached_ledger_employees = None
        if task_type == 'all' and 'application' in results:
            cached_financial = results['application'].get('_cached_financial')
            cached_ledger_employees = results['application'].get('_cached_ledger_employees')

        output_path = work_dir / f'{company_name}_給与支給総額計算.xlsx'
        status = run_wage_calculation(
            resource_folder=work_dir,
            company_name=company_name,
            output_path=output_path,
            extractor=extractor,
            cached_financial=cached_financial,
            cached_ledger_employees=cached_ledger_employees,
            fiscal_month_override=fiscal_month_override,
            selection_override=selection_override,
        )
        results['wage'] = {
            'status': status.status,
            'message': status.message,
            'output_path': output_path if status.status == '完了' else None,
            'pl_selected_filename': getattr(status, 'pl_selected_filename', ''),
            'pl_selected_end': getattr(status, 'pl_selected_end', ''),
            'pl_selection_warnings': getattr(status, 'pl_selection_warnings', []),
        }

    if task_type == 'per_employee_wage':
        if progress_callback:
            progress_callback('一人当たり給与支給総額を計算中...')

        output_path = work_dir / f'{company_name}_一人当たり給与支給総額.xlsx'
        status = run_wage_calculation(
            resource_folder=work_dir,
            company_name=company_name,
            output_path=output_path,
            extractor=extractor,
            fiscal_month_override=fiscal_month_override,
            selection_override=selection_override,
            per_employee_only=True,
        )
        results['per_employee_wage'] = {
            'status': status.status,
            'message': status.message,
            'output_path': output_path if status.status == '完了' else None,
        }

    if task_type == 'bonus_wage_ledger_creation':
        if progress_callback:
            progress_callback('加点判定用の賃金台帳を作成中...')

        bonus_template = _find_bonus_wage_ledger_template(template_dir)
        if bonus_template is None:
            results['bonus_wage_ledger_creation'] = {
                'status': 'エラー',
                'message': (
                    '加点判定用賃金台帳テンプレートが見つかりません。'
                    '`ツール/加点判定用賃金台帳テンプレート.xlsx` を配置してください。'
                ),
            }
        else:
            output_path = work_dir / f'{company_name}_加点判定用賃金台帳.xlsx'
            status = run_bonus_wage_ledger_creation(
                resource_folder=work_dir,
                company_name=company_name,
                template_path=bonus_template,
                output_path=output_path,
                extractor=extractor,
                prefecture=prefecture,
                application_ym=application_ym,
                selection_override=selection_override,
            )
            results['bonus_wage_ledger_creation'] = {
                'status': status.status,
                'message': status.message,
                'output_path': output_path if status.status == '完了' else None,
            }

    if task_type == 'bonus':
        if progress_callback:
            progress_callback('加点判定用賃金台帳を読み取り中...')

        results['bonus'] = _run_bonus_judgment(
            work_dir, company_name, prefecture, template_dir,
            template_type=template_type,
            application_ym=application_ym,
            selection_override=selection_override,
        )

    if task_type == 'wage_ledger_creation':
        if progress_callback:
            progress_callback('賃金台帳を Document AI で読み取り中...')

        # 賃金台帳テンプレートを解決（ツール/ または template_dir 直下）
        wage_template_path = _find_wage_ledger_template(template_dir)
        if wage_template_path is None:
            results['wage_ledger_creation'] = {
                'status': 'エラー',
                'message': (
                    '賃金台帳テンプレートが見つかりません。'
                    '`ツール/賃金台帳テンプレート.xlsx` を配置してください。'
                ),
            }
        else:
            # 個人事業主テンプレ選択時の雇用形態正規化を切り替え
            is_kojin = (template_type == 'インボイス枠_個人_2026')
            output_path = work_dir / f'{company_name}_賃金台帳一覧.xlsx'
            status = run_wage_ledger_conversion(
                resource_folder=work_dir,
                company_name=company_name,
                template_path=wage_template_path,
                output_path=output_path,
                extractor=extractor,
                fiscal_month_override=fiscal_month_override,
                is_kojin=is_kojin,
                selection_override=selection_override,
                bonus_paid_months=bonus_paid_months,
            )
            results['wage_ledger_creation'] = {
                'status': status.status,
                'message': status.message,
                'output_path': output_path if status.status == '完了' else None,
            }

    return results


def _find_bonus_ledger(
    work_dir: Path,
    selection_override: dict[str, list[Path]] | None = None,
) -> Path | None:
    """work_dir から加点判定用賃金台帳（専用シート『加点判定用明細』）を1つ探す。"""
    detector = FileDetector(work_dir, selection_override=selection_override)
    for p in detector.get_all('wage_ledger'):
        if is_bonus_wage_ledger(p):
            return p
    return None


def _run_bonus_judgment(
    work_dir: Path,
    company_name: str,
    prefecture: str,
    template_dir: Path,
    template_type: str = '',
    application_ym: tuple[int, int] | None = None,
    selection_override: dict[str, list[Path]] | None = None,
) -> dict:
    """加点判定を実行（専用の「加点判定用賃金台帳」を入力に、AI 再抽出なしで判定）。

    入力は『加点判定用賃金台帳の作成』タスクが生成（または手動記入）した専用テンプレ
    （シート『加点判定用明細』）。①テンプレは申請枠で1枚選ぶ
    （通常枠＝補助率引上げ・加点措置①用、それ以外＝加点措置①用）。②は全枠共通。
    """
    ledger_path = _find_bonus_ledger(work_dir, selection_override)
    if ledger_path is None:
        return {
            'status': 'エラー',
            'message': (
                '加点判定用賃金台帳（専用テンプレ）が見つかりません。'
                'まず「加点判定用賃金台帳の作成」タスクで作成するか、'
                '`加点判定用賃金台帳テンプレート.xlsx` に記入してアップロードしてください。'
            ),
        }

    try:
        ledger = read_bonus_wage_ledger(ledger_path)
    except Exception as e:
        return {'status': 'エラー',
                'message': f'加点判定用賃金台帳の読み取りに失敗しました: {str(e)}'}

    # 台帳に未入力なら UI 値で補完（台帳の入力値があればそちらを優先）
    if not ledger.prefecture and prefecture:
        ledger.prefecture = prefecture
    if ledger.application_ym is None and application_ym:
        ledger.application_ym = application_ym

    if not ledger.employees:
        return {'status': 'エラー',
                'message': '加点判定用賃金台帳から従業員データを読み取れませんでした。'}

    try:
        result = judge_bonus_points(ledger)

        # 申請枠で①テンプレを1枚選ぶ（通常枠＝補助率引上げ①用 / それ以外＝加点措置①用）。
        want_hojoritsu = template_type.startswith('通常枠')
        bonus_dir = template_dir / '補助金加点'
        bonus1_file = None
        bonus2_file = None
        if bonus_dir.exists():
            for bp in bonus_dir.iterdir():
                if bp.suffix.lower() != '.xlsx' or bp.name.startswith('~$'):
                    continue
                name = bp.name
                if '加点措置②' in name:
                    bonus2_file = bp
                elif '加点措置①' in name:
                    is_hojoritsu = '補助率' in name
                    if want_hojoritsu and is_hojoritsu:
                        bonus1_file = bp
                    elif (not want_hojoritsu) and (not is_hojoritsu):
                        bonus1_file = bp

        output_files = {}
        if bonus1_file is not None:
            label1 = '補助率引上げ・加点措置①' if want_hojoritsu else '加点措置①'
            out1 = work_dir / f'{company_name}_{label1}_結果.xlsx'
            fill_bonus_sheet_1(bonus1_file, out1, result)
            output_files['bonus1'] = out1
        if bonus2_file is not None:
            out2 = work_dir / f'{company_name}_加点措置②_結果.xlsx'
            fill_bonus_sheet_2(bonus2_file, out2, result)
            output_files['bonus2'] = out2

        return {
            'status': '完了',
            'message': f'従業員{len(ledger.employees)}名の加点判定用賃金台帳を分析しました。',
            'result': result,
            'output_files': output_files,
            'employee_count': len(ledger.employees),
        }

    except Exception as e:
        return {'status': 'エラー',
                'message': f'処理中にエラーが発生しました: {str(e)}'}


# ── ヘッダー ──
st.markdown('<div class="main-header">📋 補助金書類自動作成ツール</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">IT導入補助金の申請書類をAIで自動作成します</div>', unsafe_allow_html=True)

# API接続状態
if CLAUDE_API_KEY:
    st.success('✅ Claude API: 接続済み')
else:
    st.error('❌ Claude API: 未接続（.envファイルにCLAUDE_API_KEYを設定してください）')
    st.stop()

# ── 使い方ガイド ──
st.markdown("""
<div class="how-it-works">
<strong>使い方（3ステップ）</strong><br>
① 左のサイドバーで <strong>会社名</strong> と <strong>テンプレート種別</strong> を選択<br>
② 下のアップロード欄に <strong>資料ファイルをまとめてドラッグ&ドロップ</strong><br>
③ <strong>「処理開始」ボタン</strong> を押して完成ファイルをダウンロード
</div>
""", unsafe_allow_html=True)

# ── サイドバー ──
with st.sidebar:
    st.header('⚙️ 設定')

    company_name_input = st.text_input(
        '会社名（必須）',
        placeholder='例: ○○株式会社',
        help='正式名称でなくてもOK。出力ファイル名に使われます。',
    )
    # ファイル名に使えない文字（/ \ : * ? " < > | 等）が含まれているとパスエラーになる。
    # 入力直後にサニタイズして、以降の output_path 生成すべてに安全に使う。
    company_name = _safe_company_name(company_name_input)
    if company_name_input and company_name != company_name_input:
        st.caption(
            f'⚠ ファイル名に使えない文字を「-」に置換しました: `{company_name}`'
        )

    template_label = st.selectbox(
        'テンプレート種別',
        list(TEMPLATE_OPTIONS.keys()),
        help='申請する補助金の枠を選択してください。',
    )
    template_type = TEMPLATE_OPTIONS[template_label]

    task_label = st.selectbox(
        '実行タスク',
        list(TASK_OPTIONS.keys()),
        help='【推奨フロー】① 賃金台帳の作成 → ② 一人当たり給与支給総額 → ③ 申請書作成。'
             '賃金台帳の作成：賃金台帳PDFをツール規格のExcelに自動変換'
             '（Document AI使用、手書きPDF以外はそのままお任せ可）。'
             '一人当たり給与支給総額：賃金台帳のみで計算（決算書PDFは参照しない／決算書由来の項目シートも削除）。'
             '出力Excelを人間がチェックして数値を確定する工程。'
             '申請書作成：ヒアリングシート+各種PDFから申請書を自動作成。'
             '「給与支給総額計算」「従業員別明細」「生産性指標」シートも申請書Excelに内包する。'
             '賃金台帳は AI で再抽出せず決定論パーサーで直読するため、'
             '上記②と同じ計算ロジックで R215/R216 を算出する（数値は一致）。'
             '中途入退社が多い案件は注意喚起を出すが自動転記は実行する。'
             '加点判定：賃金台帳から加点措置の対象かを判定。',
    )
    task_type = TASK_OPTIONS[task_label]

    # 決算月（必須）— ユーザー指定で賃金台帳の対象期間を確定 + AI推定誤りを照合
    # 2026-05 方針: AI 推定に任せず、ユーザーが明示的に指定する運用に変更
    _FISCAL_MONTH_OPTIONS = ['（選択してください）'] + [f'{i}月' for i in range(1, 13)]
    fiscal_month_label = st.selectbox(
        '決算月（必須）',
        _FISCAL_MONTH_OPTIONS,
        help='決算期末の月を必ず指定してください。'
             '賃金台帳の対象12ヶ月が確定し、決算書のAI誤読も照合できます。',
    )
    fiscal_month_override: int | None = None
    if fiscal_month_label != _FISCAL_MONTH_OPTIONS[0]:
        fiscal_month_override = int(fiscal_month_label.replace('月', ''))

    # 賞与の支給月（任意）— 「賃金台帳の作成」タスクで非暦年決算のときだけ表示。
    # 年間集計表など「支給日の無い賞与」を対象事業年度の12ヶ月で正しく絞り込むためのヒント。
    bonus_paid_months: list[tuple[int, int]] | None = None
    if task_type == 'wage_ledger_creation' and fiscal_month_override not in (None, 12):
        _bonus_str = st.text_input(
            '賞与の支給月（任意・yyyy/mm カンマ区切り）',
            value='',
            help='台帳に賞与の支給年月が無い（年間集計表など）非暦年決算のとき、'
                 '夏・冬など各回の支給年月を入力すると対象事業年度で自動補正します。'
                 '例: 2024/12, 2025/07。空欄なら従来どおり（支給日不明なら要確認の警告を表示）。',
        )
        bonus_paid_months = _parse_bonus_months_input(_bonus_str)
        if _bonus_str and not bonus_paid_months:
            st.warning('賞与の支給月は yyyy/mm をカンマ区切りで入力してください（例: 2024/12, 2025/07）')

    # 製造原価ありフラグ — 製造業向け。チェック時、AI に「製造原価報告書が存在する」ヒントを注入
    # 決算書PDFを参照しないタスクでは非表示
    if task_type in ('per_employee_wage', 'bonus', 'wage_ledger_creation',
                     'bonus_wage_ledger_creation'):
        has_cost_report_hint = False
    else:
        has_cost_report_hint = st.checkbox(
            '製造原価報告書あり（製造業向け）',
            value=False,
            help='製造業のお客様で「製造原価報告書」を提出いただいている場合はチェック。'
                 'AIの読み落としを防ぎ、損益計算書＋製造原価を統合して人件費を算出します。'
                 '（資料に製造原価報告書PDFがあれば自動検出されるため、'
                 '通常は自動検出に任せて構いません）',
        )

    # Drive 格納オプション（データソースが Drive のときのみ有効化される）
    upload_to_drive = st.checkbox(
        '結果を選択した Drive フォルダに格納',
        value=False,
        help='処理完了後、生成された Excel をローカルに残さず Drive フォルダへ自動アップロードします。'
             'データソースが「Google Drive」かつ顧客フォルダ選択時のみ有効。'
             '同名ファイルがあれば上書きされます。',
    )

    # 加点判定・加点判定用台帳の作成では都道府県と交付申請月が必要
    application_ym: tuple[int, int] | None = None
    if task_type in ('bonus', 'bonus_wage_ledger_creation'):
        from hojokin.config import MIN_WAGE_MAP
        prefecture = st.selectbox(
            '事業場の都道府県',
            [''] + list(MIN_WAGE_MAP.keys()),
            help='加点判定の最低賃金比較に使用します（事業場の所在地・会社で1つ）。',
        )
        _app_month_str = st.text_input(
            '交付申請月（yyyy/mm）',
            value='',
            help='加点措置②の比較対象＝この前月（直近月）。例: 2026/06。'
                 '加点判定用賃金台帳の C3 にも書き込まれます。',
        )
        application_ym = _parse_app_month_input(_app_month_str)
        if _app_month_str and application_ym is None:
            st.warning('交付申請月は yyyy/mm 形式で入力してください（例: 2026/06）')
    else:
        prefecture = ''

    st.divider()

    # データソース選択
    _has_local_creds = bool(_DRIVE_CREDS and Path(_DRIVE_CREDS).exists())
    # secrets.toml が無いローカルでの StreamlitSecretNotFoundError をヘルパーで吸収
    _has_cloud_creds = _has_streamlit_secret('gcp_service_account')
    drive_available = (_has_local_creds or _has_cloud_creds) and bool(_DRIVE_PARENT_ID)
    # Secrets経由の場合もPARENT_IDを取得
    if not _DRIVE_PARENT_ID and _has_cloud_creds:
        try:
            drive_available = bool(st.secrets.get('drive_parent_folder_id', ''))
        except Exception as e:
            logger.warning(
                f'Streamlit Secrets の drive_parent_folder_id 取得に失敗: {e}',
                exc_info=True,
            )
    if drive_available:
        data_source = st.radio(
            'データソース',
            ['ファイルアップロード', 'Google Drive'],
            help='Google Driveから直接ファイルを取得できます。',
        )
    else:
        data_source = 'ファイルアップロード'

    st.divider()

    st.markdown('**処理の目安**')
    st.caption('所要時間: 約1〜3分（PDF量により変動）')
    st.caption('API利用料: 約20〜90円/社（PDF量により変動）')
    st.caption('└ 賃金台帳は Excel / CSV のみ対応（PDF は受け付けません）')
    st.caption('└ Opus 4.8（文章生成）＋ Sonnet 4.6（PDF抽出）の従量課金。確定額は Anthropic の請求でご確認ください')
    st.caption('実行前に「案件規模の予想」で詳細目安が表示されます')

# ── ファイル判別ヘルパー（ローカル/Drive 共通） ──
_FILE_CATEGORIES = [
    ('hearing',     'ヒアリングシート',        ['ヒアリング']),
    ('registry',    '履歴事項全部証明書',      ['履歴事項']),
    ('pl',          '損益計算書 / 決算報告書', ['損益計算書', '決算報告書', '決算書']),
    ('cost_report', '製造原価報告書',          ['製造原価報告書', '原価報告書']),
    ('tax',         '納税証明書',              ['納税証明']),
    ('estimate',    '見積書',                  ['見積']),
    ('wage_report', '賃金状況報告シート',      ['賃金状況報告']),
    # pipeline.FileDetector.PATTERNS と整合: 給与ソフト出力は「給与台帳」表記が多いため両対応
    ('wage_ledger', '賃金台帳 / 給与台帳',     ['賃金台帳', '給与台帳']),
]

_REQUIRED_CATS_BY_TASK = {
    'application':           {'hearing', 'registry', 'pl'},
    'wage':                  {'wage_ledger'},
    'per_employee_wage':     {'wage_ledger'},
    'bonus':                 {'wage_ledger'},
    'bonus_wage_ledger_creation': {'wage_ledger'},
    'wage_ledger_creation':  {'wage_ledger'},
    'all':                   {'hearing', 'registry', 'pl'},
}

# カテゴリ別の許可拡張子（pipeline.FileDetector.ALLOWED_EXTS と整合）。
# UI 側でも事前にこのフィルタを適用しないと、PDF だけアップした賃金台帳が
# 「必須あり」判定で通って実行後に skipped → 給与/加点が無データで失敗する。
_UI_ALLOWED_EXTS = {
    'hearing':     {'.xlsx', '.xlsm'},
    'registry':    {'.pdf'},
    'pl':          {'.pdf'},
    'cost_report': {'.pdf'},
    'tax':         {'.pdf'},
    'estimate':    {'.xlsx', '.xlsm', '.pdf'},
    'wage_report': {'.xlsx', '.xlsm'},
    'wage_ledger': {'.xlsx', '.xlsm', '.csv'},  # 2026-05 方針: PDF/.xls 除外
}

# 「賃金台帳の作成」タスク専用: 賃金台帳PDF も受け付ける
_UI_ALLOWED_EXTS_WAGE_LEDGER_CREATION = {
    **_UI_ALLOWED_EXTS,
    'wage_ledger': {'.xlsx', '.xlsm', '.csv', '.pdf'},
}


def _get_ui_allowed_exts(task: str | None) -> dict:
    """タスクに応じた UI 許可拡張子テーブルを返す。

    「賃金台帳の作成」「加点判定用賃金台帳の作成」タスクは賃金台帳カテゴリで PDF を許可する
    （生の賃金台帳/給与明細 PDF から AI 抽出するため）。
    他タスクは PDF 賃金台帳を弾く（ローカル変換運用のまま）。
    """
    if task in ('wage_ledger_creation', 'bonus_wage_ledger_creation'):
        return _UI_ALLOWED_EXTS_WAGE_LEDGER_CREATION
    return _UI_ALLOWED_EXTS


def _analyze_files(file_names, task):
    """ファイル名リストからタスク別の判別結果を計算"""
    required_cats = _REQUIRED_CATS_BY_TASK.get(task, set())
    allowed_table = _get_ui_allowed_exts(task)

    detected = {cat: [] for cat, _, _ in _FILE_CATEGORIES}
    unmatched = []
    for name in file_names:
        # NFD（macOS の濁点分離形式）でも比較が通るよう NFC 化してから判定
        name_nfc = unicodedata.normalize('NFC', name)
        ext = Path(name_nfc).suffix.lower()
        matched = False
        for cat, _, keywords in _FILE_CATEGORIES:
            if any(kw in name_nfc for kw in keywords):
                # 拡張子が許可外なら検出に加えず unmatched 行きにする
                # （後段で「必須あり」判定が誤って通るのを防ぐ）
                allowed = allowed_table.get(cat)
                if allowed is not None and ext not in allowed:
                    unmatched.append(name)
                    matched = True
                    break
                detected[cat].append(name)
                matched = True
                break
        if not matched:
            unmatched.append(name)

    checks = [
        (cat, display, keywords, cat in required_cats)
        for cat, display, keywords in _FILE_CATEGORIES
    ]
    missing_required = [
        display for cat, display, _, required in checks
        if required and not detected[cat]
    ]
    return {
        'checks': checks,
        'detected': detected,
        'unmatched': unmatched,
        'missing_required': missing_required,
        'all_required_ok': len(missing_required) == 0,
    }


def _check_size_warnings(
    file_size_pairs: list[tuple[str, int]],
    task: str | None = None,
) -> list[str]:
    """ファイル(名前, バイト数)から大容量・大量警告を作る。

    閾値（実害が出る前のソフト警告レベル）:
      - PDF 30MB超: API 残高消費が大きい・タイムアウトリスク
      - 賃金台帳ファイル合計 25MB超: AI 抽出に長時間かかる可能性（警告）
      - 賃金台帳ファイル 8件超: 個人別ファイル多数 → AI 抽出 max_tokens 不足の可能性
      - 単一 Excel/CSV 5MB超: 想定外に大きく、誤ったファイルの可能性
      - 賃金台帳が PDF: ツール側で処理されないため Excel/CSV 変換誘導
    """
    warnings = []
    PDF_LIMIT = 30 * 1024 * 1024
    EXCEL_CSV_LIMIT = 5 * 1024 * 1024
    WAGE_TOTAL_LIMIT = 25 * 1024 * 1024
    WAGE_COUNT_LIMIT = 8
    # 2026-05 方針: 賃金台帳は Excel/CSV のみ処理（PDF は ALLOWED_EXTS で弾かれる）
    # FileDetector.ALLOWED_EXTS['wage_ledger'] と整合させる
    WAGE_EXTS = {'.xlsx', '.xlsm', '.csv'}

    # pipeline.FileDetector.PATTERNS['wage_ledger'] と整合（給与ソフト出力は「給与台帳」表記が多い）
    wage_keywords = ('賃金台帳', '給与台帳')
    wage_files = []
    wage_pdf_names = []
    for name, size in file_size_pairs:
        if not name or size is None or size <= 0:
            continue  # サイズ不明や空ファイルはスキップ（誤警告防止）
        n = unicodedata.normalize('NFC', name)
        ext = Path(n).suffix.lower()
        if ext == '.pdf' and size > PDF_LIMIT:
            warnings.append(
                f'📦 大容量PDF: {n} ({size/1024/1024:.1f}MB) — '
                f'AI 抽出に時間がかかる可能性があります（タイムアウト 480秒）'
            )
        if ext in ('.xlsx', '.xlsm', '.csv') and size > EXCEL_CSV_LIMIT:
            warnings.append(
                f'📦 大容量{ext}: {n} ({size/1024/1024:.1f}MB) — '
                f'想定外に大きいファイルです。誤ったファイルでないか確認してください'
            )
        # 賃金台帳カウントは処理対象拡張子のみ（.docx等の誤検出を防ぐ）
        if any(kw in n for kw in wage_keywords):
            if ext in WAGE_EXTS:
                wage_files.append((n, size))
            elif ext == '.pdf':
                wage_pdf_names.append(n)

    # 賃金台帳が PDF で投入されている場合（Excel/CSV 変換誘導）
    # ただし「賃金台帳の作成」タスクは PDF を主入力として処理するため警告を出さない
    if wage_pdf_names and task != 'wage_ledger_creation':
        warnings.append(
            f'📄 賃金台帳が PDF 形式で含まれています（{len(wage_pdf_names)}件）— '
            f'ツールでは処理されません。Excel / CSV に変換してから再投入してください'
        )

    if wage_files:
        total_size = sum(s for _, s in wage_files)
        # 申請書作成（application）は賃金台帳を AI で再抽出しない（決定論パーサー一本）
        # ため、サイズ/件数の AI 関連警告は不要。
        # 「両方（all）」は application 側で決定論パーサーが0件返した場合、
        # wage 側で AI フォールバックが発動する可能性があるため警告は出す（保守的）。
        wage_uses_ai = task != 'application'
        if wage_uses_ai and total_size > WAGE_TOTAL_LIMIT:
            warnings.append(
                f'📦 賃金台帳合計サイズが大きいです（{total_size/1024/1024:.1f}MB, '
                f'{len(wage_files)}ファイル） — AI 抽出が長時間化する可能性'
            )
        if wage_uses_ai and len(wage_files) > WAGE_COUNT_LIMIT:
            warnings.append(
                f'📂 賃金台帳ファイルが多数あります（{len(wage_files)}件） — '
                f'従業員数が多い場合は AI 出力が途中で切れる可能性があります（max_tokens 16384）'
            )

    return warnings


def _estimate_case_scale(
    file_size_pairs: list[tuple[str, int]],
    task: str | None = None,
) -> dict | None:
    """案件規模・処理時間・APIコストを推定する。

    現運用（2026-05〜）では賃金台帳は Excel/CSV のみ処理されるため、
    AI 抽出のコスト・時間を支配するのは主に **PDF 群**（履歴事項 / 損益
    計算書 / 納税証明 / 見積書 / 製造原価報告書）。それに賃金台帳
    Excel/CSV の AI 抽出（USE_AI_WAGE_EXTRACTION=true 時）が乗る。

    ユーザー（坂平さん）が実行前に「これは何分かかりそう」「いくらぐらい」を
    把握できるよう、控えめ（多め見積もり）の数値を返す。

    Returns: {
        'scale': '小型' | '中型' | '大型',
        'time_label': '約1〜2分',
        'cost_label': '約20〜60円',
        'pdf_total_mb': float,
        'wage_excel_count': int,
        'wage_pdf_count': int,  # 不正投入検出用（>0 なら警告表示）
        'note': str,  # ユーザー向け補足
    }
    None: 推定不能（ファイルなし）
    """
    if not file_size_pairs:
        return None

    WAGE_KEYWORDS = ('賃金台帳', '給与台帳')
    WAGE_EXTS = {'.xlsx', '.xlsm', '.csv'}

    pdf_total = 0
    wage_excel_count = 0
    wage_pdf_count = 0
    for name, size in file_size_pairs:
        if not name or size is None or size <= 0:
            continue
        n = unicodedata.normalize('NFC', name)
        ext = Path(n).suffix.lower()
        is_wage = any(kw in n for kw in WAGE_KEYWORDS)
        if ext == '.pdf':
            if is_wage:
                # 賃金台帳PDFは ALLOWED_EXTS で弾かれて処理対象外
                # （カウントだけ取って警告表示に使う）
                wage_pdf_count += 1
            else:
                pdf_total += size
        elif ext in WAGE_EXTS and is_wage:
            wage_excel_count += 1

    pdf_total_mb = pdf_total / 1024 / 1024

    # サイズクラス判定（処理対象 PDF の合計サイズ）
    # 経験則: 通常案件で 3〜8MB、原価報告書ありで 5〜12MB
    if pdf_total_mb >= 12:
        scale = '大型'
    elif pdf_total_mb >= 5:
        scale = '中型'
    else:
        scale = '小型'

    # ── 処理時間予想 ──
    # 内訳の経験則:
    #   PDF 抽出（履歴/PL/税/見積/原価）: 各 10〜30 秒
    #   AI 判断（高加点項目の総合判定）: 20〜40 秒
    #   賃金台帳 Excel/CSV 抽出: 5〜20 秒
    #   テンプレ書き込み・後処理: 5〜10 秒
    common_min = 30   # 共通処理（楽観）
    common_max = 90   # 共通処理（悲観）
    pdf_min_sec = pdf_total_mb * 4
    pdf_max_sec = pdf_total_mb * 10
    wage_excel_sec_min = wage_excel_count * 5
    wage_excel_sec_max = wage_excel_count * 15
    total_min_sec = common_min + pdf_min_sec + wage_excel_sec_min
    total_max_sec = common_max + pdf_max_sec + wage_excel_sec_max

    if total_max_sec < 90:
        time_label = '約 30秒〜1分'
    elif total_max_sec < 180:
        time_label = f'約 1〜{max(2, round(total_max_sec/60))}分'
    else:
        time_label = f'約 {max(1, round(total_min_sec/60))}〜{round(total_max_sec/60)}分'

    # ── APIコスト予想（円、Sonnet 4.6 / 1USD=150円） ──
    # 経験則:
    #   PDF 抽出: 1MB あたり 2〜5円
    #   共通固定費（AI 判断・後処理）: 約 10〜15円
    #   賃金台帳 Excel/CSV AI 抽出: 1ファイルあたり 3〜8円
    # 申請書作成タスク（application）は賃金台帳を AI で再抽出しない
    # （決定論パーサー一本）ため、wage AI コストは 0。
    common_cost = 10
    pdf_cost_min = pdf_total_mb * 2
    pdf_cost_max = pdf_total_mb * 5
    if task == 'application':
        # application: 賃金台帳は決定論パーサー一本（AI 不使用）
        wage_cost_min = 0
        wage_cost_max = 0
    else:
        # all: application が決定論パーサーで成功すれば wage 側もキャッシュ経由で AI 不使用、
        #      ただし application で 0 件返ると wage 側で AI フォールバックが走る。
        #      コスト目安は保守的に AI 使用前提で出しておく（過大評価でも実害は小）。
        wage_cost_min = wage_excel_count * 3
        wage_cost_max = wage_excel_count * 8
    total_cost_min = common_cost + pdf_cost_min + wage_cost_min
    total_cost_max = common_cost + pdf_cost_max + wage_cost_max

    cost_label = f'約 {max(15, int(total_cost_min))}〜{max(25, int(total_cost_max))}円'

    # 補足メッセージ
    notes: list[str] = []
    # 「賃金台帳の作成」タスクは PDF を主入力として処理するため除外メッセージを出さない
    if wage_pdf_count > 0 and task != 'wage_ledger_creation':
        notes.append(
            f'賃金台帳PDFが {wage_pdf_count} 件含まれていますが、'
            'ツールでは処理されません。Excel/CSV に変換してから再投入してください'
        )
    if scale == '大型':
        notes.append('PDF 量が多めのため、処理時間が長くなる可能性があります。画面を閉じずにお待ちください')

    return {
        'scale': scale,
        'time_label': time_label,
        'cost_label': cost_label,
        'pdf_total_mb': pdf_total_mb,
        'wage_excel_count': wage_excel_count,
        'wage_pdf_count': wage_pdf_count,
        'note': ' / '.join(notes) if notes else '',
    }


def _render_case_scale_estimate(
    file_size_pairs: list[tuple[str, int]],
    task: str | None = None,
):
    """案件規模・処理時間・APIコスト予想を UI に表示。"""
    est = _estimate_case_scale(file_size_pairs, task=task)
    if not est:
        return

    scale_emoji = {'小型': '🟢', '中型': '🟡', '大型': '🟠'}
    emoji = scale_emoji.get(est['scale'], '📊')

    with st.expander(
        f'{emoji} 案件規模の予想: **{est["scale"]}** — '
        f'処理時間 {est["time_label"]} / APIコスト {est["cost_label"]}',
        expanded=(est['scale'] == '大型'),
    ):
        st.markdown(
            f'- **規模**: {emoji} {est["scale"]}（処理対象PDF合計 {est["pdf_total_mb"]:.1f}MB '
            f'/ 賃金台帳Excel/CSV {est["wage_excel_count"]}件）\n'
            f'- **処理時間目安**: {est["time_label"]}\n'
            f'- **APIコスト目安**: {est["cost_label"]}'
        )
        if est['note']:
            for n in est['note'].split(' / '):
                st.info(n)
        st.caption(
            'コスト・時間は上方寄りに見積もった目安です。'
            '実際はこれより安く済むこともありますが、'
            'PDF量が多い案件では上振れすることもあります。'
            '確定額は Anthropic の請求でご確認ください。'
        )


def _render_file_check_result(result, total_count):
    """判別結果の表示（ローカル/Drive 共通）"""
    if result['all_required_ok']:
        st.success(
            f'ファイルチェック OK — 必須ファイルがすべて揃っています'
            f'（{total_count}件）'
        )
    else:
        st.error(
            f'必須ファイルが不足しています: '
            f'**{"、".join(result["missing_required"])}**'
        )

    with st.expander('ファイル判別結果（詳細）', expanded=not result['all_required_ok']):
        for cat, display, _, required in result['checks']:
            files = result['detected'][cat]
            if files:
                st.markdown(f'✅ **{display}** → `{"`, `".join(files)}`')
            elif required:
                st.markdown(
                    f'❌ **{display}** — **未検出（必須）** '
                    'ファイル名にキーワードが含まれているか確認してください'
                )
                if cat == 'wage_ledger':
                    st.markdown(
                        '&ensp;💡 PDFしか無い場合は Excel に変換してアップロードしてください。'
                        '変換は **手元の PC の Claude Code（CC）+ `wagebook-convert` Skill** で行います。'
                        'ページ上部の「📘 賃金台帳の作成手順（CC向け Skill）」expander にセットアップ手順と依頼方法を載せています。'
                    )
            else:
                st.markdown(f'➖ {display} — なし（任意）')

        if result['unmatched']:
            st.markdown('---')
            st.markdown('**判別できなかったファイル:**')
            for name in result['unmatched']:
                st.markdown(f'&ensp; ⚠️ `{name}`（キーワードなし → 処理対象外）')


# 差し替え UI に出すカテゴリ（実 Drive 27案件の調査結果に基づく）。
# - pl: 複数期混在 48% / ファイル名タイポも頻発 → 差し替え必要性が圧倒的に高い
# - wage_ledger: 個人別×複数 / 期別混在 / 空テンプレ混入 → 同上
# 他カテゴリ（registry/tax 等）も複数候補が出ることはあるが、内容は「同名重複アップロード」
# が大半で自動選定（先頭採用）で実害なし。UI に出してもノイズになるため非表示。
_OVERRIDE_UI_CATS = {'pl', 'wage_ledger'}
# 「賃金台帳の作成」タスクでは決算書は使わず、賃金台帳PDFと履歴事項PDFを差し替え対象にする
_OVERRIDE_UI_CATS_WAGE_LEDGER_CREATION = {'wage_ledger', 'registry'}


def _get_override_ui_cats(task: str | None) -> set[str]:
    """タスクに応じた差し替え UI 対象カテゴリを返す。"""
    if task == 'wage_ledger_creation':
        return _OVERRIDE_UI_CATS_WAGE_LEDGER_CREATION
    return _OVERRIDE_UI_CATS

# 複数選択（multiselect）させるカテゴリ。それ以外は単一選択（selectbox）。
_MULTI_SELECT_CATS = {'wage_ledger'}


def _categorize_for_ui(
    file_names: list[str],
    task: str | None = None,
) -> dict[str, dict[str, list[str]]]:
    """ファイル名をカテゴリ別に「推奨候補」「その他候補」「全候補」に分類する。

    - recommended: ファイル名キーワード一致 + 拡張子許可 → 自動検出と同じ判定
    - others:      キーワード未一致 だが 拡張子は許可される候補（タイポ救済枠）
    - all:         recommended + others（差し替え UI の候補プール）

    例: PL カテゴリで「PL_R7.pdf」はキーワード（決算書/損益計算書）を含まないが
    `_UI_ALLOWED_EXTS['pl'] = {'.pdf'}` を満たすので others に入る。

    task='wage_ledger_creation' のときは賃金台帳PDFも許可される。
    """
    allowed_table = _get_ui_allowed_exts(task)
    result: dict[str, dict[str, list[str]]] = {
        cat: {'recommended': [], 'others': [], 'all': []}
        for cat, _, _ in _FILE_CATEGORIES
    }

    # 「推奨候補」: 既存 _analyze_files と同じロジックで一意決定（最初に一致したカテゴリのみ）
    for name in file_names:
        name_nfc = unicodedata.normalize('NFC', name)
        ext = Path(name_nfc).suffix.lower()
        for cat, _, keywords in _FILE_CATEGORIES:
            if any(kw in name_nfc for kw in keywords):
                allowed = allowed_table.get(cat)
                if allowed is not None and ext not in allowed:
                    break  # 拡張子NG → 他カテゴリも試さない（_analyze_files と整合）
                result[cat]['recommended'].append(name)
                break

    # 「その他候補」: キーワード一致しないが拡張子は許可（タイポ救済）。
    # 1ファイルが複数カテゴリの「その他」に同時所属し得る（PDF は registry/pl/tax/estimate/cost_report の全候補）
    # ただし recommended に既に入っているカテゴリにはその名前を入れない（重複排除）
    for name in file_names:
        name_nfc = unicodedata.normalize('NFC', name)
        ext = Path(name_nfc).suffix.lower()
        for cat, _, _kw in _FILE_CATEGORIES:
            allowed = allowed_table.get(cat)
            if allowed is not None and ext not in allowed:
                continue
            if name in result[cat]['recommended']:
                continue
            result[cat]['others'].append(name)

    # all = recommended + others
    for cat in result:
        result[cat]['all'] = result[cat]['recommended'] + result[cat]['others']

    return result


# 「自動検出」ラベルは動的に組み立てる（_auto_selected_for_display の結果を埋め込む）
_OVERRIDE_EXCLUDE_LABEL = '─ 使わない（対象外）─'
_OVERRIDE_SEP_RECOMMENDED = '── 推奨候補（自動検出ヒット）──'
_OVERRIDE_SEP_OTHERS = '── その他のファイル（タイポ救済用）──'


def _auto_selected_for_display(category: str, recommended: list[str]) -> str | None:
    """カテゴリの自動選定結果（UI 表示用）を 1 件返す。

    pipeline の選定ロジックを完全再現はしない（DRY 違反になるため簡易版）。
    PL は「ファイル名から期末年月を抽出して最新」、それ以外は recommended[0]。
    候補が無ければ None。
    """
    if not recommended:
        return None
    if category == 'pl':
        try:
            from hojokin.pipeline import (
                _parse_fiscal_end_from_filename,
                _parse_fiscal_year_from_filename,
            )
        except Exception:
            return recommended[0]
        # 月ありで取れるものは (年,月) で最新を選ぶ
        with_date = [
            (name, _parse_fiscal_end_from_filename(name))
            for name in recommended
        ]
        valid = [(n, ym) for n, ym in with_date if ym is not None]
        if valid:
            return max(valid, key=lambda t: t[1])[0]
        # 「令和N年度」のような月なし名は年だけで順位付け（前年度を先頭に出さない）。
        # pipeline の get_pl_latest（決算月指定時）の実選択と表示を一致させる。
        with_year = [
            (n, _parse_fiscal_year_from_filename(n)) for n in recommended
        ]
        valid_year = [(n, y) for n, y in with_year if y is not None]
        if valid_year:
            return max(valid_year, key=lambda t: t[1])[0]
    return recommended[0]


def _render_file_selection_override(
    file_names: list[str], case_key: str, task: str | None = None,
) -> dict[str, list[str] | None]:
    """検出カテゴリ別の候補ファイルを差し替え可能な UI で表示し、選択結果を返す。

    候補は「推奨候補（キーワード一致）」と「その他のファイル（許可拡張子のみ）」
    の両セクションに分けて並べる。タイポ等でキーワードが含まれないファイルも
    その他セクションから割り当てて処理対象にできる。

    Returns:
        dict[category, list[str] | None]
            - None: 自動検出に従う（override なし）
            - list[str]: ユーザー指定（[] は「対象外」）
    """
    cat_info = _categorize_for_ui(file_names, task=task)
    # 差し替え UI 表示対象はタスクごとに変える:
    #   - 通常: pl と wage_ledger
    #   - 賃金台帳の作成: wage_ledger と registry（履歴事項PDF も差し替え可能に）
    override_cats = _get_override_ui_cats(task)
    visible = [
        (cat, display)
        for cat, display, _ in _FILE_CATEGORIES
        if cat in override_cats and cat_info[cat]['all']
    ]
    if not visible:
        return {}

    overrides: dict[str, list[str] | None] = {}
    # selectbox の「セパレータ行」をユーザーが選べないよう、選ばれたら自動扱いに戻す
    SEP_LABELS = {_OVERRIDE_SEP_RECOMMENDED, _OVERRIDE_SEP_OTHERS}

    # タスクに応じてタイトルを変える（決算書を使わないタスクで「決算書」と表示しない）
    if task == 'wage_ledger_creation':
        _expander_title = '▶ 賃金台帳・履歴事項を差し替える（必要な場合のみ）'
    else:
        _expander_title = '▶ 決算書・賃金台帳を差し替える（必要な場合のみ）'
    with st.expander(_expander_title, expanded=False):
        st.caption(
            '通常は **自動検出のまま** で OK。'
            '誤選択リスクが高いのは決算書（複数期分混在）と賃金台帳（個人別×複数 / 空テンプレ混入）'
            'なので、差し替えはこの 2 カテゴリのみに絞っています。'
            '\n\n'
            'ファイル名にキーワード（決算書 / 賃金台帳・給与台帳）が含まれないファイルも'
            '「その他のファイル」セクションから割り当てできます（タイポ救済用）。'
        )
        for cat, display in visible:
            recommended = cat_info[cat]['recommended']
            others = cat_info[cat]['others']
            sel_key = f'override_{case_key}_{cat}'

            if cat in _MULTI_SELECT_CATS:
                # 複数選択（賃金台帳など）。
                # 推奨候補は default 全選択、その他は default 未選択（タイポ救済時のみ手動追加）。
                options = recommended + others
                default = list(recommended)
                label_suffix = (
                    f'（推奨{len(recommended)} / その他{len(others)}）'
                    if others else f'（{len(recommended)}件）'
                )
                selected = st.multiselect(
                    f'{display}{label_suffix}',
                    options=options,
                    default=default,
                    key=sel_key,
                    help=(
                        '推奨候補（自動検出ヒット）は既定でチェック済み。'
                        'その他のファイルは既定で未チェック（必要なら手動で追加）。'
                        'チェックを外したファイルは処理から除外されます。'
                    ),
                )
                if list(selected) == default:
                    overrides[cat] = None  # 既定 → 自動扱い
                else:
                    overrides[cat] = list(selected)
            else:
                # 単一選択。推奨/その他をセパレータで分割表示。
                # 「自動検出」ラベルには実際に選定されるファイル名を埋め込んで
                # 「今何が選ばれているか」を可視化する。
                # セパレータ行は disabled 不可なので、選ばれたら自動扱いに戻す（後段ガード）。
                auto_name = _auto_selected_for_display(cat, recommended)
                auto_label = (
                    f'（自動検出: {auto_name}）'
                    if auto_name else '（自動検出に従う・候補なし）'
                )
                opts: list[str] = [auto_label]
                if recommended:
                    opts.append(_OVERRIDE_SEP_RECOMMENDED)
                    opts.extend(recommended)
                if others:
                    opts.append(_OVERRIDE_SEP_OTHERS)
                    opts.extend(others)
                opts.append(_OVERRIDE_EXCLUDE_LABEL)

                selected = st.selectbox(
                    display,
                    options=opts,
                    key=sel_key,
                    help=(
                        '推奨候補（自動検出ヒット）またはその他のファイル（タイポ救済）から選択。'
                        '「使わない」で対象外指定できます。'
                    ),
                )
                if selected in SEP_LABELS or selected == auto_label:
                    overrides[cat] = None
                elif selected == _OVERRIDE_EXCLUDE_LABEL:
                    overrides[cat] = []
                else:
                    overrides[cat] = [selected]
    return overrides


def _build_path_override(
    name_override: dict[str, list[str] | None] | None,
    work_dir: Path,
) -> dict[str, list[Path]] | None:
    """ファイル名ベースの override を work_dir 内の Path に解決する。

    存在しないファイル名は無視してログに警告を残す（NFC 正規化済みファイル名で照合）。
    実質的な上書きが何も無い場合は None を返す（pipeline 側で自動検出が動く）。
    """
    if not name_override:
        return None
    path_override: dict[str, list[Path]] = {}
    for cat, names in name_override.items():
        if names is None:
            continue  # 自動検出を維持
        paths: list[Path] = []
        for name in names:
            name_nfc = unicodedata.normalize('NFC', name)
            candidate = work_dir / name_nfc
            if candidate.exists():
                paths.append(candidate)
            else:
                logger.warning(
                    f'手動選択ファイルが work_dir に見つかりません: '
                    f'{name_nfc} (カテゴリ: {cat})'
                )
        path_override[cat] = paths
    return path_override if path_override else None


def _check_required_by_names(file_names, task, name_override=None):
    """タスクに応じた必須ファイルが揃っているかチェック。

    判定優先順:
    1. ユーザーが UI で override に明示指定したファイルがあれば、それを充足とみなす
       （タイポでキーワード未一致のファイルでも、ユーザーが手動でカテゴリ割当すれば OK）
    2. それ以外は従来通り、ファイル名キーワード + 拡張子フィルタで自動検出を確認

    拡張子フィルタ（_UI_ALLOWED_EXTS）も適用するため、賃金台帳PDFだけ
    アップした状態では can_run を有効にしない（実行後の skipped 失敗を防ぐ）。
    """
    if not file_names:
        return False
    # NFD（macOS の濁点分離形式）でも比較が通るよう NFC 化してから判定
    names_nfc = [unicodedata.normalize('NFC', n) for n in file_names]
    required_cats = _REQUIRED_CATS_BY_TASK.get(task, set())
    allowed_table = _get_ui_allowed_exts(task)
    name_override = name_override or {}

    _SENTINEL = object()  # 「override に key が存在しない」を None と区別するための番兵
    for cat, _, keywords in _FILE_CATEGORIES:
        if cat not in required_cats:
            continue
        allowed = allowed_table.get(cat)

        # 1) ユーザー override に明示エントリがあれば優先
        override_val = name_override.get(cat, _SENTINEL)
        if override_val is not _SENTINEL and override_val is not None:
            # 空リストは「対象外」明示 → 必須カテゴリでは即座に未充足判定
            if not override_val:
                return False
            # ファイル指定あり → 拡張子要件を満たしているか確認
            ok = all(
                allowed is None or Path(n).suffix.lower() in allowed
                for n in override_val
            )
            if not ok:
                return False
            continue  # 充足

        # 2) override 未指定（key なし or None）→ 自動検出: キーワード + 拡張子
        def _name_ok(name: str) -> bool:
            if not any(kw in name for kw in keywords):
                return False
            if allowed is None:
                return True
            return Path(name).suffix.lower() in allowed

        if not any(_name_ok(name) for name in names_nfc):
            return False
    return True


# ── ファイル入力 ──
st.markdown(
    '<span class="step-number">1</span>'
    '<span class="step-title">資料を準備</span>',
    unsafe_allow_html=True,
)

# 賃金台帳の作成手順（CC向け Skill）— 手元の Claude Code に Skill をインストールして使う運用。
# Skill 本体は .claude/skills/wagebook-convert/ で管理（git）、配布は ZIP 経由（Drive）。
# 過去事故：暦年と事業年度が混在した独自フォーマットで作って R215/R216 が誤集計された
# Skill ZIP の更新日（build_skill_zip.py 実行後に手動更新）
_WAGEBOOK_SKILL_VERSION = '2026-06-04'
# Drive 共有リンク（wagebook-convert.zip を配置した Drive ファイル URL）
# 配布先: マイドライブ/補助金ツール/wagebook-convert.zip （カラフルボックス株式会社グループ閲覧可）
_WAGEBOOK_SKILL_ZIP_URL = 'https://drive.google.com/file/d/1VxD0y4l7DPb9qK7eDfOpWUDQvMZ9eBBF/view?usp=sharing'
with st.expander(
    '🔧 はじめての方へ：CC環境セットアップ（Python・1回だけ）',
    expanded=False,
):
    st.markdown(
        'スキルを使う前に、**このPCで1回だけ**実施してください。'
        '手元の Claude Code（CC）に下のプロンプトを貼り付けて実行するだけです。'
        'Python 3.8 以上が無ければ CC が導入を試みます。\n\n'
        '※ スキル本体（ZIP）を更新しても、このセットアップの再実行は不要です。'
    )
    st.code(
        '【補助金・賃金台帳スキル：CC環境の初回セットアップ（このPCで1回だけ）】\n'
        'あなた（Claude Code）は、賃金台帳変換スキルを動かす環境をこのPCに用意します。\n'
        '順に実行し、最後に結果だけ報告してください。\n'
        '\n'
        '1. Python を探す：`python --version` / `python3 --version`、WSLがあれば `wsl python3 --version` も試す。\n'
        '   Python 3.8 以上 が見つかればそれを使う（新しいほど良い）。\n'
        '2. 3.8以上が見つからない場合は導入する：\n'
        '   - Windowsネイティブ：`winget install -e --id Python.Python.3.12` を実行し、入れ直して再確認。\n'
        '   - winget が無い/失敗する場合：python.org から Python 3.12 を入れて再実行する旨をユーザーに伝えて停止。\n'
        '   - WSLのみの環境：WSL側の python3 が 3.8 以上ならそれを使う。\n'
        '3. openpyxl を導入：見つかった Python で `python -m pip install openpyxl`\n'
        '   （WSLの python3 を使う場合は `pip3 install openpyxl --break-system-packages`）。\n'
        '4. 動作確認：`python -c "import openpyxl; print(\'OK\', openpyxl.__version__)"`（使う Python に合わせる）。\n'
        '5. 報告：使った Python のパスとバージョン、openpyxl のバージョン、最終結果（OK/NG）。\n'
        '\n'
        '※ これは1回だけ。スキル本体（wagebook-convert.zip）を更新しても、このセットアップの再実行は不要です。',
        language='markdown',
    )

with st.expander(
    '📘 賃金台帳の作成手順（CC向け Skill） — 賃金台帳を作る/直すときはここを確認',
    expanded=False,
):
    st.info(
        '**この作業は手元の PC の Claude Code（CC）に行わせます。**\n\n'
        '人がやる作業は「自分の PC の CC に `wagebook-convert` Skill をインストールして、賃金台帳の変換を依頼する」だけ。\n'
        '具体的な変換手順・テンプレート・検証チェックリスト・サンプルは Skill 内に同梱されています。'
    )
    st.warning(
        '**📋 対応フォーマット**：Excel / CSV / テキストPDF が推奨です。\n\n'
        '画像PDF（特に **手書き** スキャン）は OCR 精度が落ちるため、'
        'CC が抽出した後に **必ず人が PDF原本と全数値を照合してください**。'
        '顧客には可能な限り Excel/CSV 形式での提出を依頼するのが理想です。'
    )

    st.markdown('### 🔧 初回セットアップ（1回だけ実施）')
    st.markdown('※ 先に上の「🔧 はじめての方へ：CC環境セットアップ（Python）」を1回済ませてください。')
    st.markdown(
        f'1. [`wagebook-convert.zip` をダウンロード]({_WAGEBOOK_SKILL_ZIP_URL})\n'
        '2. ZIP を展開し、フォルダ `wagebook-convert/` を以下に配置：\n'
        '   - Windows: `C:\\\\Users\\\\<ユーザー名>\\\\.claude\\\\skills\\\\wagebook-convert\\\\`\n'
        '   - macOS: `~/.claude/skills/wagebook-convert/`\n'
        '   - **`.claude` の下に `skills` フォルダが無ければ、自分で作成してください**'
        '（初めての方は無いのが普通です）。`skills` の中に展開した `wagebook-convert/` を丸ごと置き、'
        '`...\\\\skills\\\\wagebook-convert\\\\SKILL.md` の形になれば配置完了です。\n'
        '3. Claude Code を再起動\n'
        '4. CC に `/wagebook-convert` と打って Skill 名が候補に出れば成功\n\n'
        f'**最新版: {_WAGEBOOK_SKILL_VERSION}**（手元の Skill が古い場合は再ダウンロードして上書き）'
    )

    st.markdown('### 👤 CC への依頼方法（毎回）')
    st.markdown(
        'Skill インストール後は、手元の Claude Code に **下のプロンプトを貼り付け、`［　］` を埋めて送るだけ** です。'
    )
    st.code(
        '賃金台帳を補助金ツール用の Excel に変換してください。\n'
        '\n'
        '■ 必須\n'
        '・会社名：［　　　　　　　　］\n'
        '・賃金台帳PDF：［ローカルのファイル名/パス、または Drive共有URL。\n'
        '　　　　　　　　複数ファイルに分かれている場合は全部挙げる］\n'
        '・決算月（1〜12の数字）：［　］月\n'
        '・法人／個人事業主：［法人 or 個人事業主］\n'
        '\n'
        '■ 強く推奨（特に法人）\n'
        '・履歴事項全部証明書PDF：［ファイル名/URL、なければ「なし」］\n'
        '\n'
        '■ 任意\n'
        '・賞与の支給月（台帳に「令和7年1回」等としかなく支給月が不明なときだけ）：［わかれば］',
        language='markdown',
    )
    st.markdown(
        '- **決算月** は数字だけでOK（**決算書そのものは渡さなくて大丈夫**）。未指定だと事業年度ズレ事故の原因に。\n'
        '- **法人／個人** は **個人事業主のときだけ重要**（事業主を除外・専従者を算入）。法人なら「法人」とだけ。\n'
        '- **履歴事項全部証明書** は無くても変換できますが、**法人では役員の自動判定に使うため強く推奨**'
        '（無いと従業員数・給与支給総額が過大に出ることがあります）。\n'
        '- 通常はこのプロンプトを送るだけで Skill が自動的に起動します。'
        'もし起動しない場合は、先頭に `/wagebook-convert` と入力してください。'
    )

    st.warning(
        '**変換結果がおかしい／Skill が起動しないとき**\n\n'
        f'1. **まず手元の Skill が最新版（{_WAGEBOOK_SKILL_VERSION}）か確認してください。** '
        '古ければ、上の「初回セットアップ」のリンクから `wagebook-convert.zip` をもう一度ダウンロードし、'
        '`wagebook-convert/` を上書きして Claude Code を再起動。'
        '不具合のときは「手元が古い」可能性が高いので、まずこれを試します。\n'
        '2. **最新版でもまだおかしいときは、羽根に共有してください。** '
        '会社名・症状（例：給与支給総額が通勤手当ぶん少ない／人数が多すぎる 等）・CC が出力した Excel を添えてください。'
    )

    st.caption(
        '※ Skill 内に手順書（SKILL.md）・テンプレート・サンプル変換例・検証チェックリスト・実テスト 11 案件のケーススタディが同梱されています。'
        'CC は必要に応じて参照しながら作業します。'
        '手書き／画像PDF の場合は §4.4 のゲート判定で読み取り困難と判定したらデータ転記を停止し、代替ソース取得を打診します（低品質データを出さない設計）。'
    )

# Drive連携用の変数
drive_folder_id = None
drive_files_to_download = []
# 差し替え UI が積み上げるカテゴリ別ファイル名選択（None=自動、[]=対象外、list=明示指定）
selection_override_names: dict[str, list[str] | None] = {}

if data_source == 'Google Drive':
    # ── Google Drive モード ──
    st.caption('Google Driveの顧客フォルダからファイルを自動取得します。')

    client = _get_drive_client()
    if client is None:
        st.error('Drive接続に失敗しました。認証情報を確認してください。')
    else:
        # PARENT_IDをSecretsからも取得
        parent_id = _DRIVE_PARENT_ID
        if not parent_id:
            try:
                parent_id = st.secrets.get('drive_parent_folder_id', '')
            except Exception as e:
                logger.warning(
                    f'Streamlit Secrets の drive_parent_folder_id 取得に失敗: {e}',
                    exc_info=True,
                )

        # 顧客フォルダ一覧（モジュールレベルのキャッシュ関数を使用）
        folders = _cached_list_drive_folders(parent_id)
        # 表示はラベル、選択値は folder ID（同名フォルダ対策）。None は未選択 sentinel。
        folder_id_to_name = {f['id']: f['name'] for f in folders}
        folder_options: list = [None] + [f['id'] for f in folders]

        selected_folder_id = st.selectbox(
            '顧客フォルダを選択',
            folder_options,
            format_func=lambda fid: '（選択してください）' if fid is None else folder_id_to_name.get(fid, fid),
            help='Driveの2026フォルダ直下の顧客フォルダ一覧です。',
        )

        drive_folder_id = None  # サブフォルダ未選択 / 顧客未選択 の場合は None で実行ガード

        if selected_folder_id is not None:
            selected_folder_name = folder_id_to_name[selected_folder_id]
            parent_folder_id = selected_folder_id

            # サブフォルダ（案件単位）があるかチェック
            # 例: 紹介会社/代理店の親フォルダ直下に複数の顧客企業案件フォルダが並ぶケース
            sub_folders = _cached_list_drive_folders(parent_folder_id)

            if sub_folders:
                # サブフォルダがある場合は **必ず案件を選ばせる**（直下のみ使用は再帰取得で
                # 全案件混在になるため、選択肢として提供しない）
                sub_id_to_name = {f['id']: f['name'] for f in sub_folders}
                sub_options: list = [None] + [f['id'] for f in sub_folders]
                selected_sub_id = st.selectbox(
                    '案件フォルダを選択',
                    sub_options,
                    format_func=lambda fid: (
                        '（案件を選択してください）' if fid is None
                        else sub_id_to_name.get(fid, fid)
                    ),
                    help=(
                        f'「{selected_folder_name}」の下に {len(sub_folders)}件の案件フォルダがあります。'
                        '対象の案件を選んでください。'
                    ),
                )
                if selected_sub_id is not None:
                    drive_folder_id = selected_sub_id
            else:
                # サブフォルダなし = フラットな顧客フォルダ → 親フォルダ直下のファイルを使用
                drive_folder_id = parent_folder_id

        # ファイル一覧取得（drive_folder_id が確定している時のみ）
        all_files = (
            _cached_list_drive_files_recursive(drive_folder_id)
            if drive_folder_id else []
        )

        if all_files:
            drive_files_to_download = all_files

            drive_analysis = _analyze_files([f['name'] for f in all_files], task_type)
            _render_file_check_result(drive_analysis, len(all_files))

            # 検出されたファイルの差し替え UI（自動検出が誤ったときの保険）
            selection_override_names = _render_file_selection_override(
                [f['name'] for f in all_files],
                case_key=f'drive_{drive_folder_id}',
                task=task_type,
            )
            # 容量・件数の事前警告（Drive ファイルのサイズは f['size'] が文字列の場合があるので int 変換）
            size_pairs = []
            for f in all_files:
                sz = f.get('size', 0)
                try:
                    sz = int(sz) if sz else 0
                except (TypeError, ValueError):
                    sz = 0
                size_pairs.append((f['name'], sz))
            # 案件規模・処理時間・APIコスト予想（実行前にユーザーに把握してもらう）
            _render_case_scale_estimate(size_pairs, task=task_type)
            for w in _check_size_warnings(size_pairs, task=task_type):
                st.warning(w)

            with st.expander('ファイル一覧（Drive上の場所）', expanded=False):
                for f in all_files:
                    loc = f.get('folder_name', 'ルート')
                    st.text(f'  [{loc}] {f["name"]}')

        elif drive_folder_id:
            # フォルダ選択済みだがファイル0件
            st.warning('このフォルダにはファイルがありません。')
        # drive_folder_id が None（顧客/案件未選択）の場合は何も表示しない

    uploaded_files = None

    if task_type in ('application', 'all'):
        st.markdown('---')
        template_file = st.file_uploader(
            '申請フォーマット（任意）',
            accept_multiple_files=False,
            type=['xlsx'],
            key='template_uploader_drive',
            help='ツール名を選択済みの原本をアップロードするとそのファイルを使用します。未アップロード時はツール同梱の原本を使用します。',
        )
        if template_file is not None:
            st.caption(f'申請フォーマット: `{template_file.name}`')
    else:
        template_file = None

else:
    # ── ファイルアップロードモード ──
    st.caption('ファイルはファイル名のキーワードで自動判別されます。該当キーワードがないファイルは無視されます。')

    # タスク別にファイルカードを表示
    # (カテゴリ, 表示名, 形式, キーワード, 例, 表示するタスク, 必須のタスク)
    _file_cards = [
        ('hearing',     'ヒアリングシート',        'Excel',     ['ヒアリング'],
         'ヒアリングシート_○○株式会社.xlsx',       {'application', 'all'},          {'application', 'all'}),
        ('registry',    '履歴事項全部証明書',      'PDF',       ['履歴事項'],
         '履歴事項全部証明書_○○様.pdf',           {'application', 'all'},          {'application', 'all'}),
        ('pl',          '損益計算書 / 決算報告書', 'PDF',       ['損益計算書', '決算報告書', '決算書'],
         '42期 決算報告書.pdf',                    {'application', 'wage', 'all'},  {'application', 'all'}),
        ('wage_ledger', '賃金台帳 / 給与台帳',     'Excel/CSV',     ['賃金台帳', '給与台帳'],
         '賃金台帳_2025年度.xlsx',                 {'wage', 'bonus'},              {'wage', 'bonus'}),
        ('cost_report', '製造原価報告書',          'PDF',       ['製造原価報告書', '原価報告書'],
         '製造原価報告書.pdf',                     {'application', 'wage', 'all'},  set()),
        ('tax',         '納税証明書',              'PDF',       ['納税証明'],
         '納税証明書(その1)_○○様.pdf',            {'application', 'all'},          set()),
        ('estimate',    '見積書',                  'Excel/PDF', ['見積'],
         'お見積書_○○.pdf',                       {'application', 'all'},          set()),
        ('wage_report', '賃金状況報告シート',      'Excel',     ['賃金状況報告'],
         '賃金状況報告シート.xlsx',                 {'wage', 'all'},                set()),
    ]

    # 現在のタスクに関連するカードのみ表示
    visible_cards = [c for c in _file_cards if task_type in c[5]]
    # タスクに応じて必須/任意を判定して分割
    required_cards = [c for c in visible_cards if task_type in c[6]]
    optional_cards = [c for c in visible_cards if task_type not in c[6]]

    col1, col2 = st.columns(2)

    def _render_card(card, is_required):
        name, fmt, keywords, example = card[1], card[2], card[3], card[4]
        badge = '必須' if is_required else 'あれば'
        css_class = 'file-required' if is_required else 'file-optional'
        badge_class = 'badge-required' if is_required else 'badge-optional'
        kw_html = ' '.join(f'<span class="keyword-tag">{kw}</span>' for kw in keywords)
        return (
            f'<div class="file-card {css_class}">'
            f'<span class="{badge_class}">{badge}</span><br>'
            f'<strong>{name}</strong>（{fmt}）<br>'
            f'{kw_html} がファイル名に含まれること<br>'
            f'<small>例: {example}</small>'
            f'</div>'
        )

    with col1:
        st.markdown('\n'.join(_render_card(c, True) for c in required_cards), unsafe_allow_html=True)

    with col2:
        st.markdown('\n'.join(_render_card(c, False) for c in optional_cards), unsafe_allow_html=True)

    with st.expander('その他の注意事項'):
        st.markdown("""
- キーワードが含まれないファイルもアップロード可。**「使用するファイルを差し替える」expander**から手動でカテゴリ割当てできます（タイポ救済）
- 決算書が複数期分ある場合、**ファイル名の「第N期」「令和N年M月」「YYYY-MM」**等から直近期を自動選定します
- 関係ないファイルが混ざっていても問題ありません
- テンプレート選択（通常枠/インボイス枠）とテンプレート原本の種類を**一致**させてください
        """)

    uploaded_files = st.file_uploader(
        'ここにファイルをまとめてドラッグ&ドロップ（複数選択可）',
        accept_multiple_files=True,
        type=['pdf', 'xlsx', 'xls', 'csv'],
        key='file_uploader',
    )

    # アップロード済みファイルのチェックリスト（タスク別に必須/任意を切り替え）
    if uploaded_files:
        upload_analysis = _analyze_files([f.name for f in uploaded_files], task_type)
        _render_file_check_result(upload_analysis, len(uploaded_files))

        # 検出されたファイルの差し替え UI（同名で別期の決算書混在などへの保険）
        # case_key にはファイル名+サイズの安定ハッシュ（sha1）を使用。
        # 1) sha1 はプロセス再起動でも値が変わらない（hash() は PYTHONHASHSEED でランダム化）
        # 2) サイズも組み込むことで「別案件で同名 generic ファイルが衝突」を回避
        _upload_signature = '|'.join(
            f'{f.name}:{f.size}' for f in sorted(uploaded_files, key=lambda x: x.name)
        )
        _upload_case_key = (
            f'upload_{hashlib.sha1(_upload_signature.encode("utf-8")).hexdigest()[:16]}'
        )
        selection_override_names = _render_file_selection_override(
            [f.name for f in uploaded_files],
            case_key=_upload_case_key,
            task=task_type,
        )
        # 案件規模・処理時間・APIコスト予想
        size_pairs_upload = [(f.name, f.size) for f in uploaded_files]
        _render_case_scale_estimate(size_pairs_upload, task=task_type)
        # 容量・件数の事前警告（処理は続行可能、ユーザーに確認を促すだけ）
        size_warnings = _check_size_warnings(size_pairs_upload, task=task_type)
        for w in size_warnings:
            st.warning(w)

    if task_type in ('application', 'all'):
        st.markdown('---')
        template_file = st.file_uploader(
            '申請フォーマット（任意）',
            accept_multiple_files=False,
            type=['xlsx'],
            key='template_uploader',
            help='ツール名を選択済みの原本をアップロードするとそのファイルを使用します。未アップロード時はツール同梱の原本を使用します。',
        )
        if template_file is not None:
            st.caption(f'申請フォーマット: `{template_file.name}`')
    else:
        template_file = None

# ── 処理実行 ──
st.markdown(
    '<span class="step-number">2</span>'
    '<span class="step-title">処理実行</span>',
    unsafe_allow_html=True,
)

has_files = bool(uploaded_files)
has_drive_files = bool(drive_files_to_download)
# 差し替え UI でユーザーがカテゴリ割当しているファイルも「必須充足」とみなす
# （タイポ救済: キーワード未一致でも override に入っていれば OK）
has_required = (
    _check_required_by_names(
        [f.name for f in uploaded_files], task_type,
        name_override=selection_override_names,
    )
    if has_files else False
)
has_drive_required = (
    _check_required_by_names(
        [f['name'] for f in drive_files_to_download], task_type,
        name_override=selection_override_names,
    )
    if has_drive_files else False
)

# 必須ファイル不足時はボタンを押せないようにする（不完全ファイルでのAPI消費防止）
if data_source == 'Google Drive':
    has_data = has_drive_files
    required_ok = has_drive_required
else:
    has_data = has_files
    required_ok = has_required

_is_bonus_task = task_type in ('bonus', 'bonus_wage_ledger_creation')

can_run = bool(company_name) and has_data and required_ok
if _is_bonus_task:
    can_run = can_run and bool(prefecture)
# 決算月の指定を必須化（2026-05 方針）。賃金台帳の対象12ヶ月を確定するために必要。
# ただし加点系タスクは暦月固定（R6/10〜R7/9＋申請直近月）で判定するため決算月は不要。
if not _is_bonus_task:
    can_run = can_run and (fiscal_month_override is not None)

if not company_name:
    st.warning('⬅️ サイドバーで会社名を入力してください')
elif not _is_bonus_task and fiscal_month_override is None:
    st.warning('⬅️ サイドバーで決算月を選択してください（賃金台帳の対象期間を確定するため必須）')
elif _is_bonus_task and not prefecture:
    st.warning('⬅️ サイドバーで事業場の都道府県を選択してください')
elif data_source == 'Google Drive' and not has_drive_files:
    st.warning('⬅️ サイドバーで顧客フォルダを選択してください')
elif data_source != 'Google Drive' and not has_files:
    st.warning('⬆️ 資料ファイルをアップロードしてください')
elif (data_source == 'Google Drive' and not has_drive_required) \
        or (data_source != 'Google Drive' and not has_required):
    st.warning('⬆️ 必須ファイルが不足しています。ファイル判別結果を確認してください')
else:
    source_label = 'Google Drive' if data_source == 'Google Drive' else 'アップロード'
    if task_type == 'bonus':
        st.info(
            f'**{company_name}** の加点判定用賃金台帳を読み取り、加点措置①②を判定します'
            f'（{source_label}）— 準備OKです\n\n'
            '※ 入力は「加点判定用賃金台帳の作成」で作った専用台帳（シート『加点判定用明細』）です。'
        )
    elif task_type == 'bonus_wage_ledger_creation':
        msg = (
            f'**{company_name}** の賃金台帳/給与明細から「加点判定用賃金台帳」を作成します'
            f'（{source_label}）— 準備OKです\n\n'
            '※ 基本給と所定労働時間を抽出し、令和6年10月〜令和7年9月＋交付申請直近月の'
            '時間換算給与ベースの台帳に変換します。出力後は内容を必ず人手で確認してください。'
        )
        if application_ym is None:
            msg += '\n\n⚠️ 交付申請月（yyyy/mm）が未入力です。加点措置②の直近月列が空になります。'
        st.info(msg)
    elif task_type == 'per_employee_wage':
        st.info(
            f'**{company_name}** の賃金台帳から一人当たり給与支給総額を算定します'
            f'（{source_label}）— 準備OKです\n\n'
            '※決算書PDFは参照されません。決算書由来の8項目は出力Excelから自動削除されます。'
        )
    elif task_type == 'wage_ledger_creation':
        st.info(
            f'**{company_name}** の賃金台帳 PDF/Excel/CSV をツール規格の Excel に変換します'
            f'（{source_label}）— 準備OKです\n\n'
            '※ Document AI で抽出します（Sonnet 画像直送のフォールバックは無効）。'
            '手書きPDF は処理を続行しつつ警告を出します。'
            '出力ファイルはこのまま給与計算/加点判定の入力にも使えます。'
        )
    else:
        st.info(f'**{company_name}** の書類を **{template_label}** で作成します（{source_label}）— 準備OKです')

if st.button('処理開始', type='primary', disabled=not can_run, use_container_width=True):
    # 一時ディレクトリに保存
    with tempfile.TemporaryDirectory() as tmpdir:
        work_dir = Path(tmpdir)

        # ファイル保存（データソースに応じて）— 保存名は NFC 統一
        if data_source == 'Google Drive' and drive_files_to_download:
            with st.spinner('Google Driveからファイルをダウンロード中...'):
                from hojokin.drive_client import GoogleFormatNotSupportedError
                client = _get_drive_client()
                saved = []
                skipped: list[tuple[str, str]] = []
                for f in drive_files_to_download:
                    name = _nfc_filename(f['name'])
                    dest = work_dir / name
                    try:
                        saved_path = client.download_file(
                            f['id'], dest, mime_type=f.get('mimeType'),
                        )
                        saved.append(saved_path.name)
                    except GoogleFormatNotSupportedError as e:
                        # 対応外のGoogle形式（フォーム/図面/サイト/解決失敗ショートカット等）は
                        # スキップして処理続行。Excel/PDF 等の主要書類が揃っていれば申請書は作れる。
                        skipped.append((name, str(e)))
                        logger.warning(f'Drive ダウンロードスキップ: {name} ({e})')
                st.caption(
                    f'{len(saved)}件のファイルをダウンロードしました'
                    + (f'（{len(skipped)}件スキップ）' if skipped else '')
                )
                if skipped:
                    with st.expander(
                        f'⚠ ダウンロードできなかったファイル {len(skipped)}件',
                        expanded=False,
                    ):
                        for name, reason in skipped:
                            st.text(f'  {name}  — {reason}')
                        st.caption(
                            '上記は Google フォーム / 図面 / サイト等の対応外形式です。'
                            '主要書類（Excel/PDF）が揃っていれば申請書は作成できます。'
                        )
        else:
            saved = save_uploaded_files(uploaded_files, work_dir)

        if template_file is not None:
            template_dest = work_dir / _nfc_filename(template_file.name)
            template_dest.write_bytes(template_file.getvalue())
            template_dir = work_dir
        else:
            template_dir = Path(__file__).parent

        # ファイル名ベースの override 選択を、work_dir 内の Path に解決する
        selection_override = _build_path_override(
            selection_override_names, work_dir,
        )

        # ファイル検出プレビュー（summary 表示用）— UI と同じ override を反映
        detector = FileDetector(work_dir, selection_override=selection_override)

        # 処理実行
        if task_type == 'wage_ledger_creation':
            spinner_msg = '賃金台帳PDFを Document AI で抽出中...（1〜3分かかります）'
        elif task_type == 'bonus_wage_ledger_creation':
            spinner_msg = '加点判定用の賃金台帳を作成中...（1〜3分かかります）'
        elif task_type == 'bonus':
            spinner_msg = '加点判定用賃金台帳を読み取り中...'
        elif task_type == 'per_employee_wage':
            spinner_msg = '賃金台帳を分析中...'
        else:
            spinner_msg = 'AIが資料を読み取り中...（1〜3分かかります）'
        with st.spinner(spinner_msg):
            results = run_processing(
                company_name=company_name,
                template_type=template_type,
                task_type=task_type,
                work_dir=work_dir,
                template_dir=template_dir,
                prefecture=prefecture,
                application_ym=application_ym,
                fiscal_month_override=fiscal_month_override,
                has_cost_report_hint=has_cost_report_hint,
                selection_override=selection_override,
                bonus_paid_months=bonus_paid_months,
            )

        # Drive 格納（オプション）— Drive ソース選択 + チェックON + 格納先フォルダ確定時のみ
        drive_upload_links: dict[str, str] = {}
        drive_upload_errors: list[str] = []
        if (
            upload_to_drive
            and data_source == 'Google Drive'
            and drive_folder_id
        ):
            try:
                with st.spinner('Driveへアップロード中...'):
                    client = _get_drive_client()
                    # work_dir 内の出力ファイル（*.xlsx）を全てアップロード
                    for task_name, result in results.items():
                        out = result.get('output_path')
                        if out and out.exists():
                            res = client.upload_file(out, drive_folder_id)
                            drive_upload_links[out.name] = res.get('webViewLink', '')
                        for fname, fpath in result.get('extra_files', {}).items():
                            if fpath.exists():
                                res = client.upload_file(fpath, drive_folder_id)
                                drive_upload_links[fname] = res.get('webViewLink', '')
                        for key, fpath in result.get('output_files', {}).items():
                            if isinstance(fpath, Path) and fpath.exists():
                                res = client.upload_file(fpath, drive_folder_id)
                                drive_upload_links[fpath.name] = res.get('webViewLink', '')
                st.success(
                    f'✅ Driveへ {len(drive_upload_links)} ファイルを格納しました'
                )
            except Exception as e:
                drive_upload_errors.append(str(e))
                logger.warning(f'Driveアップロード失敗: {e}', exc_info=True)
                st.warning(
                    f'⚠ Driveアップロードに失敗しました: {e}\n'
                    'ローカルダウンロードボタンから結果を取得できます。'
                )

        # 結果をsession_stateに保存（画面再描画後も残る）
        session_results = {}
        # Drive アップロード結果を全タスク共通で session_state に保存
        st.session_state['drive_upload_links'] = drive_upload_links
        for task_name, result in results.items():
            entry = {
                'status': result['status'],
                'message': result['message'],
                'empty_cells': result.get('empty_cells', []),
                'confidence_warnings': result.get('confidence_warnings', []),
                'file_data': None,
                'file_name': None,
                'extra_files': {},
                'bonus_result': None,
                'bonus_files': {},
                # 直近年度として選定された決算書（UI 明示用）
                'pl_selected_filename': result.get('pl_selected_filename', ''),
                'pl_selected_end': result.get('pl_selected_end', ''),
                'pl_selection_warnings': result.get('pl_selection_warnings', []),
            }
            if result.get('output_path') and result['output_path'].exists():
                with open(result['output_path'], 'rb') as f:
                    entry['file_data'] = f.read()
                entry['file_name'] = result['output_path'].name

            # 追加ファイル（賃金台帳一覧等）
            for fname, fpath in result.get('extra_files', {}).items():
                if fpath.exists():
                    with open(fpath, 'rb') as f:
                        entry['extra_files'][fname] = f.read()

            # 加点判定の結果
            if task_name == 'bonus' and result.get('result'):
                br = result['result']
                entry['bonus_result'] = {
                    'bonus1_eligible': br.bonus1_eligible,
                    'bonus1_months_met': br.bonus1_months_met,
                    'bonus1_details': br.bonus1_details,
                    'bonus2_eligible': br.bonus2_eligible,
                    'bonus2_min_wage_july': br.bonus2_min_wage_july,
                    'bonus2_min_wage_latest': br.bonus2_min_wage_latest,
                    'bonus2_diff': br.bonus2_diff,
                    'bonus2_latest_label': (br.bonus2_latest_detail or {}).get('label', ''),
                    'prefecture': br.prefecture,
                    'min_wage_r6': br.min_wage_r6,
                    'min_wage_r7': br.min_wage_r7,
                    'application_ym': br.application_ym,
                    'latest_ym': br.latest_ym,
                    'notes': br.notes,
                    'employee_count': result.get('employee_count', 0),
                }
                # 加点シートのファイルデータ
                for key, path in result.get('output_files', {}).items():
                    if path.exists():
                        with open(path, 'rb') as f:
                            entry['bonus_files'][key] = {
                                'data': f.read(),
                                'name': path.name,
                            }

            session_results[task_name] = entry

        st.session_state['last_results'] = session_results
        st.session_state['last_company'] = company_name
        st.session_state['last_template'] = template_label
        st.session_state['last_time'] = datetime.now().strftime('%Y-%m-%d %H:%M')
        st.session_state['last_detector_summary'] = detector.summary()

# ── 結果表示（session_stateから復元） ──
if 'last_results' in st.session_state:
    st.markdown('---')
    st.markdown(
        '<span class="step-number">3</span>'
        '<span class="step-title">結果・ダウンロード</span>',
        unsafe_allow_html=True,
    )
    st.caption(
        f'処理日時: {st.session_state["last_time"]} | '
        f'会社名: {st.session_state["last_company"]} | '
        f'テンプレート: {st.session_state["last_template"]}'
    )

    with st.expander('検出されたファイル（デバッグ用）'):
        st.code(st.session_state.get('last_detector_summary', ''))

    for task_name, result in st.session_state['last_results'].items():
        task_display_map = {
            'application': '📝 申請書作成',
            'wage': '💰 給与支給総額計算',
            'per_employee_wage': '👤 一人当たり給与支給総額（賃金台帳のみ）',
            'bonus': '📊 加点判定',
            'bonus_wage_ledger_creation': '📑 加点判定用賃金台帳の作成（PDF→専用Excel）',
            'wage_ledger_creation': '📑 賃金台帳の作成（PDF→ツール規格Excel）',
        }
        task_display = task_display_map.get(task_name, task_name)

        if result['status'] == '完了':
            # 賃金台帳の読み取り警告がメッセージに含まれていれば警告表示、それ以外は成功表示
            if '⚠' in result['message']:
                st.warning(f'{task_display}: 完了（一部警告あり）— {result["message"]}')
            else:
                st.success(f'{task_display}: 完了 — {result["message"]}')

            # 直近年度として使用した決算書ファイルを明示（誤選定の早期発見用）
            _pl_name = result.get('pl_selected_filename')
            if _pl_name:
                _pl_end = result.get('pl_selected_end')
                _suffix = f'（推定期末: {_pl_end}）' if _pl_end else ''
                st.info(
                    f"📄 直近年度として **『{_pl_name}』** を使用しました{_suffix}。"
                    "違う期の決算書が必要な場合は、ファイル名に「令和N年」「YYYY年」など"
                    "年情報を含めて Drive を更新してください。"
                )
            for _w in result.get('pl_selection_warnings', []):
                st.warning(_w)

            # Drive 格納時のリンクを表示（ダウンロードボタンと併設）
            drive_links = st.session_state.get('drive_upload_links') or {}

            if result['file_data']:
                st.download_button(
                    label=f'⬇️ {result["file_name"]} をダウンロード',
                    data=result['file_data'],
                    file_name=result['file_name'],
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    use_container_width=True,
                    key=f'download_{task_name}',
                )
                if result['file_name'] in drive_links and drive_links[result['file_name']]:
                    st.markdown(
                        f'📂 Drive で開く: [{result["file_name"]}]({drive_links[result["file_name"]]})'
                    )

            # 追加ファイル（賃金台帳一覧等）
            for fname, fdata in result.get('extra_files', {}).items():
                st.download_button(
                    label=f'⬇️ {fname} をダウンロード',
                    data=fdata,
                    file_name=fname,
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    use_container_width=True,
                    key=f'download_extra_{fname}',
                )
                if fname in drive_links and drive_links[fname]:
                    st.markdown(f'📂 Drive で開く: [{fname}]({drive_links[fname]})')

            # 加点判定の結果表示
            if task_name == 'bonus' and result.get('bonus_result'):
                br = result['bonus_result']

                st.markdown(
                    f"**事業場所在地:** {br['prefecture']}"
                    f"（地域別最低賃金 改定前: {br['min_wage_r6']}円 → 改定後(R7): {br['min_wage_r7']}円）"
                )

                st.caption(
                    f"従業員{br.get('employee_count', 0)}名（役員は判定母数から除外）｜"
                    '時間換算給与＝基本給÷月間所定労働時間で算定（暦月固定：'
                    '加点①＝令和6年10月〜令和7年9月、加点②＝令和7年7月 と 交付申請直近月）'
                )

                # 判定上の注意（対象月の欠落・所在地/申請月の未入力など）を必ず surfacing
                for _note in br.get('notes', []):
                    st.warning(f'⚠️ {_note}')

                col_b1, col_b2 = st.columns(2)
                # 加点措置①（公式名: 補助率引上げ・加点措置① ／ 加点項目14・補助率2/3トリガー）
                with col_b1:
                    st.caption('加点措置①（加点項目14｜通常枠では補助率1/2→**2/3** のトリガー）')
                    if br['bonus1_eligible']:
                        st.success(
                            f"**① 対象** "
                            f"({len(br['bonus1_months_met'])}か月が30%以上を達成／3か月必要)"
                        )
                    else:
                        st.warning(
                            f"**① 対象外** "
                            f"({len(br['bonus1_months_met'])}か月/3か月必要)"
                        )
                    st.caption('判定: R7改定後最低賃金 **未満**で雇用の従業員が全従業員の30%以上の月')

                    with st.expander('月別詳細（R7改定後未満の人数／全従業員）'):
                        for d in br['bonus1_details']:
                            if d.get('total', 0) > 0:
                                mark = '○' if d['meets_30pct'] else '×'
                                st.text(
                                    f"{d['label']}: {d['under_r7']}/{d['total']}名 "
                                    f"= {d['ratio']*100:.1f}% {mark}"
                                )
                            else:
                                st.text(f"{d['label']}: データなし")

                # 加点措置②（公式名: 加点措置② ／ 加点項目15）
                with col_b2:
                    st.caption('加点措置②（加点項目15）')
                    if br['bonus2_eligible']:
                        st.success(f"**② 対象** (差額 {br['bonus2_diff']:.0f}円 ≥ 63円)")
                    else:
                        st.warning(f"**② 対象外** (差額 {br['bonus2_diff']:.0f}円 < 63円)")
                    _latest_lbl = br.get('bonus2_latest_label') or '交付申請直近月'
                    st.caption('判定: 交付申請直近月の事業場内最低賃金 ≥ 令和7年7月＋63円')
                    st.text(f"令和7年7月 最低時給: {br['bonus2_min_wage_july']:.0f}円")
                    st.text(f"{_latest_lbl} 最低時給: {br['bonus2_min_wage_latest']:.0f}円")

                # 加点シートダウンロード
                for key, file_info in result.get('bonus_files', {}).items():
                    label_map = {
                        'bonus1': '加点措置①シート',
                        'bonus2': '加点措置②シート',
                    }
                    st.download_button(
                        label=f"⬇️ {label_map.get(key, key)} をダウンロード",
                        data=file_info['data'],
                        file_name=file_info['name'],
                        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                        use_container_width=True,
                        key=f'download_{key}',
                    )

                # B4: 加点項目の全体像（自動判定2項目＋手動確認項目）
                with st.expander('📋 加点項目の全体像（自動判定は①②のみ・他は手動確認）'):
                    st.markdown(BONUS_ITEMS_REFERENCE_MD)

            # 空セル表示
            if result.get('empty_cells'):
                with st.expander(f'⚠️ 未入力セル（{len(result["empty_cells"])}件） — 手動確認が必要'):
                    for cell in result['empty_cells']:
                        st.text(cell)

            # Phase 4: 低信頼項目の確認キュー（自動転記できなかった項目を一覧）
            warnings_list = result.get('confidence_warnings') or []
            if warnings_list:
                with st.expander(
                    f'📋 確認キュー（{len(warnings_list)}件） — '
                    f'AI が自信を持って抽出できなかった項目です',
                    expanded=True,
                ):
                    st.caption(
                        '以下の項目は信頼度が低いため**空欄にしてあります**。'
                        '元の決算書を見て手動で入力してください。'
                    )
                    for w in warnings_list:
                        st.markdown(
                            f'**{w["label"]}** （取得元: {w["source"]}）  \n'
                            f'　▶ AI抽出値: `{w["value"]}` （**未転記**）  \n'
                            f'　▶ 警告理由: {w["reason"]}'
                        )
        else:
            st.error(f'{task_display}: {result["message"]}')

    # ── 人間チェックリスト ──
    st.markdown('---')
    st.markdown(
        '<span class="step-number">4</span>'
        '<span class="step-title">ダウンロード後の確認事項</span>',
        unsafe_allow_html=True,
    )
    st.warning('AIが自動生成した内容です。提出前に必ず以下の項目を確認してください。')

    st.markdown("""
**申請内容シート（AIが読み取り・生成した項目）**

| 確認項目 | 確認ポイント | よくあるミス |
|---|---|---|
| **役員情報** | 氏名・役職が正しいか、退任済みの人が含まれていないか | 同一人物が重複して登録される |
| **本店所在地** | 履歴事項と一致しているか（抹消線の旧住所になっていないか） | 移転前の住所が入る |
| **設立年月日** | 正しい日付か | 和暦/西暦の変換ミス |
| **業種コード（4桁）** | 実際の主要事業と一致しているか | 類似業種の取り違え |
| **事業内容（255文字）** | 内容に違和感がないか、ツール名が正しいか | AIが実態と異なる記述をする |
| **財務数値** | 売上高・営業利益・経常利益が決算書と一致しているか | 桁の読み間違い |
| **減価償却費** | 販管費と原価報告書の合計になっているか | 片方だけ拾っている |
| **賃上げ関連** | 表明方法・賃上げ幅がお客さんの実態に合っているか | デフォルト値のまま |

**転記シート（ヒアリングシートから転記した項目）**

| 確認項目 | 確認ポイント |
|---|---|
| **電話番号** | 先頭の0が消えていないか |
| **従業員数** | 正規雇用・パート等の内訳が正しいか |
| **メールアドレス** | 全角文字が混入していないか |

**給与支給総額計算シート**

| 確認項目 | 確認ポイント |
|---|---|
| **給料手当・雑給・賞与** | 決算書（販管費内訳書）の数値と一致しているか |
| **従業員数・労働時間** | 空欄になっていないか（手入力が必要な場合あり） |
    """)

    st.info('確認が完了したら、申請内容シートの手順に沿ってgBizIDから申請を進めてください。')

    # 結果クリアボタン
    if st.button('結果をクリア', type='secondary'):
        del st.session_state['last_results']
        del st.session_state['last_company']
        del st.session_state['last_template']
        del st.session_state['last_time']
        if 'last_detector_summary' in st.session_state:
            del st.session_state['last_detector_summary']
        st.rerun()

# ── フッター ──
st.markdown('---')
st.caption(f'補助金書類自動作成ツール v0.2.53 | カラフルボックス株式会社')
