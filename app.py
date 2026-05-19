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
import unicodedata
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
)
from hojokin.wage_reader import (
    read_wage_ledgers, judge_bonus_points,
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
    '申請書作成のみ': 'application',
    '給与計算のみ': 'wage',
    '加点判定（賃金台帳）': 'bonus',
    '両方（申請書 + 給与計算）': 'all',
}

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
    fiscal_month_override: int | None = None,
    has_cost_report_hint: bool = False,
    selection_override: dict[str, list[Path]] | None = None,
):
    """メイン処理を実行"""
    results = {}

    # Extractor作成（加点判定もPDF/CSV対応のためAI経路を使用）
    extractor = None
    if task_type in ('application', 'wage', 'all', 'bonus'):
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

    if task_type == 'bonus':
        if progress_callback:
            progress_callback('賃金台帳を読み取り中...')

        results['bonus'] = _run_bonus_judgment(
            work_dir, company_name, prefecture, template_dir,
            extractor=extractor,
            selection_override=selection_override,
        )

    return results


def _run_bonus_judgment(
    work_dir: Path,
    company_name: str,
    prefecture: str,
    template_dir: Path,
    extractor=None,
    cached_ledger_employees: list | None = None,
    selection_override: dict[str, list[Path]] | None = None,
) -> dict:
    """加点判定を実行。

    cached_ledger_employees があれば再利用してAPI呼出をスキップする。
    なければ FileDetector で賃金台帳ファイル（Excel/CSV）を検索して AI 経路で読み取る。
    selection_override が渡されたら自動検出より優先する。
    """
    # キャッシュ優先（all + bonus を同時実行する将来拡張に備えた経路）
    if cached_ledger_employees:
        employees = cached_ledger_employees
    else:
        # FileDetector 経由で賃金台帳ファイルを取得（手動選択 override にも対応）。
        # 拡張子フィルタ・出力ファイル除外・NFC 正規化は detector 側で実施済み。
        detector = FileDetector(work_dir, selection_override=selection_override)
        wage_files = detector.get_all('wage_ledger')

        if not wage_files:
            return {
                'status': 'エラー',
                'message': (
                    '賃金台帳ファイルが見つかりません。Excel/CSV をアップロードしてください。'
                    'ファイル名に「賃金台帳」または「給与台帳」を含めてください。'
                ),
            }

        try:
            employees = read_wage_ledgers(wage_files, extractor=extractor)
        except Exception as e:
            return {
                'status': 'エラー',
                'message': f'賃金台帳の読み取り中にエラーが発生しました: {str(e)}',
            }

        if not employees:
            return {
                'status': 'エラー',
                'message': '賃金台帳からデータを読み取れませんでした。ファイルの形式を確認してください。',
            }

    try:
        result = judge_bonus_points(employees, prefecture)

        # 加点措置シートのテンプレートを探して自動入力
        bonus_dir = template_dir / '補助金加点'
        output_files = {}

        if bonus_dir.exists():
            for bp in bonus_dir.iterdir():
                if '加点措置①' in bp.name and bp.suffix == '.xlsx':
                    out = work_dir / f'{company_name}_加点措置①_結果.xlsx'
                    fill_bonus_sheet_1(bp, out, result)
                    output_files['bonus1'] = out
                elif '加点措置②' in bp.name and bp.suffix == '.xlsx':
                    out = work_dir / f'{company_name}_加点措置②_結果.xlsx'
                    fill_bonus_sheet_2(bp, out, result)
                    output_files['bonus2'] = out

        return {
            'status': '完了',
            'message': f'従業員{len(employees)}名の賃金台帳を分析しました。',
            'result': result,
            'output_files': output_files,
        }

    except Exception as e:
        return {
            'status': 'エラー',
            'message': f'処理中にエラーが発生しました: {str(e)}',
        }


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
        help='申請書作成：ヒアリングシート+各種PDFから申請書を自動作成。給与計算：損益計算書+賃金データから給与支給総額を計算。加点判定：賃金台帳から加点措置の対象かを判定。',
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

    # 製造原価ありフラグ — 製造業向け。チェック時、AI に「製造原価報告書が存在する」ヒントを注入
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

    # 加点判定の場合は都道府県が必要
    if task_type == 'bonus':
        from hojokin.config import MIN_WAGE_MAP
        prefecture = st.selectbox(
            '事業場の都道府県',
            [''] + list(MIN_WAGE_MAP.keys()),
            help='加点判定に必要です。事業場の所在地の都道府県を選択してください。',
        )
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
    st.caption('所要時間: 約1〜10分（案件規模により変動）')
    st.caption('API利用料: 約30〜300円/社（案件規模により変動）')
    st.caption('└ 賃金台帳PDFが大きい案件では数百円〜になることがあります')
    st.caption('└ Sonnet 4.6 の従量課金。確定額は Anthropic の請求でご確認ください')
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
    ('wage_ledger', '賃金台帳',                ['賃金台帳']),
]

_REQUIRED_CATS_BY_TASK = {
    'application': {'hearing', 'registry', 'pl'},
    'wage':        {'wage_ledger'},
    'bonus':       {'wage_ledger'},
    'all':         {'hearing', 'registry', 'pl'},
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


def _analyze_files(file_names, task):
    """ファイル名リストからタスク別の判別結果を計算"""
    required_cats = _REQUIRED_CATS_BY_TASK.get(task, set())

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
                allowed = _UI_ALLOWED_EXTS.get(cat)
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


def _check_size_warnings(file_size_pairs: list[tuple[str, int]]) -> list[str]:
    """ファイル(名前, バイト数)から大容量・大量警告と処理時間目安を作る。

    閾値（実害が出る前のソフト警告レベル）:
      - PDF 30MB超: API 残高消費が大きい・タイムアウトリスク
      - 賃金台帳PDF合計 6MB超: 処理時間目安を表示（情報レベル）
        → 7MB 前後から処理時間が顕著に伸びるため、6MB 超で事前案内
      - 賃金台帳ファイル合計 25MB超: AI 抽出に長時間かかる可能性（警告）
      - 賃金台帳ファイル 8件超: 個人別ファイル多数 → AI 抽出 max_tokens 不足の可能性
      - 単一 Excel/CSV 5MB超: 想定外に大きく、誤ったファイルの可能性
    """
    warnings = []
    PDF_LIMIT = 30 * 1024 * 1024
    EXCEL_CSV_LIMIT = 5 * 1024 * 1024
    WAGE_PDF_NOTICE_LIMIT = 6 * 1024 * 1024  # 賃金台帳PDFがこのサイズ超で処理時間目安を表示
    WAGE_TOTAL_LIMIT = 25 * 1024 * 1024
    WAGE_COUNT_LIMIT = 8
    # 2026-05 方針変更: 賃金台帳は Excel/CSV のみ。PDF は集計対象外
    # FileDetector.ALLOWED_EXTS['wage_ledger'] と整合させる
    WAGE_EXTS = {'.xlsx', '.xlsm', '.csv'}

    wage_keywords = ('賃金台帳',)
    wage_files = []
    wage_pdf_files = []
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
        if any(kw in n for kw in wage_keywords) and ext in WAGE_EXTS:
            wage_files.append((n, size))
            if ext == '.pdf':
                wage_pdf_files.append((n, size))

    # 賃金台帳PDF合計サイズに応じた処理時間目安（情報レベル）
    # 経験則: PDF 1MB あたり約 6 ページ、1 ページあたり 2〜3 秒の API 処理。
    # 7MB → 約 1.5〜2分 / 10MB → 2〜3分 / 14MB → 3〜5分（タイムアウト 480秒に近づく）
    if wage_pdf_files:
        pdf_total = sum(s for _, s in wage_pdf_files)
        if pdf_total > WAGE_PDF_NOTICE_LIMIT:
            mb = pdf_total / 1024 / 1024
            est_min_sec = mb * 12   # 1MB ≈ 12 秒（楽観値）
            est_max_sec = mb * 25   # 1MB ≈ 25 秒（悲観値）
            est_min = max(1, round(est_min_sec / 60))
            est_max = max(est_min + 1, round(est_max_sec / 60))
            warnings.append(
                f'⏱ 賃金台帳PDFが大きめです（{mb:.1f}MB, {len(wage_pdf_files)}ファイル）— '
                f'処理時間目安: 約{est_min}〜{est_max}分。'
                f'可能であれば Excel / CSV 形式で提供いただくと数秒〜数十秒で完了します'
            )

    if wage_files:
        total_size = sum(s for _, s in wage_files)
        if total_size > WAGE_TOTAL_LIMIT:
            warnings.append(
                f'📦 賃金台帳合計サイズが大きいです（{total_size/1024/1024:.1f}MB, '
                f'{len(wage_files)}ファイル） — AI 抽出が長時間化する可能性'
            )
        if len(wage_files) > WAGE_COUNT_LIMIT:
            warnings.append(
                f'📂 賃金台帳ファイルが多数あります（{len(wage_files)}件） — '
                f'従業員数が多い場合は AI 出力が途中で切れる可能性があります（max_tokens 16384）'
            )

    return warnings


def _estimate_case_scale(file_size_pairs: list[tuple[str, int]]) -> dict | None:
    """案件規模・処理時間・APIコストを推定する。

    ユーザー（坂平さん）が実行前に「これは何分かかりそう」「いくらぐらい」が
    把握できるよう、控えめ（多め見積もり）の数値を返す。

    Returns: {
        'scale': '小型' | '中型' | '大型' | '超大型',
        'time_label': '約1〜3分',
        'cost_label': '約30〜80円',
        'pre_split': bool,  # 事前分割発動の見込み
        'pdf_total_mb': float,
        'wage_pdf_mb': float,
        'note': str,  # ユーザー向け補足（推奨ファイル形式等）
    }
    None: 推定不能（ファイルなし）
    """
    if not file_size_pairs:
        return None

    PRE_SPLIT_BYTES = 4 * 1024 * 1024  # 賃金台帳PDF 4MB超で事前分割発動

    # 賃金台帳PDFのサイズ
    wage_pdf_total = 0
    pdf_total = 0
    for name, size in file_size_pairs:
        if not name or size is None or size <= 0:
            continue
        n = unicodedata.normalize('NFC', name)
        ext = Path(n).suffix.lower()
        if ext != '.pdf':
            continue
        pdf_total += size
        if '賃金台帳' in n:
            wage_pdf_total += size

    wage_pdf_mb = wage_pdf_total / 1024 / 1024
    pdf_total_mb = pdf_total / 1024 / 1024

    # サイズクラス判定（賃金台帳PDFのサイズが支配的）
    if wage_pdf_mb >= 10:
        scale = '超大型'
    elif wage_pdf_mb >= 5:
        scale = '大型'
    elif wage_pdf_mb >= 2:
        scale = '中型'
    else:
        scale = '小型'

    # 事前分割発動見込み（4MB超で発動）
    pre_split = wage_pdf_total > PRE_SPLIT_BYTES

    # ── 処理時間予想 ──
    # 控えめ（多め見積もり）の値を出す。
    # 内訳の経験則:
    #   履歴事項抽出 ~30秒、PL ~30秒、納税証明 ~20秒、見積 ~20秒、AI判断 ~30秒
    #   = 共通 ~2分 (上振れで 3分)
    #   賃金台帳: PDF 1MB あたり 12〜25秒（経験値）
    #   事前分割発動時: API呼出が +1回 → +1〜2分
    common_min = 60   # 共通処理 1分（楽観）
    common_max = 180  # 共通処理 3分（悲観）
    wage_min = wage_pdf_mb * 12
    wage_max = wage_pdf_mb * 25
    if pre_split:
        wage_max += 90  # 事前分割発動時は最大1.5分追加
        wage_min += 30
    total_min_sec = common_min + wage_min
    total_max_sec = common_max + wage_max

    if total_max_sec < 120:
        time_label = f'約 30秒〜{round(total_max_sec/60)+1}分'
    elif total_max_sec < 600:
        time_label = f'約 {max(1, round(total_min_sec/60))}〜{round(total_max_sec/60)}分'
    else:
        time_label = f'約 {round(total_min_sec/60)}〜{round(total_max_sec/60)}分（PDFが大きいため長めです）'

    # ── APIコスト予想（円、Sonnet 4.6 / 1USD=150円） ──
    # 実態に合わせ上方寄りに見積もる（過去案件で表示より高くなる傾向があったため）
    # 共通処理（履歴/PL/税/見積/AI判断）: ~30円
    # 賃金台帳: PDF 1MB あたり ~8〜15円
    # 事前分割発動時: 賃金台帳側が +50〜80%（API呼出 +1回）
    common_cost = 30
    wage_cost_min = wage_pdf_mb * 8
    wage_cost_max = wage_pdf_mb * 15
    if pre_split:
        wage_cost_max *= 1.8  # 分割発動でAPIコール+1回
        wage_cost_min *= 1.5
    total_cost_min = common_cost + wage_cost_min
    total_cost_max = common_cost + wage_cost_max

    # さらに上方寄りに（最大値を1.5倍に切り上げ）
    cost_label = f'約 {max(30, int(total_cost_min))}〜{int(total_cost_max * 1.5)}円'

    # 補足メッセージ
    notes: list[str] = []
    if pre_split:
        notes.append('賃金台帳PDFが大きいため、自動で分割処理します（処理時間+1〜2分）')
    if scale in ('大型', '超大型'):
        notes.append('Excel/CSV形式の賃金台帳があれば数十秒で完了します')
    if scale == '超大型':
        notes.append('処理時間が長くなる可能性があるため、画面を閉じずにお待ちください')

    return {
        'scale': scale,
        'time_label': time_label,
        'cost_label': cost_label,
        'pre_split': pre_split,
        'pdf_total_mb': pdf_total_mb,
        'wage_pdf_mb': wage_pdf_mb,
        'note': ' / '.join(notes) if notes else '',
    }


def _render_case_scale_estimate(file_size_pairs: list[tuple[str, int]]):
    """案件規模・処理時間・APIコスト予想を UI に表示。"""
    est = _estimate_case_scale(file_size_pairs)
    if not est:
        return

    scale_emoji = {'小型': '🟢', '中型': '🟡', '大型': '🟠', '超大型': '🔴'}
    emoji = scale_emoji.get(est['scale'], '📊')

    with st.expander(
        f'{emoji} 案件規模の予想: **{est["scale"]}** — '
        f'処理時間 {est["time_label"]} / APIコスト {est["cost_label"]}',
        expanded=(est['scale'] in ('大型', '超大型')),
    ):
        st.markdown(
            f'- **規模**: {emoji} {est["scale"]}（賃金台帳PDF {est["wage_pdf_mb"]:.1f}MB '
            f'/ 全PDF {est["pdf_total_mb"]:.1f}MB）\n'
            f'- **処理時間目安**: {est["time_label"]}\n'
            f'- **APIコスト目安**: {est["cost_label"]}\n'
            f'- **事前PDF分割**: {"✅ 発動見込み（精度向上のため）" if est["pre_split"] else "発動なし（小型のため不要）"}'
        )
        if est['note']:
            for n in est['note'].split(' / '):
                st.info(n)
        st.caption(
            'コスト・時間は上方寄りに見積もった目安です。'
            '実際はこれより安く済むこともありますが、'
            'PDFが多い・賃金台帳が大きい案件では上振れすることもあります。'
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
                        '[変換手順（CC向け）](https://github.com/hane-colorfulbox/hojokin/blob/main/docs/%E8%B3%83%E9%87%91%E5%8F%B0%E5%B8%B3%E5%A4%89%E6%8F%9B%E6%89%8B%E9%A0%86_CC%E5%90%91%E3%81%91.md) ・ '
                        '[空テンプレート](https://github.com/hane-colorfulbox/hojokin/blob/main/%E3%83%84%E3%83%BC%E3%83%AB/%E8%B3%83%E9%87%91%E5%8F%B0%E5%B8%B3%E3%83%86%E3%83%B3%E3%83%97%E3%83%AC%E3%83%BC%E3%83%88.xlsx)'
                    )
            else:
                st.markdown(f'➖ {display} — なし（任意）')

        if result['unmatched']:
            st.markdown('---')
            st.markdown('**判別できなかったファイル:**')
            for name in result['unmatched']:
                st.markdown(f'&ensp; ⚠️ `{name}`（キーワードなし → 処理対象外）')


# 複数選択（multiselect）させるカテゴリ。それ以外は単一選択（selectbox）。
_MULTI_SELECT_CATS = {'wage_ledger'}


def _categorize_for_ui(file_names: list[str]) -> dict[str, list[str]]:
    """ファイル名を _FILE_CATEGORIES のキーワード + _UI_ALLOWED_EXTS で分類。

    `_analyze_files` の分類ロジックと整合させる（差し替え UI と判別結果表示で
    候補ファイルが食い違うと混乱するため）。
    """
    candidates: dict[str, list[str]] = {cat: [] for cat, _, _ in _FILE_CATEGORIES}
    for name in file_names:
        name_nfc = unicodedata.normalize('NFC', name)
        ext = Path(name_nfc).suffix.lower()
        for cat, _, keywords in _FILE_CATEGORIES:
            if any(kw in name_nfc for kw in keywords):
                allowed = _UI_ALLOWED_EXTS.get(cat)
                if allowed is not None and ext not in allowed:
                    break  # 拡張子NG → 他カテゴリは試さない（_analyze_files と同じ挙動）
                candidates[cat].append(name)
                break
    return candidates


_OVERRIDE_AUTO_LABEL = '（自動検出に従う）'
_OVERRIDE_EXCLUDE_LABEL = '─ 使わない（対象外）─'


def _render_file_selection_override(
    file_names: list[str], case_key: str,
) -> dict[str, list[str] | None]:
    """検出カテゴリ別の候補ファイルを差し替え可能な UI で表示し、選択結果を返す。

    Returns:
        dict[category, list[str] | None]
            - None: 自動検出に従う（override なし）
            - list[str]: ユーザー指定（[] は「対象外」）
    """
    candidates = _categorize_for_ui(file_names)
    visible = [(cat, display) for cat, display, _ in _FILE_CATEGORIES if candidates[cat]]
    if not visible:
        return {}

    overrides: dict[str, list[str] | None] = {}

    with st.expander('▶ 使用するファイルを差し替える（必要な場合のみ）', expanded=False):
        st.caption(
            '通常は **自動検出のまま** で OK。'
            '同じカテゴリに該当するファイルが複数あって自動選定が誤っているとき、'
            'または特定のファイル（空テンプレ・別期分など）を除外したいときだけ変更してください。'
        )
        for cat, display in visible:
            cat_files = candidates[cat]
            sel_key = f'override_{case_key}_{cat}'

            if cat in _MULTI_SELECT_CATS:
                # 複数選択（賃金台帳など）: チェックを外せば除外。既定は全選択。
                selected = st.multiselect(
                    f'{display}（複数選択可）',
                    options=cat_files,
                    default=cat_files,
                    key=sel_key,
                    help='チェックを外したファイルは処理から除外されます。',
                )
                if list(selected) == cat_files:
                    overrides[cat] = None  # 全選択=デフォルト → 自動扱い
                else:
                    overrides[cat] = list(selected)
            else:
                options = [_OVERRIDE_AUTO_LABEL] + cat_files + [_OVERRIDE_EXCLUDE_LABEL]
                selected = st.selectbox(
                    display,
                    options=options,
                    key=sel_key,
                    help='候補が複数ある場合の差し替え・「使わない」指定ができます。',
                )
                if selected == _OVERRIDE_AUTO_LABEL:
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


def _check_required_by_names(file_names, task):
    """タスクに応じた必須ファイルが揃っているかチェック

    拡張子フィルタ（_UI_ALLOWED_EXTS）も適用するため、賃金台帳PDFだけ
    アップした状態では can_run を有効にしない（実行後の skipped 失敗を防ぐ）。
    """
    if not file_names:
        return False
    # NFD（macOS の濁点分離形式）でも比較が通るよう NFC 化してから判定
    names_nfc = [unicodedata.normalize('NFC', n) for n in file_names]
    required_cats = _REQUIRED_CATS_BY_TASK.get(task, set())
    for cat, _, keywords in _FILE_CATEGORIES:
        if cat not in required_cats:
            continue
        allowed = _UI_ALLOWED_EXTS.get(cat)

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
            _render_case_scale_estimate(size_pairs)
            for w in _check_size_warnings(size_pairs):
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
        ('wage_ledger', '賃金台帳',               'Excel/CSV',     ['賃金台帳'],
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
- キーワードが含まれないファイルは**無視されます**（エラーにはなりません）
- 決算書が2期分ある場合、**サイズの大きい方**が自動選択されます
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
        # case_key にはファイル名集合のハッシュを使用 → 同じ顧客の同じ資料セットなら
        # 選択状態を維持、別案件をドロップすれば自動でリセット
        _upload_case_key = (
            f'upload_{hash(tuple(sorted(f.name for f in uploaded_files)))}'
        )
        selection_override_names = _render_file_selection_override(
            [f.name for f in uploaded_files],
            case_key=_upload_case_key,
        )
        # 案件規模・処理時間・APIコスト予想
        size_pairs_upload = [(f.name, f.size) for f in uploaded_files]
        _render_case_scale_estimate(size_pairs_upload)
        # 容量・件数の事前警告（処理は続行可能、ユーザーに確認を促すだけ）
        size_warnings = _check_size_warnings(size_pairs_upload)
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
has_required = (
    _check_required_by_names([f.name for f in uploaded_files], task_type)
    if has_files else False
)
has_drive_required = (
    _check_required_by_names([f['name'] for f in drive_files_to_download], task_type)
    if has_drive_files else False
)

# 必須ファイル不足時はボタンを押せないようにする（不完全ファイルでのAPI消費防止）
if data_source == 'Google Drive':
    has_data = has_drive_files
    required_ok = has_drive_required
else:
    has_data = has_files
    required_ok = has_required

can_run = bool(company_name) and has_data and required_ok
if task_type == 'bonus':
    can_run = can_run and bool(prefecture)
# 決算月の指定を必須化（2026-05 方針）
# 賃金台帳の対象12ヶ月を確定するためにユーザー指定が必要
can_run = can_run and (fiscal_month_override is not None)

if not company_name:
    st.warning('⬅️ サイドバーで会社名を入力してください')
elif fiscal_month_override is None:
    st.warning('⬅️ サイドバーで決算月を選択してください（賃金台帳の対象期間を確定するため必須）')
elif task_type == 'bonus' and not prefecture:
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
        st.info(f'**{company_name}** の賃金台帳を分析して加点判定を行います（{source_label}）— 準備OKです')
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
        spinner_msg = '賃金台帳を分析中...' if task_type == 'bonus' else 'AIが資料を読み取り中...（1〜3分かかります）'
        with st.spinner(spinner_msg):
            results = run_processing(
                company_name=company_name,
                template_type=template_type,
                task_type=task_type,
                work_dir=work_dir,
                template_dir=template_dir,
                prefecture=prefecture,
                fiscal_month_override=fiscal_month_override,
                has_cost_report_hint=has_cost_report_hint,
                selection_override=selection_override,
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
                    'prefecture': br.prefecture,
                    'min_wage_r6': br.min_wage_r6,
                    'min_wage_r7': br.min_wage_r7,
                    'employee_count': len(br.employees),
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
            'bonus': '📊 加点判定',
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

                st.markdown(f"**事業場所在地:** {br['prefecture']}（R6最低賃金: {br['min_wage_r6']}円 → R7: {br['min_wage_r7']}円）")

                # 加点措置①
                col_b1, col_b2 = st.columns(2)
                with col_b1:
                    if br['bonus1_eligible']:
                        st.success(f"**加点措置①: 対象** ({len(br['bonus1_months_met'])}か月が条件達成)")
                    else:
                        st.warning(f"**加点措置①: 対象外** ({len(br['bonus1_months_met'])}か月/3か月必要)")

                    with st.expander('月別詳細'):
                        for d in br['bonus1_details']:
                            if d['total'] > 0:
                                mark = '○' if d['meets_30pct'] else '×'
                                st.text(f"{d['month']}: {d['under_r7']}/{d['total']}名 = {d['ratio']*100:.1f}% {mark}")

                # 加点措置②
                with col_b2:
                    if br['bonus2_eligible']:
                        st.success(f"**加点措置②: 対象** (差額 {br['bonus2_diff']:.0f}円 >= 63円)")
                    else:
                        st.warning(f"**加点措置②: 対象外** (差額 {br['bonus2_diff']:.0f}円 < 63円)")
                    st.text(f"7月最低時給: {br['bonus2_min_wage_july']:.0f}円")
                    st.text(f"直近月最低時給: {br['bonus2_min_wage_latest']:.0f}円")

                # 加点シートダウンロード
                for key, file_info in result.get('bonus_files', {}).items():
                    label_map = {'bonus1': '加点措置①シート', 'bonus2': '加点措置②シート'}
                    st.download_button(
                        label=f"⬇️ {label_map.get(key, key)} をダウンロード",
                        data=file_info['data'],
                        file_name=file_info['name'],
                        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                        use_container_width=True,
                        key=f'download_{key}',
                    )

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
st.caption(f'補助金書類自動作成ツール v0.1.4 | カラフルボックス株式会社')
