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
from pathlib import Path
from datetime import datetime

import streamlit as st

# .env読み込み
from dotenv import load_dotenv
load_dotenv()

# パッケージパス追加
sys.path.insert(0, str(Path(__file__).parent))

from hojokin.ai_extractor import create_extractor
from hojokin.config import CLAUDE_API_KEY, detect_prefecture
from hojokin.pipeline import (
    FileDetector, run_application_transfer, run_wage_calculation,
)
from hojokin.wage_reader import (
    read_wage_ledger, judge_bonus_points,
    fill_bonus_sheet_1, fill_bonus_sheet_2,
)

# Drive連携（認証情報がある場合のみ）
_drive_client = None
_DRIVE_CREDS = os.getenv('GOOGLE_SERVICE_ACCOUNT_JSON', '')
_DRIVE_PARENT_ID = os.getenv('DRIVE_PARENT_FOLDER_ID', '')


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
    try:
        if 'gcp_service_account' in st.secrets:
            _drive_client = DriveClient(
                credentials_dict=dict(st.secrets['gcp_service_account']),
            )
            return _drive_client
    except Exception:
        pass

    return None

# ── 定数 ──
TEMPLATE_OPTIONS = {
    '通常枠 2026（法人）': '通常枠_2026',
    'インボイス枠 2026（法人）': 'インボイス枠_2026',
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
    """テンプレートファイルを検索（v2を優先）"""
    import unicodedata
    keywords = {
        '通常枠_2026': ['原本', '通常枠', '2026'],
        'インボイス枠_2026': ['原本', 'インボイス', '2026'],
    }
    kws = keywords.get(template_type, [])
    candidates = []
    for p in base_dir.iterdir():
        name = unicodedata.normalize('NFC', p.name)
        if p.suffix == '.xlsx' and all(kw in name for kw in kws) and not name.startswith('~$'):
            candidates.append(p)
    if not candidates:
        # ツール/サブフォルダも探す
        tool_dir = base_dir / 'ツール'
        if tool_dir.exists():
            for p in tool_dir.iterdir():
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
    """アップロードファイルを一時ディレクトリに保存"""
    saved = []
    for f in uploaded_files:
        dest = target_dir / f.name
        dest.write_bytes(f.getvalue())
        saved.append(f.name)
    return saved


def run_processing(
    company_name: str,
    template_type: str,
    task_type: str,
    work_dir: Path,
    template_dir: Path,
    progress_callback=None,
    prefecture: str = '',
):
    """メイン処理を実行"""
    results = {}

    # Extractor作成（加点判定のみの場合はAPI不要）
    extractor = None
    if task_type in ('application', 'wage', 'all'):
        extractor = create_extractor(CLAUDE_API_KEY)

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
            )
            results['application'] = {
                'status': status.status,
                'message': status.message,
                'output_path': output_path if status.status == '完了' else None,
                'empty_cells': status.empty_cells,
            }

    if task_type in ('wage', 'all'):
        if progress_callback:
            progress_callback('給与支給総額を計算中...')

        output_path = work_dir / f'{company_name}_給与支給総額計算.xlsx'
        status = run_wage_calculation(
            resource_folder=work_dir,
            company_name=company_name,
            output_path=output_path,
            extractor=extractor,
        )
        results['wage'] = {
            'status': status.status,
            'message': status.message,
            'output_path': output_path if status.status == '完了' else None,
        }

    if task_type == 'bonus':
        if progress_callback:
            progress_callback('賃金台帳を読み取り中...')

        results['bonus'] = _run_bonus_judgment(work_dir, company_name, prefecture, template_dir)

    return results


def _run_bonus_judgment(work_dir: Path, company_name: str, prefecture: str, template_dir: Path) -> dict:
    """加点判定を実行"""
    # 賃金台帳ファイルを探す
    wage_file = None
    for f in work_dir.iterdir():
        if f.suffix in ('.xlsx', '.xls') and not f.name.startswith('~$'):
            if '賃金' in f.name or '給与' in f.name or '明細' in f.name or 'wage' in f.name.lower():
                wage_file = f
                break
    if wage_file is None:
        # 賃金台帳のキーワードがなくてもExcelならとりあえず読む
        for f in work_dir.iterdir():
            if f.suffix in ('.xlsx', '.xls') and not f.name.startswith('~$'):
                wage_file = f
                break

    if wage_file is None:
        return {
            'status': 'エラー',
            'message': '賃金台帳ファイルが見つかりません。Excelファイルをアップロードしてください。',
        }

    try:
        employees = read_wage_ledger(wage_file)
        if not employees:
            return {
                'status': 'エラー',
                'message': '賃金台帳からデータを読み取れませんでした。シートの形式を確認してください。',
            }

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

    company_name = st.text_input(
        '会社名（必須）',
        placeholder='例: 京のお肉処弘',
        help='正式名称でなくてもOK。出力ファイル名に使われます。',
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
    _has_cloud_creds = 'gcp_service_account' in st.secrets if hasattr(st, 'secrets') else False
    drive_available = (_has_local_creds or _has_cloud_creds) and bool(_DRIVE_PARENT_ID)
    # Secrets経由の場合もPARENT_IDを取得
    if not _DRIVE_PARENT_ID and _has_cloud_creds:
        try:
            drive_available = bool(st.secrets.get('drive_parent_folder_id', ''))
        except Exception:
            pass
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
    st.caption('所要時間: 約1〜3分')
    st.caption('API利用料: 約7〜10円/社')

# ── ファイル入力 ──
st.markdown(
    '<span class="step-number">1</span>'
    '<span class="step-title">資料を準備</span>',
    unsafe_allow_html=True,
)

# Drive連携用の変数
drive_folder_id = None
drive_files_to_download = []

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
            except Exception:
                pass

        @st.cache_data(ttl=60)
        def _list_customer_folders():
            c = _get_drive_client()
            return c.list_folders(parent_id) if c else []

        folders = _list_customer_folders()
        folder_names = ['（選択してください）'] + [f['name'] for f in folders]

        selected_folder_name = st.selectbox(
            '顧客フォルダを選択',
            folder_names,
            help='Driveの2026フォルダ直下の顧客フォルダ一覧です。',
        )

        if selected_folder_name != '（選択してください）':
            selected_folder = next(f for f in folders if f['name'] == selected_folder_name)
            drive_folder_id = selected_folder['id']

            # フォルダ内のファイル一覧（サブフォルダ含む）
            @st.cache_data(ttl=30)
            def _list_folder_files(folder_id):
                c = _get_drive_client()
                return c.list_files_recursive(folder_id) if c else []

            all_files = _list_folder_files(drive_folder_id)

            if all_files:
                st.success(f'{len(all_files)}件のファイルが見つかりました')
                with st.expander('ファイル一覧', expanded=True):
                    for f in all_files:
                        loc = f.get('folder_name', 'ルート')
                        st.text(f'  [{loc}] {f["name"]}')
                drive_files_to_download = all_files
            else:
                st.warning('このフォルダにはファイルがありません。')

    uploaded_files = None
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
        ('wage_ledger', '賃金台帳',               'Excel',     ['賃金台帳'],
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
        type=['pdf', 'xlsx', 'xls'],
        key='file_uploader',
    )

    # アップロード済みファイルのチェックリスト（タスク別に必須/任意を切り替え）
    if uploaded_files:
        # タスク別の必須カテゴリ
        _REQUIRED_BY_TASK = {
            'application': {'hearing', 'registry', 'pl'},
            'wage':        {'wage_ledger'},
            'bonus':       {'wage_ledger'},
            'all':         {'hearing', 'registry', 'pl'},
        }
        _required_cats = _REQUIRED_BY_TASK.get(task_type, set())

        FILE_CHECKS = [
            ('hearing',     'ヒアリングシート',     ['ヒアリング'],                          'hearing' in _required_cats),
            ('registry',    '履歴事項全部証明書',   ['履歴事項'],                            'registry' in _required_cats),
            ('pl',          '損益計算書 / 決算報告書', ['損益計算書', '決算報告書', '決算書'], 'pl' in _required_cats),
            ('cost_report', '製造原価報告書',       ['製造原価報告書', '原価報告書'],         False),
            ('tax',         '納税証明書',           ['納税証明'],                            False),
            ('estimate',    '見積書',               ['見積'],                                False),
            ('wage_report', '賃金状況報告シート',   ['賃金状況報告'],                        'wage_report' in _required_cats),
            ('wage_ledger', '賃金台帳',             ['賃金台帳'],                            'wage_ledger' in _required_cats),
        ]

        detected = {cat: [] for cat, *_ in FILE_CHECKS}
        unmatched = []
        for f in uploaded_files:
            matched = False
            for cat, _, keywords, _ in FILE_CHECKS:
                if any(kw in f.name for kw in keywords):
                    detected[cat].append(f.name)
                    matched = True
                    break
            if not matched:
                unmatched.append(f.name)

        missing_required = [
            display for cat, display, _, required in FILE_CHECKS
            if required and not detected[cat]
        ]
        all_required_ok = len(missing_required) == 0

        if all_required_ok:
            st.success(f'ファイルチェック OK — 必須ファイルがすべて揃っています（{len(uploaded_files)}件アップロード済み）')
        else:
            st.error(f'必須ファイルが不足しています: **{"、".join(missing_required)}**')

        with st.expander('ファイル判別結果（詳細）', expanded=not all_required_ok):
            for cat, display, _, required in FILE_CHECKS:
                files = detected[cat]
                if files:
                    st.markdown(f'✅ **{display}** → `{"`, `".join(files)}`')
                elif required:
                    st.markdown(f'❌ **{display}** — **未検出（必須）** ファイル名にキーワードが含まれているか確認してください')
                else:
                    st.markdown(f'➖ {display} — なし（任意）')

            if unmatched:
                st.markdown('---')
                st.markdown('**判別できなかったファイル:**')
                for name in unmatched:
                    st.markdown(f'&ensp; ⚠️ `{name}`（キーワードなし → 処理対象外）')

    # テンプレート原本
    st.markdown('---')
    template_file = st.file_uploader(
        'テンプレート原本（初回のみ必要。サーバーに保存されるので2回目以降は不要です）',
        accept_multiple_files=False,
        type=['xlsx'],
        key='template_uploader',
        help='「【原本_法人】企業名_○○枠_法人2026.xlsx」のファイルをアップロードしてください。',
    )

# ── 処理実行 ──
st.markdown(
    '<span class="step-number">2</span>'
    '<span class="step-title">処理実行</span>',
    unsafe_allow_html=True,
)

# 必須ファイルチェック（タスク別）
_REQUIRED_KEYWORDS_BY_TASK = {
    'application': {
        'hearing': ['ヒアリング'],
        'registry': ['履歴事項'],
        'pl': ['損益計算書', '決算報告書', '決算書'],
    },
    'wage': {
        'wage_ledger': ['賃金台帳'],
    },
    'bonus': {
        'wage_ledger': ['賃金台帳'],
    },
    'all': {
        'hearing': ['ヒアリング'],
        'registry': ['履歴事項'],
        'pl': ['損益計算書', '決算報告書', '決算書'],
    },
}

def _check_required(files, task):
    """タスクに応じた必須ファイルが全て揃っているかチェック"""
    if not files:
        return False
    required = _REQUIRED_KEYWORDS_BY_TASK.get(task, {})
    names = [f.name for f in files]
    for cat, keywords in required.items():
        if not any(any(kw in name for kw in keywords) for name in names):
            return False
    return True

has_files = bool(uploaded_files)
has_drive_files = bool(drive_files_to_download)
has_required = _check_required(uploaded_files, task_type) if has_files else False

if data_source == 'Google Drive':
    if task_type == 'bonus':
        can_run = bool(company_name) and has_drive_files and bool(prefecture)
    else:
        can_run = bool(company_name) and has_drive_files
else:
    if task_type == 'bonus':
        can_run = bool(company_name) and has_files and bool(prefecture)
    else:
        can_run = bool(company_name) and has_files

if not company_name:
    st.warning('⬅️ サイドバーで会社名を入力してください')
elif task_type == 'bonus' and not prefecture:
    st.warning('⬅️ サイドバーで事業場の都道府県を選択してください')
elif data_source == 'Google Drive' and not has_drive_files:
    st.warning('⬅️ サイドバーで顧客フォルダを選択してください')
elif data_source != 'Google Drive' and not has_files:
    st.warning('⬆️ 資料ファイルをアップロードしてください')
elif data_source != 'Google Drive' and not has_required:
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

        # ファイル保存（データソースに応じて）
        if data_source == 'Google Drive' and drive_files_to_download:
            with st.spinner('Google Driveからファイルをダウンロード中...'):
                client = _get_drive_client()
                saved = []
                for f in drive_files_to_download:
                    dest = work_dir / f['name']
                    client.download_file(f['id'], dest)
                    saved.append(f['name'])
                st.caption(f'{len(saved)}件のファイルをダウンロードしました')
        else:
            saved = save_uploaded_files(uploaded_files, work_dir)

        # テンプレートディレクトリ
        template_dir = Path(__file__).parent
        if template_file:
            # アップロードされたテンプレートを保存
            template_dest = work_dir / template_file.name
            template_dest.write_bytes(template_file.getvalue())
            # メインディレクトリにもコピー（次回用）
            main_copy = Path(__file__).parent / template_file.name
            if not main_copy.exists():
                main_copy.write_bytes(template_file.getvalue())

        # ファイル検出プレビュー
        detector = FileDetector(work_dir)

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
            )

        # 結果をsession_stateに保存（画面再描画後も残る）
        session_results = {}
        for task_name, result in results.items():
            entry = {
                'status': result['status'],
                'message': result['message'],
                'empty_cells': result.get('empty_cells', []),
                'file_data': None,
                'file_name': None,
                'bonus_result': None,
                'bonus_files': {},
            }
            if result.get('output_path') and result['output_path'].exists():
                with open(result['output_path'], 'rb') as f:
                    entry['file_data'] = f.read()
                entry['file_name'] = result['output_path'].name

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
            st.success(f'{task_display}: 完了 — {result["message"]}')

            if result['file_data']:
                st.download_button(
                    label=f'⬇️ {result["file_name"]} をダウンロード',
                    data=result['file_data'],
                    file_name=result['file_name'],
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    use_container_width=True,
                    key=f'download_{task_name}',
                )

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
st.caption(f'補助金書類自動作成ツール v0.1.0 | カラフルボックス株式会社')
