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
from hojokin.config import CLAUDE_API_KEY
from hojokin.pipeline import (
    FileDetector, run_application_transfer, run_wage_calculation,
)

# ── 定数 ──
TEMPLATE_OPTIONS = {
    '通常枠 2026（法人）': '通常枠_2026',
    'インボイス枠 2026（法人）': 'インボイス枠_2026',
}

TASK_OPTIONS = {
    '申請書作成のみ': 'application',
    '給与計算のみ': 'wage',
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
        font-size: 2rem;
        font-weight: bold;
        margin-bottom: 0.5rem;
    }
    .sub-header {
        color: #666;
        margin-bottom: 2rem;
    }
    .status-box {
        padding: 1rem;
        border-radius: 0.5rem;
        margin: 0.5rem 0;
    }
    .stFileUploader > div > div {
        padding: 2rem;
    }
</style>
""", unsafe_allow_html=True)


def find_template(base_dir: Path, template_type: str) -> Path | None:
    """テンプレートファイルを検索"""
    keywords = {
        '通常枠_2026': ['原本', '通常枠', '2026'],
        'インボイス枠_2026': ['原本', 'インボイス', '2026'],
    }
    kws = keywords.get(template_type, [])
    for p in base_dir.iterdir():
        if p.suffix == '.xlsx' and all(kw in p.name for kw in kws) and not p.name.startswith('~$'):
            return p
    return None


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
):
    """メイン処理を実行"""
    results = {}

    # Extractor作成
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

    return results


# ── メイン画面 ──
st.markdown('<div class="main-header">補助金書類自動作成ツール</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">IT導入補助金の申請書類をAIで自動作成します</div>', unsafe_allow_html=True)

# API接続状態
if CLAUDE_API_KEY:
    st.success('Claude API: 接続済み')
else:
    st.error('Claude API: 未接続（.envファイルにCLAUDE_API_KEYを設定してください）')
    st.stop()

# ── サイドバー ──
with st.sidebar:
    st.header('設定')

    company_name = st.text_input(
        '会社名（必須）',
        placeholder='例: 京のお肉処弘',
    )

    template_label = st.selectbox(
        'テンプレート',
        list(TEMPLATE_OPTIONS.keys()),
    )
    template_type = TEMPLATE_OPTIONS[template_label]

    task_label = st.selectbox(
        '実行タスク',
        list(TASK_OPTIONS.keys()),
    )
    task_type = TASK_OPTIONS[task_label]

    st.divider()
    st.caption('API利用料: 約7〜10円/社')

# ── ファイルアップロード ──
st.header('1. 資料をアップロード')

col1, col2 = st.columns(2)

with col1:
    st.subheader('必須ファイル')
    st.markdown("""
    - **ヒアリングシート**（Excel）
    - **履歴事項全部証明書**（PDF）
    - **損益計算書**（PDF）
    """)

with col2:
    st.subheader('あれば追加')
    st.markdown("""
    - 納税証明書（PDF）
    - 見積書（Excel/PDF）
    - 賃金状況報告シート（Excel）
    """)

uploaded_files = st.file_uploader(
    'ファイルをドラッグ&ドロップ（複数可）',
    accept_multiple_files=True,
    type=['pdf', 'xlsx', 'xls'],
    key='file_uploader',
)

# テンプレート原本
st.markdown('---')
template_file = st.file_uploader(
    'テンプレート原本（初回のみ必要。2回目以降はサーバーに保存されます）',
    accept_multiple_files=False,
    type=['xlsx'],
    key='template_uploader',
)

# ── 処理実行 ──
st.header('2. 処理実行')

can_run = bool(company_name) and bool(uploaded_files)

if not company_name:
    st.warning('サイドバーで会社名を入力してください')
elif not uploaded_files:
    st.warning('資料ファイルをアップロードしてください')

if st.button('処理開始', type='primary', disabled=not can_run, use_container_width=True):
    # 一時ディレクトリに保存
    with tempfile.TemporaryDirectory() as tmpdir:
        work_dir = Path(tmpdir)

        # ファイル保存
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
        with st.expander('検出されたファイル', expanded=True):
            st.code(detector.summary())

        # 処理実行
        progress_bar = st.progress(0, text='処理を開始します...')
        status_text = st.empty()

        def update_progress(msg):
            status_text.info(msg)

        progress_bar.progress(10, text='APIに送信中...')

        results = run_processing(
            company_name=company_name,
            template_type=template_type,
            task_type=task_type,
            work_dir=work_dir,
            template_dir=template_dir,
            progress_callback=update_progress,
        )

        progress_bar.progress(100, text='処理完了')

        # ── 結果表示 ──
        st.header('3. 結果')

        for task_name, result in results.items():
            task_display = '申請書作成' if task_name == 'application' else '給与支給総額計算'

            if result['status'] == '完了':
                st.success(f'{task_display}: 完了')

                if result.get('output_path') and result['output_path'].exists():
                    with open(result['output_path'], 'rb') as f:
                        st.download_button(
                            label=f'{result["output_path"].name} をダウンロード',
                            data=f.read(),
                            file_name=result['output_path'].name,
                            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                            use_container_width=True,
                        )

                # 空セル表示
                if result.get('empty_cells'):
                    with st.expander(f'未入力セル（{len(result["empty_cells"])}件）'):
                        for cell in result['empty_cells']:
                            st.text(cell)
            else:
                st.error(f'{task_display}: {result["message"]}')

# ── フッター ──
st.markdown('---')
st.caption(f'補助金書類自動作成ツール v0.1.0 | カラフルボックス株式会社')
