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
        help='申請書作成：ヒアリングシート+各種PDFから申請書を自動作成。給与計算：損益計算書+賃金データから給与支給総額を計算。',
    )
    task_type = TASK_OPTIONS[task_label]

    st.divider()

    st.markdown('**処理の目安**')
    st.caption('所要時間: 約1〜3分')
    st.caption('API利用料: 約7〜10円/社')

# ── ファイルアップロード ──
st.markdown(
    '<span class="step-number">1</span>'
    '<span class="step-title">資料をアップロード</span>',
    unsafe_allow_html=True,
)
st.caption('ファイルはファイル名のキーワードで自動判別されます。該当キーワードがないファイルは無視されます。')

col1, col2 = st.columns(2)

with col1:
    st.markdown("""
<div class="file-card file-required">
<span class="badge-required">必須</span><br>
<strong>ヒアリングシート</strong>（Excel）<br>
<span class="keyword-tag">ヒアリング</span> がファイル名に含まれること<br>
<small>例: ヒアリングシート_○○株式会社.xlsx</small>
</div>

<div class="file-card file-required">
<span class="badge-required">必須</span><br>
<strong>履歴事項全部証明書</strong>（PDF）<br>
<span class="keyword-tag">履歴事項</span> がファイル名に含まれること<br>
<small>例: 履歴事項全部証明書_○○様.pdf</small>
</div>

<div class="file-card file-required">
<span class="badge-required">必須</span><br>
<strong>損益計算書 / 決算報告書</strong>（PDF）<br>
<span class="keyword-tag">損益計算書</span> <span class="keyword-tag">決算報告書</span> <span class="keyword-tag">決算書</span> のいずれか<br>
<small>例: 42期 決算報告書.pdf</small>
</div>
""", unsafe_allow_html=True)

with col2:
    st.markdown("""
<div class="file-card file-optional">
<span class="badge-optional">あれば</span><br>
<strong>納税証明書</strong>（PDF）<br>
<span class="keyword-tag">納税証明</span> がファイル名に含まれること<br>
<small>例: 納税証明書(その1)_○○様.pdf</small>
</div>

<div class="file-card file-optional">
<span class="badge-optional">あれば</span><br>
<strong>見積書</strong>（Excel/PDF）<br>
<span class="keyword-tag">見積</span> がファイル名に含まれること<br>
<small>例: お見積書_○○.pdf</small>
</div>

<div class="file-card file-optional">
<span class="badge-optional">あれば</span><br>
<strong>賃金状況報告シート</strong>（Excel）<br>
<span class="keyword-tag">賃金状況報告</span> がファイル名に含まれること<br>
<small>例: 賃金状況報告シート.xlsx</small>
</div>
""", unsafe_allow_html=True)

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

# アップロード済みファイルのチェックリスト
if uploaded_files:
    # ファイル判別ルール: (カテゴリ, 表示名, キーワードリスト, 必須フラグ)
    FILE_CHECKS = [
        ('hearing',     'ヒアリングシート',     ['ヒアリング'],                          True),
        ('registry',    '履歴事項全部証明書',   ['履歴事項'],                            True),
        ('pl',          '損益計算書 / 決算報告書', ['損益計算書', '決算報告書', '決算書'], True),
        ('tax',         '納税証明書',           ['納税証明'],                            False),
        ('estimate',    '見積書',               ['見積'],                                False),
        ('wage_report', '賃金状況報告シート',   ['賃金状況報告'],                        False),
    ]

    # 各カテゴリにマッチしたファイルを集計
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

    # 必須チェック
    missing_required = [
        display for cat, display, _, required in FILE_CHECKS
        if required and not detected[cat]
    ]
    all_required_ok = len(missing_required) == 0

    # チェック結果を表示
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

# 必須ファイルチェック（アップロード済みの場合のみ）
_required_keywords = {
    'hearing': ['ヒアリング'],
    'registry': ['履歴事項'],
    'pl': ['損益計算書', '決算報告書', '決算書'],
}

def _check_required(files):
    """必須ファイルが全て揃っているかチェック"""
    if not files:
        return False
    names = [f.name for f in files]
    for cat, keywords in _required_keywords.items():
        if not any(any(kw in name for kw in keywords) for name in names):
            return False
    return True

has_files = bool(uploaded_files)
has_required = _check_required(uploaded_files) if has_files else False
can_run = bool(company_name) and has_files

if not company_name:
    st.warning('⬅️ サイドバーで会社名を入力してください')
elif not has_files:
    st.warning('⬆️ 資料ファイルをアップロードしてください')
elif not has_required:
    st.warning('⬆️ 必須ファイルが不足しています。ファイル判別結果を確認してください')
else:
    st.info(f'**{company_name}** の書類を **{template_label}** で作成します — 準備OKです')

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
        with st.expander('検出されたファイル（デバッグ用）'):
            st.code(detector.summary())

        # 処理実行
        progress_bar = st.progress(0, text='処理を開始します...')
        status_text = st.empty()

        def update_progress(msg):
            status_text.info(msg)

        progress_bar.progress(10, text='AIが資料を読み取り中...')

        results = run_processing(
            company_name=company_name,
            template_type=template_type,
            task_type=task_type,
            work_dir=work_dir,
            template_dir=template_dir,
            progress_callback=update_progress,
        )

        progress_bar.progress(100, text='処理完了!')

        # ── 結果表示 ──
        st.markdown(
            '<span class="step-number">3</span>'
            '<span class="step-title">結果・ダウンロード</span>',
            unsafe_allow_html=True,
        )

        for task_name, result in results.items():
            task_display = '📝 申請書作成' if task_name == 'application' else '💰 給与支給総額計算'

            if result['status'] == '完了':
                st.success(f'{task_display}: 完了')

                if result.get('output_path') and result['output_path'].exists():
                    with open(result['output_path'], 'rb') as f:
                        st.download_button(
                            label=f'⬇️ {result["output_path"].name} をダウンロード',
                            data=f.read(),
                            file_name=result['output_path'].name,
                            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                            use_container_width=True,
                        )

                # 空セル表示
                if result.get('empty_cells'):
                    with st.expander(f'⚠️ 未入力セル（{len(result["empty_cells"])}件） — 手動確認が必要'):
                        for cell in result['empty_cells']:
                            st.text(cell)
            else:
                st.error(f'{task_display}: {result["message"]}')

# ── フッター ──
st.markdown('---')
st.caption(f'補助金書類自動作成ツール v0.1.0 | カラフルボックス株式会社')
