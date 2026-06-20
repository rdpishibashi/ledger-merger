import tempfile
import zipfile
from datetime import datetime

import streamlit as st

from utils.ledger_finder import find_ledger_files, reconcile_missing_folders
from utils.ledger_merger import build_merged_workbook

st.set_page_config(page_title="図面親子管理台帳 統合ツール", page_icon="📑", layout="wide")

st.title("図面親子管理台帳 統合ツール")
st.write(
    "DXF-diff-manager の複数の出力フォルダにある台帳ファイル（Diff List シート）を"
    "1つの Excel ファイルに統合します。"
)

st.subheader("ZIPファイルをアップロード")
st.caption("DXF-diff-manager の出力フォルダ群を ZIP 化してアップロードしてください（複数可）。")
zip_files = st.file_uploader(
    "ZIPファイル", type=["zip"], accept_multiple_files=True
)

has_input = bool(zip_files)

run = st.button("統合実行", type="primary", disabled=not has_input)

if run:
    total_sources = len(zip_files)

    progress_placeholder = st.empty()
    progress_bar = progress_placeholder.progress(0.0, text="処理を開始しています...")

    all_entries = []
    all_missing_folders = []

    for idx, zip_file in enumerate(zip_files, start=1):
        with tempfile.TemporaryDirectory() as tmp_dir:
            zip_file.seek(0)
            try:
                with zipfile.ZipFile(zip_file) as zf:
                    zf.extractall(tmp_dir)
            except zipfile.BadZipFile:
                st.warning(f"ZIPファイルとして開けませんでした: {zip_file.name}")
                progress_bar.progress(idx / total_sources, text=f"{idx}/{total_sources}件を処理中...（{zip_file.name}）")
                continue
            entries, missing_folders = find_ledger_files(tmp_dir)
            all_entries.extend(entries)
            all_missing_folders.extend(missing_folders)

        progress_bar.progress(idx / total_sources, text=f"{idx}/{total_sources}件を処理中...（{zip_file.name}）")

    progress_placeholder.empty()
    all_missing_folders = reconcile_missing_folders(all_entries, all_missing_folders)

    if not all_entries:
        st.error("有効な台帳ファイルが見つかりませんでした。")
        st.session_state.pop("merged_bytes", None)
        if all_missing_folders:
            with st.expander(f"⚠️ 台帳ファイルが見つからなかったフォルダ（{len(all_missing_folders)}件）"):
                for name in all_missing_folders:
                    st.write(f"- {name}")
    else:
        all_entries.sort(key=lambda e: e.package_name)
        merged_bytes = build_merged_workbook(all_entries)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        st.session_state["merged_bytes"] = merged_bytes
        st.session_state["merged_filename"] = f"統合_図面親子管理台帳_{timestamp}.xlsx"
        st.session_state["merged_count"] = len(all_entries)
        st.session_state["merged_missing_folders"] = all_missing_folders

if "merged_bytes" in st.session_state:
    st.success(f"{st.session_state['merged_count']}個のDiff Packageを統合しました。")
    missing_folders = st.session_state.get("merged_missing_folders") or []
    if missing_folders:
        with st.expander(f"⚠️ 台帳ファイルが見つからなかったフォルダ（{len(missing_folders)}件）"):
            for name in missing_folders:
                st.write(f"- {name}")
    st.download_button(
        "統合Excelをダウンロード",
        data=st.session_state["merged_bytes"],
        file_name=st.session_state["merged_filename"],
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        width="stretch",
    )
