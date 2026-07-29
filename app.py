import io
import tempfile
import zipfile

import streamlit as st

from utils.group_summary_builder import build_group_workbooks
from utils.ledger_finder import find_ledger_files, reconcile_missing_folders
from utils.ledger_merger import build_merged_workbook
from utils.master_ledger_builder import build_master_workbook, read_master_rows

st.set_page_config(page_title="図面親子管理台帳 統合ツール", page_icon="📑", layout="wide")

st.title("図面親子管理台帳 統合ツール")
st.write(
    "DXF-diff-manager の複数の出力フォルダにある台帳ファイル（Diff List シート）を"
    "1つの Excel ファイルに統合します。"
)

zip_uploader_version = st.session_state.get("zip_uploader_version", 0)
use_last_master = st.session_state.get("use_last_master", False)

st.subheader("ZIPファイルをアップロード")
st.caption("DXF-diff-manager の出力フォルダ群を ZIP 化してアップロードしてください（複数可）。")
zip_files = st.file_uploader(
    "ZIPファイル",
    type=["zip"],
    accept_multiple_files=True,
    key=f"zip_uploader_{zip_uploader_version}",
)

st.subheader("統合図面管理台帳.xlsx をアップロード")
master_upload = None
master_bytes_override = None
if use_last_master and st.session_state.get("master_bytes"):
    st.info("前回作成した統合図面管理台帳.xlsxを自動的に使用します（再アップロード不要）。")
    if st.button("別のファイルをアップロードし直す"):
        st.session_state["use_last_master"] = False
        st.rerun()
    master_bytes_override = st.session_state["master_bytes"]
else:
    st.caption(
        "前回ダウンロードした統合図面管理台帳.xlsxをアップロードしてください。今回の結果を"
        "マージ（同じChild-Parentは上書き）した最新版を出力します。"
    )
    master_upload = st.file_uploader("統合図面管理台帳.xlsx", type=["xlsx"], key="master_upload")

has_input = bool(zip_files) and (master_upload is not None or master_bytes_override is not None)

merge_done = "final_zip_bytes" in st.session_state
run = st.button("統合実行", type="secondary" if merge_done else "primary", disabled=not has_input)

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
        st.session_state.pop("final_zip_bytes", None)
        if all_missing_folders:
            with st.expander(f"⚠️ 台帳ファイルが見つからなかったフォルダ（{len(all_missing_folders)}件）"):
                for name in all_missing_folders:
                    st.write(f"- {name}")
    else:
        all_entries.sort(key=lambda e: e.package_name)

        previous_master_rows = None
        if master_bytes_override is not None:
            previous_master_rows = read_master_rows(master_bytes_override)
            if previous_master_rows is None:
                st.warning(
                    "前回作成した統合図面管理台帳.xlsxを読み込めなかったため、"
                    "今回分のデータのみで作成します。"
                )
        elif master_upload is not None:
            master_upload.seek(0)
            previous_master_rows = read_master_rows(master_upload.read())
            if previous_master_rows is None:
                st.warning(
                    f"アップロードされた統合図面管理台帳.xlsx（{master_upload.name}）を"
                    "読み込めなかったため、今回分のデータのみで作成します。"
                )

        merged_bytes = build_merged_workbook(all_entries)
        master_bytes = build_master_workbook(all_entries, previous_master_rows=previous_master_rows)
        group_files = build_group_workbooks(all_entries)

        final_zip_buffer = io.BytesIO()
        with zipfile.ZipFile(final_zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
            zf.writestr("図形変更量詳細.xlsx", merged_bytes)
            zf.writestr("統合図面管理台帳.xlsx", master_bytes)
            for filename, data in sorted(group_files.items()):
                zf.writestr(f"指番_モジュール_サイド別集計/{filename}", data)

        st.session_state["final_zip_bytes"] = final_zip_buffer.getvalue()
        st.session_state["master_bytes"] = master_bytes
        st.session_state["merged_count"] = len(all_entries)
        st.session_state["merged_missing_folders"] = all_missing_folders
        st.session_state["group_summary_count"] = len(group_files)
        st.session_state["downloaded_once"] = False
        st.rerun()

if "final_zip_bytes" in st.session_state:
    st.success(f"{st.session_state['merged_count']}個のDiff Packageを統合しました。")
    if st.session_state.get("group_summary_count"):
        st.caption(f"指番_モジュール_サイド単位の集計ファイルが{st.session_state['group_summary_count']}件生成されました。")
    missing_folders = st.session_state.get("merged_missing_folders") or []
    if missing_folders:
        with st.expander(f"⚠️ 台帳ファイルが見つからなかったフォルダ（{len(missing_folders)}件）"):
            for name in missing_folders:
                st.write(f"- {name}")
    downloaded = st.download_button(
        "統合台帳をダウンロード",
        data=st.session_state["final_zip_bytes"],
        file_name="統合図面台帳.zip",
        mime="application/zip",
        type="primary",
    )
    if downloaded:
        st.session_state["downloaded_once"] = True

    if st.session_state.get("downloaded_once"):
        if st.button("新規統合の実行", type="primary"):
            st.session_state["use_last_master"] = True
            st.session_state["zip_uploader_version"] = zip_uploader_version + 1
            for key in (
                "final_zip_bytes",
                "merged_count",
                "merged_missing_folders",
                "group_summary_count",
                "downloaded_once",
            ):
                st.session_state.pop(key, None)
            st.rerun()
