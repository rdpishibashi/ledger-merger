"""実データを使った Ledger-merger の回帰テスト。

DXF-diff-manager の実出力フォルダ
/Users/ryozo/Dropbox/Workspace/diff_PairC_ME25-7102-0_ME25-9606-0_405_01 を直接参照する。
このフォルダが存在しない環境ではテストをスキップする。
"""

import os
import shutil
import sys

import openpyxl
import pytest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", ".."))

from utils.ledger_finder import LedgerEntry, find_ledger_files, reconcile_missing_folders
from utils.ledger_merger import OUTPUT_HEADERS, build_merged_workbook

REAL_DATA_ROOT = "/Users/ryozo/Dropbox/Workspace/diff_PairC_ME25-7102-0_ME25-9606-0_405_01"
SAMPLE_LEDGER_DIR = os.path.join(
    REAL_DATA_ROOT,
    "dxf_diff_results_PairC_ZC00_405_ME25-9606-0_01_ME25-9606-0_02",
)

pytestmark = pytest.mark.skipif(
    not os.path.isdir(REAL_DATA_ROOT), reason="実データフォルダが見つかりません"
)


def test_find_ledger_files_count_matches_actual_files():
    expected = 0
    for dirpath, _dirnames, filenames in os.walk(REAL_DATA_ROOT):
        expected += filenames.count("図面親子管理台帳.xlsx")

    entries, missing_folders = find_ledger_files(REAL_DATA_ROOT)

    assert len(entries) == expected
    assert expected > 0
    assert missing_folders == []


def test_non_ledger_filenames_excluded_and_not_reported_as_missing():
    """diff_labels.xlsx / unchanged_labels.xlsx は台帳候補から除外され、
    台帳が別途存在するフォルダは「台帳が見つからないフォルダ」にも出てこない。"""
    entries, missing_folders = find_ledger_files(REAL_DATA_ROOT)

    ledger_paths = {entry.source_path for entry in entries}
    assert all(os.path.basename(p) not in {"diff_labels.xlsx", "unchanged_labels.xlsx"} for p in ledger_paths)
    assert missing_folders == []


def test_folder_with_only_non_ledger_files_is_reported_as_missing_by_name_only(tmp_path):
    """diff_labels.xlsx / unchanged_labels.xlsx しか無いフォルダは「台帳が見つからない
    フォルダ」として basename のみで報告され、台帳が存在するフォルダは混在しても無視されない。"""
    ok_dir = tmp_path / "dxf_diff_results_PairC_OK_01"
    ok_dir.mkdir()
    shutil.copy(os.path.join(SAMPLE_LEDGER_DIR, "図面親子管理台帳.xlsx"), ok_dir / "図面親子管理台帳.xlsx")

    missing_dir = tmp_path / "dxf_diff_results_PairC_MISSING_01"
    missing_dir.mkdir()
    shutil.copy(os.path.join(SAMPLE_LEDGER_DIR, "diff_labels.xlsx"), missing_dir / "diff_labels.xlsx")
    shutil.copy(os.path.join(SAMPLE_LEDGER_DIR, "unchanged_labels.xlsx"), missing_dir / "unchanged_labels.xlsx")

    entries, missing_folders = find_ledger_files(str(tmp_path))

    assert len(entries) == 1
    assert entries[0].package_name == "dxf_diff_results_PairC_OK_01"
    assert missing_folders == ["dxf_diff_results_PairC_MISSING_01"]


def test_rows_without_diff_stats_are_excluded():
    """Diff List には「差分抽出ペア数」と同数の行（実際に差分抽出された行）だけが
    含まれる。Total Entities が空欄の行（差分未抽出の図番ペアの関係記録）は除外される。

    実データのフォルダ dxf_diff_results_PairC_ZM00_405_ME25-7102-0_02_ME25-9606-0_01 は
    Diff List シートに25行あるが、実際に差分抽出されたのは5行のみ（Summaryシートの
    「差分抽出ペア数」=5 と一致）で、残り20行は統計が空欄の関係記録である。"""
    entries, _missing_folders = find_ledger_files(REAL_DATA_ROOT)

    for entry in entries:
        assert len(entry.diff_list_rows) == entry.summary_values["差分抽出ペア数"]
        assert all(row[-1] is not None for row in entry.diff_list_rows)  # Total Entities

    target = next(
        e for e in entries
        if e.package_name == "dxf_diff_results_PairC_ZM00_405_ME25-7102-0_02_ME25-9606-0_01"
    )
    assert len(target.diff_list_rows) == 5


def test_filtered_rows_entity_sums_match_summary_exactly():
    """除外後に残った行の Deleted/Added/Diff/Unchanged/Total Entities の合計が、
    Summary シートの対応する合計値と1件単位の差もなく完全に一致することを、
    件数だけでなく数値レベルで全フォルダにわたって検証する。"""
    from utils.ledger_finder import DIFF_LIST_HEADERS

    cols = {h: i for i, h in enumerate(DIFF_LIST_HEADERS)}
    label_by_col = {
        "Deleted Entities": "削除図形数 合計",
        "Added Entities": "追加図形数 合計",
        "Diff Entities": "差分図形数 合計",
        "Unchanged Entities": "変更なし図形数 合計",
        "Total Entities": "総図形数 合計",
    }

    entries, _missing_folders = find_ledger_files(REAL_DATA_ROOT)
    assert entries

    for entry in entries:
        for col_name, summary_label in label_by_col.items():
            computed = sum(row[cols[col_name]] for row in entry.diff_list_rows)
            assert computed == entry.summary_values[summary_label], (
                f"{entry.package_name}: {summary_label} 不一致"
            )


def test_filtered_rows_preserve_original_order_and_values():
    """除外後に残る行は、元の Diff List シートの該当行を順序・値ともに
    そのまま保持している（並び替えやデータ欠落が無いことの往復確認）。"""
    from utils.ledger_finder import _TOTAL_ENTITIES_COL

    entries, _missing_folders = find_ledger_files(REAL_DATA_ROOT)
    for entry in entries:
        wb = openpyxl.load_workbook(entry.source_path, data_only=True, read_only=True)
        try:
            src_rows = list(wb["Diff List"].iter_rows(values_only=True))[1:]
        finally:
            wb.close()
        expected = [row for row in src_rows if row[_TOTAL_ENTITIES_COL] is not None]
        assert list(entry.diff_list_rows) == expected


def test_reconcile_missing_folders_removes_names_found_elsewhere():
    """異なる入力ソース（複数ZIP等）の結果を集約した際、同名フォルダが片方で台帳あり・
    もう片方で台帳なしと判定されても、「台帳が見つからなかったフォルダ」一覧には
    矛盾して現れない（成功した名前を優先する）。重複した欠落名も1つにまとめる。"""
    entries = [LedgerEntry(package_name="A", source_path="x", diff_list_rows=[], summary_values={})]
    missing = ["A", "B", "B", "C"]

    reconciled = reconcile_missing_folders(entries, missing)

    assert reconciled == ["B", "C"]


def test_macosx_mirror_folder_not_reported_as_missing(tmp_path):
    """macOS Finder/ditto で ZIP 化すると同名の __MACOSX/<folder>/._ファイル名 ミラーが
    作られる。これが本物のフォルダと同名の偽フォルダとして誤検出されないことを確認する
    （実際に発生した不具合: 台帳が見つかるフォルダ名がそのまま「見つからなかったフォルダ」
    にも重複して表示されていた）。"""
    real_dir = tmp_path / "dxf_diff_results_PairC_OK_01"
    real_dir.mkdir()
    shutil.copy(os.path.join(SAMPLE_LEDGER_DIR, "図面親子管理台帳.xlsx"), real_dir / "図面親子管理台帳.xlsx")

    mirror_dir = tmp_path / "__MACOSX" / "dxf_diff_results_PairC_OK_01"
    mirror_dir.mkdir(parents=True)
    (mirror_dir / "._図面親子管理台帳.xlsx").write_bytes(b"\x00\x05\x16\x07")  # AppleDouble ダミー

    entries, missing_folders = find_ledger_files(str(tmp_path))

    assert len(entries) == 1
    assert entries[0].package_name == "dxf_diff_results_PairC_OK_01"
    assert missing_folders == []


def test_merged_workbook_structure():
    entries, _missing_folders = find_ledger_files(REAL_DATA_ROOT)
    merged_bytes = build_merged_workbook(entries)

    wb = openpyxl.load_workbook(__import__("io").BytesIO(merged_bytes))
    ws = wb["Diff List"]

    header = tuple(c.value for c in ws[1])
    assert header == OUTPUT_HEADERS
    assert len(header) == 22

    row_idx = 2
    for entry in entries:
        for row_in_block in range(len(entry.diff_list_rows)):
            row = ws[row_idx]
            package_cell = row[0]
            summary_cells = row[13:]

            if row_in_block == 0:
                assert package_cell.font.color.rgb == "FF000000"
                assert all(c.value is not None for c in summary_cells)
            else:
                assert package_cell.font.color.rgb == "FFA6A6A6"
                assert all(c.value is None for c in summary_cells)

            assert package_cell.value == entry.package_name
            row_idx += 1

    assert row_idx - 2 == sum(len(e.diff_list_rows) for e in entries)
