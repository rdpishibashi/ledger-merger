"""utils.master_ledger_builder のユニットテスト。"""

import io
import os
import sys
from datetime import datetime

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", ".."))

import openpyxl

from utils.ledger_finder import DIFF_LIST_HEADERS, LedgerEntry
from utils.master_ledger_builder import (
    MASTER_SHEET_NAME,
    SUMMARY_HEADERS,
    SUMMARY_SHEET_NAME,
    WORK_MASTER_HEADERS,
    WORK_MASTER_SHEET_NAME,
    build_master_workbook,
    compute_summary_rows,
    extract_unique_child_parent_rows,
    extract_unique_work_master_rows,
    read_master_rows,
    read_summary_rows,
    read_work_master_rows,
)


def _row(child, parent, recorded_date, deleted=1, added=2, diff=3, unchanged=4, total=5):
    return (child, parent, "RevUp", "T", "S", recorded_date, None, deleted, added, diff, unchanged, total)


def test_extract_unique_child_parent_rows_dedupes_within_entries():
    """同じChild-Parentペアが複数エントリにまたがる場合はどれか1件を採用する。"""
    entry_a = LedgerEntry(
        package_name="p1", source_path="a.xlsx",
        diff_list_rows=[_row("C1", "P1", datetime(2026, 7, 1), deleted=1)],
        summary_values={},
    )
    entry_b = LedgerEntry(
        package_name="p2", source_path="b.xlsx",
        diff_list_rows=[_row("C1", "P1", datetime(2026, 7, 1), deleted=1), _row("C2", "P2", datetime(2026, 7, 1))],
        summary_values={},
    )

    unique = extract_unique_child_parent_rows([entry_a, entry_b])

    assert set(unique.keys()) == {("C1", "P1"), ("C2", "P2")}


def test_build_master_workbook_has_three_sheets_in_order():
    """package_name が指番命名規則に一致しない場合、MasterシートにはChild-Parentが
    含まれるが、Work Master・SummaryシートはEmpty相当（指番を逆算できないため）。"""
    entry = LedgerEntry(
        package_name="p1", source_path="a.xlsx",
        diff_list_rows=[
            _row("C2", "P2", datetime(2026, 7, 1)),
            _row("C1", "P1", datetime(2026, 7, 1)),
        ],
        summary_values={},
    )

    wb_bytes = build_master_workbook([entry])
    wb = openpyxl.load_workbook(io.BytesIO(wb_bytes))

    assert wb.sheetnames == [MASTER_SHEET_NAME, WORK_MASTER_SHEET_NAME, SUMMARY_SHEET_NAME]
    ws = wb[MASTER_SHEET_NAME]
    header = tuple(c.value for c in ws[1])
    assert header == DIFF_LIST_HEADERS

    children = [ws.cell(row=r, column=1).value for r in range(2, ws.max_row + 1)]
    assert children == ["C1", "C2"]

    wm_ws = wb[WORK_MASTER_SHEET_NAME]
    assert tuple(c.value for c in wm_ws[1]) == WORK_MASTER_HEADERS
    assert wm_ws.max_row == 1  # ヘッダーのみ（指番を逆算できるエントリが無いため）

    summary_ws = wb[SUMMARY_SHEET_NAME]
    assert tuple(c.value for c in summary_ws[1]) == SUMMARY_HEADERS
    assert summary_ws.max_row == 1  # ヘッダーのみ（同上）


def test_build_master_workbook_overwrites_matching_pair_and_keeps_others():
    """前回のMasterに存在する行のうち、今回のデータに無いChild-Parentは残す。
    同じペアは今回のデータで上書きする。"""
    previous_entry = LedgerEntry(
        package_name="prev", source_path="prev.xlsx",
        diff_list_rows=[
            _row("C1", "P1", datetime(2026, 7, 1), deleted=100),
            _row("C_OLD_ONLY", "P_OLD", datetime(2026, 7, 1), deleted=999),
        ],
        summary_values={},
    )
    previous_bytes = build_master_workbook([previous_entry])
    previous_rows = read_master_rows(previous_bytes)

    new_entry = LedgerEntry(
        package_name="new", source_path="new.xlsx",
        diff_list_rows=[_row("C1", "P1", datetime(2026, 7, 9), deleted=1)],
        summary_values={},
    )

    merged_bytes = build_master_workbook([new_entry], previous_master_rows=previous_rows)
    wb = openpyxl.load_workbook(io.BytesIO(merged_bytes))
    ws = wb[MASTER_SHEET_NAME]

    rows_by_child = {
        row[0]: row for row in ws.iter_rows(min_row=2, values_only=True)
    }
    assert set(rows_by_child.keys()) == {"C1", "C_OLD_ONLY"}
    assert rows_by_child["C1"][7] == 1  # 上書きされた新しい値（Deleted Entities）
    assert rows_by_child["C_OLD_ONLY"][7] == 999  # 旧データがそのまま保持される


def test_read_master_rows_returns_none_for_invalid_file():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "SomeOtherSheet"
    ws.append(["not", "a", "master", "sheet"])
    buf = io.BytesIO()
    wb.save(buf)

    assert read_master_rows(buf.getvalue()) is None


# --- Work Master ---

_MATCHING_PACKAGE_NAME = "dxf_diff_results_PairA_ME24-1001-0_ZC00_405_01"
_MATCHING_PACKAGE_NAME_2 = "dxf_diff_results_TypeB_ME24-1001-0_ZMF1_405_02"
_UNRESOLVABLE_PACKAGE_NAME = "some_manually_named_folder"


def test_extract_unique_work_master_rows_dedupes_and_excludes_unresolvable_sashiban():
    """指番を逆算できるエントリのみ (指番,Child,Parent) でユニーク化する。
    逆算できないエントリ（package_nameが命名規則に一致しない）は対象外。"""
    resolvable = LedgerEntry(
        package_name=_MATCHING_PACKAGE_NAME, source_path="a.xlsx",
        diff_list_rows=[_row("C1", "P1", datetime(2026, 7, 1))],
        summary_values={},
    )
    duplicate_pair_different_entry = LedgerEntry(
        package_name=_MATCHING_PACKAGE_NAME_2, source_path="b.xlsx",
        diff_list_rows=[_row("C1", "P1", datetime(2026, 7, 1))],
        summary_values={},
    )
    unresolvable = LedgerEntry(
        package_name=_UNRESOLVABLE_PACKAGE_NAME, source_path="c.xlsx",
        diff_list_rows=[_row("C2", "P2", datetime(2026, 7, 1))],
        summary_values={},
    )

    unique = extract_unique_work_master_rows([resolvable, duplicate_pair_different_entry, unresolvable])

    assert set(unique.keys()) == {("ME24-1001-0", "C1", "P1")}


def test_extract_unique_work_master_rows_excludes_relation_and_includes_sashiban():
    entry = LedgerEntry(
        package_name=_MATCHING_PACKAGE_NAME, source_path="a.xlsx",
        diff_list_rows=[_row("C1", "P1", datetime(2026, 7, 1), deleted=10, added=20, diff=30, unchanged=40, total=50)],
        summary_values={},
    )

    unique = extract_unique_work_master_rows([entry])
    row = unique[("ME24-1001-0", "C1", "P1")]

    assert row == ("ME24-1001-0", "C1", "P1", "T", "S", datetime(2026, 7, 1), None, 10, 20, 30, 40, 50)
    assert "RevUp" not in row  # Relation列は含まれない


def test_build_master_workbook_work_master_populated_for_matching_package_names():
    entry_1 = LedgerEntry(
        package_name=_MATCHING_PACKAGE_NAME, source_path="a.xlsx",
        diff_list_rows=[_row("C2", "P2", datetime(2026, 7, 1))],
        summary_values={},
    )
    entry_2 = LedgerEntry(
        package_name=_UNRESOLVABLE_PACKAGE_NAME, source_path="b.xlsx",
        diff_list_rows=[_row("C1", "P1", datetime(2026, 7, 1))],
        summary_values={},
    )

    wb_bytes = build_master_workbook([entry_1, entry_2])
    wb = openpyxl.load_workbook(io.BytesIO(wb_bytes))
    wm_ws = wb[WORK_MASTER_SHEET_NAME]

    rows = list(wm_ws.iter_rows(min_row=2, values_only=True))
    assert len(rows) == 1  # 指番不明の entry_2 は含まれない
    assert rows[0][:3] == ("ME24-1001-0", "C2", "P2")


def test_build_master_workbook_work_master_overwrites_matching_key_and_keeps_others():
    """前回のWork Masterに存在する行のうち、今回のデータに無い(指番,Child,Parent)は残す。
    同じキーは今回のデータで上書きする（Masterシートと同じ蓄積方式）。"""
    previous_entry = LedgerEntry(
        package_name=_MATCHING_PACKAGE_NAME, source_path="prev.xlsx",
        diff_list_rows=[
            _row("C1", "P1", datetime(2026, 7, 1), deleted=100),
            _row("C_OLD_ONLY", "P_OLD", datetime(2026, 7, 1), deleted=999),
        ],
        summary_values={},
    )
    previous_bytes = build_master_workbook([previous_entry])
    previous_work_master_rows = read_work_master_rows(previous_bytes)
    assert previous_work_master_rows is not None

    new_entry = LedgerEntry(
        package_name=_MATCHING_PACKAGE_NAME, source_path="new.xlsx",
        diff_list_rows=[_row("C1", "P1", datetime(2026, 7, 9), deleted=1)],
        summary_values={},
    )

    merged_bytes = build_master_workbook([new_entry], previous_work_master_rows=previous_work_master_rows)
    wb = openpyxl.load_workbook(io.BytesIO(merged_bytes))
    wm_ws = wb[WORK_MASTER_SHEET_NAME]

    rows_by_child = {row[1]: row for row in wm_ws.iter_rows(min_row=2, values_only=True)}
    assert set(rows_by_child.keys()) == {"C1", "C_OLD_ONLY"}
    assert rows_by_child["C1"][7] == 1  # 上書きされた新しい値（Deleted Entities）
    assert rows_by_child["C_OLD_ONLY"][7] == 999  # 旧データがそのまま保持される


def test_read_work_master_rows_returns_none_when_sheet_missing():
    """Work Masterシートが無い旧バージョンの統合図面管理台帳.xlsxをアップロードしても
    None を返し（例外にならず）、呼び出し側で今回分のみの新規作成にフォールバックできる。"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = MASTER_SHEET_NAME
    ws.append(DIFF_LIST_HEADERS)
    buf = io.BytesIO()
    wb.save(buf)

    assert read_work_master_rows(buf.getvalue()) is None


def test_read_work_master_rows_returns_none_for_invalid_file():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "SomeOtherSheet"
    ws.append(["not", "a", "work", "master", "sheet"])
    buf = io.BytesIO()
    wb.save(buf)

    assert read_work_master_rows(buf.getvalue()) is None


# --- Summary ---

def test_compute_summary_rows_values():
    """削除/追加/変更/図形総数・変更率・差分ペア総数・流用率が定義通り算出される。
    本テストのデータは全行 Relation='RevUp' のため完全新規図面は0件（完全新規図面の
    除外・カウントの検証は test_compute_summary_rows_excludes_brand_new_from_pair_count
    参照）。指番図面総数は指番_モジュール_サイド別集計と同じ「アップロード図面総数」
    TOTAL値。"""
    entry = LedgerEntry(
        package_name=_MATCHING_PACKAGE_NAME, source_path="a.xlsx",
        diff_list_rows=[
            _row("C1", "P1", datetime(2026, 7, 1), deleted=10, added=20, diff=30, unchanged=40, total=100),
            _row("C2", "P2", datetime(2026, 7, 1), deleted=5, added=5, diff=10, unchanged=10, total=50),
        ],
        summary_values={"入力図面総数": 8, "差分抽出ペア数": 2},
    )

    rows = compute_summary_rows([entry], run_timestamp=datetime(2026, 7, 31, 12, 0, 0))

    assert len(rows) == 1
    (sashiban, deleted, added, changed, entity_total, change_rate, pair_count,
     brand_new_count, input_total, reuse_rate, brand_new_rate, ts) = rows[0]
    assert sashiban == "ME24-1001-0"
    assert (deleted, added, changed) == (15, 25, 40)  # 10+5, 20+5, 40=15+25
    assert entity_total == 150  # 100+50
    assert change_rate == 40 / 150
    assert pair_count == 2
    assert brand_new_count == 0
    assert input_total == 8  # LedgerEntry.summary_values["入力図面総数"]から
    assert reuse_rate == 2 / 8
    assert brand_new_rate == 0.0
    assert ts == datetime(2026, 7, 31, 12, 0, 0)


def test_compute_summary_rows_excludes_brand_new_from_pair_count():
    """Relation='完全新規図面' の行は「差分ペア総数」から除外され、代わりに
    「完全新規図面数」（Childユニーク数）としてカウントされる（2026-08、
    DXF-diff-manager自身の「差分抽出ペア数」の定義〈完全新規図面を含まない〉に
    揃えるための変更）。エンティティ統計（削除/追加/変更/図形総数）は完全新規図面の
    行も含めたまま合計する。"""
    entry = LedgerEntry(
        package_name=_MATCHING_PACKAGE_NAME, source_path="a.xlsx",
        diff_list_rows=[
            _row("C1", "P1", datetime(2026, 7, 1), deleted=10, added=20, diff=30, unchanged=40, total=100),
            (
                "C2", "none", "完全新規図面", "T", "S", datetime(2026, 7, 1), None,
                "n/a", 30, "n/a", "n/a", 30,
            ),
        ],
        summary_values={"入力図面総数": 4, "差分抽出ペア数": 1},
    )

    rows = compute_summary_rows([entry], run_timestamp=datetime.now())

    assert len(rows) == 1
    (_sashiban, deleted, added, changed, entity_total, _change_rate, pair_count,
     brand_new_count, input_total, reuse_rate, brand_new_rate, _ts) = rows[0]
    assert pair_count == 1  # 完全新規図面（C2）を除外し、通常ペア（C1）のみ
    assert brand_new_count == 1
    assert (deleted, added, changed, entity_total) == (10, 50, 60, 130)  # 完全新規図面の行も合算対象
    assert input_total == 4
    assert reuse_rate == 1 / 4
    assert brand_new_rate == 1 / 4


def test_compute_summary_rows_excludes_unresolvable_sashiban():
    entry = LedgerEntry(
        package_name=_UNRESOLVABLE_PACKAGE_NAME, source_path="a.xlsx",
        diff_list_rows=[_row("C1", "P1", datetime(2026, 7, 1))],
        summary_values={"入力図面総数": 1, "差分抽出ペア数": 1},
    )

    rows = compute_summary_rows([entry], run_timestamp=datetime.now())

    assert rows == []


def test_compute_summary_rows_treats_non_numeric_entity_values_as_zero():
    """'n/a'（完全新規図面等）が混在しても、Summaryの合計は常に数値になる。"""
    entry = LedgerEntry(
        package_name=_MATCHING_PACKAGE_NAME, source_path="a.xlsx",
        diff_list_rows=[
            _row("C1", "P1", datetime(2026, 7, 1), deleted='n/a', added=100, diff='n/a', unchanged='n/a', total=100),
        ],
        summary_values={"入力図面総数": 1, "差分抽出ペア数": 1},
    )

    rows = compute_summary_rows([entry], run_timestamp=datetime.now())
    _sashiban, deleted, added, changed, entity_total, *_rest = rows[0]

    assert deleted == 0
    assert added == 100
    assert changed == 100  # 0 + 100
    assert entity_total == 100


def test_build_master_workbook_summary_appends_without_deduping_by_sashiban():
    """Summaryシートはキー単位のマージを行わず、前回分の末尾に今回分を単純追記する
    （同じ指番の行が複数回登場しうる）。"""
    entry = LedgerEntry(
        package_name=_MATCHING_PACKAGE_NAME, source_path="a.xlsx",
        diff_list_rows=[_row("C1", "P1", datetime(2026, 7, 1))],
        summary_values={"入力図面総数": 1, "差分抽出ペア数": 1},
    )

    first_bytes = build_master_workbook([entry])
    previous_summary_rows = read_summary_rows(first_bytes)
    assert len(previous_summary_rows) == 1

    second_bytes = build_master_workbook([entry], previous_summary_rows=previous_summary_rows)
    wb = openpyxl.load_workbook(io.BytesIO(second_bytes))
    summary_ws = wb[SUMMARY_SHEET_NAME]

    rows = list(summary_ws.iter_rows(min_row=2, values_only=True))
    assert len(rows) == 2  # 1回目の行がそのまま残り、2回目の行が追記される
    assert all(row[0] == "ME24-1001-0" for row in rows)


def test_read_summary_rows_returns_empty_list_when_sheet_missing():
    """Summaryシートが無い旧バージョンの統合図面管理台帳.xlsxをアップロードしても
    空リストを返す（呼び出し側で今回分のみ追記できる）。"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = MASTER_SHEET_NAME
    ws.append(DIFF_LIST_HEADERS)
    buf = io.BytesIO()
    wb.save(buf)

    assert read_summary_rows(buf.getvalue()) == []


def test_summary_percent_columns_are_formatted_as_percent():
    entry = LedgerEntry(
        package_name=_MATCHING_PACKAGE_NAME, source_path="a.xlsx",
        diff_list_rows=[_row("C1", "P1", datetime(2026, 7, 1))],
        summary_values={"入力図面総数": 1, "差分抽出ペア数": 1},
    )

    wb_bytes = build_master_workbook([entry])
    wb = openpyxl.load_workbook(io.BytesIO(wb_bytes))
    summary_ws = wb[SUMMARY_SHEET_NAME]

    change_rate_col = SUMMARY_HEADERS.index("図形変更率 [%]") + 1
    reuse_rate_col = SUMMARY_HEADERS.index("流用率 [%]") + 1
    brand_new_rate_col = SUMMARY_HEADERS.index("新規作成率 [%]") + 1
    assert summary_ws.cell(row=2, column=change_rate_col).number_format == "0.00%"
    assert summary_ws.cell(row=2, column=reuse_rate_col).number_format == "0.00%"
    assert summary_ws.cell(row=2, column=brand_new_rate_col).number_format == "0.00%"
