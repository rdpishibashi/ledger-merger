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
    build_master_workbook,
    extract_unique_child_parent_rows,
    read_master_rows,
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


def test_build_master_workbook_sorted_by_child_with_single_sheet():
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

    assert wb.sheetnames == [MASTER_SHEET_NAME]
    ws = wb[MASTER_SHEET_NAME]
    header = tuple(c.value for c in ws[1])
    assert header == DIFF_LIST_HEADERS

    children = [ws.cell(row=r, column=1).value for r in range(2, ws.max_row + 1)]
    assert children == ["C1", "C2"]


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
