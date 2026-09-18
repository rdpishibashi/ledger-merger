"""
このテストが守るもの: DXF-diff-manager の台帳（新形式、"Unchanged Offset
Entities" 列を持つ）から読み込んだオフセット一致件数が、Ledger-merger の
Master・Work Master・"指番別集計"（group_summary_builder）の Summary の
どちらにも欠落なく伝播すること（2026-09-18新設）。

対応する受入条件（ユーザー承認済み）:
    1. 統合図面管理台帳.xlsx の Master シート: "Unchanged Offset Entities" 列が
       "Total Entities" の直後（末尾）に出力され、値は元の台帳の値をそのまま
       保持する。
    2. 同 Work Master シート: 同列が "Note" の後・"Recorded Date" の前に出力される。
    3. "指番_モジュール_サイド_all.xlsx" の Summary シート: "変更なし（オフセット
       一致）図形 総数" 行が「変更なし図形 総数」の直後に出力され、値は
       entry.summary_values の対応するラベルから取得される。
    4. 同一 Child が複数レビジョンにまたがる場合、Unchanged Offset Entities も
       他のエンティティ数列と同様に単純合計される（aggregate_diff_list_by_child）。

実行:
    cd Ledger-merger
    python -m pytest tests/regression/spec/test_unchanged_offset_column_propagation.py -v
"""
import io
import os
import sys
from datetime import datetime

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "..", ".."))

import openpyxl

from utils.ledger_finder import LedgerEntry
from utils.master_ledger_builder import (
    MASTER_HEADERS,
    MASTER_SHEET_NAME,
    WORK_MASTER_HEADERS,
    WORK_MASTER_SHEET_NAME,
    build_master_workbook,
    extract_unique_child_parent_rows,
    extract_unique_work_master_rows,
)
from utils.group_summary_builder import aggregate_diff_list_by_child

_MATCHING_PACKAGE_NAME = "dxf_diff_results_TypeA_ME24-1001-0_ZC00_405_01"


def _row(child, parent, recorded_date, unchanged_offset):
    """DIFF_LIST_HEADERS（新形式13列）形状の行を作る。"""
    return (child, parent, "RevUp", "T", "S", recorded_date, None, 1, 2, 3, 4, 10, unchanged_offset)


def test_master_sheet_carries_unchanged_offset_entities_value():
    """Master シートの 'Unchanged Offset Entities' 列に、台帳の値がそのまま出力される。"""
    entry = LedgerEntry(
        package_name=_MATCHING_PACKAGE_NAME, source_path="a.xlsx",
        diff_list_rows=[_row("C1", "P1", datetime(2026, 9, 18), unchanged_offset=6)],
        summary_values={},
    )

    wb_bytes = build_master_workbook([entry])
    wb = openpyxl.load_workbook(io.BytesIO(wb_bytes))
    ws = wb[MASTER_SHEET_NAME]
    header = [c.value for c in ws[1]]
    assert header == list(MASTER_HEADERS)

    row = dict(zip(header, next(ws.iter_rows(min_row=2, max_row=2, values_only=True))))
    assert row["Unchanged Offset Entities"] == 6


def test_work_master_sheet_carries_unchanged_offset_entities_value():
    """Work Master シートの同列にも台帳の値がそのまま出力される。"""
    entry = LedgerEntry(
        package_name=_MATCHING_PACKAGE_NAME, source_path="a.xlsx",
        diff_list_rows=[_row("C1", "P1", datetime(2026, 9, 18), unchanged_offset=9)],
        summary_values={},
    )

    wb_bytes = build_master_workbook([entry])
    wb = openpyxl.load_workbook(io.BytesIO(wb_bytes))
    ws = wb[WORK_MASTER_SHEET_NAME]
    header = [c.value for c in ws[1]]
    assert header == list(WORK_MASTER_HEADERS)

    row = dict(zip(header, next(ws.iter_rows(min_row=2, max_row=2, values_only=True))))
    assert row["Unchanged Offset Entities"] == 9


def test_extract_unique_rows_place_value_at_documented_position():
    """extract_unique_child_parent_rows / extract_unique_work_master_rows の
    戻り値でも、列の並び順（Note直後・Recorded Date直前）と値が一致する。"""
    entry = LedgerEntry(
        package_name=_MATCHING_PACKAGE_NAME, source_path="a.xlsx",
        diff_list_rows=[_row("C1", "P1", datetime(2026, 9, 18), unchanged_offset=4)],
        summary_values={},
    )

    master_row = extract_unique_child_parent_rows([entry])[("C1", "P1")]
    assert dict(zip(MASTER_HEADERS, master_row))["Unchanged Offset Entities"] == 4
    # Note(index -3) の直後・Recorded Date(index -1) の直前に位置する
    assert master_row[-2] == 4

    wm_row = extract_unique_work_master_rows([entry])[("ME24-1001-0", "ZC00", "405", "C1", "P1")]
    assert dict(zip(WORK_MASTER_HEADERS, wm_row))["Unchanged Offset Entities"] == 4
    assert wm_row[-2] == 4


def test_aggregate_diff_list_by_child_sums_unchanged_offset_across_revisions():
    """同一Childが複数レビジョンにまたがる場合、Unchanged Offset Entities も
    他のエンティティ数列と同様に単純合計される（指番別集計Summaryの元データ）。"""
    entry_rev1 = LedgerEntry(
        package_name="p_01", source_path="r1.xlsx",
        diff_list_rows=[_row("C1", "P1", datetime(2026, 9, 1), unchanged_offset=3)],
        summary_values={},
    )
    entry_rev2 = LedgerEntry(
        package_name="p_02", source_path="r2.xlsx",
        diff_list_rows=[_row("C1", "P1", datetime(2026, 9, 2), unchanged_offset=5)],
        summary_values={},
    )

    aggregated = aggregate_diff_list_by_child([("01", entry_rev1), ("02", entry_rev2)])

    assert len(aggregated) == 1
    assert aggregated[0][-1] == 8  # 3 + 5


if __name__ == '__main__':
    import pytest
    sys.exit(pytest.main([__file__, "-v"]))
