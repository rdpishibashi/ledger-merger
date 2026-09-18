"""utils.master_ledger_builder のユニットテスト。"""

import io
import os
import sys
from datetime import datetime

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", ".."))

import openpyxl

from utils.ledger_finder import DIFF_LIST_HEADERS, LedgerEntry
from utils.master_ledger_builder import (
    MASTER_HEADERS,
    MASTER_SHEET_NAME,
    SUMMARY_HEADERS,
    SUMMARY_SHEET_NAME,
    WORK_MASTER_HEADERS,
    WORK_MASTER_SHEET_NAME,
    WORK_MASTER_STRING_LABELS,
    build_master_workbook,
    compute_summary_rows,
    extract_unique_child_parent_rows,
    extract_unique_work_master_rows,
    read_master_rows,
    read_work_master_rows,
)


def _row(child, parent, recorded_date, deleted=1, added=2, diff=3, unchanged=4, total=5, unchanged_offset=0):
    return (child, parent, "RevUp", "T", "S", recorded_date, None, deleted, added, diff, unchanged, total, unchanged_offset)


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
    assert header == MASTER_HEADERS

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
    assert rows_by_child["C1"][6] == 1  # 上書きされた新しい値（Deleted Entities）
    assert rows_by_child["C_OLD_ONLY"][6] == 999  # 旧データがそのまま保持される


def test_read_master_rows_returns_none_for_invalid_file():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "SomeOtherSheet"
    ws.append(["not", "a", "master", "sheet"])
    buf = io.BytesIO()
    wb.save(buf)

    assert read_master_rows(buf.getvalue()) is None


# --- Work Master ---

_MATCHING_PACKAGE_NAME = "dxf_diff_results_TypeA_ME24-1001-0_ZC00_405_01"
_MATCHING_PACKAGE_NAME_2 = "dxf_diff_results_TypeB_ME24-1001-0_ZMF1_405_02"
_UNRESOLVABLE_PACKAGE_NAME = "some_manually_named_folder"


def test_extract_unique_work_master_rows_dedupes_and_excludes_unresolvable_sashiban():
    """指番を逆算できるエントリのみ (指番,モジュール,サイド,Child,Parent) で
    ユニーク化する。逆算できないエントリ（package_nameが命名規則に一致しない）は
    対象外。同一 (指番,Child,Parent) でもモジュール/サイドが異なれば別行として残る
    （2026-08、実データで同一指番+Child+Parentが複数モジュールに跨る例を確認した
    ための仕様）。"""
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

    assert set(unique.keys()) == {
        ("ME24-1001-0", "ZC00", "405", "C1", "P1"),
        ("ME24-1001-0", "ZMF1", "405", "C1", "P1"),
    }


def test_extract_unique_work_master_rows_includes_relation_and_sashiban():
    """列順は Sashiban, Module, Side, Child, Parent, Relation, Title, Subtitle,
    Diff Type, Deleted/Added/Diff/Unchanged/Total Entities, Note, Recorded Date。
    Relation は 2026-09 に追加（元の Diff List 行の値をそのまま持つ）。Diff Type は
    package_name（"dxf_diff_results_TypeA_..."）の "A" が入る。"""
    entry = LedgerEntry(
        package_name=_MATCHING_PACKAGE_NAME, source_path="a.xlsx",
        diff_list_rows=[_row("C1", "P1", datetime(2026, 7, 1), deleted=10, added=20, diff=30, unchanged=40, total=50)],
        summary_values={},
    )

    unique = extract_unique_work_master_rows([entry])
    row = unique[("ME24-1001-0", "ZC00", "405", "C1", "P1")]

    assert row == (
        "ME24-1001-0", "ZC00", "405", "C1", "P1", "RevUp", "T", "S", "A", 10, 20, 30, 40, 50, None, 0,
        datetime(2026, 7, 1),
    )
    assert dict(zip(WORK_MASTER_HEADERS, row))["Relation"] == "RevUp"


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
    assert rows[0][:5] == ("ME24-1001-0", "ZC00", "405", "C2", "P2")


def test_build_master_workbook_work_master_overwrites_matching_key_and_keeps_others():
    """前回のWork Masterに存在する行のうち、今回のデータに無い
    (指番,モジュール,サイド,Child,Parent) は残す。同じキーは今回のデータで
    上書きする（Masterシートと同じ蓄積方式）。"""
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

    child_col = WORK_MASTER_HEADERS.index("Child")
    deleted_col = WORK_MASTER_HEADERS.index("Deleted Entities")
    rows_by_child = {row[child_col]: row for row in wm_ws.iter_rows(min_row=2, values_only=True)}
    assert set(rows_by_child.keys()) == {"C1", "C_OLD_ONLY"}
    assert rows_by_child["C1"][deleted_col] == 1  # 上書きされた新しい値（Deleted Entities）
    assert rows_by_child["C_OLD_ONLY"][deleted_col] == 999  # 旧データがそのまま保持される


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


# --- Work Master: 文字列型・左寄せ（2026-09） ---

def test_work_master_string_columns_are_left_aligned_and_stringified():
    """Sashiban/Module/Side/Child/Parent/Title/Subtitleの7列は文字列型・左寄せで
    出力される。"""
    entry = LedgerEntry(
        package_name=_MATCHING_PACKAGE_NAME, source_path="a.xlsx",
        diff_list_rows=[_row("C1", "P1", datetime(2026, 7, 1))],
        summary_values={},
    )

    wb_bytes = build_master_workbook([entry])
    wb = openpyxl.load_workbook(io.BytesIO(wb_bytes))
    wm_ws = wb[WORK_MASTER_SHEET_NAME]

    assert WORK_MASTER_STRING_LABELS == ("Sashiban", "Module", "Side", "Child", "Parent", "Title", "Subtitle")
    for label in WORK_MASTER_STRING_LABELS:
        col = WORK_MASTER_HEADERS.index(label) + 1
        cell = wm_ws.cell(row=2, column=col)
        assert isinstance(cell.value, str)
        assert cell.alignment.horizontal == "left"


def test_work_master_string_columns_heal_numeric_drift_from_previous_upload():
    """前回アップロードされたWork Masterの該当列が数値型で保存されていた場合でも、
    今回の出力では文字列に正規化される（前回分のみで新規データが無い場合も含む）。"""
    previous_work_master_rows = {
        ("ME24-1001-0", "ZC00", 405, "C_OLD", "P_OLD"): (
            "ME24-1001-0", "ZC00", 405, "C_OLD", "P_OLD", "RevUp", "T", "S", "A",
            1, 2, 3, 4, 5, None, 0, datetime(2026, 7, 1),
        ),
    }

    wb_bytes = build_master_workbook([], previous_work_master_rows=previous_work_master_rows)
    wb = openpyxl.load_workbook(io.BytesIO(wb_bytes))
    wm_ws = wb[WORK_MASTER_SHEET_NAME]

    side_col = WORK_MASTER_HEADERS.index("Side") + 1
    cell = wm_ws.cell(row=2, column=side_col)
    assert cell.value == "405"
    assert isinstance(cell.value, str)
    assert cell.alignment.horizontal == "left"


# --- Summary（2026-09、Work Master集計への変更） ---

def test_compute_summary_rows_values():
    """削除/追加/変更/図形総数・変更率・変更図面総数が定義通り算出される。
    本テストのデータは全行 Parent='P1'/'P2'（"none"ではない）のため完全新規図面は0件
    （完全新規図面の除外・カウントの検証は
    test_compute_summary_rows_counts_brand_new_via_parent_none 参照）。
    差分方式（"dxf_diff_results_TypeA_..."から逆算される"A"）は2026-08追加。"""
    entry = LedgerEntry(
        package_name=_MATCHING_PACKAGE_NAME, source_path="a.xlsx",
        diff_list_rows=[
            _row("C1", "P1", datetime(2026, 7, 1), deleted=10, added=20, diff=30, unchanged=40, total=100),
            _row("C2", "P2", datetime(2026, 7, 1), deleted=5, added=5, diff=10, unchanged=10, total=50),
        ],
        summary_values={},
    )
    work_master = extract_unique_work_master_rows([entry])

    rows = compute_summary_rows(work_master)

    assert len(rows) == 1
    (sashiban, diff_type, deleted, added, changed, entity_total, change_rate, pair_count,
     brand_new_count, latest_date) = rows[0]
    assert sashiban == "ME24-1001-0"
    assert diff_type == "A"
    assert (deleted, added, changed) == (15, 25, 40)  # 10+5, 20+5, 40=15+25
    assert entity_total == 150  # 100+50
    assert change_rate == 40 / 150
    assert pair_count == 2
    assert brand_new_count == 0
    assert latest_date == datetime(2026, 7, 1)


def test_compute_summary_rows_counts_brand_new_via_parent_none():
    """Parent=="none" の行は「変更図面総数」から除外され、代わりに「完全新規図面数」
    （Childユニーク数）としてカウントされる。Work MasterにはRelation列が無いため、
    DXF-diff-manager自身の完全新規図面の表現（Parent="none"）で判定する。
    エンティティ統計（削除/追加/変更/図形総数）は完全新規図面の行も含めたまま合計する。"""
    entry = LedgerEntry(
        package_name=_MATCHING_PACKAGE_NAME, source_path="a.xlsx",
        diff_list_rows=[
            _row("C1", "P1", datetime(2026, 7, 1), deleted=10, added=20, diff=30, unchanged=40, total=100),
            ("C2", "none", "完全新規図面", "T", "S", datetime(2026, 7, 1), None, 0, 30, 0, 0, 30, 0),
        ],
        summary_values={},
    )
    work_master = extract_unique_work_master_rows([entry])

    rows = compute_summary_rows(work_master)

    assert len(rows) == 1
    (_sashiban, _diff_type, deleted, added, changed, entity_total, _change_rate, pair_count,
     brand_new_count, _latest_date) = rows[0]
    assert pair_count == 1  # 完全新規図面（C2）を除外し、通常ペア（C1）のみ
    assert brand_new_count == 1
    assert (deleted, added, changed, entity_total) == (10, 50, 60, 130)  # 完全新規図面の行も合算対象


def test_compute_summary_rows_excludes_unresolvable_sashiban():
    """指番を逆算できないエントリはそもそもWork Masterに含まれないため、
    Summaryにも現れない。"""
    entry = LedgerEntry(
        package_name=_UNRESOLVABLE_PACKAGE_NAME, source_path="a.xlsx",
        diff_list_rows=[_row("C1", "P1", datetime(2026, 7, 1))],
        summary_values={},
    )
    work_master = extract_unique_work_master_rows([entry])

    assert compute_summary_rows(work_master) == []


def test_compute_summary_rows_treats_non_numeric_entity_values_as_zero():
    """非数値のエンティティ値が混在しても、Summaryの合計は常に数値になる
    （utils.ledger_finderの読み込み時点でn/aは既に0へ正規化されるため通常は
    到達しないが、防御的に確認する）。"""
    entry = LedgerEntry(
        package_name=_MATCHING_PACKAGE_NAME, source_path="a.xlsx",
        diff_list_rows=[
            _row("C1", "P1", datetime(2026, 7, 1), deleted='n/a', added=100, diff='n/a', unchanged='n/a', total=100),
        ],
        summary_values={},
    )
    work_master = extract_unique_work_master_rows([entry])

    rows = compute_summary_rows(work_master)
    _sashiban, _diff_type, deleted, added, changed, entity_total, *_rest = rows[0]

    assert deleted == 0
    assert added == 100
    assert changed == 100  # 0 + 100
    assert entity_total == 100


def test_compute_summary_rows_date_is_max_recorded_date_in_group():
    """「日付」はグループ内のRecorded Dateの最大値になる（実行時刻ではない）。"""
    work_master = {
        ("S1", "M1", "SD1", "C1", "P1"): (
            "S1", "M1", "SD1", "C1", "P1", "RevUp", "T", "S", "A", 1, 1, 2, 1, 3, None, 0, datetime(2026, 7, 1),
        ),
        ("S1", "M1", "SD1", "C2", "P2"): (
            "S1", "M1", "SD1", "C2", "P2", "RevUp", "T", "S", "A", 1, 1, 2, 1, 3, None, 0, datetime(2026, 7, 20),
        ),
    }

    rows = compute_summary_rows(work_master)

    assert len(rows) == 1
    assert rows[0][SUMMARY_HEADERS.index("日付")] == datetime(2026, 7, 20)


def test_build_master_workbook_summary_is_recomputed_from_merged_work_master():
    """Summaryはキー単位の追記ログではなく、マージ済み（前回分+今回分）の
    Work Master全体を(指番,差分方式)単位で毎回再集計する。前回のみのChildも
    今回の集計に含まれ、同じキーは重複せず1行にまとまる。"""
    previous_entry = LedgerEntry(
        package_name=_MATCHING_PACKAGE_NAME, source_path="prev.xlsx",
        diff_list_rows=[_row("C_OLD", "P_OLD", datetime(2026, 7, 1), deleted=1, added=1, diff=2, unchanged=1, total=3)],
        summary_values={},
    )
    previous_bytes = build_master_workbook([previous_entry])
    previous_work_master_rows = read_work_master_rows(previous_bytes)

    new_entry = LedgerEntry(
        package_name=_MATCHING_PACKAGE_NAME, source_path="new.xlsx",
        diff_list_rows=[_row("C_NEW", "P_NEW", datetime(2026, 7, 9), deleted=1, added=1, diff=2, unchanged=1, total=3)],
        summary_values={},
    )

    merged_bytes = build_master_workbook([new_entry], previous_work_master_rows=previous_work_master_rows)
    wb = openpyxl.load_workbook(io.BytesIO(merged_bytes))
    summary_ws = wb[SUMMARY_SHEET_NAME]

    rows = list(summary_ws.iter_rows(min_row=2, values_only=True))
    assert len(rows) == 1  # 同一(指番,差分方式)は1行にまとまる（追記されない）
    pair_count_col = SUMMARY_HEADERS.index("変更図面総数")
    assert rows[0][pair_count_col] == 2  # 前回のC_OLDと今回のC_NEWの両方を含む


def test_summary_headers_no_longer_include_removed_columns():
    """指番図面総数・流用率[%]・新規作成率[%]は2026-09に削除された。"""
    assert "指番図面総数" not in SUMMARY_HEADERS
    assert "流用率 [%]" not in SUMMARY_HEADERS
    assert "新規作成率 [%]" not in SUMMARY_HEADERS
    assert SUMMARY_HEADERS == (
        "指番", "差分方式", "削除図形総数", "追加図形総数", "変更図形総数", "図形総数",
        "図形変更率 [%]", "変更図面総数", "完全新規図面数", "日付",
    )


def test_summary_percent_columns_are_formatted_as_percent():
    entry = LedgerEntry(
        package_name=_MATCHING_PACKAGE_NAME, source_path="a.xlsx",
        diff_list_rows=[_row("C1", "P1", datetime(2026, 7, 1))],
        summary_values={},
    )

    wb_bytes = build_master_workbook([entry])
    wb = openpyxl.load_workbook(io.BytesIO(wb_bytes))
    summary_ws = wb[SUMMARY_SHEET_NAME]

    change_rate_col = SUMMARY_HEADERS.index("図形変更率 [%]") + 1
    assert summary_ws.cell(row=2, column=change_rate_col).number_format == "0.00%"
