"""仕様確認（Spec Regression）: Work Master・Summary シートのデータ形式変更。

対応する受入条件（2026-08-05 のユーザー依頼）:
    Work Master:
        - "Note" を "Total Entities" の後ろへ移動。
        - "Recorded Date" を最後尾へ移動。
        - "Diff Type" を "Deleted Entities" の直前に追加。
    Summary:
        - "差分方式" を "指番" の直後に追加。
    "Diff Type"/"差分方式" は Diff Package（出力フォルダ名。
    "dxf_diff_results_Type{A|B|C}_{指番}(_{モジュール}_{サイド}_*)"）の
    "Type"/"Pair" 直後の1文字から逆算する。

境界・例外の意図:
    - 同一指番内で差分方式が異なる場合（実データでは未確認・理論上のケース）、
      Summaryでは (指番, 差分方式) の組ごとに別行へ分ける
      （ユーザー確認済み仕様。値が誤って混ざらないようにするため）。
    - Work Masterの行自体は差分方式をキーに含めない（1エントリ=1差分方式のため
      行内で曖昧さは生じない。Recorded Date が最も新しい行が丸ごと採用される
      既存の蓄積規則にそのまま従う）。
"""

import os
import sys
from datetime import datetime

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "..", ".."))

from utils.ledger_finder import LedgerEntry
from utils.master_ledger_builder import (
    SUMMARY_HEADERS,
    WORK_MASTER_HEADERS,
    compute_summary_rows,
    extract_unique_work_master_rows,
)

_TYPE_A_PACKAGE = "dxf_diff_results_TypeA_ME24-1001-0_ZC00_405"
_TYPE_B_PACKAGE = "dxf_diff_results_TypeB_ME24-1001-0_ZM00_405"


def _row(child, parent, relation, recorded_date, deleted=1, added=2, diff=3, unchanged=4, total=5, note=None):
    return (child, parent, relation, "T", "S", recorded_date, note, deleted, added, diff, unchanged, total)


def test_work_master_headers_column_order():
    assert WORK_MASTER_HEADERS == (
        "Sashiban", "Module", "Side", "Child", "Parent", "Title", "Subtitle", "Diff Type",
        "Deleted Entities", "Added Entities", "Diff Entities", "Unchanged Entities",
        "Total Entities", "Note", "Recorded Date",
    )


def test_summary_headers_column_order():
    assert SUMMARY_HEADERS == (
        "指番", "差分方式", "削除図形総数", "追加図形総数", "変更図形総数", "図形総数",
        "図形変更率 [%]", "差分ペア総数", "完全新規図面数", "指番図面総数",
        "流用率 [%]", "新規作成率 [%]", "日付",
    )


def test_work_master_row_has_diff_type_and_moved_note_recorded_date():
    entry = LedgerEntry(
        package_name=_TYPE_A_PACKAGE, source_path="a.xlsx",
        diff_list_rows=[_row("C1", "P1", "RevUp", datetime(2026, 8, 5), note="メモ")],
        summary_values={},
    )

    row = extract_unique_work_master_rows([entry])[("ME24-1001-0", "ZC00", "405", "C1", "P1")]

    assert dict(zip(WORK_MASTER_HEADERS, row)) == {
        "Sashiban": "ME24-1001-0", "Module": "ZC00", "Side": "405", "Child": "C1", "Parent": "P1",
        "Title": "T", "Subtitle": "S", "Diff Type": "A",
        "Deleted Entities": 1, "Added Entities": 2, "Diff Entities": 3, "Unchanged Entities": 4,
        "Total Entities": 5, "Note": "メモ", "Recorded Date": datetime(2026, 8, 5),
    }


def test_summary_row_has_diff_type_after_sashiban():
    entry = LedgerEntry(
        package_name=_TYPE_A_PACKAGE, source_path="a.xlsx",
        diff_list_rows=[_row("C1", "P1", "RevUp", datetime(2026, 8, 5))],
        summary_values={"入力図面総数": 1, "差分抽出ペア数": 1},
    )

    rows = compute_summary_rows([entry], run_timestamp=datetime(2026, 8, 5))

    assert len(rows) == 1
    assert dict(zip(SUMMARY_HEADERS, rows[0]))["指番"] == "ME24-1001-0"
    assert dict(zip(SUMMARY_HEADERS, rows[0]))["差分方式"] == "A"


def test_summary_splits_into_separate_rows_when_diff_type_differs_within_sashiban():
    """同一指番でも差分方式が異なれば、Summaryは (指番,差分方式) ごとに別行になる
    （ユーザー確認済み仕様。実データでは未確認の理論上のケース）。"""
    entry_type_a = LedgerEntry(
        package_name=_TYPE_A_PACKAGE, source_path="a.xlsx",
        diff_list_rows=[_row("C1", "P1", "RevUp", datetime(2026, 8, 5), deleted=10, added=20, total=100)],
        summary_values={"入力図面総数": 4, "差分抽出ペア数": 1},
    )
    entry_type_b = LedgerEntry(
        package_name=_TYPE_B_PACKAGE, source_path="b.xlsx",
        diff_list_rows=[_row("C2", "P2", "RevUp", datetime(2026, 8, 5), deleted=1, added=2, total=10)],
        summary_values={"入力図面総数": 2, "差分抽出ペア数": 1},
    )

    rows = compute_summary_rows([entry_type_a, entry_type_b], run_timestamp=datetime(2026, 8, 5))

    assert len(rows) == 2
    by_type = {row[1]: row for row in rows}
    assert set(by_type.keys()) == {"A", "B"}
    assert by_type["A"][0] == by_type["B"][0] == "ME24-1001-0"  # 指番は同じ
    # 各行の集計値は自分のグループ分のみ（混ざらない）
    assert by_type["A"][2] == 10  # 削除図形総数
    assert by_type["B"][2] == 1
    assert by_type["A"][9] == 4  # 指番図面総数
    assert by_type["B"][9] == 2
