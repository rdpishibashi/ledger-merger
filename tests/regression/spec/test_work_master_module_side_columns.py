"""仕様確認（Spec Regression）: Work Masterシートに Module・Side 列を追加。

対応する受入条件（2026-08-04 のユーザー依頼）:
    「Work Master」シートの "Sashiban" と "Child" の間に "Module"・"Side" 列を
    追加する。

境界・例外の意図:
    - Work Masterのユニーク化キーは (Sashiban, Module, Side, Child, Parent) に
      拡張する。実データ（Archive.zip、2026-08-04検証）で、同一 (指番, Child,
      Parent) が異なるモジュール（ZC00・ZM00）の両方に記録される例が3件確認
      されたため、モジュール/サイドを含めないと片方が消えてしまう。
    - Summaryシートのキーは従来どおり (Sashiban, Child, Parent) のまま据え置く。
      Summaryにもモジュール/サイドを含めると、「差分ペア総数」がDXF-diff-manager
      自身の「差分抽出ペア数」の定義（グループごとの合計）とずれる
      （実データ検証: ME24-1001-0 で 51→52 に変化）。
"""

import os
import sys
from datetime import datetime

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "..", ".."))

from utils.ledger_finder import LedgerEntry
from utils.master_ledger_builder import (
    WORK_MASTER_HEADERS,
    compute_summary_rows,
    extract_unique_work_master_rows,
)

_ZC00_PACKAGE = "dxf_diff_results_TypeA_ME24-1001-0_ZC00_405"
_ZM00_PACKAGE = "dxf_diff_results_TypeA_ME24-1001-0_ZM00_405"


def _row(child, parent, relation, recorded_date, deleted=1, added=2, diff=3, unchanged=4, total=5):
    return (child, parent, relation, "T", "S", recorded_date, None, deleted, added, diff, unchanged, total)


def test_work_master_headers_place_module_side_between_sashiban_and_child():
    assert WORK_MASTER_HEADERS[:5] == ("Sashiban", "Module", "Side", "Child", "Parent")


def test_work_master_keeps_both_rows_when_same_child_parent_spans_two_modules():
    """同一 (指番,Child,Parent) が異なるモジュールに記録されている場合、Work Master
    では両方が別行として残る（モジュール/サイドをキーに含めるため）。"""
    entry_zc00 = LedgerEntry(
        package_name=_ZC00_PACKAGE, source_path="a.xlsx",
        diff_list_rows=[_row("C1", "P1", "RevUp", datetime(2026, 8, 1))],
        summary_values={},
    )
    entry_zm00 = LedgerEntry(
        package_name=_ZM00_PACKAGE, source_path="b.xlsx",
        diff_list_rows=[_row("C1", "P1", "RevUp", datetime(2026, 8, 1))],
        summary_values={},
    )

    unique = extract_unique_work_master_rows([entry_zc00, entry_zm00])

    assert set(unique.keys()) == {
        ("ME24-1001-0", "ZC00", "405", "C1", "P1"),
        ("ME24-1001-0", "ZM00", "405", "C1", "P1"),
    }
    assert unique[("ME24-1001-0", "ZC00", "405", "C1", "P1")][:5] == ("ME24-1001-0", "ZC00", "405", "C1", "P1")
    assert unique[("ME24-1001-0", "ZM00", "405", "C1", "P1")][:5] == ("ME24-1001-0", "ZM00", "405", "C1", "P1")


def test_summary_pair_count_unaffected_by_module_side_key_change():
    """Work Masterのキー拡張後も、Summaryの「差分ペア総数」は従来どおり
    (指番,Child,Parent) 単位でユニーク化した値のまま変わらない（モジュール違いの
    重複は1ペアとして数える。DXF-diff-manager自身の「差分抽出ペア数」の定義に
    揃えるための仕様）。"""
    entry_zc00 = LedgerEntry(
        package_name=_ZC00_PACKAGE, source_path="a.xlsx",
        diff_list_rows=[_row("C1", "P1", "RevUp", datetime(2026, 8, 1, 10, 0))],
        summary_values={},
    )
    entry_zm00 = LedgerEntry(
        package_name=_ZM00_PACKAGE, source_path="b.xlsx",
        diff_list_rows=[_row("C1", "P1", "RevUp", datetime(2026, 8, 1, 11, 0))],
        summary_values={},
    )

    summary_rows = compute_summary_rows([entry_zc00, entry_zm00], run_timestamp=datetime(2026, 8, 4))

    assert len(summary_rows) == 1
    pair_count = summary_rows[0][6]  # 差分ペア総数
    assert pair_count == 1  # モジュール違いでも(Child,Parent)は1ペアとして数える
