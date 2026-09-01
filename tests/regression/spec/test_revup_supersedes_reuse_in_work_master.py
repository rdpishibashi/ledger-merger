"""仕様確認（Spec Regression）: Work Master で「RevUp」が「流用」に優先する。

対応する受入条件（2026-09-01 のユーザー依頼）:
    「Work Master に『流用』と『RevUp』の違いがある同じ図面のデータがある。
      Drawing-genealogy プロジェクトと同様に、『RevUp』『流用』の2つの関係が
      ある場合は『RevUp』を採用して『流用』は不採用（削除）してほしい」

    Drawing-genealogy 側の対応実装:
    `utils/graph_builder.py::GraphBuilder._reuse_pairs_to_delete()`
    （ある Child が RevUp 種別の入力エッジを持つ場合、その Child への 流用 エッジを
      すべて落とす）。

境界・例外の意図:
    - 判定単位は **(Sashiban, Module, Side, Child)**。指番をまたいで判定しない——
      実データ検証（2026-09-01、Archive.zip + 統合図面管理台帳_old.xlsx）で、
      指番 PE25-9601-0 の10行が「別指番 NE24-0062-0 に RevUp がある」ことを理由に
      削除され、PE25-9601-0 側ではその図面の流用元が一切たどれなくなることを確認した。
      Work Master は指番ごとの作業台帳のため、その指番の中に RevUp が実在する場合
      のみ流用を落とす（実データでの削除件数: 指番横断=23行、指番内=13行）。
    - 完全新規図面（Parent="none"）の行は削除対象にしない。規則の対象は流用のみで、
      Drawing-genealogy も Relation が '流用' の行だけを落とす。
    - Master シートには適用しない。Master は Relation 列を持つ関係の記録そのもの
      であり、かつこの判定に使う Relation の参照元でもあるため。
    - Work Master は Relation 列を持たないため、判定にはマージ済み Master の
      Relation を (Child, Parent) で引く。これにより、前回台帳から引き継いだ
      （＝この規則が無かった頃に書かれた）流用の行も、次回の統合時に間引かれる。
"""

import io
import os
import sys
from datetime import datetime

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "..", ".."))

import openpyxl

from utils.ledger_finder import LedgerEntry
from utils.master_ledger_builder import (
    MASTER_SHEET_NAME,
    SUMMARY_HEADERS,
    SUMMARY_SHEET_NAME,
    WORK_MASTER_SHEET_NAME,
    build_master_workbook,
    read_master_rows,
    read_work_master_rows,
)

_PACKAGE = "dxf_diff_results_TypeA_NE24-0062-0_ZM00_405"
_OTHER_SASHIBAN_PACKAGE = "dxf_diff_results_TypeA_PE25-9601-0_ZM00_405"


def _row(child, parent, relation, recorded_date, deleted=1, added=2, diff=3, unchanged=4, total=5):
    return (child, parent, relation, "T", "S", recorded_date, None, deleted, added, diff, unchanged, total)


def _work_master_keys(wb_bytes):
    wb = openpyxl.load_workbook(io.BytesIO(wb_bytes))
    return [tuple(row[:5]) for row in wb[WORK_MASTER_SHEET_NAME].iter_rows(min_row=2, values_only=True)]


def test_reuse_row_dropped_when_same_child_also_has_revup():
    """同一 (指番,モジュール,サイド,Child) に RevUp と 流用 の両方があれば、
    流用の行だけが落ちる。"""
    entry = LedgerEntry(
        package_name=_PACKAGE, source_path="NE24-0062-0_ZM00_405.xlsx",
        diff_list_rows=[
            _row("EE3273-608-32B", "EE3273-608-32A", "RevUp", datetime(2026, 8, 1)),
            _row("EE3273-608-32B", "EE3273-608-24B", "流用", datetime(2026, 8, 1)),
        ],
        summary_values={},
    )

    keys = _work_master_keys(build_master_workbook([entry]))

    assert keys == [("NE24-0062-0", "ZM00", "405", "EE3273-608-32B", "EE3273-608-32A")]


def test_reuse_row_kept_when_child_has_no_revup():
    """RevUp が無ければ流用の行はそのまま残る（唯一の関係を消さない）。"""
    entry = LedgerEntry(
        package_name=_PACKAGE, source_path="NE24-0062-0_ZM00_405.xlsx",
        diff_list_rows=[_row("EE3273-608-32B", "EE3273-608-24B", "流用", datetime(2026, 8, 1))],
        summary_values={},
    )

    keys = _work_master_keys(build_master_workbook([entry]))

    assert keys == [("NE24-0062-0", "ZM00", "405", "EE3273-608-32B", "EE3273-608-24B")]


def test_reuse_row_kept_when_revup_belongs_to_another_sashiban():
    """別の指番に RevUp があっても、その指番の流用の行は落とさない
    （指番をまたいで削除すると、その指番で流用元がたどれなくなる）。"""
    revup_entry = LedgerEntry(
        package_name=_PACKAGE, source_path="NE24-0062-0_ZM00_405.xlsx",
        diff_list_rows=[_row("EE3273-608-32B", "EE3273-608-32A", "RevUp", datetime(2026, 8, 1))],
        summary_values={},
    )
    reuse_entry = LedgerEntry(
        package_name=_OTHER_SASHIBAN_PACKAGE, source_path="PE25-9601-0_ZM00_405.xlsx",
        diff_list_rows=[_row("EE3273-608-32B", "EE3273-608-24B", "流用", datetime(2026, 8, 1))],
        summary_values={},
    )

    keys = _work_master_keys(build_master_workbook([revup_entry, reuse_entry]))

    assert ("PE25-9601-0", "ZM00", "405", "EE3273-608-32B", "EE3273-608-24B") in keys
    assert ("NE24-0062-0", "ZM00", "405", "EE3273-608-32B", "EE3273-608-32A") in keys
    assert len(keys) == 2


def test_brand_new_row_is_not_dropped_by_revup():
    """完全新規図面（Parent="none"）は削除対象にしない（対象は流用のみ）。"""
    entry = LedgerEntry(
        package_name=_PACKAGE, source_path="NE24-0062-0_ZM00_405.xlsx",
        diff_list_rows=[
            _row("EE5526-405-16B", "EE5526-405-16A", "RevUp", datetime(2026, 8, 1)),
            _row("EE5526-405-16B", "none", "完全新規図面", datetime(2026, 8, 1)),
        ],
        summary_values={},
    )

    keys = _work_master_keys(build_master_workbook([entry]))

    assert set(keys) == {
        ("NE24-0062-0", "ZM00", "405", "EE5526-405-16B", "EE5526-405-16A"),
        ("NE24-0062-0", "ZM00", "405", "EE5526-405-16B", "none"),
    }


def test_master_sheet_keeps_both_relations():
    """Master は関係の記録そのものなので、流用の行も残す（Work Masterのみ間引く）。"""
    entry = LedgerEntry(
        package_name=_PACKAGE, source_path="NE24-0062-0_ZM00_405.xlsx",
        diff_list_rows=[
            _row("EE3273-608-32B", "EE3273-608-32A", "RevUp", datetime(2026, 8, 1)),
            _row("EE3273-608-32B", "EE3273-608-24B", "流用", datetime(2026, 8, 1)),
        ],
        summary_values={},
    )

    wb = openpyxl.load_workbook(io.BytesIO(build_master_workbook([entry])))
    master_keys = [(r[0], r[1]) for r in wb[MASTER_SHEET_NAME].iter_rows(min_row=2, values_only=True)]

    assert set(master_keys) == {
        ("EE3273-608-32B", "EE3273-608-32A"),
        ("EE3273-608-32B", "EE3273-608-24B"),
    }


def test_legacy_reuse_row_from_previous_ledger_is_dropped_on_next_merge():
    """この規則が無かった頃の台帳から引き継いだ流用の行も、次回の統合で間引かれる
    （判定はマージ済みMasterのRelationを引くため、前回分にも遡って効く）。"""
    legacy_entry = LedgerEntry(
        package_name=_PACKAGE, source_path="NE24-0062-0_ZM00_405.xlsx",
        diff_list_rows=[_row("EE3273-608-32B", "EE3273-608-24B", "流用", datetime(2026, 8, 1))],
        summary_values={},
    )
    legacy_bytes = build_master_workbook([legacy_entry])
    assert len(_work_master_keys(legacy_bytes)) == 1  # 前回時点では流用のみで残っている

    revup_entry = LedgerEntry(
        package_name=_PACKAGE, source_path="NE24-0062-0_ZM00_405.xlsx",
        diff_list_rows=[_row("EE3273-608-32B", "EE3273-608-32A", "RevUp", datetime(2026, 9, 1))],
        summary_values={},
    )

    merged_bytes = build_master_workbook(
        [revup_entry],
        previous_master_rows=read_master_rows(legacy_bytes),
        previous_work_master_rows=read_work_master_rows(legacy_bytes),
    )

    assert _work_master_keys(merged_bytes) == [
        ("NE24-0062-0", "ZM00", "405", "EE3273-608-32B", "EE3273-608-32A"),
    ]


def test_summary_counts_exclude_dropped_reuse_rows():
    """Summaryは間引き後のWork Masterから集計する（表示と集計値を一致させる）。"""
    entry = LedgerEntry(
        package_name=_PACKAGE, source_path="NE24-0062-0_ZM00_405.xlsx",
        diff_list_rows=[
            _row("EE3273-608-32B", "EE3273-608-32A", "RevUp", datetime(2026, 8, 1),
                 deleted=10, added=20, total=100),
            _row("EE3273-608-32B", "EE3273-608-24B", "流用", datetime(2026, 8, 1),
                 deleted=7, added=7, total=70),
        ],
        summary_values={},
    )

    wb = openpyxl.load_workbook(io.BytesIO(build_master_workbook([entry])))
    rows = list(wb[SUMMARY_SHEET_NAME].iter_rows(min_row=2, values_only=True))

    assert len(rows) == 1
    summary = dict(zip(SUMMARY_HEADERS, rows[0]))
    assert summary["変更図面総数"] == 1  # 流用の行を含めれば2になる
    assert summary["削除図形総数"] == 10  # 流用の行(7)は加算されない
    assert summary["図形総数"] == 100
