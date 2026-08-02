"""仕様確認（Spec Regression）: 指番_モジュール_サイド単位のレビジョン横断集計Excel
（"{指番}_{モジュール}_{サイド}_all.xlsx"）の生成。

対応する受入条件（2026-07-28 のユーザー依頼、添付レイアウト参照ファイル
ME24-1001-0_ZC00_405_all.xlsx との突き合わせで確認）:
    - Summaryシート: TOTAL列 + レビジョン列（"01","02",...）。カウント系9項目のうち
      7項目はレビジョンごとの値をそのまま横に並べ、TOTAL列は単純合計。
      図形変更率[%]・流用率[%]の2項目のみ、TOTAL列は単純合計ではなく
      TOTAL(分子)/TOTAL(分母)で再計算する。
    - Diff Listシート: Diff Package列・9項目合計列は含めない。同一Childが複数
      レビジョンにまたがる場合、Deleted/Added/Diff/Unchanged/Total Entitiesの
      5列は単純合計する。
    - 同一出力フォルダに複数の有効な台帳がある場合（実データで確認済みのケース）、
      Recorded Dateが最新のものだけを採用する。

このテストは、ユーザー提供の実際の参照ファイル（ME24-1001-0_ZC00_405_all.xlsx、
2026-07-28受領）のSummaryシートの値と完全一致することを検証する
（tests/fixtures/dxf_diff_manager_output/dxf_diff_results_PairA_ME24-1001-0_ZC00_405_01〜04
が、その参照ファイルの生成元となった実データそのもの）。

2026-08、DXF-diff-manager Summaryシートに追加された「完全新規図面数」「新規作成率 [%]」
（差分抽出ペア数の直下・流用率[%]の直下）の2行を EXPECTED_SUMMARY_ROWS に追加した。
上記の参照ファイル（この2指標が追加される前のもの）にこの2行は存在しないため、
このフィクスチャ（旧形式の台帳）から算出される値は全レビジョン・TOTALとも0になる
（utils.ledger_finder.OPTIONAL_SUMMARY_LABELS が無い旧形式台帳の既定値。参照ファイル
自体の値は変更していない）。
"""

import io
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "..", ".."))

import openpyxl
import pytest

from utils.group_summary_builder import build_group_workbooks
from utils.ledger_finder import find_ledger_files

REAL_DATA_ROOT = os.path.join(
    os.path.dirname(__file__), "..", "..", "fixtures", "dxf_diff_manager_output"
)

# ユーザー提供の参照ファイル ME24-1001-0_ZC00_405_all.xlsx（2026-07-28）の
# Summaryシートの値をそのまま書き写したもの。
EXPECTED_SUMMARY_ROWS = [
    (None, None, "TOTAL", '"01"', '"02"', '"03"', '"04"'),
    ("エンティティ統計", "削除図形 総数", 518, 222, 233, 10, 53),
    (None, "追加図形 総数", 11652, 2867, 3072, 2835, 2878),
    (None, "変更（追加+削除）図形 総数", 902, 272, 488, 28, 114),
    (None, "変更なし図形 総数", 67705, 31813, 4216, 20708, 10968),
    (None, "アップロード図面 図形総数", 79875, 34902, 7521, 23553, 13899),
    (None, "図形変更率 [%]", 0.01129264475743349, 0.007793249670505988, 0.0648849886983114,
     0.0011888082197596909, 0.008202028922944096),
    ("図面統計", "アップロード図面総数", 113, 31, 29, 27, 26),
    (None, "差分抽出ペア数", 9, 4, 2, 2, 1),
    (None, "完全新規図面数", 0, 0, 0, 0, 0),  # 旧形式フィクスチャは新指標を持たないため0
    (None, "流用率 [%]", 0.07964601769911504, 0.12903225806451613, 0.06896551724137931,
     0.07407407407407407, 0.038461538461538464),
    (None, "新規作成率 [%]", 0.0, 0.0, 0.0, 0.0, 0.0),  # 同上
]


def test_group_summary_matches_user_provided_reference_file():
    entries, _missing = find_ledger_files(REAL_DATA_ROOT)
    files = build_group_workbooks(entries)

    wb = openpyxl.load_workbook(io.BytesIO(files["ME24-1001-0_ZC00_405_all.xlsx"]), data_only=True)
    ws = wb["Summary"]

    actual_rows = [tuple(c.value for c in row) for row in ws.iter_rows(min_row=1, max_row=len(EXPECTED_SUMMARY_ROWS))]

    for actual, expected in zip(actual_rows, EXPECTED_SUMMARY_ROWS):
        assert len(actual) == len(expected)
        for a, e in zip(actual, expected):
            if isinstance(e, float):
                assert a == pytest.approx(e), f"{expected[1]}: {a} != {e}"
            else:
                assert a == e, f"{expected[1]}: {a} != {e}"


def test_group_summary_diff_list_excludes_diff_package_and_total_columns():
    """"Diff Package" 列・合計欄（9項目）は Diff List シートに含まれない。"""
    entries, _missing = find_ledger_files(REAL_DATA_ROOT)
    files = build_group_workbooks(entries)

    wb = openpyxl.load_workbook(io.BytesIO(files["ME24-1001-0_ZC00_405_all.xlsx"]))
    header = [c.value for c in wb["Diff List"][1]]

    assert "Diff Package" not in header
    assert len(header) == 12
