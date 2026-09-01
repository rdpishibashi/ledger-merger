"""
Bugfix regression: 2026-09, アップロードされた前回の統合図面管理台帳.xlsxに
Sashiban/Module/Side/Child のいずれかが「空欄のセル」または「文字列ではなく
数値として保存されたセル」（例: サイド "405" が数値として保存）を持つ行が
混在していると、Master・Work Masterシート生成時のソートでクラッシュしていた
（ユーザー実機で2回発生・報告。1回目: 空欄セル→None、2回目: 数値セル→int）。

不具合: build_master_workbook() 内の _write_ledger_sheet() 呼び出しに渡す
sort_key が、キーの各要素をそのまま比較していた。今回分のエントリは
parse_sashiban_module_side() が返す値が必ず文字列（未入力時も "na"）だが、
アップロードされた前回台帳（read_master_rows()/read_work_master_rows() が
openpyxl でセルの生値をそのまま読む）に空欄セルや数値セルがあると、
そのキー要素が None や int/float になる。今回分の文字列キーと型の異なる
前回分のキーが混在した状態で sorted() が比較を試みると、Python 3 では
型の異なる値同士の大小比較ができずクラッシュする:
    1回目: TypeError: '<' not supported between instances of 'str' and 'NoneType'
    2回目: TypeError: '<' not supported between instances of 'str' and 'int'

修正の変遷: 当初 `k[0] or ''` で None のみ防御したが、int混在の再発報告を
受けて不十分と判明。`_sort_str()`（None→''、それ以外は str() で文字列化）に
一本化し、型に関わらず必ず文字列同士の比較になるようにした（Master・
Work Master 両方）。

保証したいこと:
- Master・Work Master いずれも、前回分に None や int/float を含むキーが
  混在していてもクラッシュせず生成できること。
- 型に関わらず文字列として比較され、他の値との比較で例外にならないこと
  （並び順自体の厳密な妥当性は問わない——異常データの並び順まで厳密に保証する
  必要はなく、「クラッシュしない」ことが本テストの主眼）。
"""
import io
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "..", ".."))

import openpyxl

from utils.master_ledger_builder import build_master_workbook, MASTER_SHEET_NAME, WORK_MASTER_SHEET_NAME


def test_work_master_sort_survives_none_in_previous_rows():
    """previous_work_master_rows に None を含むキーが混在してもクラッシュしない。"""
    previous_work_master_rows = {
        ("AA11-1111-1", None, "405", "C1", "P1"):
            ("AA11-1111-1", None, "405", "C1", "P1", "T", "S", "A", 1, 2, 3, 4, 10, None, None),
        ("AA11-1111-1", "ZM00", "405", "C2", "P2"):
            ("AA11-1111-1", "ZM00", "405", "C2", "P2", "T", "S", "A", 1, 2, 3, 4, 10, None, None),
    }

    data = build_master_workbook([], previous_work_master_rows=previous_work_master_rows)

    wb = openpyxl.load_workbook(io.BytesIO(data))
    ws = wb[WORK_MASTER_SHEET_NAME]
    rows = list(ws.iter_rows(values_only=True))
    assert len(rows) == 3  # ヘッダー + 2行
    children = {row[3] for row in rows[1:]}
    assert children == {"C1", "C2"}


def test_work_master_sort_survives_int_side_mixed_with_str_side():
    """previous_work_master_rows に、Sideが数値(int)で保存された行と文字列で
    保存された行が混在していてもクラッシュしない（2026-09 再発報告のケース）。
    """
    previous_work_master_rows = {
        ("AA11-1111-1", "ZM00", 405, "C1", "P1"):
            ("AA11-1111-1", "ZM00", 405, "C1", "P1", "T", "S", "A", 1, 2, 3, 4, 10, None, None),
        ("AA11-1111-1", "ZM00", "405", "C2", "P2"):
            ("AA11-1111-1", "ZM00", "405", "C2", "P2", "T", "S", "A", 1, 2, 3, 4, 10, None, None),
    }

    data = build_master_workbook([], previous_work_master_rows=previous_work_master_rows)

    wb = openpyxl.load_workbook(io.BytesIO(data))
    ws = wb[WORK_MASTER_SHEET_NAME]
    rows = list(ws.iter_rows(values_only=True))
    assert len(rows) == 3  # ヘッダー + 2行
    children = {row[3] for row in rows[1:]}
    assert children == {"C1", "C2"}


def test_master_sort_survives_none_in_previous_rows():
    """previous_master_rows に Child が None のキーが混在してもクラッシュしない。"""
    previous_master_rows = {
        (None, "P1"): (None, "P1", "流用", "T", "S", "A", 1, 2, 3, 4, 10, None, None),
        ("C2", "P2"): ("C2", "P2", "流用", "T", "S", "A", 1, 2, 3, 4, 10, None, None),
    }

    data = build_master_workbook([], previous_master_rows=previous_master_rows)

    wb = openpyxl.load_workbook(io.BytesIO(data))
    ws = wb[MASTER_SHEET_NAME]
    rows = list(ws.iter_rows(values_only=True))
    assert len(rows) == 3  # ヘッダー + 2行


def test_master_sort_survives_int_child_mixed_with_str_child():
    """previous_master_rows に、Childが数値(int)のキーと文字列のキーが
    混在していてもクラッシュしない。"""
    previous_master_rows = {
        (123, "P1"): (123, "P1", "流用", "T", "S", "A", 1, 2, 3, 4, 10, None, None),
        ("C2", "P2"): ("C2", "P2", "流用", "T", "S", "A", 1, 2, 3, 4, 10, None, None),
    }

    data = build_master_workbook([], previous_master_rows=previous_master_rows)

    wb = openpyxl.load_workbook(io.BytesIO(data))
    ws = wb[MASTER_SHEET_NAME]
    rows = list(ws.iter_rows(values_only=True))
    assert len(rows) == 3  # ヘッダー + 2行


def test_both_sheets_survive_combined_none_int_and_string_keys():
    """MasterとWork Master双方に前回分のNone・int混在キーがあっても
    同時に処理してクラッシュしない。"""
    previous_master_rows = {
        (None, None): (None, None, None, None, None, None, None, None, None, None, None, None, None),
        (123, 456): (123, 456, None, None, None, None, None, None, None, None, None, None, None),
    }
    previous_work_master_rows = {
        (None, None, None, None, None):
            (None, None, None, None, None, None, None, None, None, None, None, None, None, None, None),
        ("AA11-1111-1", "ZM00", 405, 789, "P9"):
            ("AA11-1111-1", "ZM00", 405, 789, "P9", None, None, None, None, None, None, None, None, None, None),
    }

    data = build_master_workbook(
        [], previous_master_rows=previous_master_rows,
        previous_work_master_rows=previous_work_master_rows,
    )
    assert isinstance(data, bytes) and len(data) > 0
