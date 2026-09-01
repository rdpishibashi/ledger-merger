"""
Bugfix regression (予防的対応): 2026-09, DXF-diff-manager が Summary シートの
「図面統計」欄（入力図面総数・差分抽出ペア数・流用率 [%]・完全新規図面数・
新規作成率 [%]）を全削除した仕様変更への追随。

不具合（対応しなかった場合に発生していたこと）: ledger_finder.SUMMARY_LABELS は
これら3項目（入力図面総数・差分抽出ペア数・流用率 [%]）を必須条件に含んでいた。
DXF-diff-manager が2026-09にこれらをSummaryシートから削除すると、以降
DXF-diff-manager が生成する台帳ファイルはすべて「有効な台帳が見つからない」と
黙って判定され（例外にもならず、単に None を返す）、Ledger-merger の統合対象から
静かに抜け落ちるところだった（2026-08の "Diff List"→"Master" シート名改名の際に
実際に発生しかけた失敗パターンと同種）。

修正: SUMMARY_LABELS（台帳判定の必須条件）からこの3項目を外し、
OPTIONAL_SUMMARY_LABELS（存在すれば取り込むが必須にしない）へ移した。

保証したいこと:
- 「図面統計」を持たない新形式の台帳（2026-09以降のDXF-diff-manager出力）も、
  引き続き有効な台帳として検出されること。
- 「図面統計」を持つ旧形式の台帳は、従来どおり検出され、値も取得できること
  （後方互換）。
- 新形式の台帳では、入力図面総数・差分抽出ペア数・流用率 [%] が
  summary_values に含まれないこと（app.py 側の警告表示・0扱いフォールバックの
  前提となる）。
"""
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "..", ".."))

from openpyxl import Workbook

from utils.ledger_finder import find_ledger_files, _try_load_ledger

_HEADERS = [
    "Child", "Parent", "Relation", "Title", "Subtitle", "Recorded Date", "Note",
    "Deleted Entities", "Added Entities", "Diff Entities", "Unchanged Entities",
    "Total Entities",
]

# 2026-09以降の新形式: エンティティ統計のみ（図面統計を持たない）
_NEW_FORMAT_SUMMARY_ROWS = [
    ("エンティティ統計", None),
    ("削除図形 総数", 1), ("追加図形 総数", 2), ("変更（追加+削除）図形 総数", 3),
    ("変更なし図形 総数", 4), ("流用先図面 図形総数", 5),
    ("図形変更率 [%]", 0.6),
]

# 旧形式: 図面統計を含む
_OLD_FORMAT_SUMMARY_ROWS = _NEW_FORMAT_SUMMARY_ROWS + [
    ("図面統計", None),
    ("流用先図面総数", 1), ("差分抽出ペア数", 1), ("完全新規図面数", 0),
    ("流用率 [%]", 1.0), ("新規作成率 [%]", 0.0),
]


def _build_ledger_file(path, summary_rows):
    path.parent.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    data_ws = wb.active
    data_ws.title = "Master"
    data_ws.append(_HEADERS)
    data_ws.append(["C1", "P1", "RevUp", "T", "S", None, None, 1, 2, 3, 4, 10])

    summary_ws = wb.create_sheet("Summary")
    for row in summary_rows:
        summary_ws.append(row)

    wb.save(path)


def test_new_format_without_drawing_statistics_is_detected(tmp_path):
    """図面統計を持たない新形式の台帳も、有効な台帳として検出される。"""
    folder = tmp_path / "dxf_diff_results_TypeA_NEWFORMAT_01"
    _build_ledger_file(folder / "NEWFORMAT.xlsx", _NEW_FORMAT_SUMMARY_ROWS)

    entries, missing = find_ledger_files(str(tmp_path))

    assert len(entries) == 1, f"新形式の台帳が検出されなかった: missing={missing}"
    assert missing == []


def test_new_format_summary_values_lack_drawing_statistics_keys(tmp_path):
    """新形式の台帳では、入力図面総数・差分抽出ペア数・流用率 [%] が
    summary_values に含まれない（呼び出し側の0扱いフォールバックの前提）。"""
    path = tmp_path / "NEWFORMAT.xlsx"
    _build_ledger_file(path, _NEW_FORMAT_SUMMARY_ROWS)

    entry = _try_load_ledger(str(path))

    assert entry is not None
    assert "入力図面総数" not in entry.summary_values
    assert "差分抽出ペア数" not in entry.summary_values
    assert "流用率 [%]" not in entry.summary_values
    # エンティティ統計は引き続き取得できる
    assert entry.summary_values["削除図形数 合計"] == 1
    assert entry.summary_values["図形変更率 [%]"] == 0.6


def test_old_format_with_drawing_statistics_still_detected(tmp_path):
    """図面統計を持つ旧形式の台帳は、従来どおり検出され値も取得できる（後方互換）。"""
    folder = tmp_path / "dxf_diff_results_TypeA_OLDFORMAT_01"
    _build_ledger_file(folder / "OLDFORMAT.xlsx", _OLD_FORMAT_SUMMARY_ROWS)

    entries, missing = find_ledger_files(str(tmp_path))

    assert len(entries) == 1
    assert missing == []
    entry = entries[0]
    assert entry.summary_values["入力図面総数"] == 1
    assert entry.summary_values["差分抽出ペア数"] == 1
    assert entry.summary_values["流用率 [%]"] == 1.0


def test_file_without_entity_statistics_still_rejected(tmp_path):
    """エンティティ統計（削除/追加/変更/変更なし/総数図形・図形変更率）を
    持たないファイルは、図面統計の有無に関わらず従来どおり無効。"""
    folder = tmp_path / "dxf_diff_results_TypeA_NOENTITY_01"
    _build_ledger_file(folder / "NOENTITY.xlsx", [("図面統計", None), ("流用先図面総数", 1)])

    entries, missing = find_ledger_files(str(tmp_path))

    assert entries == []
    assert missing == ["dxf_diff_results_TypeA_NOENTITY_01"]
