"""
Bugfix regression: 2026-08-30, ledger detection hardcoded to sheet name "Diff List"

不具合（実際に発生する前に発見・修正）: ledger_finder._try_load_ledger() はデータ
シートをシート名 "Diff List" の完全一致でのみ探していた。DXF-diff-manager は
2026-08 に出力シート名を "Diff List" から "Master" に改名する予定で、そのまま
リリースされると、以降 DXF-diff-manager が生成する台帳ファイルはすべて
「有効な台帳が見つからない」と黙って判定され（例外にもならず、単に None を返す）、
Ledger-merger の統合対象から静かに抜け落ちるところだった。

修正: ledger_finder._find_diff_list_rows() を新設し、シート名ではなくヘッダー行が
DIFF_LIST_HEADERS（12列）と完全一致するシートを名前を問わず探す方式に変更した
（DXF-diff-manager 自身の load_parent_child_master() と同じ考え方）。

保証したいこと:
- データシート名が "Diff List"（旧形式）でも "Master"（2026-08改名後の新形式）でも、
  どちらも台帳として検出されること
- シート名が別の任意の文字列でも、ヘッダー構成さえ一致すれば検出されること
- Summary シートが無い、またはヘッダー構成が一致するシートが無いファイルは、
  従来通り無効な台帳として扱われること
"""
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "..", ".."))

from openpyxl import Workbook

from utils.ledger_finder import find_ledger_files

_HEADERS = [
    "Child", "Parent", "Relation", "Title", "Subtitle", "Recorded Date", "Note",
    "Deleted Entities", "Added Entities", "Diff Entities", "Unchanged Entities",
    "Total Entities",
]

_SUMMARY_ROWS = [
    ("削除図形 総数", 1), ("追加図形 総数", 2), ("変更（追加+削除）図形 総数", 3),
    ("変更なし図形 総数", 4), ("アップロード図面 図形総数", 5),
    ("図形変更率 [%]", 0.6), ("アップロード図面総数", 1),
    ("差分抽出ペア数", 1), ("流用率 [%]", 1.0),
]


def _build_ledger_file(path, data_sheet_name):
    path.parent.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    data_ws = wb.active
    data_ws.title = data_sheet_name
    data_ws.append(_HEADERS)
    data_ws.append(["C1", "P1", "RevUp", "T", "S", None, None, 1, 2, 3, 4, 10])

    summary_ws = wb.create_sheet("Summary")
    for row in _SUMMARY_ROWS:
        summary_ws.append(row)

    wb.save(path)


def test_detects_ledger_with_legacy_diff_list_sheet_name(tmp_path):
    folder = tmp_path / "dxf_diff_results_TypeA_LEGACY_01"
    _build_ledger_file(folder / "LEGACY.xlsx", "Diff List")

    entries, missing = find_ledger_files(str(tmp_path))

    assert len(entries) == 1
    assert missing == []


def test_detects_ledger_with_renamed_master_sheet_name(tmp_path):
    """2026-08改名後、DXF-diff-managerの出力データシート名は 'Master' になる。"""
    folder = tmp_path / "dxf_diff_results_TypeA_RENAMED_01"
    _build_ledger_file(folder / "RENAMED.xlsx", "Master")

    entries, missing = find_ledger_files(str(tmp_path))

    assert len(entries) == 1
    assert missing == []


def test_detects_ledger_with_arbitrary_data_sheet_name(tmp_path):
    """シート名そのものには依存しない（列構成のみで判定する）ことの確認。"""
    folder = tmp_path / "dxf_diff_results_TypeA_ARBITRARY_01"
    _build_ledger_file(folder / "ARBITRARY.xlsx", "SomeOtherName")

    entries, missing = find_ledger_files(str(tmp_path))

    assert len(entries) == 1
    assert missing == []


def test_file_without_summary_sheet_still_rejected(tmp_path):
    """Summaryシートが無いファイルは、データシート名に関わらず従来通り無効。"""
    folder = tmp_path / "dxf_diff_results_TypeA_NOSUMMARY_01"
    path = folder / "NOSUMMARY.xlsx"
    path.parent.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = "Master"
    ws.append(_HEADERS)
    ws.append(["C1", "P1", "RevUp", "T", "S", None, None, 1, 2, 3, 4, 10])
    wb.save(path)

    entries, missing = find_ledger_files(str(tmp_path))

    assert entries == []
    assert missing == ["dxf_diff_results_TypeA_NOSUMMARY_01"]


def test_file_with_no_matching_header_sheet_still_rejected(tmp_path):
    """どのシートのヘッダーもDIFF_LIST_HEADERSと一致しないファイルは、
    従来通り無効な台帳として扱われる。"""
    folder = tmp_path / "dxf_diff_results_TypeA_NOMATCH_01"
    path = folder / "NOMATCH.xlsx"
    path.parent.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = "Master"
    ws.append(["Foo", "Bar"])
    ws.append([1, 2])
    summary_ws = wb.create_sheet("Summary")
    for row in _SUMMARY_ROWS:
        summary_ws.append(row)
    wb.save(path)

    entries, missing = find_ledger_files(str(tmp_path))

    assert entries == []
    assert missing == ["dxf_diff_results_TypeA_NOMATCH_01"]
