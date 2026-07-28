"""DXF-diff-manager の出力フォルダ名（"指番_モジュール_サイド"単位）で LedgerEntry を
グルーピングし、レビジョン横断の集計Excel（"{指番}_{モジュール}_{サイド}_all.xlsx"）を
生成するモジュール。Streamlit には依存しない。

DXF-diff-manager のZIPダウンロードファイル名の命名規則
（dxf_diff_results_Type{A/B/C}_{指番}_{モジュール}_{サイド}_{リビジョン}）に依存する。
2026-07-28以前に手動で "dxf_diff_results_Pair{A/B/C}_..." と命名されたフォルダ
（DXF-diff-manager がファイル名自動生成に対応する前の実データ）も引き続き解釈できる
よう、"Pair"/"Type" どちらのトークンも受け付ける。
"""

import io
import re
from collections import defaultdict
from datetime import datetime

from openpyxl import Workbook
from openpyxl.styles import Font

from utils.ledger_finder import DIFF_LIST_HEADERS

GROUP_REVISION_PATTERN = re.compile(r'^dxf_diff_results_(?:Pair|Type)[A-Za-z]_(?P<group>.+)_(?P<revision>\d+)$')

_CHILD_COL = DIFF_LIST_HEADERS.index("Child")
_RECORDED_DATE_COL = DIFF_LIST_HEADERS.index("Recorded Date")
_ENTITY_COLS = ("Deleted Entities", "Added Entities", "Diff Entities", "Unchanged Entities", "Total Entities")

# Summaryシートの行構成（セクション見出し, 項目ラベル）。項目ラベルは DXF-diff-manager
# 自身のSummaryシートの表記（Type A: all_in_one）をそのまま使う。Type B/C
# （流用先図面総数/流用先図面 図形総数）の文言には対応しない（既知の制約）。
SUMMARY_ROWS = (
    ("エンティティ統計", "削除図形 総数"),
    (None, "追加図形 総数"),
    (None, "変更（追加+削除）図形 総数"),
    (None, "変更なし図形 総数"),
    (None, "アップロード図面 図形総数"),
    (None, "図形変更率 [%]"),
    ("図面統計", "アップロード図面総数"),
    (None, "差分抽出ペア数"),
    (None, "流用率 [%]"),
)

# Summaryシートの項目ラベル → LedgerEntry.summary_values のキー（utils.ledger_finder の
# SUMMARY_LABELS、Ledger-merger自身の統合Excel列名）への対応。
_CANONICAL_BY_DISPLAY_LABEL = {
    "削除図形 総数": "削除図形数 合計",
    "追加図形 総数": "追加図形数 合計",
    "変更（追加+削除）図形 総数": "差分図形数 合計",
    "変更なし図形 総数": "変更なし図形数 合計",
    "アップロード図面 図形総数": "総図形数 合計",
    "図形変更率 [%]": "図形変更率 [%]",
    "アップロード図面総数": "入力図面総数",
    "差分抽出ペア数": "差分抽出ペア数",
    "流用率 [%]": "流用率 [%]",
}

_COUNT_LABELS = (
    "削除図形 総数", "追加図形 総数", "変更（追加+削除）図形 総数", "変更なし図形 総数",
    "アップロード図面 図形総数", "アップロード図面総数", "差分抽出ペア数",
)
_PERCENT_LABELS = {"図形変更率 [%]", "流用率 [%]"}


def parse_group_and_revision(package_name):
    """DXF-diff-managerのZIPダウンロードファイル名の命名規則
    (dxf_diff_results_Type{A/B/C}_{指番}_{モジュール}_{サイド}_{リビジョン}、
    または旧手動命名の dxf_diff_results_Pair{A/B/C}_...) に従うフォルダ名から、
    グループキー（指番_モジュール_サイド）とレビジョン番号を取り出す。
    一致しない場合は None を返す（この機能の対象外として扱う）。
    """
    match = GROUP_REVISION_PATTERN.match(package_name)
    if not match:
        return None
    return match['group'], match['revision']


def _max_recorded_date(entry):
    dates = [row[_RECORDED_DATE_COL] for row in entry.diff_list_rows if row[_RECORDED_DATE_COL] is not None]
    return max(dates) if dates else datetime.min


def group_entries(entries):
    """LedgerEntry のリストを、指番_モジュール_サイど ごと・レビジョンごとにまとめる。

    同一フォルダ（同一 package_name）に複数の有効な台帳が見つかった場合は、Diff List
    内の最大 Recorded Date が最も新しいものだけを採用する（差分抽出のやり直しで古い
    実行結果がフォルダに残っていた場合の取り違え防止。2026-07-28 の実データで実際に
    確認したケース: 図番抽出に失敗した古い実行結果と、成功した新しい実行結果が
    同じフォルダに混在していた）。

    Returns:
        dict[group_key, list[(revision, LedgerEntry)]]（レビジョン文字列の昇順）
    """
    by_folder = defaultdict(list)
    for entry in entries:
        parsed = parse_group_and_revision(entry.package_name)
        if parsed is None:
            continue
        group_key, revision = parsed
        by_folder[(group_key, revision, entry.package_name)].append(entry)

    groups = defaultdict(list)
    for (group_key, revision, _package_name), folder_entries in by_folder.items():
        latest_entry = max(folder_entries, key=_max_recorded_date)
        groups[group_key].append((revision, latest_entry))

    for group_key in groups:
        groups[group_key].sort(key=lambda item: item[0])

    return dict(groups)


def aggregate_diff_list_by_child(revision_entries):
    """revision_entries（[(revision, LedgerEntry), ...]）の全 diff_list_rows を
    Child ごとに集計する。

    Deleted/Added/Diff/Unchanged/Total Entities の5列は全レビジョンにわたって
    単純合計する（'n/a' の値は数値でないためスキップし、ある Child の全出現が
    'n/a' の列はそのまま 'n/a' として残す）。非数値列（Parent/Relation/Title/
    Subtitle/Recorded Date/Note）は、Recorded Date が最も新しい行の値を採用する。

    Returns:
        list[tuple]（DIFF_LIST_HEADERS 12列、Childの昇順）
    """
    rows_by_child = defaultdict(list)
    for _revision, entry in revision_entries:
        for row in entry.diff_list_rows:
            rows_by_child[row[_CHILD_COL]].append(row)

    aggregated = []
    for child in sorted(rows_by_child.keys()):
        rows = rows_by_child[child]
        latest_row = max(rows, key=lambda r: r[_RECORDED_DATE_COL] or datetime.min)

        out_row = list(latest_row)
        for col_name in _ENTITY_COLS:
            col_idx = DIFF_LIST_HEADERS.index(col_name)
            numeric_values = [r[col_idx] for r in rows if isinstance(r[col_idx], (int, float))]
            out_row[col_idx] = sum(numeric_values) if numeric_values else 'n/a'
        aggregated.append(tuple(out_row))

    return aggregated


def _revision_value(entry, display_label):
    return entry.summary_values.get(_CANONICAL_BY_DISPLAY_LABEL[display_label])


def _total_value(revision_entries, display_label):
    if display_label == "図形変更率 [%]":
        total_diff = _total_value(revision_entries, "変更（追加+削除）図形 総数")
        total_entities = _total_value(revision_entries, "アップロード図面 図形総数")
        return (total_diff / total_entities) if total_entities else 0.0
    if display_label == "流用率 [%]":
        total_pairs = _total_value(revision_entries, "差分抽出ペア数")
        total_drawings = _total_value(revision_entries, "アップロード図面総数")
        return (total_pairs / total_drawings) if total_drawings else 0.0
    return sum(_revision_value(entry, display_label) for _revision, entry in revision_entries)


def build_group_workbook(group_key, revision_entries):
    """1つのグループ（指番_モジュール_サイド）について、レビジョン横断の
    Summary + Diff List を持つ統合Excel（bytes）を生成する。"""
    wb = Workbook()

    # --- Summary シート（先に追加してタブ順を先頭にする） ---
    summary_ws = wb.active
    summary_ws.title = "Summary"

    revisions = [revision for revision, _entry in revision_entries]
    summary_ws.append([None, None, "TOTAL"] + [f'"{r}"' for r in revisions])
    for cell in summary_ws[1]:
        cell.font = Font(bold=True)

    for section, label in SUMMARY_ROWS:
        total = _total_value(revision_entries, label)
        row = [section, label, total] + [_revision_value(entry, label) for _revision, entry in revision_entries]
        summary_ws.append(row)

        row_idx = summary_ws.max_row
        number_format = "0.00%" if label in _PERCENT_LABELS else "#,##0"
        for col_idx in range(3, 3 + 1 + len(revisions)):
            summary_ws.cell(row=row_idx, column=col_idx).number_format = number_format

    summary_ws.column_dimensions["A"].width = 18
    summary_ws.column_dimensions["B"].width = 24
    for col_idx in range(3, 3 + 1 + len(revisions)):
        summary_ws.column_dimensions[summary_ws.cell(row=1, column=col_idx).column_letter].width = 12

    # --- Diff List シート（"Diff Package"・Summary9項目列は含めない。Childごとに集計済み） ---
    diff_ws = wb.create_sheet("Diff List")
    diff_ws.append(DIFF_LIST_HEADERS)
    for cell in diff_ws[1]:
        cell.font = Font(bold=True)
    diff_ws.freeze_panes = "A2"

    recorded_date_col = DIFF_LIST_HEADERS.index("Recorded Date") + 1
    for row in aggregate_diff_list_by_child(revision_entries):
        diff_ws.append(row)
        diff_ws.cell(row=diff_ws.max_row, column=recorded_date_col).number_format = "YYYY-MM-DD HH:MM:SS"

    for col_idx, header in enumerate(DIFF_LIST_HEADERS, start=1):
        width = max(len(str(header)) + 2, 12)
        diff_ws.column_dimensions[diff_ws.cell(row=1, column=col_idx).column_letter].width = width

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()


def build_group_workbooks(entries):
    """entries から検出できる全グループについて、"{指番}_{モジュール}_{サイド}_all.xlsx"
    を生成する。

    Returns:
        dict[filename, bytes]（グループキーの昇順）
    """
    groups = group_entries(entries)
    return {
        f"{group_key}_all.xlsx": build_group_workbook(group_key, groups[group_key])
        for group_key in sorted(groups)
    }
