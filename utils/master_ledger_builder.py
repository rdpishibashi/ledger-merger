"""LedgerEntry の Diff List 行を "Child"-"Parent" ペア単位でユニーク化し、
"統合図面管理台帳.xlsx"（"Master" シートのみ）を生成・蓄積するモジュール。
Streamlit には依存しない。

前回ダウンロードした "統合図面管理台帳.xlsx" を今回の実行時にアップロードすると、
今回新たに得られた Child-Parent ペアのデータで、同じペアの既存行を上書きしつつ
（新しいペアは追加）マージする。"""

import io

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font

from utils.ledger_finder import DIFF_LIST_HEADERS

MASTER_SHEET_NAME = "Master"

CENTER_ALIGNMENT = Alignment(horizontal="center")
_ENTITY_LABELS = ("Deleted Entities", "Added Entities", "Diff Entities", "Unchanged Entities", "Total Entities")
_CHILD_COL = DIFF_LIST_HEADERS.index("Child")
_PARENT_COL = DIFF_LIST_HEADERS.index("Parent")
_RECORDED_DATE_COL = DIFF_LIST_HEADERS.index("Recorded Date")


def extract_unique_child_parent_rows(entries):
    """LedgerEntry のリストから、"Child"-"Parent" ペアでユニーク化した
    DIFF_LIST_HEADERS 12列のデータを返す。同じペアが複数エントリにまたがる場合は
    どれか1件を採用する（データはどれも同じはずのため）。

    Returns:
        dict[(child, parent), tuple]
    """
    unique = {}
    for entry in entries:
        for row in entry.diff_list_rows:
            key = (row[_CHILD_COL], row[_PARENT_COL])
            unique.setdefault(key, row)
    return unique


def read_master_rows(file_bytes):
    """アップロードされた統合図面管理台帳.xlsx（Masterシートのみ）から
    (child, parent) をキーとする行の辞書を読み込む。シート構成が想定と異なる
    （壊れている、別ファイル等）場合は None を返す。
    """
    try:
        wb = load_workbook(io.BytesIO(file_bytes), data_only=True, read_only=True)
    except Exception:
        return None

    try:
        if MASTER_SHEET_NAME not in wb.sheetnames:
            return None
        ws = wb[MASTER_SHEET_NAME]
        rows = list(ws.iter_rows(values_only=True))
        if not rows or tuple(rows[0]) != DIFF_LIST_HEADERS:
            return None
        return {(row[_CHILD_COL], row[_PARENT_COL]): tuple(row) for row in rows[1:]}
    finally:
        wb.close()


def build_master_workbook(entries, previous_master_rows=None):
    """今回のDiff Listデータ（Child-Parentユニーク化）と、アップロードされた前回の
    Masterシート内容（previous_master_rows、無ければ None）をマージし（同じ
    Child-Parentは今回のデータで上書き、無ければ追加）、Child昇順でソートした
    "統合図面管理台帳.xlsx"（Masterシートのみ）を bytes で返す。
    """
    current = extract_unique_child_parent_rows(entries)
    combined = dict(previous_master_rows or {})
    combined.update(current)

    wb = Workbook()
    ws = wb.active
    ws.title = MASTER_SHEET_NAME
    ws.append(DIFF_LIST_HEADERS)
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = CENTER_ALIGNMENT
    ws.freeze_panes = "A2"

    recorded_date_col = _RECORDED_DATE_COL + 1
    entity_cols = [DIFF_LIST_HEADERS.index(label) + 1 for label in _ENTITY_LABELS]

    for key in sorted(combined.keys(), key=lambda k: k[0]):
        ws.append(combined[key])
        row_idx = ws.max_row
        ws.cell(row=row_idx, column=recorded_date_col).number_format = "YYYY-MM-DD HH:MM:SS"
        for col in entity_cols:
            cell = ws.cell(row=row_idx, column=col)
            cell.number_format = "#,##0"
            cell.alignment = CENTER_ALIGNMENT

    for col_idx, header in enumerate(DIFF_LIST_HEADERS, start=1):
        width = max(len(str(header)) + 2, 12)
        ws.column_dimensions[ws.cell(row=1, column=col_idx).column_letter].width = width

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()
