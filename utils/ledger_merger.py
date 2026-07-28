"""複数の台帳ファイル（LedgerEntry）の Diff List を1つの Excel に統合するモジュール。
Streamlit には依存しない。"""

import io

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font

from utils.ledger_finder import DIFF_LIST_HEADERS, SUMMARY_LABELS

FIRST_ROW_FONT = Font(color="FF000000")
OTHER_ROW_FONT = Font(color="FFA6A6A6")

CENTER_ALIGNMENT = Alignment(horizontal="center")

PERCENT_LABELS = {"図形変更率 [%]", "流用率 [%]"}
ENTITY_LABELS = ("Deleted Entities", "Added Entities", "Diff Entities", "Unchanged Entities", "Total Entities")

OUTPUT_HEADERS = ("Diff Package",) + DIFF_LIST_HEADERS + SUMMARY_LABELS
RECORDED_DATE_COL = OUTPUT_HEADERS.index("Recorded Date") + 1
ENTITY_COLS = [OUTPUT_HEADERS.index(label) + 1 for label in ENTITY_LABELS]


def build_merged_workbook(entries):
    """LedgerEntry のリストから統合済み Excel を生成し bytes で返す。"""
    wb = Workbook()
    ws = wb.active
    ws.title = "Diff List"

    ws.append(OUTPUT_HEADERS)
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = CENTER_ALIGNMENT
    ws.freeze_panes = "A2"

    row_idx = 2
    for entry in entries:
        for row_in_block, diff_row in enumerate(entry.diff_list_rows):
            row = [entry.package_name, *diff_row]
            if row_in_block == 0:
                row += [entry.summary_values[label] for label in SUMMARY_LABELS]
            else:
                row += [None] * len(SUMMARY_LABELS)
            ws.append(row)

            package_cell = ws.cell(row=row_idx, column=1)
            package_cell.font = FIRST_ROW_FONT if row_in_block == 0 else OTHER_ROW_FONT

            ws.cell(row=row_idx, column=RECORDED_DATE_COL).number_format = "YYYY-MM-DD HH:MM:SS"

            # "* Entities" 列（数値/'n/a' 混在）はカンマ区切り＋中央揃いで表示位置を揃える
            for col in ENTITY_COLS:
                cell = ws.cell(row=row_idx, column=col)
                cell.number_format = "#,##0"
                cell.alignment = CENTER_ALIGNMENT

            if row_in_block == 0:
                for offset, label in enumerate(SUMMARY_LABELS):
                    col = len(DIFF_LIST_HEADERS) + 2 + offset
                    cell = ws.cell(row=row_idx, column=col)
                    cell.number_format = "0.00%" if label in PERCENT_LABELS else "#,##0"
                    cell.alignment = CENTER_ALIGNMENT

            row_idx += 1

    for col_idx, header in enumerate(OUTPUT_HEADERS, start=1):
        width = max(len(str(header)) + 2, 12)
        ws.column_dimensions[ws.cell(row=1, column=col_idx).column_letter].width = width

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()
