"""複数の台帳ファイル（LedgerEntry）の Diff List を1つの Excel に統合するモジュール。
Streamlit には依存しない。"""

import io

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font

from utils.group_summary_builder import parse_sashiban_module_side
from utils.ledger_finder import DIFF_LIST_HEADERS

FIRST_ROW_FONT = Font(color="FF000000")
OTHER_ROW_FONT = Font(color="FFA6A6A6")

CENTER_ALIGNMENT = Alignment(horizontal="center")

ENTITY_LABELS = ("Deleted Entities", "Added Entities", "Diff Entities", "Unchanged Entities", "Total Entities")

# LedgerEntry.summary_values の正規キー（utils.ledger_finder.SUMMARY_LABELS の一部）→
# 図形変更量詳細.xlsx に出力する表示列名。エンティティ統計5項目のみに絞り、図面統計系
# （入力図面総数・差分抽出ペア数・流用率 [%]）は含めない（2026-08、ユーザー要望により
# 図面統計は指番_モジュール_サイド別集計・統合図面管理台帳側の指標に一本化し、こちらの
# 変更量詳細は図形の変更内容に絞った。併せて「差分図形数 合計」→「変更図形数 合計」、
# 「総図形数 合計」→「図形総数 合計」に表示名を変更）。
_MERGED_SUMMARY_COLUMNS = (
    ("削除図形数 合計", "削除図形数 合計"),
    ("追加図形数 合計", "追加図形数 合計"),
    ("差分図形数 合計", "変更図形数 合計"),
    ("総図形数 合計", "図形総数 合計"),
    ("図形変更率 [%]", "図形変更率 [%]"),
)
_MERGED_SUMMARY_CANONICAL_KEYS = tuple(canonical for canonical, _display in _MERGED_SUMMARY_COLUMNS)
_MERGED_SUMMARY_DISPLAY_LABELS = tuple(display for _canonical, display in _MERGED_SUMMARY_COLUMNS)

PERCENT_LABELS = {"図形変更率 [%]"}

# "Sashiban"/"Module"/"Side" は台帳ファイル名を主・Diff Package（DXF-diff-manager
# 出力フォルダ名）を従として parse_sashiban_module_side() で逆算した指番/モジュール/
# サイド。"Diff Package" 自体は
# 参照情報として最終列に残す（2026-07-31、ユーザー要望によりChildの前に指番系3列を
# 追加し、Diff Packageを先頭から最終列へ移動）。
OUTPUT_HEADERS = ("Sashiban", "Module", "Side") + DIFF_LIST_HEADERS + _MERGED_SUMMARY_DISPLAY_LABELS + ("Diff Package",)
RECORDED_DATE_COL = OUTPUT_HEADERS.index("Recorded Date") + 1
ENTITY_COLS = [OUTPUT_HEADERS.index(label) + 1 for label in ENTITY_LABELS]
DIFF_PACKAGE_COL = OUTPUT_HEADERS.index("Diff Package") + 1
_SUMMARY_COLS = [OUTPUT_HEADERS.index(label) + 1 for label in _MERGED_SUMMARY_DISPLAY_LABELS]


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
        sashiban, module, side = parse_sashiban_module_side(entry.package_name, entry.source_path)
        for row_in_block, diff_row in enumerate(entry.diff_list_rows):
            row = [sashiban, module, side, *diff_row]
            if row_in_block == 0:
                row += [entry.summary_values[canonical] for canonical in _MERGED_SUMMARY_CANONICAL_KEYS]
            else:
                row += [None] * len(_MERGED_SUMMARY_CANONICAL_KEYS)
            row.append(entry.package_name)
            ws.append(row)

            package_cell = ws.cell(row=row_idx, column=DIFF_PACKAGE_COL)
            package_cell.font = FIRST_ROW_FONT if row_in_block == 0 else OTHER_ROW_FONT

            ws.cell(row=row_idx, column=RECORDED_DATE_COL).number_format = "YYYY-MM-DD HH:MM:SS"

            # "* Entities" 列（数値/'n/a' 混在）はカンマ区切り＋中央揃いで表示位置を揃える
            for col in ENTITY_COLS:
                cell = ws.cell(row=row_idx, column=col)
                cell.number_format = "#,##0"
                cell.alignment = CENTER_ALIGNMENT

            if row_in_block == 0:
                for col, label in zip(_SUMMARY_COLS, _MERGED_SUMMARY_DISPLAY_LABELS):
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
