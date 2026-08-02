"""LedgerEntry の Diff List 行をユニーク化し、"統合図面管理台帳.xlsx"
（"Master"→"Work Master"→"Summary" の3シート）を生成・蓄積するモジュール。
Streamlit には依存しない。

"Master" は "Child"-"Parent" ペア単位、"Work Master" は指番（Diff Package から
parse_sashiban_module_side() で逆算）ごとの "Child"-"Parent" ペア単位でユニーク化する。
前回ダウンロードした "統合図面管理台帳.xlsx" を今回の実行時にアップロードすると、
今回新たに得られたデータで既存行とマージする（Master・Work Masterともに同じ蓄積方式）。
同じキーが前回・今回の両方にある場合、または今回の入力内で複数フォルダにまたがる場合は、
"Recorded Date" が最も新しい行を採用する（2026-08、同一 Child-Parent ペアが複数の
DXF-diff-manager出力フォルダに異なる実行時刻で記録されるケースが実データで確認された
ための変更。それまでは先勝ち〈今回データが常に前回を上書き〉だった）。

"Summary" は指番ごとに、今回アップロードしたZIPのデータのみから算出した集計値を
1行として毎回追記していく（Master/Work Masterのようなキー単位のマージ・上書きは
行わない）。同じ指番の履歴を実行日時ごとに追うための単純な追記ログ。"""

import io
from datetime import datetime

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font

from utils.group_summary_builder import aggregate_input_drawing_totals_by_sashiban, parse_sashiban_module_side
from utils.ledger_finder import DIFF_LIST_HEADERS

MASTER_SHEET_NAME = "Master"
WORK_MASTER_SHEET_NAME = "Work Master"
SUMMARY_SHEET_NAME = "Summary"

# DXF-diff-manager 自身の Relation 表記と完全一致させる（"完全新規図面-changed" は
# 含めない。model/master_ledger.py の save_master_to_bytes() の brand_new_mask と同じ判定）。
BRAND_NEW_RELATION = "完全新規図面"

CENTER_ALIGNMENT = Alignment(horizontal="center")
_ENTITY_LABELS = ("Deleted Entities", "Added Entities", "Diff Entities", "Unchanged Entities", "Total Entities")
_CHILD_COL = DIFF_LIST_HEADERS.index("Child")
_PARENT_COL = DIFF_LIST_HEADERS.index("Parent")
_RELATION_COL = DIFF_LIST_HEADERS.index("Relation")
_RECORDED_DATE_COL = DIFF_LIST_HEADERS.index("Recorded Date")
_DELETED_COL = DIFF_LIST_HEADERS.index("Deleted Entities")
_ADDED_COL = DIFF_LIST_HEADERS.index("Added Entities")
_TOTAL_COL = DIFF_LIST_HEADERS.index("Total Entities")

# Work Master は Relation を除く DIFF_LIST_HEADERS の前に Sashiban（指番）を付与した構成
_NON_RELATION_INDICES = [i for i, h in enumerate(DIFF_LIST_HEADERS) if h != "Relation"]
WORK_MASTER_HEADERS = ("Sashiban",) + tuple(DIFF_LIST_HEADERS[i] for i in _NON_RELATION_INDICES)
_WM_RECORDED_DATE_COL = WORK_MASTER_HEADERS.index("Recorded Date")

# Summaryシートは指番ごとの実行時点のスナップショットを追記するログ形式。
# 列名はユーザー指定のとおり日本語（既存のMaster/Work Masterの英語列名とは別扱い）。
# 「完全新規図面数」「新規作成率 [%]」は2026-08追加（DXF-diff-manager Summaryシートの
# 対応する2指標と同じ相対位置：差分ペア総数の直下・流用率[%]の直下）。
SUMMARY_HEADERS = (
    "指番", "削除図形総数", "追加図形総数", "変更図形総数", "図形総数",
    "図形変更率 [%]", "差分ペア総数", "完全新規図面数", "指番図面総数",
    "流用率 [%]", "新規作成率 [%]", "日付",
)
_SUMMARY_PERCENT_LABELS = {"図形変更率 [%]", "流用率 [%]", "新規作成率 [%]"}
_SUMMARY_COUNT_LABELS = (
    "削除図形総数", "追加図形総数", "変更図形総数", "図形総数",
    "差分ペア総数", "完全新規図面数", "指番図面総数",
)
_SUMMARY_DATE_COL = SUMMARY_HEADERS.index("日付")


def _recorded_date_or_min(value):
    """Recorded Date セルの値を比較可能にする。datetime でなければ（None・壊れた
    値等）datetime.min を返し、実際の日時を持つ行が常に優先されるようにする。"""
    return value if isinstance(value, datetime) else datetime.min


def extract_unique_child_parent_rows(entries):
    """LedgerEntry のリストから、"Child"-"Parent" ペアでユニーク化した
    DIFF_LIST_HEADERS 12列のデータを返す。同じペアが複数エントリにまたがる場合は
    "Recorded Date" が最も新しい行を採用する（2026-08、同一ペアが複数の
    DXF-diff-manager出力フォルダに異なる実行時刻で記録される実データケースを
    確認したための変更。それまでは先勝ちだった）。

    Returns:
        dict[(child, parent), tuple]
    """
    unique = {}
    for entry in entries:
        for row in entry.diff_list_rows:
            key = (row[_CHILD_COL], row[_PARENT_COL])
            existing = unique.get(key)
            if existing is None or (
                _recorded_date_or_min(row[_RECORDED_DATE_COL])
                > _recorded_date_or_min(existing[_RECORDED_DATE_COL])
            ):
                unique[key] = row
    return unique


def _extract_unique_work_master_entries(entries):
    """(sashiban, child, parent) -> (sashiban, diff_list_row) の辞書を返す内部
    共有ヘルパー。diff_list_row は DIFF_LIST_HEADERS 12列（Relationを含む）。
    Diff Package（出力フォルダ名）から指番を逆算できないエントリは対象外とする。
    同じキーが複数エントリにまたがる場合は "Recorded Date" が最も新しい行を採用する
    （extract_unique_child_parent_rows と同じ規則）。

    extract_unique_work_master_rows()（Work Master出力用にRelationを除いた形へ
    変換する）と compute_summary_rows()（Relationを用いて完全新規図面数・
    差分ペア総数を算出する）の両方から使う。Work Master自体はRelationを持たない
    ため、この中間形式でRelationを保持しておく必要がある。
    """
    unique = {}
    for entry in entries:
        sashiban, _module, _side = parse_sashiban_module_side(entry.package_name)
        if sashiban is None:
            continue
        for row in entry.diff_list_rows:
            key = (sashiban, row[_CHILD_COL], row[_PARENT_COL])
            existing = unique.get(key)
            if existing is None or (
                _recorded_date_or_min(row[_RECORDED_DATE_COL])
                > _recorded_date_or_min(existing[1][_RECORDED_DATE_COL])
            ):
                unique[key] = (sashiban, row)
    return unique


def extract_unique_work_master_rows(entries):
    """LedgerEntry のリストから、指番ごとに "Child"-"Parent" ペアでユニーク化した
    WORK_MASTER_HEADERS 12列のデータを返す。Diff Package（出力フォルダ名）から
    指番を逆算できないエントリは対象外とする。同じキーが複数エントリにまたがる
    場合は "Recorded Date" が最も新しい行を採用する
    （_extract_unique_work_master_entries 参照）。

    Returns:
        dict[(sashiban, child, parent), tuple]
    """
    return {
        key: (sashiban,) + tuple(row[i] for i in _NON_RELATION_INDICES)
        for key, (sashiban, row) in _extract_unique_work_master_entries(entries).items()
    }


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


def read_work_master_rows(file_bytes):
    """アップロードされた統合図面管理台帳.xlsxのWork Masterシートから
    (sashiban, child, parent) をキーとする行の辞書を読み込む。シートが存在しない
    （旧バージョンで作成されたファイル等）・構成が想定と異なる場合は None を返す
    （呼び出し側は今回分のみで新規作成する）。
    """
    try:
        wb = load_workbook(io.BytesIO(file_bytes), data_only=True, read_only=True)
    except Exception:
        return None

    try:
        if WORK_MASTER_SHEET_NAME not in wb.sheetnames:
            return None
        ws = wb[WORK_MASTER_SHEET_NAME]
        rows = list(ws.iter_rows(values_only=True))
        if not rows or tuple(rows[0]) != WORK_MASTER_HEADERS:
            return None
        return {(row[0], row[1], row[2]): tuple(row) for row in rows[1:]}
    finally:
        wb.close()


def read_summary_rows(file_bytes):
    """アップロードされた統合図面管理台帳.xlsxのSummaryシートから、既存の全行を
    そのまま読み込む（指番等でユニーク化しない。実行ごとのスナップショットを
    単純に追記していくログ形式のため）。シートが存在しない・構成が想定と異なる
    場合は空リストを返す（呼び出し側は今回分のみで新規作成する）。
    """
    try:
        wb = load_workbook(io.BytesIO(file_bytes), data_only=True, read_only=True)
    except Exception:
        return []

    try:
        if SUMMARY_SHEET_NAME not in wb.sheetnames:
            return []
        ws = wb[SUMMARY_SHEET_NAME]
        rows = list(ws.iter_rows(values_only=True))
        if not rows or tuple(rows[0]) != SUMMARY_HEADERS:
            return []
        return [tuple(row) for row in rows[1:]]
    finally:
        wb.close()


def _numeric_sum(rows, col_idx):
    """rows（DIFF_LIST_HEADERS形状のタプルのリスト）の col_idx 列を合計する。
    'n/a'（文字列）等の非数値は0として扱う（Summaryシートの合計列は常に数値で
    表示するため。個々の行の値取得ではなく指番単位の総数を出す用途のため、
    utils.group_summary_builder.aggregate_diff_list_by_child のような
    「全行n/aならn/aのまま残す」規約は採用しない）。
    """
    return sum(row[col_idx] for row in rows if isinstance(row[col_idx], (int, float)))


def compute_summary_rows(entries, run_timestamp):
    """今回のZIP入力（entries）のみから、指番ごとのSummary行（SUMMARY_HEADERS
    12列）を算出する。指番を逆算できないエントリは対象外（Work Masterと同じ扱い）。

    「完全新規図面数」は Relation == BRAND_NEW_RELATION の行の Child ユニーク数、
    「差分ペア総数」はそれ以外（完全新規図面を除く）の (Child, Parent) ユニーク数
    （2026-08、DXF-diff-manager 自身の「差分抽出ペア数」〈status=='complete' の
    ペア数、完全新規図面を含まない〉と定義を揃えるための変更。それまでは完全新規
    図面の行も差分ペア総数に含めていた）。削除/追加/変更/図形総数のエンティティ
    統計は、完全新規図面の行も含めたまま合計する（DXF-diff-manager 側の集計と
    同じ範囲）。

    Returns:
        list[tuple]（指番昇順）
    """
    entries_by_key = _extract_unique_work_master_entries(entries)
    rows_by_sashiban = {}
    for (sashiban, _child, _parent), (_sashiban, row) in entries_by_key.items():
        rows_by_sashiban.setdefault(sashiban, []).append(row)

    input_drawing_totals = aggregate_input_drawing_totals_by_sashiban(entries)

    summary_rows = []
    for sashiban in sorted(rows_by_sashiban.keys()):
        rows = rows_by_sashiban[sashiban]
        deleted_total = _numeric_sum(rows, _DELETED_COL)
        added_total = _numeric_sum(rows, _ADDED_COL)
        changed_total = deleted_total + added_total
        entity_total = _numeric_sum(rows, _TOTAL_COL)
        change_rate = (changed_total / entity_total) if entity_total else 0.0

        brand_new_children = {row[_CHILD_COL] for row in rows if row[_RELATION_COL] == BRAND_NEW_RELATION}
        brand_new_count = len(brand_new_children)
        pair_count = sum(1 for row in rows if row[_RELATION_COL] != BRAND_NEW_RELATION)

        input_drawing_total = input_drawing_totals.get(sashiban, 0)
        reuse_rate = (pair_count / input_drawing_total) if input_drawing_total else 0.0
        brand_new_rate = (brand_new_count / input_drawing_total) if input_drawing_total else 0.0

        summary_rows.append((
            sashiban, deleted_total, added_total, changed_total, entity_total,
            change_rate, pair_count, brand_new_count, input_drawing_total,
            reuse_rate, brand_new_rate, run_timestamp,
        ))
    return summary_rows


def _write_ledger_sheet(ws, headers, combined_rows, sort_key, recorded_date_col_idx):
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = CENTER_ALIGNMENT
    ws.freeze_panes = "A2"

    recorded_date_col = recorded_date_col_idx + 1
    entity_cols = [headers.index(label) + 1 for label in _ENTITY_LABELS]

    for key in sorted(combined_rows.keys(), key=sort_key):
        ws.append(combined_rows[key])
        row_idx = ws.max_row
        ws.cell(row=row_idx, column=recorded_date_col).number_format = "YYYY-MM-DD HH:MM:SS"
        for col in entity_cols:
            cell = ws.cell(row=row_idx, column=col)
            cell.number_format = "#,##0"
            cell.alignment = CENTER_ALIGNMENT

    for col_idx, header in enumerate(headers, start=1):
        width = max(len(str(header)) + 2, 12)
        ws.column_dimensions[ws.cell(row=1, column=col_idx).column_letter].width = width


def _write_summary_sheet(ws, previous_rows, new_rows):
    ws.append(SUMMARY_HEADERS)
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = CENTER_ALIGNMENT
    ws.freeze_panes = "A2"

    count_cols = [SUMMARY_HEADERS.index(label) + 1 for label in _SUMMARY_COUNT_LABELS]
    percent_cols = [SUMMARY_HEADERS.index(label) + 1 for label in _SUMMARY_PERCENT_LABELS]
    date_col = _SUMMARY_DATE_COL + 1

    for row in list(previous_rows) + list(new_rows):
        ws.append(row)
        row_idx = ws.max_row
        for col in count_cols:
            cell = ws.cell(row=row_idx, column=col)
            cell.number_format = "#,##0"
            cell.alignment = CENTER_ALIGNMENT
        for col in percent_cols:
            cell = ws.cell(row=row_idx, column=col)
            cell.number_format = "0.00%"
            cell.alignment = CENTER_ALIGNMENT
        ws.cell(row=row_idx, column=date_col).number_format = "YYYY-MM-DD HH:MM:SS"

    for col_idx, header in enumerate(SUMMARY_HEADERS, start=1):
        width = max(len(str(header)) + 2, 12)
        ws.column_dimensions[ws.cell(row=1, column=col_idx).column_letter].width = width


def _merge_by_recorded_date(previous_rows, new_rows, recorded_date_col_idx):
    """previous_rows・new_rows（同じキー形状の辞書）をマージする。同じキーが
    両方に存在する場合は "Recorded Date" が新しい方（同日時なら今回分）を採用する。
    前回のみに存在するキーはそのまま保持し、今回のみに存在するキーは追加する
    （2026-08、それまでの「今回データが常に前回を上書き」から変更）。
    """
    combined = dict(previous_rows or {})
    for key, new_row in (new_rows or {}).items():
        existing = combined.get(key)
        if existing is None or (
            _recorded_date_or_min(new_row[recorded_date_col_idx])
            >= _recorded_date_or_min(existing[recorded_date_col_idx])
        ):
            combined[key] = new_row
    return combined


def build_master_workbook(
    entries, previous_master_rows=None, previous_work_master_rows=None, previous_summary_rows=None,
):
    """今回のDiff Listデータ（Master: Child-Parentユニーク化、Work Master: 指番ごとの
    Child-Parentユニーク化）と、アップロードされた前回のMaster/Work Masterシート内容
    （無ければ None）をマージし、"統合図面管理台帳.xlsx"（"Master"→"Work Master"→
    "Summary"の順で3シート）を bytes で返す。同じキーが前回・今回の両方にある場合は
    "Recorded Date" が新しい方を採用する（_merge_by_recorded_date 参照）。
    Summaryシートのみキー単位のマージは行わず、今回分の指番ごとの集計行を、
    アップロードされた前回分（previous_summary_rows、無ければ空）の末尾に追記する。
    """
    combined_master = _merge_by_recorded_date(
        previous_master_rows, extract_unique_child_parent_rows(entries), _RECORDED_DATE_COL,
    )
    combined_work_master = _merge_by_recorded_date(
        previous_work_master_rows, extract_unique_work_master_rows(entries), _WM_RECORDED_DATE_COL,
    )

    new_summary_rows = compute_summary_rows(entries, run_timestamp=datetime.now())

    wb = Workbook()
    ws = wb.active
    ws.title = MASTER_SHEET_NAME
    _write_ledger_sheet(
        ws, DIFF_LIST_HEADERS, combined_master,
        sort_key=lambda k: k[0], recorded_date_col_idx=_RECORDED_DATE_COL,
    )

    wm_ws = wb.create_sheet(WORK_MASTER_SHEET_NAME)
    _write_ledger_sheet(
        wm_ws, WORK_MASTER_HEADERS, combined_work_master,
        sort_key=lambda k: (k[0], k[1]), recorded_date_col_idx=_WM_RECORDED_DATE_COL,
    )

    summary_ws = wb.create_sheet(SUMMARY_SHEET_NAME)
    _write_summary_sheet(summary_ws, previous_summary_rows or [], new_summary_rows)

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()
