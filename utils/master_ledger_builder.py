"""LedgerEntry の Diff List 行をユニーク化し、"統合図面管理台帳.xlsx"
（"Master"→"Work Master"→"Summary" の3シート）を生成・蓄積するモジュール。
Streamlit には依存しない。

"Master" は "Child"-"Parent" ペア単位、"Work Master" は指番・モジュール・サイド
（Diff Package から parse_sashiban_module_side() で逆算）ごとの "Child"-"Parent"
ペア単位でユニーク化する。前回ダウンロードした "統合図面管理台帳.xlsx" を今回の
実行時にアップロードすると、今回新たに得られたデータで既存行とマージする
（Master・Work Masterともに同じ蓄積方式）。同じキーが前回・今回の両方にある場合、
または今回の入力内で複数フォルダにまたがる場合は、"Recorded Date" が最も新しい行を
採用する（2026-08、同一 Child-Parent ペアが複数の DXF-diff-manager出力フォルダに
異なる実行時刻で記録されるケースが実データで確認されたための変更。それまでは
先勝ち〈今回データが常に前回を上書き〉だった）。

"Summary" は指番・差分方式ごとに、今回アップロードしたZIPのデータのみから算出した
集計値を1行として毎回追記していく（Master/Work Masterのようなキー単位のマージ・
上書きは行わない）。同じ指番の履歴を実行日時ごとに追うための単純な追記ログ。

"Master"・"Work Master" はいずれも Diff Type 列（Diff Package から
parse_diff_type() で逆算）を持つが、フィールド構成自体は異なる（Master は
Sashiban/Module/Side を持たず Relation を保持する。Work Master はその逆）。"""

import io
from datetime import datetime

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font

from utils.group_summary_builder import (
    aggregate_input_drawing_totals_by_sashiban_and_diff_type,
    parse_diff_type,
    parse_sashiban_module_side,
)
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
_TITLE_COL = DIFF_LIST_HEADERS.index("Title")
_SUBTITLE_COL = DIFF_LIST_HEADERS.index("Subtitle")
_RECORDED_DATE_COL = DIFF_LIST_HEADERS.index("Recorded Date")
_NOTE_COL = DIFF_LIST_HEADERS.index("Note")
_DELETED_COL = DIFF_LIST_HEADERS.index("Deleted Entities")
_ADDED_COL = DIFF_LIST_HEADERS.index("Added Entities")
_DIFF_COL = DIFF_LIST_HEADERS.index("Diff Entities")
_UNCHANGED_COL = DIFF_LIST_HEADERS.index("Unchanged Entities")
_TOTAL_COL = DIFF_LIST_HEADERS.index("Total Entities")

# Master の列構成（2026-08、Work Masterと同じ並び替えパターンを適用。Sashiban・
# Module・Side は追加しない——Master は指番を問わず Child-Parent 単位で全体を
# ユニーク化するシートであり、Work Masterとはフィールド構成が異なる。ユーザー確認
# 済み：DIFF_LIST_HEADERSの並びをそのまま転記する方式から、明示的な列順の指定に
# 変更し、Subtitleと Deleted Entitiesの間に Diff Type を追加、Note・Recorded Date
# を Total Entities の後ろへ移動）。
MASTER_HEADERS = (
    "Child", "Parent", "Relation", "Title", "Subtitle", "Diff Type",
    "Deleted Entities", "Added Entities", "Diff Entities", "Unchanged Entities",
    "Total Entities", "Note", "Recorded Date",
)
_MASTER_DIFF_TYPE_COL = MASTER_HEADERS.index("Diff Type")
_MASTER_RECORDED_DATE_COL = MASTER_HEADERS.index("Recorded Date")

# Work Master の列構成（2026-08、Sashiban と Child の間に Module・Side を追加した後、
# さらに Subtitle と Deleted Entities の間に Diff Type を追加し、Note・Recorded Date を
# Total Entities の後ろへ移動——DIFF_LIST_HEADERS の並びをそのまま転記する方式から、
# 明示的な列順の指定に変更）。
WORK_MASTER_HEADERS = (
    "Sashiban", "Module", "Side", "Child", "Parent", "Title", "Subtitle", "Diff Type",
    "Deleted Entities", "Added Entities", "Diff Entities", "Unchanged Entities",
    "Total Entities", "Note", "Recorded Date",
)
_WM_RECORDED_DATE_COL = WORK_MASTER_HEADERS.index("Recorded Date")

# Summaryシートは指番・差分方式ごとの実行時点のスナップショットを追記するログ形式。
# 列名はユーザー指定のとおり日本語（既存のMaster/Work Masterの英語列名とは別扱い）。
# 「完全新規図面数」「新規作成率 [%]」は2026-08追加（DXF-diff-manager Summaryシートの
# 対応する2指標と同じ相対位置：差分ペア総数の直下・流用率[%]の直下）。「差分方式」も
# 2026-08追加（指番の直後。Diff Package から parse_diff_type() で逆算。同一指番内で
# 差分方式が異なる場合は別行に分ける——ユーザー確認済み仕様）。
SUMMARY_HEADERS = (
    "指番", "差分方式", "削除図形総数", "追加図形総数", "変更図形総数", "図形総数",
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
    MASTER_HEADERS 13列のデータを返す。同じペアが複数エントリにまたがる場合は
    "Recorded Date" が最も新しい行を採用する（2026-08、同一ペアが複数の
    DXF-diff-manager出力フォルダに異なる実行時刻で記録される実データケースを
    確認したための変更。それまでは先勝ちだった）。Diff Type は Diff Package
    （出力フォルダ名）から parse_diff_type() で逆算する（指番の解決可否に関わらず
    全エントリが対象——Master は Work Master と異なり指番不明のエントリも含むため）。

    Returns:
        dict[(child, parent), tuple]
    """
    unique = {}
    for entry in entries:
        diff_type = parse_diff_type(entry.package_name)
        for row in entry.diff_list_rows:
            key = (row[_CHILD_COL], row[_PARENT_COL])
            candidate = (
                row[_CHILD_COL], row[_PARENT_COL], row[_RELATION_COL], row[_TITLE_COL], row[_SUBTITLE_COL],
                diff_type, row[_DELETED_COL], row[_ADDED_COL], row[_DIFF_COL], row[_UNCHANGED_COL],
                row[_TOTAL_COL], row[_NOTE_COL], row[_RECORDED_DATE_COL],
            )
            existing = unique.get(key)
            if existing is None or (
                _recorded_date_or_min(candidate[_MASTER_RECORDED_DATE_COL])
                > _recorded_date_or_min(existing[_MASTER_RECORDED_DATE_COL])
            ):
                unique[key] = candidate
    return unique


def _extract_unique_work_master_entries(entries, key_by_module_side=False):
    """(sashiban, child, parent) または (sashiban, module, side, child, parent) ->
    (sashiban, module, side, diff_type, diff_list_row) の辞書を返す内部共有ヘルパー。
    diff_list_row は DIFF_LIST_HEADERS 12列（Relationを含む）。台帳ファイル名を主・
    出力フォルダ名を従として指番を逆算できないエントリは対象外とする
    （parse_sashiban_module_side() 参照。ミスタイプ等で台帳ファイル名の命名規則にも
    一致しない場合は find_entries_with_unresolved_sashiban() で検出できる）。
    diff_type は Diff Package（出力フォルダ名）のみから parse_diff_type() で逆算する
    （台帳ファイル名には含まれないため。一致しない場合は None）。同じキーが複数
    エントリにまたがる場合は "Recorded Date" が最も新しい行を採用する
    （extract_unique_child_parent_rows と同じ規則）。

    extract_unique_work_master_rows()（Work Master出力用にRelationを除いた形へ
    変換する。key_by_module_side=True で呼ぶ）と compute_summary_rows()（Relationを
    用いて完全新規図面数・差分ペア総数を算出する。key_by_module_side=False の既定値
    のまま呼ぶ）の両方から使う。Work Master自体はRelationを持たないため、この中間
    形式でRelationを保持しておく必要がある。

    key_by_module_side: True にすると、同一 (指番, Child, Parent) が複数のモジュール/
    サイドに跨る場合（実データで確認済み）にそれぞれ別行として残す。Summary側は
    key_by_module_side=False（既定）のまま据え置く——含めると「差分ペア総数」が
    DXF-diff-manager 自身の「差分抽出ペア数」の定義（グループごとの合計）とずれる
    ため（2026-08、Work Master に Module/Side 列を追加した際の実データ検証で判明）。
    """
    unique = {}
    for entry in entries:
        sashiban, module, side = parse_sashiban_module_side(entry.package_name, entry.source_path)
        if sashiban is None:
            continue
        diff_type = parse_diff_type(entry.package_name)
        for row in entry.diff_list_rows:
            key = (
                (sashiban, module, side, row[_CHILD_COL], row[_PARENT_COL])
                if key_by_module_side
                else (sashiban, row[_CHILD_COL], row[_PARENT_COL])
            )
            existing = unique.get(key)
            if existing is None or (
                _recorded_date_or_min(row[_RECORDED_DATE_COL])
                > _recorded_date_or_min(existing[4][_RECORDED_DATE_COL])
            ):
                unique[key] = (sashiban, module, side, diff_type, row)
    return unique


def extract_unique_work_master_rows(entries):
    """LedgerEntry のリストから、指番・モジュール・サイドごとに "Child"-"Parent"
    ペアでユニーク化した WORK_MASTER_HEADERS 15列のデータを返す。Diff Package
    （出力フォルダ名）から指番を逆算できないエントリは対象外とする。同じキーが
    複数エントリにまたがる場合は "Recorded Date" が最も新しい行を採用する
    （_extract_unique_work_master_entries 参照）。

    Returns:
        dict[(sashiban, module, side, child, parent), tuple]
    """
    result = {}
    for key, (sashiban, module, side, diff_type, row) in _extract_unique_work_master_entries(
        entries, key_by_module_side=True,
    ).items():
        result[key] = (
            sashiban, module, side, row[_CHILD_COL], row[_PARENT_COL],
            row[_TITLE_COL], row[_SUBTITLE_COL], diff_type,
            row[_DELETED_COL], row[_ADDED_COL], row[_DIFF_COL], row[_UNCHANGED_COL], row[_TOTAL_COL],
            row[_NOTE_COL], row[_RECORDED_DATE_COL],
        )
    return result


def read_master_rows(file_bytes):
    """アップロードされた統合図面管理台帳.xlsx（Masterシートのみ）から
    (child, parent) をキーとする行の辞書を読み込む。シート構成が想定と異なる
    （壊れている、別ファイル等。旧12列形式〈Diff Type追加・Note/Recorded Date
    移動前〉も含む）場合は None を返す。先頭2列（Child, Parent）の位置は列構成
    変更の前後で変わらないため、キーの取り出し方自体は変更不要。
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
        if not rows or tuple(rows[0]) != MASTER_HEADERS:
            return None
        return {(row[_CHILD_COL], row[_PARENT_COL]): tuple(row) for row in rows[1:]}
    finally:
        wb.close()


def read_work_master_rows(file_bytes):
    """アップロードされた統合図面管理台帳.xlsxのWork Masterシートから
    (sashiban, module, side, child, parent) をキーとする行の辞書を読み込む。
    シートが存在しない（旧バージョンで作成されたファイル等）・構成が想定と異なる
    場合は None を返す（呼び出し側は今回分のみで新規作成する。Work Master が
    Diff Type・Note/Recorded Date の位置を持たない旧12列・14列形式だった場合も
    ヘッダー不一致で None になり、警告なしに今回分のみへフォールバックする——
    設計判断によりこの移行時の蓄積データ読み捨ては許容している）。先頭5列
    （Sashiban, Module, Side, Child, Parent）の位置は列追加・並び替えの前後で
    変わらないため、キーの取り出し方（row[0]〜row[4]）自体は変更不要。
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
        return {(row[0], row[1], row[2], row[3], row[4]): tuple(row) for row in rows[1:]}
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
    """今回のZIP入力（entries）のみから、(指番, 差分方式) ごとのSummary行
    （SUMMARY_HEADERS 13列）を算出する。指番を逆算できないエントリは対象外
    （Work Masterと同じ扱い）。差分方式は Diff Package から parse_diff_type() で
    逆算し、同一指番内で異なる差分方式が混在する場合は別行に分ける（2026-08、
    ユーザー確認済み仕様。実データでは指番内で差分方式が揃っているケースしか
    確認していないが、混在した場合も指番と差分方式の組ごとに集計を分離することで
    値が誤って混ざらないようにする）。

    「完全新規図面数」は Relation == BRAND_NEW_RELATION の行の Child ユニーク数、
    「差分ペア総数」はそれ以外（完全新規図面を除く）の (Child, Parent) ユニーク数
    （2026-08、DXF-diff-manager 自身の「差分抽出ペア数」〈status=='complete' の
    ペア数、完全新規図面を含まない〉と定義を揃えるための変更。それまでは完全新規
    図面の行も差分ペア総数に含めていた）。削除/追加/変更/図形総数のエンティティ
    統計は、完全新規図面の行も含めたまま合計する（DXF-diff-manager 側の集計と
    同じ範囲）。

    Returns:
        list[tuple]（指番昇順、同一指番内は差分方式昇順。差分方式が逆算できない
        〈None〉場合は同一指番内の末尾に回る）
    """
    entries_by_key = _extract_unique_work_master_entries(entries)
    rows_by_sashiban_type = {}
    for (sashiban, _child, _parent), (_sashiban, _module, _side, diff_type, row) in entries_by_key.items():
        rows_by_sashiban_type.setdefault((sashiban, diff_type), []).append(row)

    input_drawing_totals = aggregate_input_drawing_totals_by_sashiban_and_diff_type(entries)

    summary_rows = []
    for sashiban, diff_type in sorted(rows_by_sashiban_type.keys(), key=lambda k: (k[0], k[1] is None, k[1] or "")):
        rows = rows_by_sashiban_type[(sashiban, diff_type)]
        deleted_total = _numeric_sum(rows, _DELETED_COL)
        added_total = _numeric_sum(rows, _ADDED_COL)
        changed_total = deleted_total + added_total
        entity_total = _numeric_sum(rows, _TOTAL_COL)
        change_rate = (changed_total / entity_total) if entity_total else 0.0

        brand_new_children = {row[_CHILD_COL] for row in rows if row[_RELATION_COL] == BRAND_NEW_RELATION}
        brand_new_count = len(brand_new_children)
        pair_count = sum(1 for row in rows if row[_RELATION_COL] != BRAND_NEW_RELATION)

        input_drawing_total = input_drawing_totals.get((sashiban, diff_type), 0)
        reuse_rate = (pair_count / input_drawing_total) if input_drawing_total else 0.0
        brand_new_rate = (brand_new_count / input_drawing_total) if input_drawing_total else 0.0

        summary_rows.append((
            sashiban, diff_type, deleted_total, added_total, changed_total, entity_total,
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
    """今回のDiff Listデータ（Master: Child-Parentユニーク化、Work Master: 指番・
    モジュール・サイドごとのChild-Parentユニーク化）と、アップロードされた前回の
    Master/Work Masterシート内容（無ければ None）をマージし、"統合図面管理台帳.xlsx"
    （"Master"→"Work Master"→"Summary"の順で3シート）を bytes で返す。同じキーが
    前回・今回の両方にある場合は "Recorded Date" が新しい方を採用する
    （_merge_by_recorded_date 参照）。Summaryシートのみキー単位のマージは行わず、
    今回分の指番ごとの集計行を、アップロードされた前回分（previous_summary_rows、
    無ければ空）の末尾に追記する。
    """
    combined_master = _merge_by_recorded_date(
        previous_master_rows, extract_unique_child_parent_rows(entries), _MASTER_RECORDED_DATE_COL,
    )
    combined_work_master = _merge_by_recorded_date(
        previous_work_master_rows, extract_unique_work_master_rows(entries), _WM_RECORDED_DATE_COL,
    )

    new_summary_rows = compute_summary_rows(entries, run_timestamp=datetime.now())

    wb = Workbook()
    ws = wb.active
    ws.title = MASTER_SHEET_NAME
    # 並び順は Diff Type → Child（2026-08、ユーザー要望により Diff Type をソート
    # キーに追加。Diff Type はキー〈child, parent〉ではなく行の値側にあるため、
    # combined_master から都度引いて判定する。None〈逆算不可〉は末尾に回す）。
    _write_ledger_sheet(
        ws, MASTER_HEADERS, combined_master,
        sort_key=lambda k: (
            combined_master[k][_MASTER_DIFF_TYPE_COL] is None,
            combined_master[k][_MASTER_DIFF_TYPE_COL] or "",
            k[0],
        ),
        recorded_date_col_idx=_MASTER_RECORDED_DATE_COL,
    )

    wm_ws = wb.create_sheet(WORK_MASTER_SHEET_NAME)
    _write_ledger_sheet(
        wm_ws, WORK_MASTER_HEADERS, combined_work_master,
        sort_key=lambda k: (k[0], k[1], k[2], k[3]), recorded_date_col_idx=_WM_RECORDED_DATE_COL,
    )

    summary_ws = wb.create_sheet(SUMMARY_SHEET_NAME)
    _write_summary_sheet(summary_ws, previous_summary_rows or [], new_summary_rows)

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()
