"""LedgerEntry の Diff List 行をユニーク化し、"統合図面管理台帳.xlsx"
（"Master"→"Work Master"→"Summary" の3シート）を生成・蓄積するモジュール。
Streamlit には依存しない。

"Master" は "Child"-"Parent" ペア単位、"Work Master" は指番・モジュール・サイド
（Diff Package から parse_sashiban_module_side() で逆算）ごとの "Child"-"Parent"
ペア単位でユニーク化する。前回ダウンロードした "統合図面管理台帳.xlsx" を今回の
実行時にアップロードすると、今回新たに得られたデータで既存行とマージする
（Master・Work Masterともに同じ蓄積方式）。同じキーが前回・今回の両方にある場合、
または今回の入力内で複数フォルダにまたがる場合は、"Recorded Date" が最も新しい行を
採用する（同一 Child-Parent ペアが複数の DXF-diff-manager出力フォルダに異なる
実行時刻で記録されることがあるため、常に最新の実行結果を残す）。

"Summary" は、マージ済み（＝前回分と今回分を合わせた累積状態の）"Work Master" を
(指番, 差分方式) 単位で集計したスナップショットを毎回全再計算する（2026-09、
旧仕様「今回バッチのみを追記するログ」から変更。ユーザー要求）。Work Master 自体が
前回分とマージ済み・累積済みのため、Summary を毎回再計算しても実行履歴は失われない
——常に「その時点の統合図面管理台帳全体を集計した最新スナップショット」になる。

"Master"・"Work Master" はいずれも Diff Type 列（Diff Package から
parse_diff_type() で逆算）を持つが、フィールド構成自体は異なる（Master は
Sashiban/Module/Side を持たず Relation を保持する。Work Master はその逆——
Relation の代わりに、完全新規図面の判定には DXF-diff-manager 自身の規約
（流用元なしの場合 Parent="none"）を使う）。"""

import io
from datetime import datetime

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font

from utils.group_summary_builder import (
    parse_diff_type,
    parse_sashiban_module_side,
)
from utils.ledger_finder import (
    ADDED_COL,
    CHILD_COL,
    DELETED_COL,
    DIFF_COL,
    ENTITY_LABELS,
    NOTE_COL,
    PARENT_COL,
    RECORDED_DATE_COL,
    RELATION_COL,
    SUBTITLE_COL,
    TITLE_COL,
    TOTAL_COL,
    UNCHANGED_COL,
)

MASTER_SHEET_NAME = "Master"
WORK_MASTER_SHEET_NAME = "Work Master"
SUMMARY_SHEET_NAME = "Summary"

CENTER_ALIGNMENT = Alignment(horizontal="center")
LEFT_ALIGNMENT = Alignment(horizontal="left")

# Master の列構成。Sashiban・Module・Side は含まない——Master は指番を問わず
# Child-Parent 単位で全体をユニーク化するシートであり、Work Masterとはフィールド
# 構成が異なる。
MASTER_HEADERS = (
    "Child", "Parent", "Relation", "Title", "Subtitle", "Diff Type",
    "Deleted Entities", "Added Entities", "Diff Entities", "Unchanged Entities",
    "Total Entities", "Note", "Recorded Date",
)
_MASTER_RELATION_COL = MASTER_HEADERS.index("Relation")
_MASTER_DIFF_TYPE_COL = MASTER_HEADERS.index("Diff Type")
_MASTER_RECORDED_DATE_COL = MASTER_HEADERS.index("Recorded Date")

# DXF-diff-manager が Relation 列に書く関係の種別。同じ図面（Child）に対して
# 「RevUp」と「流用」の両方が記録されている場合、Work Master では RevUp を採用し
# 流用の行を落とす（Drawing-genealogy の GraphBuilder._reuse_pairs_to_delete()
# と同じ規則。2026-09、ユーザー要求）。
REVUP_RELATION = "RevUp"
REUSE_RELATION = "流用"

# Work Master の列構成。Master と異なり Sashiban・Module・Side を Child の前に持つ。
# "Relation" は 2026-09 に追加（ユーザー要求）。Master と同じく Parent の直後に置く。
# 旧形式（Relation列なしの15列）でアップロードされた台帳も読めるようにしてある
# （read_work_master_rows 参照。黙って前回分を捨てないため）。
WORK_MASTER_HEADERS = (
    "Sashiban", "Module", "Side", "Child", "Parent", "Relation", "Title", "Subtitle", "Diff Type",
    "Deleted Entities", "Added Entities", "Diff Entities", "Unchanged Entities",
    "Total Entities", "Note", "Recorded Date",
)
# Relation列を追加する前の列構成（2026-09以前にダウンロードされた台帳）。
LEGACY_WORK_MASTER_HEADERS = tuple(h for h in WORK_MASTER_HEADERS if h != "Relation")
_WM_SASHIBAN_COL = WORK_MASTER_HEADERS.index("Sashiban")
_WM_CHILD_COL = WORK_MASTER_HEADERS.index("Child")
_WM_PARENT_COL = WORK_MASTER_HEADERS.index("Parent")
_WM_RELATION_COL = WORK_MASTER_HEADERS.index("Relation")
_WM_DIFF_TYPE_COL = WORK_MASTER_HEADERS.index("Diff Type")
_WM_DELETED_COL = WORK_MASTER_HEADERS.index("Deleted Entities")
_WM_ADDED_COL = WORK_MASTER_HEADERS.index("Added Entities")
_WM_TOTAL_COL = WORK_MASTER_HEADERS.index("Total Entities")
_WM_RECORDED_DATE_COL = WORK_MASTER_HEADERS.index("Recorded Date")

# Work Master 上で文字列型・左寄せに統一する列（2026-09、ユーザー要求）。
# Sashiban/Module/Side/Child は本来すべて文字列だが、_sort_str() のdocstringで
# 説明している型ドリフト（前回アップロードファイルに数値として保存されている
# ことがある）と同じ理由で、書き込み時にも明示的に str() へ正規化する。
WORK_MASTER_STRING_LABELS = ("Sashiban", "Module", "Side", "Child", "Parent", "Title", "Subtitle")
_WM_STRING_COL_INDEXES = tuple(WORK_MASTER_HEADERS.index(label) for label in WORK_MASTER_STRING_LABELS)

# Work Master のユニークキー（Sashiban, Module, Side, Child, Parent）を構成する列。
# WORK_MASTER_HEADERS の先頭5列であり、read_work_master_rows() が読み込み時に
# 作るキーと同じ並び。
WORK_MASTER_KEY_LABELS = ("Sashiban", "Module", "Side", "Child", "Parent")
_WM_KEY_COL_INDEXES = tuple(WORK_MASTER_HEADERS.index(label) for label in WORK_MASTER_KEY_LABELS)

# DXF-diff-manager 自身の完全新規図面（流用元なし）の表現。Work Master には
# Relation 列が無いため、完全新規図面の判定にはこの値との比較を使う
# （model/master_ledger.py の `parent_value = parent if parent else 'none'` 参照）。
BRAND_NEW_PARENT = "none"

# Summaryシートは、マージ済みWork Masterを (指番, 差分方式) 単位で集計した
# スナップショット（2026-09、旧仕様「実行ごとの追記ログ」から変更）。
# 列名はユーザー指定のとおり日本語（既存のMaster/Work Masterの英語列名とは別扱い）。
# 「差分方式」は指番の直後（Work Master の Diff Type 列をそのまま使う）。
SUMMARY_HEADERS = (
    "指番", "差分方式", "削除図形総数", "追加図形総数", "変更図形総数", "図形総数",
    "図形変更率 [%]", "変更図面総数", "完全新規図面数", "日付",
)
_SUMMARY_PERCENT_LABELS = {"図形変更率 [%]"}
_SUMMARY_COUNT_LABELS = (
    "削除図形総数", "追加図形総数", "変更図形総数", "図形総数",
    "変更図面総数", "完全新規図面数",
)
_SUMMARY_DATE_COL = SUMMARY_HEADERS.index("日付")


def _recorded_date_or_min(value):
    """Recorded Date セルの値を比較可能にする。datetime でなければ（None・壊れた
    値等）datetime.min を返し、実際の日時を持つ行が常に優先されるようにする。"""
    return value if isinstance(value, datetime) else datetime.min


def _sort_str(value):
    """ソートキーの各要素を文字列に正規化する。

    Sashiban/Module/Side/Child は本来すべて文字列だが、アップロードされた前回の
    統合図面管理台帳.xlsx（previous_master_rows/previous_work_master_rows。
    openpyxl でセルの生値をそのまま読む）に、数字だけのセル（例: サイド "405"）が
    テキストではなく数値として保存されていた場合、int/float で返ってくることが
    ある（Excel上での手編集・別ツールでの再保存等が原因になりうる。2026-09
    ユーザー報告で実際に発生: `TypeError: '<' not supported between instances of
    'str' and 'int'`）。None（空欄セル）と合わせて、ここで一律 str に変換してから
    比較することでクラッシュを防ぐ。
    """
    return '' if value is None else str(value)


def _normalize_work_master_row(row):
    """Work Master 行の文字列列（WORK_MASTER_STRING_LABELS の7列）を str に正規化する。

    新規に算出した行（extract_unique_work_master_rows() の戻り値）だけでなく、
    アップロードされた前回の統合図面管理台帳.xlsx 由来の行（openpyxl でセルの
    生値をそのまま読むため、Excel上での手編集等で int/float が混入しうる）にも
    適用する。これにより、過去に数値として保存された値も次回以降の出力では
    文字列に健全化される（_sort_str() が対処したのと同種の型ドリフト。
    2026-09、ユーザー要求により Work Master の該当7列は常に文字列・左寄せに
    統一する）。
    """
    row = list(row)
    for idx in _WM_STRING_COL_INDEXES:
        if row[idx] is not None:
            row[idx] = str(row[idx])
    return tuple(row)


def _normalize_work_master_rows(rows):
    """Work Master の辞書全体を、行の値だけでなく **キーも** 正規化して返す。

    **キーの正規化を省いてはいけない（省くと重複行が出る）**: 前回の統合図面管理台帳
    .xlsx の Work Master シートでサイドが数値として保存されていると（実データで確認:
    237行すべての Side が int の 405）、read_work_master_rows() が作るキーは
    `('ME24-1001-0','ZC00',405,...)`、今回分のキーは `(...,'405',...)` となり、
    _merge_by_recorded_date() が同一行と認識できず**前回分と今回分が両方残る**
    （2026-09、ユーザー報告: 474行中237件が重複。Summaryの集計値も倍になった）。
    値だけを str 化しても、マージはその前に失敗しており、表示上は
    「キー列が同一に見える重複行」になる。

    キーは正規化後の行から作り直すことで、キーと値の整合を構造的に保証する。
    正規化の結果、複数の行が同一キーへ収束した場合は "Recorded Date" が最も新しい
    行を採用する（_merge_by_recorded_date と同じ規則）。

    根拠テスト: tests/regression/bugfix/test_work_master_key_type_drift_duplicates.py
    """
    normalized = {}
    for row in (rows or {}).values():
        row = _normalize_work_master_row(row)
        key = tuple(row[idx] for idx in _WM_KEY_COL_INDEXES)
        existing = normalized.get(key)
        if existing is None or (
            _recorded_date_or_min(row[_WM_RECORDED_DATE_COL])
            >= _recorded_date_or_min(existing[_WM_RECORDED_DATE_COL])
        ):
            normalized[key] = row
    return normalized


def _relation_lookup(master_rows):
    """Master の行（Relation列を持つ）から (child, parent) -> Relation の辞書を作る。

    旧形式（Relation列が無い15列）でアップロードされた Work Master 行の Relation を
    補完するために使う（_backfill_work_master_relation 参照）。Master は前回分・
    今回分をマージ済みで、かつ Work Master より対象が広い（指番を逆算できない
    エントリも含む）ため、Work Master の全行に対応する Relation を引ける。

    キーは Work Master 側（正規化済みで必ず str）と突き合わせるため、Master 側も
    同じ規則で str 化する（型ドリフト対策。_normalize_work_master_rows 参照）。
    """
    lookup = {}
    for row in (master_rows or {}).values():
        child = row[CHILD_COL] if row[CHILD_COL] is None else str(row[CHILD_COL])
        parent = row[PARENT_COL] if row[PARENT_COL] is None else str(row[PARENT_COL])
        lookup[(child, parent)] = row[_MASTER_RELATION_COL]
    return lookup


def _backfill_work_master_relation(work_master, relation_by_child_parent):
    """Relation が空の Work Master 行（旧形式の台帳から読み込んだ行）を、Master の
    Relation で埋める。

    これにより、Relation列が無かった頃にダウンロードされた台帳を再アップロードして
    も、次回の出力では Relation が入り、RevUp/流用の判定
    （_drop_reuse_rows_superseded_by_revup）も前回分に遡って効く。Master に該当
    ペアが無い場合は空のままにする（判定対象外として素通しする＝消さない）。
    """
    filled = {}
    for key, row in work_master.items():
        if row[_WM_RELATION_COL] is None:
            relation = relation_by_child_parent.get((row[_WM_CHILD_COL], row[_WM_PARENT_COL]))
            row = row[:_WM_RELATION_COL] + (relation,) + row[_WM_RELATION_COL + 1:]
        filled[key] = row
    return filled


def _drop_reuse_rows_superseded_by_revup(rows, relation_col_idx, group_key):
    """同じ図面に「RevUp」と「流用」の両方の関係がある場合、「流用」の行を落として
    「RevUp」の行だけを残す（2026-09、ユーザー要求。Drawing-genealogy の
    GraphBuilder._reuse_pairs_to_delete() と同じ考え方）。Master・Work Master の
    両方に適用する。

    group_key: 「同じ図面」をどの単位で見るかを決める関数（キー -> グループ）。
        - Master: Child のみ（指番の列を持たないシートのため、これが唯一の単位。
          Drawing-genealogy が台帳全体を Child 単位で見るのと同じ）。
        - Work Master: (Sashiban, Module, Side, Child)。**指番をまたいで判定しない**——
          その指番には RevUp が記録されていないのに他の指番の RevUp を根拠に唯一の
          関係行が消えるのを避けるため（実データ検証: 指番横断で判定すると指番
          PE25-9601-0 の10行が別指番 NE24-0062-0 の RevUp を理由に削除され、
          PE25-9601-0 ではその図面の流用元が一切たどれなくなる）。

    Drawing-genealogy との違い: あちらは図番の版数から RevUp エッジを**推測**もするが
    （_infer_revision_up_edges）、ここでは台帳に実在する Relation のみで判定する
    （ユーザー要求は「2つの関係がある場合」＝両方が記録されている場合のため）。
    完全新規図面（Parent="none"）の行は削除対象にしない（規則の対象は流用のみ）。

    Returns:
        (filtered_rows, dropped_keys)
    """
    revup_groups = {
        group_key(key) for key, row in rows.items() if row[relation_col_idx] == REVUP_RELATION
    }
    dropped = {
        key for key, row in rows.items()
        if row[relation_col_idx] == REUSE_RELATION and group_key(key) in revup_groups
    }
    return {key: row for key, row in rows.items() if key not in dropped}, dropped


def extract_unique_child_parent_rows(entries):
    """LedgerEntry のリストから、"Child"-"Parent" ペアでユニーク化した
    MASTER_HEADERS 13列のデータを返す。同じペアが複数エントリにまたがる場合は
    "Recorded Date" が最も新しい行を採用する。Diff Type は Diff Package
    （出力フォルダ名）から parse_diff_type() で逆算する（指番の解決可否に関わらず
    全エントリが対象——Master は Work Master と異なり指番不明のエントリも含むため）。

    Returns:
        dict[(child, parent), tuple]
    """
    unique = {}
    for entry in entries:
        diff_type = parse_diff_type(entry.package_name)
        for row in entry.diff_list_rows:
            key = (row[CHILD_COL], row[PARENT_COL])
            candidate = (
                row[CHILD_COL], row[PARENT_COL], row[RELATION_COL], row[TITLE_COL], row[SUBTITLE_COL],
                diff_type, row[DELETED_COL], row[ADDED_COL], row[DIFF_COL], row[UNCHANGED_COL],
                row[TOTAL_COL], row[NOTE_COL], row[RECORDED_DATE_COL],
            )
            existing = unique.get(key)
            if existing is None or (
                _recorded_date_or_min(candidate[_MASTER_RECORDED_DATE_COL])
                > _recorded_date_or_min(existing[_MASTER_RECORDED_DATE_COL])
            ):
                unique[key] = candidate
    return unique


def _extract_unique_work_master_entries(entries):
    """(sashiban, module, side, child, parent) ->
    (sashiban, module, side, diff_type, diff_list_row) の辞書を返す内部ヘルパー。
    diff_list_row は DIFF_LIST_HEADERS 12列（Relationを含む）。台帳ファイル名を主・
    出力フォルダ名を従として指番を逆算できないエントリは対象外とする
    （parse_sashiban_module_side() 参照。ミスタイプ等で台帳ファイル名の命名規則にも
    一致しない場合は find_entries_with_unresolved_sashiban() で検出できる）。
    diff_type は Diff Package（出力フォルダ名）のみから parse_diff_type() で逆算する
    （台帳ファイル名には含まれないため。一致しない場合は None）。同じキーが複数
    エントリにまたがる場合は "Recorded Date" が最も新しい行を採用する
    （extract_unique_child_parent_rows と同じ規則）。

    モジュール/サイドをキーに含めるため、同一 (指番, Child, Parent) が複数の
    モジュール/サイドに跨る場合はそれぞれ別行として残る（実データで確認済みの
    ケース。tests/regression/spec/test_work_master_module_side_columns.py 参照）。
    """
    unique = {}
    for entry in entries:
        sashiban, module, side = parse_sashiban_module_side(entry.package_name, entry.source_path)
        if sashiban is None:
            continue
        diff_type = parse_diff_type(entry.package_name)
        for row in entry.diff_list_rows:
            key = (sashiban, module, side, row[CHILD_COL], row[PARENT_COL])
            existing = unique.get(key)
            if existing is None or (
                _recorded_date_or_min(row[RECORDED_DATE_COL])
                > _recorded_date_or_min(existing[4][RECORDED_DATE_COL])
            ):
                unique[key] = (sashiban, module, side, diff_type, row)
    return unique


def extract_unique_work_master_rows(entries):
    """LedgerEntry のリストから、指番・モジュール・サイドごとに "Child"-"Parent"
    ペアでユニーク化した WORK_MASTER_HEADERS 16列のデータを返す。Diff Package
    （出力フォルダ名）から指番を逆算できないエントリは対象外とする。同じキーが
    複数エントリにまたがる場合は "Recorded Date" が最も新しい行を採用する
    （_extract_unique_work_master_entries 参照）。

    Relation は元の Diff List 行の値をそのまま持たせる（2026-09追加。Master と
    同じ値。RevUp/流用の判定にも使う）。

    Returns:
        dict[(sashiban, module, side, child, parent), tuple]
    """
    result = {}
    for key, (sashiban, module, side, diff_type, row) in _extract_unique_work_master_entries(
        entries,
    ).items():
        result[key] = (
            sashiban, module, side, row[CHILD_COL], row[PARENT_COL], row[RELATION_COL],
            row[TITLE_COL], row[SUBTITLE_COL], diff_type,
            row[DELETED_COL], row[ADDED_COL], row[DIFF_COL], row[UNCHANGED_COL], row[TOTAL_COL],
            row[NOTE_COL], row[RECORDED_DATE_COL],
        )
    return result


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
        if not rows or tuple(rows[0]) != MASTER_HEADERS:
            return None
        return {(row[CHILD_COL], row[PARENT_COL]): tuple(row) for row in rows[1:]}
    finally:
        wb.close()


def read_work_master_rows(file_bytes):
    """アップロードされた統合図面管理台帳.xlsxのWork Masterシートから
    (sashiban, module, side, child, parent) をキーとする行の辞書を読み込む。
    シートが存在しない・構成が想定と異なる場合は None を返す（呼び出し側は今回分
    のみで新規作成する。Masterと異なりこの場合は警告を出さない）。

    **旧形式（Relation列が無い15列。2026-09以前にダウンロードされた台帳）も受け付け、
    Relation を None として現行の16列形へ変換する**（Relation は呼び出し側が Master
    から補完する。_backfill_work_master_relation 参照）。列を増やした際に旧形式を
    弾いてしまうと、蓄積済みのWork Masterが警告も無く丸ごと捨てられるため。
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
        if not rows:
            return None
        header = tuple(rows[0])
        if header == WORK_MASTER_HEADERS:
            data_rows = [tuple(row) for row in rows[1:]]
        elif header == LEGACY_WORK_MASTER_HEADERS:
            data_rows = [
                tuple(row[:_WM_RELATION_COL]) + (None,) + tuple(row[_WM_RELATION_COL:])
                for row in rows[1:]
            ]
        else:
            return None
        return {tuple(row[idx] for idx in _WM_KEY_COL_INDEXES): row for row in data_rows}
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


def compute_summary_rows(combined_work_master):
    """マージ済みのWork Master（combined_work_master。build_master_workbook() 内で
    前回分と今回分を合わせた累積状態）を (Sashiban, Diff Type) ごとに集計し、
    Summary行（SUMMARY_HEADERS 10列）を算出する。

    Work Masterは既に (Sashiban, Module, Side, Child, Parent) でユニーク化済み
    のため、ここでの追加のユニーク化は不要——グループ内の行数がそのまま
    「完全新規図面数」「変更図面総数」の内訳になる。

    Work Master には Relation 列が無いため、完全新規図面の判定は
    DXF-diff-manager 自身の規約（流用元なしの場合 Parent="none"、
    BRAND_NEW_PARENT）を使う。「完全新規図面数」は Parent==BRAND_NEW_PARENT の
    行の Child ユニーク数、「変更図面総数」はそれ以外の行数。削除/追加/変更/
    図形総数のエンティティ統計は、完全新規図面の行も含めたまま合計する
    （DXF-diff-manager 側の集計と同じ範囲）。「図形変更率 [%]」は集計後の
    合計値から再計算する（各行の変更率の平均ではない）。「日付」はグループ内の
    Recorded Date の最大値（有効な日時を持つ行が1つも無ければ None）。

    Returns:
        list[tuple]（指番昇順、同一指番内は差分方式昇順。差分方式が逆算できない
        〈None〉場合は同一指番内の末尾に回る）
    """
    rows_by_sashiban_type = {}
    for row in combined_work_master.values():
        sashiban = row[_WM_SASHIBAN_COL]
        diff_type = row[_WM_DIFF_TYPE_COL]
        rows_by_sashiban_type.setdefault((sashiban, diff_type), []).append(row)

    summary_rows = []
    for sashiban, diff_type in sorted(
        rows_by_sashiban_type.keys(), key=lambda k: (_sort_str(k[0]), k[1] is None, _sort_str(k[1])),
    ):
        rows = rows_by_sashiban_type[(sashiban, diff_type)]
        deleted_total = _numeric_sum(rows, _WM_DELETED_COL)
        added_total = _numeric_sum(rows, _WM_ADDED_COL)
        changed_total = deleted_total + added_total
        entity_total = _numeric_sum(rows, _WM_TOTAL_COL)
        change_rate = (changed_total / entity_total) if entity_total else 0.0

        brand_new_children = {
            row[_WM_CHILD_COL] for row in rows if row[_WM_PARENT_COL] == BRAND_NEW_PARENT
        }
        brand_new_count = len(brand_new_children)
        pair_count = sum(1 for row in rows if row[_WM_PARENT_COL] != BRAND_NEW_PARENT)

        dates = [
            row[_WM_RECORDED_DATE_COL] for row in rows if isinstance(row[_WM_RECORDED_DATE_COL], datetime)
        ]
        latest_date = max(dates) if dates else None

        summary_rows.append((
            sashiban, diff_type, deleted_total, added_total, changed_total, entity_total,
            change_rate, pair_count, brand_new_count, latest_date,
        ))
    return summary_rows


def _write_ledger_sheet(ws, headers, combined_rows, sort_key, recorded_date_col_idx, left_align_labels=()):
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = CENTER_ALIGNMENT
    ws.freeze_panes = "A2"

    recorded_date_col = recorded_date_col_idx + 1
    entity_cols = [headers.index(label) + 1 for label in ENTITY_LABELS]
    diff_type_col = headers.index("Diff Type") + 1
    left_align_cols = [headers.index(label) + 1 for label in left_align_labels]

    for key in sorted(combined_rows.keys(), key=sort_key):
        ws.append(combined_rows[key])
        row_idx = ws.max_row
        ws.cell(row=row_idx, column=recorded_date_col).number_format = "YYYY-MM-DD HH:MM:SS"
        for col in entity_cols:
            cell = ws.cell(row=row_idx, column=col)
            cell.number_format = "#,##0"
            cell.alignment = CENTER_ALIGNMENT
        ws.cell(row=row_idx, column=diff_type_col).alignment = CENTER_ALIGNMENT
        for col in left_align_cols:
            ws.cell(row=row_idx, column=col).alignment = LEFT_ALIGNMENT

    for col_idx, header in enumerate(headers, start=1):
        width = max(len(str(header)) + 2, 12)
        ws.column_dimensions[ws.cell(row=1, column=col_idx).column_letter].width = width


def _write_summary_sheet(ws, new_rows):
    ws.append(SUMMARY_HEADERS)
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = CENTER_ALIGNMENT
    ws.freeze_panes = "A2"

    count_cols = [SUMMARY_HEADERS.index(label) + 1 for label in _SUMMARY_COUNT_LABELS]
    percent_cols = [SUMMARY_HEADERS.index(label) + 1 for label in _SUMMARY_PERCENT_LABELS]
    date_col = _SUMMARY_DATE_COL + 1
    diff_type_col = SUMMARY_HEADERS.index("差分方式") + 1

    for row in new_rows:
        ws.append(row)
        row_idx = ws.max_row
        ws.cell(row=row_idx, column=diff_type_col).alignment = CENTER_ALIGNMENT
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
    前回のみに存在するキーはそのまま保持し、今回のみに存在するキーは追加する。
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


def build_master_workbook(entries, previous_master_rows=None, previous_work_master_rows=None):
    """今回のDiff Listデータ（Master: Child-Parentユニーク化、Work Master: 指番・
    モジュール・サイドごとのChild-Parentユニーク化）と、アップロードされた前回の
    Master/Work Masterシート内容（無ければ None）をマージし、"統合図面管理台帳.xlsx"
    （"Master"→"Work Master"→"Summary"の順で3シート）を bytes で返す。同じキーが
    前回・今回の両方にある場合は "Recorded Date" が新しい方を採用する
    （_merge_by_recorded_date 参照）。Summaryシートはキー単位のマージではなく、
    マージ済みWork Master全体を (指番,差分方式) 単位で毎回再集計する
    （compute_summary_rows 参照。Work Master自体が累積済みのため、再集計しても
    履歴は失われない）。

    同じ図面に「RevUp」と「流用」の両方の関係がある場合は流用の行を落とす
    （_drop_reuse_rows_superseded_by_revup 参照）。Master・Work Master の両方に適用し、
    判定単位だけが異なる（Master は Child のみ、Work Master は指番・モジュール・
    サイドも含む）。
    """
    combined_master = _merge_by_recorded_date(
        previous_master_rows, extract_unique_child_parent_rows(entries), _MASTER_RECORDED_DATE_COL,
    )
    # **マージの前に**両側のキー・値を正規化する。前回分の Side が数値で保存されて
    # いると、正規化前のキーでは今回分と一致せず重複行になるため
    # （_normalize_work_master_rows のdocstring参照。順序を入れ替えないこと）。
    combined_work_master = _merge_by_recorded_date(
        _normalize_work_master_rows(previous_work_master_rows),
        _normalize_work_master_rows(extract_unique_work_master_rows(entries)),
        _WM_RECORDED_DATE_COL,
    )
    # 旧形式（Relation列なし）の台帳から来た行の Relation を Master から補完する。
    # **間引きより前に**行うこと（Relationが空のままだと判定対象外になり、前回分の
    # 流用の行が残ってしまう）。参照元は間引き前の Master（間引き後だと落とした
    # 流用ペアを引けなくなる）。
    combined_work_master = _backfill_work_master_relation(
        combined_work_master, _relation_lookup(combined_master),
    )
    # RevUpと流用が両方ある図面は流用を落とす。Summaryは間引き後のWork Masterから
    # 集計する（Work Masterの見た目と集計値を一致させるため）。
    combined_work_master, _dropped_work_master = _drop_reuse_rows_superseded_by_revup(
        combined_work_master, _WM_RELATION_COL, group_key=lambda key: key[:4],
    )
    # Master は指番の列を持たないため、判定単位は Child のみ（Drawing-genealogy が
    # 台帳全体を Child 単位で見るのと同じ）。
    combined_master, _dropped_master = _drop_reuse_rows_superseded_by_revup(
        combined_master, _MASTER_RELATION_COL, group_key=lambda key: key[0],
    )

    new_summary_rows = compute_summary_rows(combined_work_master)

    wb = Workbook()
    ws = wb.active
    ws.title = MASTER_SHEET_NAME
    # 並び順は Child → Diff Type。Child が全体を通じて昇順になることを優先し、
    # Diff Type は同一Childが複数のDiff Typeに跨る場合のタイブレークとしてのみ
    # 使う（Diff Typeを優先すると、Diff Type混在時にChild列が全体としては昇順に
    # 見えなくなる）。Diff Type はキー〈child, parent〉ではなく行の値側にあるため、
    # combined_master から都度引いて判定する。None〈逆算不可〉は同一Child内で
    # 末尾に回す。
    #
    # k[0]（Child）・Diff Type を _sort_str() で文字列に正規化しているのは、
    # アップロードされた前回の統合図面管理台帳.xlsx（previous_master_rows。
    # openpyxl でセルの生値をそのまま読む）に、空欄セル（None）や数値として
    # 保存されたセル（int/float）が混在していた場合、今回分（常に文字列）との
    # 比較でクラッシュするのを防ぐため（_sort_str() のdocstring参照。2026-09、
    # Work Master側で実際に発生した不具合と同じクラス）。
    _write_ledger_sheet(
        ws, MASTER_HEADERS, combined_master,
        sort_key=lambda k: (
            _sort_str(k[0]),
            combined_master[k][_MASTER_DIFF_TYPE_COL] is None,
            _sort_str(combined_master[k][_MASTER_DIFF_TYPE_COL]),
        ),
        recorded_date_col_idx=_MASTER_RECORDED_DATE_COL,
    )

    wm_ws = wb.create_sheet(WORK_MASTER_SHEET_NAME)
    # 各要素を _sort_str() で文字列に正規化しているのは、アップロードされた前回の
    # 統合図面管理台帳.xlsx（previous_work_master_rows）の Work Master シートに、
    # Sashiban/Module/Side/Child のいずれかが空欄（None）や数値として保存された
    # セル（int/float。例: サイド "405" が数値として保存されていた場合）を持つ
    # 行が混在していた場合、今回分（parse_sashiban_module_side() が返す値は
    # 必ず文字列）との比較でクラッシュするのを防ぐため（_sort_str() の
    # docstring参照。2026-09 ユーザー報告で実際に発生。当初 `or ''` で None のみ
    # 防御していたが、int混在の再発報告を受けて str() 正規化に強化した）。
    _write_ledger_sheet(
        wm_ws, WORK_MASTER_HEADERS, combined_work_master,
        sort_key=lambda k: tuple(_sort_str(v) for v in k[:4]),
        recorded_date_col_idx=_WM_RECORDED_DATE_COL,
        left_align_labels=WORK_MASTER_STRING_LABELS,
    )

    summary_ws = wb.create_sheet(SUMMARY_SHEET_NAME)
    _write_summary_sheet(summary_ws, new_summary_rows)

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()
