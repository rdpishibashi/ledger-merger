"""DXF-diff-manager の出力フォルダ群から台帳ファイル（Child/Parent等の列構成を持つ
データシート + Summary シートを持つ .xlsx）を再帰的に検出するモジュール。
Streamlit には依存しない。"""

import os
from dataclasses import dataclass

import openpyxl

DIFF_LIST_HEADERS = (
    "Child", "Parent", "Relation", "Title", "Subtitle", "Recorded Date", "Note",
    "Deleted Entities", "Added Entities", "Diff Entities", "Unchanged Entities",
    "Total Entities",
)

# DIFF_LIST_HEADERS 内の各列インデックス。台帳の生行（entry.diff_list_rows の要素）を
# 読み出す全モジュール（group_summary_builder.py, ledger_merger.py,
# master_ledger_builder.py）で共有する。
CHILD_COL = DIFF_LIST_HEADERS.index("Child")
PARENT_COL = DIFF_LIST_HEADERS.index("Parent")
RELATION_COL = DIFF_LIST_HEADERS.index("Relation")
TITLE_COL = DIFF_LIST_HEADERS.index("Title")
SUBTITLE_COL = DIFF_LIST_HEADERS.index("Subtitle")
RECORDED_DATE_COL = DIFF_LIST_HEADERS.index("Recorded Date")
NOTE_COL = DIFF_LIST_HEADERS.index("Note")
DELETED_COL = DIFF_LIST_HEADERS.index("Deleted Entities")
ADDED_COL = DIFF_LIST_HEADERS.index("Added Entities")
DIFF_COL = DIFF_LIST_HEADERS.index("Diff Entities")
UNCHANGED_COL = DIFF_LIST_HEADERS.index("Unchanged Entities")
TOTAL_COL = DIFF_LIST_HEADERS.index("Total Entities")

ENTITY_LABELS = ("Deleted Entities", "Added Entities", "Diff Entities", "Unchanged Entities", "Total Entities")

SUMMARY_LABELS = (
    "削除図形数 合計", "追加図形数 合計", "差分図形数 合計", "変更なし図形数 合計",
    "総図形数 合計", "図形変更率 [%]",
)

# DXF-diff-manager のバージョンによってはSummaryシートに含まれない指標。
# SUMMARY_LABELS（台帳判定の必須条件）には含めない——必須にすると、この指標を
# 持たないDXF-diff-manager出力の台帳がすべて無効判定されてしまうため。存在すれば
# summary_values に取り込むだけの任意項目とし、欠損時は呼び出し側で0として扱う。
#
# 「入力図面総数」「差分抽出ペア数」「流用率 [%]」は、DXF-diff-manager が
# 2026-09 に Summary シートの「図面統計」欄（見出し＋この3項目＋「完全新規図面数」
# 「新規作成率 [%]」）を全削除したことに伴い、ここで必須（SUMMARY_LABELS）から
# 任意（OPTIONAL_SUMMARY_LABELS）へ格下げした。新形式の台帳（この3項目を持たない）
# を必須のままにすると、台帳自体が黙って「無効」判定され統合対象から除外される
# （2026-08、"Master" へのシート名改名で実際に発生しかけた問題と同種の失敗モード）。
# 新形式の台帳では、これらに依存する下流の指標（Ledger-merger の「指番図面総数」
# 「流用率 [%]」「新規作成率 [%]」）は 0 として集計される（app.py 側で警告表示する）。
OPTIONAL_SUMMARY_LABELS = (
    "入力図面総数", "差分抽出ペア数", "流用率 [%]", "完全新規図面数", "新規作成率 [%]",
)

# SUMMARY_LABELS（Ledger-merger 自身の統合Excel出力列名。README.md 記載の契約）に対する、
# DXF-diff-manager の Summary シート側の実際のラベル文言。DXF-diff-manager 側でラベル
# 文言が変更されても統合Excelの列名は変えたくないため、ここでエイリアスとして吸収する。
# 「総図形数 合計」「入力図面総数」はペアリング方式（Type A/B/C）によって文言が変わる
# （DXF-diff-manager `model/master_ledger.py` の `save_master_to_bytes()` 参照）ため、
# 複数エイリアスを許容する。OPTIONAL_SUMMARY_LABELS の2項目もここに含め、
# _read_summary_values() で他の項目と同様に解決する（存在する場合のみ summary_values
# に入る。SUMMARY_LABELS の必須チェックには含めない）。
_SOURCE_LABEL_ALIASES = {
    "削除図形数 合計": ("削除図形 総数",),
    "追加図形数 合計": ("追加図形 総数",),
    "差分図形数 合計": ("変更（追加+削除）図形 総数",),
    "変更なし図形数 合計": ("変更なし図形 総数",),
    "総図形数 合計": ("アップロード図面 図形総数", "流用先図面 図形総数"),
    "図形変更率 [%]": ("図形変更率 [%]",),
    "入力図面総数": ("アップロード図面総数", "流用先図面総数"),
    "差分抽出ペア数": ("差分抽出ペア数",),
    "流用率 [%]": ("流用率 [%]",),
    "完全新規図面数": ("完全新規図面数",),
    "新規作成率 [%]": ("新規作成率 [%]",),
}

# DXF-diff-manager が出力する台帳以外の固定ファイル名。台帳の候補から除外する。
NON_LEDGER_FILENAMES = {"diff_labels.xlsx", "unchanged_labels.xlsx"}


@dataclass
class LedgerEntry:
    package_name: str
    source_path: str
    diff_list_rows: list
    summary_values: dict


def _read_summary_values(ws):
    raw = {}
    for row in ws.iter_rows(values_only=True):
        if not row:
            continue
        label = row[0]
        if label and len(row) > 1:
            raw[label] = row[1]

    values = {}
    for canonical, aliases in _SOURCE_LABEL_ALIASES.items():
        for alias in aliases:
            if alias in raw:
                values[canonical] = raw[alias]
                break
    return values


def _find_diff_list_rows(wb):
    """データシート（ヘッダー行が DIFF_LIST_HEADERS と完全一致するシート）を
    シート名に依存せず探し、全行を返す。

    DXF-diff-manager 側はこのシート名を過去の改修で変えてきた実績があり
    （'Diff List' → 'Master'、2026-08改名）、今後も変わりうる。固定シート名に
    依存すると、名前が変わった途端に台帳が黙って「無効」判定され統合対象から
    除外される（2026-08、'Master' への改名で実際に発生しかけた問題）。
    DXF-diff-manager 自身の load_parent_child_master()（model/master_ledger.py）
    と同じ考え方で、列構成（ヘッダーの完全一致）で判定する。

    Returns:
        list または None: 見つかったシートの全行（ヘッダー含む）。無ければ None。
    """
    for name in wb.sheetnames:
        rows = list(wb[name].iter_rows(values_only=True))
        if rows and tuple(rows[0]) == DIFF_LIST_HEADERS:
            return rows
    return None


def _try_load_ledger(path):
    """path が有効な台帳ファイルなら LedgerEntry を返す。そうでなければ None。"""
    try:
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    except Exception:
        return None

    try:
        if "Summary" not in wb.sheetnames:
            return None

        rows = _find_diff_list_rows(wb)
        if rows is None:
            return None

        # Total Entities が空欄の行は、実際には差分抽出されていない図番ペアの
        # 関係記録（親子マスター管理用）であり、Summary シートの集計にも含まれない。
        # 統合 Diff List には実際に差分抽出された行のみを含める。
        diff_list_rows = [row for row in rows[1:] if row[TOTAL_COL] is not None]
        if not diff_list_rows:
            return None

        summary_values = _read_summary_values(wb["Summary"])
        if any(label not in summary_values for label in SUMMARY_LABELS):
            return None

        package_name = os.path.basename(os.path.dirname(path))
        return LedgerEntry(
            package_name=package_name,
            source_path=path,
            diff_list_rows=diff_list_rows,
            summary_values=summary_values,
        )
    finally:
        wb.close()


def find_ledger_files(root_dir):
    """root_dir 以下を再帰的に走査し、台帳ファイルを検出する。

    「出力フォルダ」は、直下に .xlsx を持ち、かつそれより深い階層のどのフォルダにも
    .xlsx が存在しない、木構造上もっとも深い地点として判定する（サブフォルダの
    有無そのものでは判定しない）。DXF-diff-manager の出力フォルダには元DXFを格納する
    `dxf図面` 等の非xlsxサブフォルダが付随する場合があるが、そのサブフォルダ自体は
    xlsxを持たないため出力フォルダの深さ判定に影響しない。一方、ZIP展開時のラッパー
    フォルダ（直下にもxlsxがあるが、より深い階層の各出力フォルダにもxlsxがある場合）は
    中間ラッパーとして評価対象から除外する。macOSのFinder/dittoがZIP化時に作る
    `__MACOSX/` ミラーフォルダ（リソースフォーク `._*` を含む）も同名の偽フォルダとして
    誤検出されるため除外する。

    Returns:
        (entries, folders_without_ledger):
            entries: package_name でソートされた LedgerEntry のリスト。
            folders_without_ledger: 台帳に相当する Excel ファイルが見つからなかった
                出力フォルダ名（ベース名のみ、重複なし、出現順）のリスト。
    """
    entries = []
    folders_without_ledger = []
    seen_missing = set()

    xlsx_files_by_dir = {}
    for dirpath, dirnames, filenames in os.walk(root_dir):
        # macOS の Finder/ditto で ZIP 化すると、各フォルダをミラーする
        # __MACOSX/ フォルダ（リソースフォーク ._ファイル名 を含む）が作られる。
        # 同名の偽フォルダとして誤検出されるため、走査対象から除外する。
        dirnames[:] = [d for d in dirnames if d != "__MACOSX"]

        xlsx_files = [
            f for f in filenames
            if f.lower().endswith(".xlsx")
            and not f.startswith("~$")
            and not f.startswith("._")
        ]
        if xlsx_files:
            xlsx_files_by_dir[dirpath] = xlsx_files

    for dirpath, xlsx_files in xlsx_files_by_dir.items():
        deeper_prefix = dirpath + os.sep
        has_deeper_ledger_dir = any(
            other.startswith(deeper_prefix)
            for other in xlsx_files_by_dir
            if other != dirpath
        )
        if has_deeper_ledger_dir:
            # より深い階層にもxlsxを持つフォルダがある＝このフォルダは中間ラッパー。
            # 実際の出力フォルダはさらに深い階層で個別に評価される。
            continue

        candidates = [f for f in xlsx_files if f.lower() not in NON_LEDGER_FILENAMES]

        found = False
        for filename in candidates:
            entry = _try_load_ledger(os.path.join(dirpath, filename))
            if entry:
                entries.append(entry)
                found = True

        if not found:
            folder_name = os.path.basename(dirpath)
            if folder_name not in seen_missing:
                seen_missing.add(folder_name)
                folders_without_ledger.append(folder_name)

    entries.sort(key=lambda e: e.package_name)
    return entries, folders_without_ledger


def reconcile_missing_folders(entries, folders_without_ledger):
    """複数の入力ソース（複数ZIP等）の find_ledger_files() 結果を集約した後に呼ぶ。

    find_ledger_files() は1つの入力ソースごとに独立して判定するため、同じフォルダ名が
    異なる入力ソースに存在し、片方には有効な台帳があり、もう片方には無い場合、
    集約結果は「統合成功」と「台帳が見つからなかったフォルダ」の両方に同じ名前が
    矛盾して現れてしまう。ここで台帳が見つかったフォルダ名を除外し、矛盾を解消する。
    また、複数の入力ソースに渡って重複した欠落フォルダ名も除去する（出現順を維持）。
    """
    found_names = {entry.package_name for entry in entries}
    reconciled = []
    seen = set()
    for name in folders_without_ledger:
        if name in found_names or name in seen:
            continue
        seen.add(name)
        reconciled.append(name)
    return reconciled
