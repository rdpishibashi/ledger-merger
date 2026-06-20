"""DXF-diff-manager の出力フォルダ群から台帳ファイル（Diff List + Summary シートを持つ
.xlsx）を再帰的に検出するモジュール。Streamlit には依存しない。"""

import os
from dataclasses import dataclass

import openpyxl

DIFF_LIST_HEADERS = (
    "Child", "Parent", "Relation", "Title", "Subtitle", "Recorded Date", "Note",
    "Deleted Entities", "Added Entities", "Diff Entities", "Unchanged Entities",
    "Total Entities",
)

SUMMARY_LABELS = (
    "削除図形数 合計", "追加図形数 合計", "差分図形数 合計", "変更なし図形数 合計",
    "総図形数 合計", "図形変更率 [%]", "入力図面総数", "差分抽出ペア数", "流用率 [%]",
)

# DXF-diff-manager が出力する台帳以外の固定ファイル名。台帳の候補から除外する。
NON_LEDGER_FILENAMES = {"diff_labels.xlsx", "unchanged_labels.xlsx"}

_TOTAL_ENTITIES_COL = DIFF_LIST_HEADERS.index("Total Entities")


@dataclass
class LedgerEntry:
    package_name: str
    source_path: str
    diff_list_rows: list
    summary_values: dict


def _read_summary_values(ws):
    values = {}
    for row in ws.iter_rows(values_only=True):
        if not row:
            continue
        label = row[0]
        if label in SUMMARY_LABELS and len(row) > 1:
            values[label] = row[1]
    return values


def _try_load_ledger(path):
    """path が有効な台帳ファイルなら LedgerEntry を返す。そうでなければ None。"""
    try:
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    except Exception:
        return None

    try:
        if "Diff List" not in wb.sheetnames or "Summary" not in wb.sheetnames:
            return None

        ws_diff = wb["Diff List"]
        rows = list(ws_diff.iter_rows(values_only=True))
        if not rows or tuple(rows[0]) != DIFF_LIST_HEADERS:
            return None

        # Total Entities が空欄の行は、実際には差分抽出されていない図番ペアの
        # 関係記録（親子マスター管理用）であり、Summary シートの集計にも含まれない。
        # 統合 Diff List には実際に差分抽出された行のみを含める。
        diff_list_rows = [row for row in rows[1:] if row[_TOTAL_ENTITIES_COL] is not None]
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

    DXF-diff-manager の出力フォルダはサブフォルダを持たない（葉フォルダ）ため、
    サブフォルダを持つ中間フォルダ（ZIP展開時のラッパーフォルダ等）は評価対象から
    除外し、葉フォルダのみを「出力フォルダ」として扱う。macOSのFinder/dittoが
    ZIP化時に作る `__MACOSX/` ミラーフォルダ（リソースフォーク `._*` を含む）も
    同名の偽フォルダとして誤検出されるため除外する。

    Returns:
        (entries, folders_without_ledger):
            entries: package_name でソートされた LedgerEntry のリスト。
            folders_without_ledger: 台帳に相当する Excel ファイルが見つからなかった
                出力フォルダ名（ベース名のみ、重複なし、出現順）のリスト。
    """
    entries = []
    folders_without_ledger = []
    seen_missing = set()

    for dirpath, dirnames, filenames in os.walk(root_dir):
        # macOS の Finder/ditto で ZIP 化すると、各フォルダをミラーする
        # __MACOSX/ フォルダ（リソースフォーク ._ファイル名 を含む）が作られる。
        # 同名の偽フォルダとして誤検出されるため、走査対象から除外する。
        dirnames[:] = [d for d in dirnames if d != "__MACOSX"]

        if dirnames or not filenames:
            continue

        candidates = [
            f for f in filenames
            if f.lower().endswith(".xlsx")
            and not f.startswith("~$")
            and not f.startswith("._")
            and f.lower() not in NON_LEDGER_FILENAMES
        ]

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
