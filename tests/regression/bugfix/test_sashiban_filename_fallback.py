"""不具合再発防止: 台帳ファイル名を主とした指番/モジュール/サイドの逆算。

不具合ID: 2026-08-03 ユーザー報告
    「Archive.zip には4つの指番が含まれているが、統合図面管理台帳.xlsxのWork
    Master/Summaryには ME24-1001-0 の1指番しか反映されない」。

以前どう壊れていたか:
    parse_sashiban_module_side()/parse_group_and_revision() が、Diff Package
    フォルダ名（package_name）だけを見て指番・モジュール・サイドを逆算していた。
    フォルダ名は DXF-diff-manager でZIPダウンロード時にユーザーが自由編集できる
    テキスト欄に由来し、既定値のまま・または短く編集されるとモジュール/サイドが
    欠落する（実データ: フォルダ名 "dxf_diff_results_TypeA_PE25-9601-0" にモジュール/
    サイドが無い）。この場合、指番が逆算できず Work Master・Summary・
    指番_モジュール_サイド別集計から**警告なしに黙って除外**されていた。

修正後に保証したいこと:
    1. 台帳ファイル名（"{指番}_{モジュール}_{サイド}[_-suffix].xlsx"。
       DXF-diff-manager 自身の model.master_ledger.MASTER_FILENAME_PATTERN と
       同一規則）が指番・モジュール・サイドの一次情報源になり、フォルダ名に
       欠落があっても正しく解決される。
    2. リビジョン番号は常にフォルダ名からのみ取得する（ファイル名の末尾サフィックスは
       リビジョン以外の自由文字列でありうるため）。
    3. 同一フォルダ内に複数の台帳候補がある場合（差分抽出やり直しの残骸）の
       「最新Recorded Dateを採用」処理は、指番解決より先に行う（指番の解決方法を
       変えても、同一フォルダ内の取り違え防止機構は壊れない）。
    4. ファイル名・フォルダ名のどちらからも指番を特定できない場合（ミスタイプ等。
       実データ: "NE24-0062-0_ZM00_405l.xlsx" の末尾 "l" が区切り文字
       "_"/"-" を伴わず不一致）は、find_entries_with_unresolved_sashiban() で
       検出できる（app.py の警告表示に使う）。
"""

import os
import sys
from datetime import datetime

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "..", ".."))

from utils.group_summary_builder import (
    find_entries_with_unresolved_sashiban,
    group_entries,
    parse_group_and_revision,
    parse_sashiban_module_side,
)
from utils.ledger_finder import LedgerEntry

_FLAT_FOLDER = "dxf_diff_results_TypeA_PE25-9601-0"  # モジュール/サイド無し（実データ）
_FULL_FOLDER = "dxf_diff_results_TypeA_ME24-1001-0_ZC00_405_04"  # モジュール/サイド/リビジョン有り


def _row(child, parent, relation, recorded_date):
    return (child, parent, relation, "T", "S", recorded_date, None, 1, 1, 2, 1, 3)


def test_sashiban_resolved_from_filename_when_folder_name_lacks_module_side():
    """フォルダ名にモジュール/サイドが無くても、台帳ファイル名から解決できる。"""
    assert parse_sashiban_module_side(_FLAT_FOLDER, "PE25-9601-0_ZM00_405.xlsx") == (
        "PE25-9601-0", "ZM00", "405",
    )


def test_sashiban_folder_fallback_still_works_when_no_filename_given():
    """filename を渡さない場合は従来通りフォルダ名から解決する（後方互換）。"""
    assert parse_sashiban_module_side(_FULL_FOLDER) == ("ME24-1001-0", "ZC00", "405")


def test_sashiban_unresolved_when_both_folder_and_filename_mismatch():
    """フォルダ名にもファイル名にも一致しない場合（ミスタイプ等）は解決できない。

    実例: "NE24-0062-0_ZM00_405l.xlsx" — 末尾の "l" が区切り文字 "_"/"-" を
    伴わず MASTER_FILENAME_PATTERN 同等の規則に一致しない。
    """
    assert parse_sashiban_module_side(
        "dxf_diff_results_TypeA_NE24-0062-0", "NE24-0062-0_ZM00_405l.xlsx"
    ) == (None, None, None)


def test_group_and_revision_resolved_via_filename_with_no_revision():
    """フォルダ名からの解決に失敗した場合、グループキーはファイル名から解決され、
    リビジョンはフォルダ名にリビジョン情報が無いため None になる。"""
    assert parse_group_and_revision(_FLAT_FOLDER, "PE25-9601-0_ZM00_405.xlsx") == (
        "PE25-9601-0_ZM00_405", None,
    )


def test_group_and_revision_ignores_filename_suffix_for_revision():
    """フォルダ名が指番/モジュール/サイド/リビジョンすべてを持つ場合、リビジョンは
    ファイル名の末尾サフィックス（リビジョン以外の自由文字列でありうる）ではなく
    フォルダ名から取得する。"""
    assert parse_group_and_revision(_FULL_FOLDER, "ME24-1001-0_ZC00_405_backup.xlsx") == (
        "ME24-1001-0_ZC00_405", "04",
    )


def test_group_entries_resolves_via_folder_first_before_splitting_by_sashiban():
    """同一フォルダ内に複数の台帳候補があり、片方はファイル名が指番書式に一致しない
    （例: 図番抽出に失敗した古い実行結果の "na_na" 命名）場合でも、まず同一フォルダ内で
    最新Recorded Dateの候補を1つに絞ってから指番を解決する。指番解決を先にすると、
    2つの候補が別グループに分かれてしまい「同一フォルダ内の取り違え防止」が機能しなく
    なる回帰を、実装時に一度発生させた（本テストはその回帰の再発防止）。"""
    older_failed_run = LedgerEntry(
        package_name=_FULL_FOLDER,
        source_path="ME24-1001-0_na_na.xlsx",
        diff_list_rows=[_row("C1", "P1", "RevUp", datetime(2026, 7, 7))],
        summary_values={},
    )
    newer_successful_run = LedgerEntry(
        package_name=_FULL_FOLDER,
        source_path="ME24-1001-0_ZC00_405.xlsx",
        diff_list_rows=[_row("C1", "P1", "RevUp", datetime(2026, 7, 9))],
        summary_values={},
    )

    groups = group_entries([older_failed_run, newer_successful_run])

    assert list(groups.keys()) == ["ME24-1001-0_ZC00_405"]
    revisions = groups["ME24-1001-0_ZC00_405"]
    assert len(revisions) == 1
    _revision, chosen_entry = revisions[0]
    assert chosen_entry.source_path == "ME24-1001-0_ZC00_405.xlsx"


def test_find_entries_with_unresolved_sashiban():
    """指番未解決エントリ（app.pyの警告表示対象）だけが抽出される。"""
    resolved_via_filename = LedgerEntry(
        package_name=_FLAT_FOLDER, source_path="PE25-9601-0_ZM00_405.xlsx",
        diff_list_rows=[], summary_values={},
    )
    resolved_via_folder = LedgerEntry(
        package_name=_FULL_FOLDER, source_path="ME24-1001-0_ZC00_405.xlsx",
        diff_list_rows=[], summary_values={},
    )
    unresolved_mistyped = LedgerEntry(
        package_name="dxf_diff_results_TypeA_NE24-0062-0",
        source_path="NE24-0062-0_ZM00_405l.xlsx",
        diff_list_rows=[], summary_values={},
    )

    result = find_entries_with_unresolved_sashiban(
        [resolved_via_filename, resolved_via_folder, unresolved_mistyped]
    )

    assert result == [unresolved_mistyped]
