"""utils.ledger_finder のユニットテスト（実データフィクスチャ使用）。

`tests/fixtures/dxf_diff_manager_output`（コミット済み）は 2026-07-28 に実際に報告
された不具合の調査時に取得した実データの一部で、DXF-diff-manager の現行の出力形式
（Summary シートのラベル文言、`dxf図面` サブフォルダ付きの出力フォルダ構成）を
そのまま反映している。

フィクスチャ構成:
    dxf_diff_manager_output/
    ├── ME24-1001-0_ZM00_405_all.xlsx
    │       ZIP直下（ラッパー）に置かれた、有効な台帳フォーマットだが統合対象外の
    │       集約ファイル。より深い階層の各出力フォルダにもxlsxがあるため、
    │       中間ラッパーとして無視されるべき（ユーザー確認済みの仕様）。
    ├── dxf_diff_results_TypeA_ME24-1001-0_ZMF1_405_01/
    │       有効な台帳（2行）+ dxf図面/（xlsxを含まないサブフォルダ）付き。
    ├── dxf_diff_results_TypeA_ME24-1001-0_ZMB1_405_01/
    │       有効な台帳（3行）。
    ├── dxf_diff_results_TypeA_ME24-1001-0_ZMF2_405_01/
    │       diff_labels.xlsx / unchanged_labels.xlsx のみで有効な台帳が無い
    │       （「台帳が見つからなかったフォルダ」として報告されるべき）。
    └── dxf_diff_results_TypeA_ME24-1001-0_ZC00_405_01〜04/
            "ME24-1001-0_ZC00_405"（指番_モジュール_サイド）1グループ・4レビジョン分。
            `_02` のみ有効な台帳が2件（ME24-1001-0_ZC00_405.xlsx と
            ME24-1001-0_na_na.xlsx、後者は図番抽出に失敗した古い実行結果の残骸）。
            utils.group_summary_builder のグルーピング・重複排除・レビジョン横断
            集計の回帰テスト（tests/regression/spec/test_group_summary_export.py）に使う。

このファイルは元 `tests/unit/test_ledger_merger.py` から移設したもの（2026-09-02）。
同ファイルは `utils/ledger_merger.py`（図形変更量詳細.xlsx の生成。ユーザー要求により
削除）のテストと、`utils/ledger_finder.py` のテストが同居しており、モジュール削除に
伴ってファイルごと消すと ledger_finder 側の検証まで失われるため、こちらへ移した。
"""

import os
import shutil
import sys

import openpyxl

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", ".."))

from utils.ledger_finder import LedgerEntry, find_ledger_files, reconcile_missing_folders

REAL_DATA_ROOT = os.path.join(
    os.path.dirname(__file__), "..", "fixtures", "dxf_diff_manager_output"
)
SAMPLE_LEDGER_DIR = os.path.join(
    REAL_DATA_ROOT, "dxf_diff_results_TypeA_ME24-1001-0_ZMF1_405_01"
)
SAMPLE_LEDGER_FILENAME = "ME24-1001-0_ZMF1_405.xlsx"

MISSING_LEDGER_FOLDER_NAME = "dxf_diff_results_TypeA_ME24-1001-0_ZMF2_405_01"


def test_find_ledger_files_count_matches_actual_files():
    """フィクスチャ内の有効な台帳は7件（ZMF1_405_01, ZMB1_405_01, ZC00_405_01〜04。
    ただし ZC00_405_02 のみ2件）。ZM00_405_all.xlsx（ラッパー直下の集約ファイル）と
    ZMF2_405_01（台帳無し）は entries に含まれない。"""
    entries, missing_folders = find_ledger_files(REAL_DATA_ROOT)

    assert len(entries) == 7
    assert {e.package_name for e in entries} == {
        "dxf_diff_results_TypeA_ME24-1001-0_ZMF1_405_01",
        "dxf_diff_results_TypeA_ME24-1001-0_ZMB1_405_01",
        "dxf_diff_results_TypeA_ME24-1001-0_ZC00_405_01",
        "dxf_diff_results_TypeA_ME24-1001-0_ZC00_405_02",
        "dxf_diff_results_TypeA_ME24-1001-0_ZC00_405_03",
        "dxf_diff_results_TypeA_ME24-1001-0_ZC00_405_04",
    }
    assert missing_folders == [MISSING_LEDGER_FOLDER_NAME]


def test_non_ledger_filenames_excluded_and_not_reported_as_missing():
    """diff_labels.xlsx / unchanged_labels.xlsx は台帳候補から除外され、
    台帳が別途存在するフォルダは「台帳が見つからないフォルダ」にも出てこない。"""
    entries, missing_folders = find_ledger_files(REAL_DATA_ROOT)

    ledger_paths = {entry.source_path for entry in entries}
    assert all(os.path.basename(p) not in {"diff_labels.xlsx", "unchanged_labels.xlsx"} for p in ledger_paths)

    found_names = {e.package_name for e in entries}
    assert found_names.isdisjoint(missing_folders)


def test_folder_with_only_non_ledger_files_is_reported_as_missing_by_name_only(tmp_path):
    """diff_labels.xlsx / unchanged_labels.xlsx しか無いフォルダは「台帳が見つからない
    フォルダ」として basename のみで報告され、台帳が存在するフォルダは混在しても無視されない。"""
    ok_dir = tmp_path / "dxf_diff_results_TypeC_OK_01"
    ok_dir.mkdir()
    shutil.copy(
        os.path.join(SAMPLE_LEDGER_DIR, SAMPLE_LEDGER_FILENAME),
        ok_dir / SAMPLE_LEDGER_FILENAME,
    )

    missing_dir = tmp_path / "dxf_diff_results_TypeC_MISSING_01"
    missing_dir.mkdir()
    shutil.copy(os.path.join(SAMPLE_LEDGER_DIR, "diff_labels.xlsx"), missing_dir / "diff_labels.xlsx")
    shutil.copy(os.path.join(SAMPLE_LEDGER_DIR, "unchanged_labels.xlsx"), missing_dir / "unchanged_labels.xlsx")

    entries, missing_folders = find_ledger_files(str(tmp_path))

    assert len(entries) == 1
    assert entries[0].package_name == "dxf_diff_results_TypeC_OK_01"
    assert missing_folders == ["dxf_diff_results_TypeC_MISSING_01"]


def test_rows_without_diff_stats_are_excluded(tmp_path):
    """Diff List には「差分抽出ペア数」と同数の行（実際に差分抽出された行）だけが
    含まれる。Total Entities が空欄の行（差分未抽出の図番ペアの関係記録）は除外される。

    現行のコミット済みフィクスチャには空欄行を含む実データが無いため、
    DXF-diff-manager の現行フォーマットに沿った最小限の台帳ファイルをその場で
    合成して検証する（決定的なテストにするため）。"""
    from openpyxl import Workbook

    path = tmp_path / "dxf_diff_results_TypeX_ME00-0000-0_01" / "ME00-0000-0_ZZ00_000.xlsx"
    path.parent.mkdir(parents=True)

    wb = Workbook()
    diff_ws = wb.active
    diff_ws.title = "Diff List"
    diff_ws.append([
        "Child", "Parent", "Relation", "Title", "Subtitle", "Recorded Date", "Note",
        "Deleted Entities", "Added Entities", "Diff Entities", "Unchanged Entities",
        "Total Entities",
    ])
    diff_ws.append(["C1", "P1", "RevUp", "T", "S", None, None, 1, 2, 3, 4, 10])
    diff_ws.append(["C2", "P2", "親子", "T", "S", None, None, None, None, None, None, None])
    diff_ws.append(["C3", "P3", "RevUp", "T", "S", None, None, 5, 6, 11, 7, 25])

    summary_ws = wb.create_sheet("Summary")
    for row in [
        ("削除図形 総数", 6), ("追加図形 総数", 8), ("変更（追加+削除）図形 総数", 14),
        ("変更なし図形 総数", 11), ("アップロード図面 図形総数", 35),
        ("図形変更率 [%]", 0.4), ("アップロード図面総数", 3),
        ("差分抽出ペア数", 2), ("流用率 [%]", 0.66),
    ]:
        summary_ws.append(row)

    wb.save(path)

    entries, _missing_folders = find_ledger_files(str(tmp_path))
    assert len(entries) == 1
    entry = entries[0]

    assert len(entry.diff_list_rows) == entry.summary_values["差分抽出ペア数"] == 2
    assert all(row[-1] is not None for row in entry.diff_list_rows)
    assert [row[0] for row in entry.diff_list_rows] == ["C1", "C3"]


def test_filtered_rows_entity_sums_match_summary_exactly():
    """除外後に残った行の Deleted/Added/Diff/Unchanged/Total Entities の合計が、
    Summary シートの対応する合計値と1件単位の差もなく完全に一致することを、
    件数だけでなく数値レベルで検証する。完全新規図面の行は該当列が 'n/a'（文字列）に
    なるため、DXF-diff-manager 自身の集計と同様に数値以外はスキップして合計する。"""
    from utils.ledger_finder import DIFF_LIST_HEADERS

    cols = {h: i for i, h in enumerate(DIFF_LIST_HEADERS)}
    label_by_col = {
        "Deleted Entities": "削除図形数 合計",
        "Added Entities": "追加図形数 合計",
        "Diff Entities": "差分図形数 合計",
        "Unchanged Entities": "変更なし図形数 合計",
        "Total Entities": "総図形数 合計",
    }

    entries, _missing_folders = find_ledger_files(REAL_DATA_ROOT)
    assert entries

    for entry in entries:
        for col_name, summary_label in label_by_col.items():
            computed = sum(
                row[cols[col_name]] for row in entry.diff_list_rows
                if isinstance(row[cols[col_name]], (int, float))
            )
            assert computed == entry.summary_values[summary_label], (
                f"{entry.package_name}: {summary_label} 不一致"
            )


def test_filtered_rows_preserve_original_order_and_values():
    """除外後に残る行は、元の Diff List シートの該当行を順序・値ともに保持している
    （並び替えやデータ欠落が無いことの往復確認）。

    唯一の例外がエンティティ6列の `"n/a"`→`0` 正規化（2026-09、出力を数値で統一する
    ユーザー要求。`_normalize_entity_values()`）と、旧形式（12列、"Unchanged Offset
    Entities" 列が無い）実データ行を新形式13列に揃えるための末尾への 0 埋め
    （2026-09-18、_find_diff_list_rows() の後方互換パディングと同じ処理。
    このテストのサンプルファイル自体が旧形式のため必要）。
    """
    from utils.ledger_finder import ENTITY_LABELS, DIFF_LIST_HEADERS, TOTAL_COL

    entity_cols = [DIFF_LIST_HEADERS.index(label) for label in ENTITY_LABELS]

    def normalized(row):
        row = list(row)
        if len(row) < len(DIFF_LIST_HEADERS):
            row = row + [0] * (len(DIFF_LIST_HEADERS) - len(row))
        for col in entity_cols:
            if row[col] == 'n/a':
                row[col] = 0
        return tuple(row)

    entries, _missing_folders = find_ledger_files(REAL_DATA_ROOT)
    for entry in entries:
        wb = openpyxl.load_workbook(entry.source_path, data_only=True, read_only=True)
        try:
            src_rows = list(wb["Diff List"].iter_rows(values_only=True))[1:]
        finally:
            wb.close()
        expected = [normalized(row) for row in src_rows if row[TOTAL_COL] is not None]
        assert list(entry.diff_list_rows) == expected


def test_reconcile_missing_folders_removes_names_found_elsewhere():
    """異なる入力ソース（複数ZIP等）の結果を集約した際、同名フォルダが片方で台帳あり・
    もう片方で台帳なしと判定されても、「台帳が見つからなかったフォルダ」一覧には
    矛盾して現れない（成功した名前を優先する）。重複した欠落名も1つにまとめる。"""
    entries = [LedgerEntry(package_name="A", source_path="x", diff_list_rows=[], summary_values={})]
    missing = ["A", "B", "B", "C"]

    reconciled = reconcile_missing_folders(entries, missing)

    assert reconciled == ["B", "C"]


def test_macosx_mirror_folder_not_reported_as_missing(tmp_path):
    """macOS Finder/ditto で ZIP 化すると同名の __MACOSX/<folder>/._ファイル名 ミラーが
    作られる。これが本物のフォルダと同名の偽フォルダとして誤検出されないことを確認する
    （実際に発生した不具合: 台帳が見つかるフォルダ名がそのまま「見つからなかったフォルダ」
    にも重複して表示されていた）。"""
    real_dir = tmp_path / "dxf_diff_results_TypeC_OK_01"
    real_dir.mkdir()
    shutil.copy(
        os.path.join(SAMPLE_LEDGER_DIR, SAMPLE_LEDGER_FILENAME),
        real_dir / SAMPLE_LEDGER_FILENAME,
    )

    mirror_dir = tmp_path / "__MACOSX" / "dxf_diff_results_TypeC_OK_01"
    mirror_dir.mkdir(parents=True)
    (mirror_dir / f"._{SAMPLE_LEDGER_FILENAME}").write_bytes(b"\x00\x05\x16\x07")  # AppleDouble ダミー

    entries, missing_folders = find_ledger_files(str(tmp_path))

    assert len(entries) == 1
    assert entries[0].package_name == "dxf_diff_results_TypeC_OK_01"
    assert missing_folders == []
