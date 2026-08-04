"""不具合再発防止テスト: 2026-07-28 に報告された「有効な台帳ファイルが見つかりません
でした」不具合。

不具合の識別子: 2026-07-28 ユーザー報告（ZIP: ME24-1001-0.zip）

再現条件（修正前の挙動）:
    1. DXF-diff-manager の出力フォルダに `dxf図面`（元DXFを格納する非xlsxサブフォルダ）
       が付随するようになり、`find_ledger_files()` の「出力フォルダ＝サブフォルダを
       持たない葉フォルダ」という前提が崩れていた。サブフォルダを持つ出力フォルダは
       丸ごと評価対象から除外され、`dxf図面` 自体だけが唯一の走査対象となり、xlsxが
       無いため誤って「台帳が見つからなかったフォルダ」として報告されていた
       （29フォルダ中29フォルダとも本来の台帳が無視される事態）。
    2. さらに、DXF-diff-manager の Summary シートのラベル文言が変更されており
       （例:「削除図形数 合計」→「削除図形 総数」）、Ledger-merger 側の
       `SUMMARY_LABELS` 定数が追随できていなかったため、(1) を修正しても
       Summary の必須ラベルチェックで全滅していた。

修正後に保証したいこと:
    - 出力フォルダの判定は「サブフォルダの有無」ではなく「自分より深い階層にxlsxを
      含むフォルダが無い、xlsxを直下に持つフォルダ」という基準に一般化されている。
    - xlsxを含まないサブフォルダ（`dxf図面` 等）は走査対象にも missing 報告にも
      含まれない。
    - ZIP直下（ラッパー）に置かれた、有効フォーマットだが個別出力フォルダの台帳と
      重複する集約xlsxは、中間ラッパーとして無視される（entries にも missing にも
      現れない）。
    - DXF-diff-manager の現行 Summary ラベル文言が、Ledger-merger 自身の統合Excel
      列名（SUMMARY_LABELS）に正しくエイリアス変換される。

対応するフィクスチャ: tests/fixtures/dxf_diff_manager_output/
    （2026-07-28 の実データの一部。フィクスチャ構成は tests/unit/test_ledger_merger.py
    のモジュールdocstringを参照）
"""

import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "..", ".."))

from utils.ledger_finder import SUMMARY_LABELS, find_ledger_files

REAL_DATA_ROOT = os.path.join(
    os.path.dirname(__file__), "..", "..", "fixtures", "dxf_diff_manager_output"
)


def test_output_folder_with_non_xlsx_subfolder_is_detected():
    """`dxf図面` サブフォルダが付随していても、そのフォルダ自身の台帳は正しく検出され、
    `dxf図面` 自体は「台帳が見つからないフォルダ」に混入しない。"""
    entries, missing_folders = find_ledger_files(REAL_DATA_ROOT)

    zmf1_entries = [e for e in entries if e.package_name == "dxf_diff_results_TypeA_ME24-1001-0_ZMF1_405_01"]
    assert len(zmf1_entries) == 1
    assert len(zmf1_entries[0].diff_list_rows) == 2

    assert "dxf図面" not in missing_folders
    assert "PLACEHOLDER.txt" not in missing_folders


def test_wrapper_folder_loose_xlsx_excluded_from_entries_and_missing():
    """ZIP直下（ラッパー）に置かれた `ME24-1001-0_ZM00_405_all.xlsx` は、有効な
    台帳フォーマットであっても、より深い階層の各出力フォルダ自身にxlsxがあるため
    中間ラッパーとして無視され、entries にも missing_folders にも現れない
    （ユーザー確認済み: 個別出力フォルダの台帳と重複するため統合対象外とする仕様）。"""
    entries, missing_folders = find_ledger_files(REAL_DATA_ROOT)

    source_filenames = {os.path.basename(e.source_path) for e in entries}
    assert "ME24-1001-0_ZM00_405_all.xlsx" not in source_filenames

    wrapper_folder_name = os.path.basename(REAL_DATA_ROOT)
    assert wrapper_folder_name not in missing_folders


def test_summary_label_aliases_match_current_dxf_diff_manager_wording():
    """DXF-diff-manager の現行 Summary ラベル文言（例:「削除図形 総数」「アップロード図面
    図形総数」）が、Ledger-merger 自身の統合Excel列名（SUMMARY_LABELS、例:「削除図形数
    合計」「総図形数 合計」）に正しくエイリアス変換されることを確認する。"""
    entries, _missing_folders = find_ledger_files(REAL_DATA_ROOT)
    assert entries
    for entry in entries:
        assert all(label in entry.summary_values for label in SUMMARY_LABELS)
