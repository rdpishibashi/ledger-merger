"""仕様確認（無効化・2026-09）: DXF-diff-manager の Summary「図面統計」欄削除に伴う警告表示。

この機能は 2026-08 時点で「指番図面総数・流用率 [%]・新規作成率 [%] が 0 になる」
下流影響をユーザーに警告するために追加されたが、2026-09、ユーザー要求により
統合図面管理台帳.xlsx の Summary シート自体からこの3列を削除した
（tests/unit/test_master_ledger_builder.py::test_summary_headers_no_longer_include_removed_columns
参照）。値が失われる列がそもそも存在しなくなったため、この警告機能自体が
`app.py` から削除された（`merged_missing_drawing_stats` セッションキー・
対応する expander ともに廃止）。警告対象の問題自体が無くなったため、代替の
仕様テストは追加していない（新しい受入条件が無い）。このファイルはL2の記録
として保持し、削除はしない。
"""

import pytest
from streamlit.testing.v1 import AppTest

pytestmark = pytest.mark.skip(
    reason="2026-09、警告対象の「図面統計を持たない台帳」機能ごとapp.pyから削除された"
    "（Summaryシートの指番図面総数・流用率[%]・新規作成率[%]列を削除したため）。"
    "モジュールdocstring参照。"
)


def _run(**session_state_overrides):
    at = AppTest.from_file("app.py", default_timeout=10)
    for key, value in session_state_overrides.items():
        at.session_state[key] = value
    at.run()
    assert at.exception == []
    return at


def test_warning_shown_when_entries_missing_drawing_statistics():
    at = _run(
        final_zip_bytes=b"dummy-zip",
        merged_count=1,
        group_summary_count=0,
        merged_missing_folders=[],
        merged_unresolved_sashiban=[],
        merged_missing_drawing_stats=["PKG1 / NEWFORMAT.xlsx"],
        downloaded_once=False,
    )
    labels = [e.label for e in at.expander]
    assert any("図面統計を持たない台帳ファイル" in label for label in labels), labels


def test_warning_hidden_when_no_entries_missing_drawing_statistics():
    at = _run(
        final_zip_bytes=b"dummy-zip",
        merged_count=1,
        group_summary_count=0,
        merged_missing_folders=[],
        merged_unresolved_sashiban=[],
        merged_missing_drawing_stats=[],
        downloaded_once=False,
    )
    labels = [e.label for e in at.expander]
    assert not any("図面統計を持たない台帳ファイル" in label for label in labels), labels
