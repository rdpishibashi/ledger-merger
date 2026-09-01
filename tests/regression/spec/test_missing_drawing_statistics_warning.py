"""仕様確認: 2026-09、DXF-diff-manager の Summary「図面統計」欄削除に伴う警告表示。

受入条件（ユーザーが下流影響〈指番図面総数・流用率 [%]・新規作成率 [%] が 0 に
なる〉を承知のうえで全削除を選択したため、Ledger-merger 側は台帳を検出対象外に
しないことに加えて、値が失われる台帳をユーザーに明示する）:
- 「図面統計」を持たない台帳が1件以上あった場合、その一覧を警告表示する。
- 該当が無ければ警告は表示しない。

`st.file_uploader` の実アップロード操作は `AppTest` がサポートしないため、
`tests/regression/spec/test_new_merge_flow.py` と同様に `st.session_state` を
直接シードしてUIの表示条件のみを検証する。
"""

from streamlit.testing.v1 import AppTest


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
