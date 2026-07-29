"""不具合再発防止: 2026-07-29「統合台帳をダウンロード」ボタンの色がダウンロード後もprimaryのまま。

「統合実行」ボタン（`test_run_button_color_after_merge.py` 参照）と同じ状態色分け
パターンをダウンロードボタンにも適用すべきだったが、`app.py` の
`st.download_button(..., type="primary")` が固定値だったため、`downloaded_once`
が立った後も secondary（白背景）にならなかった。
`type="secondary" if download_done else "primary"` の動的計算＋ダウンロード検知時の
`st.rerun()` で修正（`app.py` 参照）。

`AppTest` は `st.download_button` の実際のクリック操作をサポートしないため、
`st.session_state["downloaded_once"]` を直接シードして `type` の分岐のみを検証する。
"""

from streamlit.testing.v1 import AppTest


def _run(**session_state_overrides):
    at = AppTest.from_file("app.py", default_timeout=10)
    at.session_state["final_zip_bytes"] = b"dummy-zip"
    at.session_state["merged_count"] = 1
    at.session_state["group_summary_count"] = 0
    at.session_state["merged_missing_folders"] = []
    for key, value in session_state_overrides.items():
        at.session_state[key] = value
    at.run()
    assert at.exception == []
    return at


def _download_button(at):
    matches = [e for e in at.get("download_button") if e.proto.label == "統合台帳をダウンロード"]
    assert len(matches) == 1
    return matches[0]


def test_download_button_is_primary_before_download():
    at = _run(downloaded_once=False)
    assert _download_button(at).proto.type == "primary"


def test_download_button_becomes_secondary_after_download():
    at = _run(downloaded_once=True)
    assert _download_button(at).proto.type == "secondary"
