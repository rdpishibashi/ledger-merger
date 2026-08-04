"""不具合再発防止: 2026-07-29「統合実行」ボタンの色が成功後も primary のまま。

streamlit スキルの「状態に応じたボタンの色分け」パターンでは、統合が成功し
ダウンロードボタンが有効になった時点で「統合実行」ボタンは secondary（白背景）に
なるはずだが、`app.py` の `run = st.button("統合実行", type="primary", ...)` が
常に固定で `type="primary"` を指定していたため反映されていなかった。
`merge_done = "final_zip_bytes" in st.session_state` に基づく動的な `type` 計算に
修正（`app.py` 参照）。

`st.file_uploader` / `st.download_button` は `streamlit.testing.v1.AppTest` が
ウィジェット操作をサポートしないため、実際のアップロード操作は
`tests/regression/README.md` にも記載のとおりブラウザでのブラックボックス確認
（2026-07-29 claude-in-chrome で実施済み）に委ね、本テストは
`st.session_state` を直接シードして「統合実行」ボタンの `type` 遷移ロジックのみを
検証する。

2026-08、`merge_done` の判定に `not zip_files` を追加（統合成功後にZIPアップローダーへ
新しいファイルを追加すると、`type` が自動的に primary に戻る）。`st.file_uploader` の
戻り値は `AppTest` でシードできないため、この追加条件（新しいZIPが選択されている場合の
遷移）は本テストではなく実ブラウザでの確認に委ねる（`tests/regression/README.md` 参照）。
"""

from streamlit.testing.v1 import AppTest


def _run(**session_state_overrides):
    at = AppTest.from_file("app.py", default_timeout=10)
    for key, value in session_state_overrides.items():
        at.session_state[key] = value
    at.run()
    assert at.exception == []
    return at


def _run_button(at):
    matches = [b for b in at.button if b.label == "統合実行"]
    assert len(matches) == 1
    return matches[0]


def test_run_button_is_primary_before_first_merge():
    at = _run()
    assert _run_button(at).proto.type == "primary"


def test_run_button_becomes_secondary_once_merge_succeeded():
    at = _run(
        final_zip_bytes=b"dummy-zip",
        merged_count=1,
        group_summary_count=0,
        merged_missing_folders=[],
    )
    assert _run_button(at).proto.type == "secondary"


def test_run_button_reverts_to_primary_after_new_merge_reset():
    # 「新規統合の実行」クリック時のリセット（final_zip_bytes 等の pop）を再現。
    at = _run(use_last_master=True, master_bytes=b"dummy-master")
    assert _run_button(at).proto.type == "primary"
