"""仕様確認: 2026-07-29「新規統合の実行」フロー。

受入条件（ユーザー要求）:
- 「統合台帳をダウンロード」実行後にのみ「新規統合の実行」ボタンを表示する。
- 「新規統合の実行」を押すと、ZIPアップロード欄はクリアされ、統合図面管理台帳.xlsx は
  直前に作成したものを自動使用し、再アップロードを求めない。
- 自動使用中でも「別のファイルをアップロードし直す」で手動アップロードに戻せる
  （ユーザー確認により追加したエスケープハッチ）。

`st.file_uploader` の実アップロード操作は `AppTest` がサポートしないため、
`st.session_state` を直接シードしてUIの分岐条件（表示/非表示、キャプション文言）を
検証する。実際のアップロード〜ダウンロード〜新規統合の一連操作は
`tests/regression/README.md` 記載のとおり、2026-07-29 に claude-in-chrome の
ブラウザ自動操作で実施済み（ZIP複数・自動使用マージ・エスケープハッチの
組み合わせを含む）。

2026-08、「新規統合の実行」の表示条件に `not zip_files` を追加（ダウンロード後に
新しいZIPが選択された状態では、次の操作を「統合実行」1つに絞るため隠す）。
`zip_files` は `AppTest` でシードできないため、この追加条件は本テストではなく
実ブラウザでの確認に委ねる（`tests/regression/README.md` 参照）。
"""

from streamlit.testing.v1 import AppTest


def _run(**session_state_overrides):
    at = AppTest.from_file("app.py", default_timeout=10)
    for key, value in session_state_overrides.items():
        at.session_state[key] = value
    at.run()
    assert at.exception == []
    return at


def test_new_merge_button_hidden_before_download():
    at = _run(
        final_zip_bytes=b"dummy-zip",
        merged_count=1,
        group_summary_count=0,
        merged_missing_folders=[],
        downloaded_once=False,
    )
    assert "新規統合の実行" not in [b.label for b in at.button]


def test_new_merge_button_shown_after_download():
    at = _run(
        final_zip_bytes=b"dummy-zip",
        merged_count=1,
        group_summary_count=0,
        merged_missing_folders=[],
        downloaded_once=True,
    )
    assert "新規統合の実行" in [b.label for b in at.button]


def test_use_last_master_hides_manual_uploader_and_shows_escape_hatch():
    at = _run(use_last_master=True, master_bytes=b"dummy-master")
    info_texts = [i.value for i in at.info]
    assert any("自動的に使用します" in t for t in info_texts)
    assert "別のファイルをアップロードし直す" in [b.label for b in at.button]
    # 手動アップロード用のキャプションは表示されない
    caption_texts = [c.value for c in at.caption]
    assert not any("前回ダウンロードした統合図面管理台帳.xlsx" in t for t in caption_texts)


def test_without_use_last_master_shows_manual_uploader_caption():
    at = _run()
    caption_texts = [c.value for c in at.caption]
    assert any("前回ダウンロードした統合図面管理台帳.xlsx" in t for t in caption_texts)
