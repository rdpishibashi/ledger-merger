# 回帰テスト一覧

## bugfix（不具合再発防止）

| 不具合ID | 対応テスト | 修正内容の要点 |
|---|---|---|
| 2026-07-28 「有効な台帳ファイルが見つかりませんでした」 | `bugfix/test_output_folder_detection.py` | (1) `find_ledger_files()` の出力フォルダ判定を「サブフォルダの有無」から「自分より深い階層にxlsxを含むフォルダが無い、xlsxを直下に持つフォルダ」に一般化し、DXF-diff-manager出力に追加された `dxf図面` サブフォルダに対応。(2) DXF-diff-manager の Summary シートラベル文言の変更に追随するエイリアス変換（`_SOURCE_LABEL_ALIASES`）を追加。詳細は各テストのdocstring参照。 |
| 2026-07-29 「統合実行」ボタンの色が成功後もprimaryのまま | `bugfix/test_run_button_color_after_merge.py` | `app.py` の「統合実行」ボタンが常に固定で `type="primary"` だったため、統合成功でダウンロードボタンが有効になっても白背景（secondary）にならなかった。`merge_done = "final_zip_bytes" in st.session_state` に基づく動的 `type` 計算＋成功時 `st.rerun()` で修正。 |
| 2026-07-29 「統合台帳をダウンロード」ボタンの色がダウンロード後もprimaryのまま | `bugfix/test_download_button_color_after_download.py` | 上記と同じ色分けパターンをダウンロードボタン自体には適用し忘れていた。`type = "secondary" if st.session_state.get("downloaded_once") else "primary"` の動的計算＋ダウンロード検知時の `st.rerun()` で修正。 |

## spec（仕様確認）

| 受入条件 | 対応テスト | 要点 |
|---|---|---|
| 指番_モジュール_サイド単位のレビジョン横断集計Excelの生成 | `spec/test_group_summary_export.py` | ユーザー提供の実際の参照ファイル（`ME24-1001-0_ZC00_405_all.xlsx`）のSummaryシートの値と完全一致することを検証。Diff Listシートに `Diff Package`・合計列が含まれないことも確認 |
| 統合実行結果を単一ZIP「統合図面台帳.zip」にまとめる | `spec/test_unified_zip_bundle.py` | ZIP内に `図形変更量詳細.xlsx`・`統合図面管理台帳.xlsx`・`指番_モジュール_サイド別集計/` が固定名で含まれること、Masterシートが `Child`-`Parent` ペアで重複なく `Child` 昇順であることを検証 |
| 「統合台帳をダウンロード」実行後に「新規統合の実行」を表示し、直前作成の統合図面管理台帳.xlsxを自動使用する | `spec/test_new_merge_flow.py` | ダウンロード前は「新規統合の実行」ボタンが非表示、ダウンロード後（`downloaded_once`）に表示されることを検証。`use_last_master` モード時は手動アップロード用キャプションが消え、自動使用の案内メッセージと「別のファイルをアップロードし直す」エスケープハッチボタンが表示されることを検証 |

## App.py（View層）の状態遷移テストについて

`bugfix/test_run_button_color_after_merge.py`・`spec/test_new_merge_flow.py` は
`streamlit.testing.v1.AppTest` を使用する。`AppTest` は `st.file_uploader` /
`st.download_button` のウィジェット操作をサポートしないため、これらのテストは
`st.session_state` を直接シードしてボタンの `type`（色）や表示/非表示の分岐条件のみを
検証する。実際のZIP・Excelアップロード〜統合実行〜ダウンロード〜新規統合の一連操作は
2026-07-29、claude-in-chromeによるブラウザ自動操作で別途動作確認済み（結果は
このテストファイルには含まれない一過性の確認のため、再現が必要な場合は
`streamlit` スキル§12またはclaude-in-chromeで同様の手順を実施すること）。

## フィクスチャ

`tests/fixtures/dxf_diff_manager_output/` — 2026-07-28 の不具合調査時に取得した実データの一部（DXF-diff-manager の現行出力形式をそのまま反映）。構成は `tests/unit/test_ledger_merger.py` のモジュールdocstring参照。

---
最終更新: 2026-07-29
