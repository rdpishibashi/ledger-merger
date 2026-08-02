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
| 図形変更量詳細.xlsxにSashiban/Module/Side列を追加、Diff Packageを最終列へ移動 | `tests/unit/test_ledger_merger.py::test_merged_workbook_structure` | 25列構成（先頭3列がSashiban/Module/Side、最終列がDiff Package）と、各列の値が`parse_sashiban_module_side()`の逆算結果と一致することを検証 |
| 統合図面管理台帳.xlsxにWork Master（指番ごとのChild-Parentユニーク化）・Summary（指番ごとの実行時点スナップショット追記ログ）シートを追加 | `spec/test_unified_zip_bundle.py`、`tests/unit/test_master_ledger_builder.py` | Work Masterが`(指番,Child,Parent)`で重複なく指番→Child昇順であること、Summaryがキー単位でマージされず前回分の末尾に単純追記されること（同じ指番の行が実行回数分増える）、Summaryの各列の算出式（削除/追加/変更図形総数・図形総数・図形変更率[%]・差分ペア総数・指番図面総数・流用率[%]）を検証。2026-07-31、当初「指番ごとの小計行をWork Masterに追加」する設計だったが、使い勝手の観点からユーザー判断により撤回し、代わりにSummaryシートとして独立させた経緯あり |
| DXF-diff-manager が2026-08にZIPダウンロードファイル名末尾のリビジョン番号を省略するようになった仕様変更への追随 | `spec/test_2026_08_dxf_diff_manager_spec_update.py` | リビジョン省略形（`..._ZC00_405`、末尾に`_数字`が無い）でも指番/モジュール/サイドを逆算でき、グループ集計Summaryはレビジョン別列を出さずTOTAL列のみになることを検証。あわせて、Master・Work Masterの蓄積マージ（前回台帳との統合／今回アップロード内での重複解決）が「今回データによる無条件上書き」ではなく「Recorded Dateが新しい方を採用」に変わったことを、前回の方が新しいケース・今回の方が新しいケースの両方向で検証 |
| DXF-diff-manager Summaryシートに追加された「完全新規図面数」「新規作成率 [%]」を統合図面管理台帳・指番_モジュール_サイド別集計に反映、差分ペア総数から完全新規図面を除外 | `tests/unit/test_master_ledger_builder.py::test_compute_summary_rows_excludes_brand_new_from_pair_count`、`spec/test_group_summary_export.py` | 統合図面管理台帳Summary（`SUMMARY_HEADERS`10列→12列）はRelation='完全新規図面'の行をChildユニーク数で「完全新規図面数」として集計し「差分ペア総数」から除外すること、エンティティ統計（削除/追加/変更/図形総数）は完全新規図面の行も含めて合計することを検証。指番_モジュール_サイド別集計Summary（`SUMMARY_ROWS`9行→11行）は、この2指標を持たない旧形式台帳では0として扱われること（例外にならないこと）を、実データフィクスチャの期待値に0を追加する形で検証 |

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
最終更新: 2026-08-02
