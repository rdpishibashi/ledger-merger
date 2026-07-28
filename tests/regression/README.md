# 回帰テスト一覧

## bugfix（不具合再発防止）

| 不具合ID | 対応テスト | 修正内容の要点 |
|---|---|---|
| 2026-07-28 「有効な台帳ファイルが見つかりませんでした」 | `bugfix/test_output_folder_detection.py` | (1) `find_ledger_files()` の出力フォルダ判定を「サブフォルダの有無」から「自分より深い階層にxlsxを含むフォルダが無い、xlsxを直下に持つフォルダ」に一般化し、DXF-diff-manager出力に追加された `dxf図面` サブフォルダに対応。(2) DXF-diff-manager の Summary シートラベル文言の変更に追随するエイリアス変換（`_SOURCE_LABEL_ALIASES`）を追加。詳細は各テストのdocstring参照。 |

## spec（仕様確認）

| 受入条件 | 対応テスト | 要点 |
|---|---|---|
| 指番_モジュール_サイド単位のレビジョン横断集計Excelの生成 | `spec/test_group_summary_export.py` | ユーザー提供の実際の参照ファイル（`ME24-1001-0_ZC00_405_all.xlsx`）のSummaryシートの値と完全一致することを検証。Diff Listシートに `Diff Package`・合計列が含まれないことも確認 |
| 統合実行結果を単一ZIP「統合図面台帳.zip」にまとめる | `spec/test_unified_zip_bundle.py` | ZIP内に `図形変更量詳細.xlsx`・`統合図面管理台帳.xlsx`・`指番_モジュール_サイド別集計/` が固定名で含まれること、Masterシートが `Child`-`Parent` ペアで重複なく `Child` 昇順であることを検証 |

## フィクスチャ

`tests/fixtures/dxf_diff_manager_output/` — 2026-07-28 の不具合調査時に取得した実データの一部（DXF-diff-manager の現行出力形式をそのまま反映）。構成は `tests/unit/test_ledger_merger.py` のモジュールdocstring参照。

---
最終更新: 2026-07-28
