# 回帰テスト一覧

## bugfix（不具合再発防止）

| 不具合ID | 対応テスト | 修正内容の要点 |
|---|---|---|
| 2026-07-28 「有効な台帳ファイルが見つかりませんでした」 | `bugfix/test_output_folder_detection.py` | (1) `find_ledger_files()` の出力フォルダ判定を「サブフォルダの有無」から「自分より深い階層にxlsxを含むフォルダが無い、xlsxを直下に持つフォルダ」に一般化し、DXF-diff-manager出力に追加された `dxf図面` サブフォルダに対応。(2) DXF-diff-manager の Summary シートラベル文言の変更に追随するエイリアス変換（`_SOURCE_LABEL_ALIASES`）を追加。詳細は各テストのdocstring参照。 |

## フィクスチャ

`tests/fixtures/dxf_diff_manager_output/` — 2026-07-28 の不具合調査時に取得した実データの一部（DXF-diff-manager の現行出力形式をそのまま反映）。構成は `tests/unit/test_ledger_merger.py` のモジュールdocstring参照。

---
最終更新: 2026-07-28
