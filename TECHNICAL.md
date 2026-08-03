# Ledger-merger 技術仕様

## 概要

DXF-diff-manager の出力フォルダ群（ZIP化してアップロード）を再帰的に走査し、各台帳の
`Diff List` シートを集約する Streamlit アプリ。統合実行のたびに、単一ZIP
`統合図面台帳.zip` として以下の3種類を出力する。

| 出力 | 内容 |
|---|---|
| `図形変更量詳細.xlsx` | 全 Diff Package を1シートに集約した統合Excel（旧名 `統合_図面親子管理台帳_*.xlsx`）。毎回フレッシュに生成される。`Child` の前に Diff Package名から逆算した `Sashiban`/`Module`/`Side` 列、最終列に `Diff Package` 自体を持つ |
| `統合図面管理台帳.xlsx` | `Master`（`Child`-`Parent` ペア単位でユニーク化）・`Work Master`（指番ごとの `Child`-`Parent` ペア単位でユニーク化）・`Summary`（指番ごとの実行時点スナップショットを追記していくログ）の3シート構成。前回分をアップロードすると、Master・Work Masterは同じキーが `Recorded Date` の新しい方で上書き（前回の方が新しければ前回を保持）・それ以外は保持して蓄積、Summaryはキー単位のマージを行わず単純追記する |
| `指番_モジュール_サイド別集計/` | DXF-diff-manager のZIPダウンロードファイル名の命名規則（`dxf_diff_results_Type{A/B/C}_{指番}_{モジュール}_{サイド}_{リビジョン}`。2026-08以降リビジョンは省略される場合がある）から検出したグループ単位で、レビジョン横断のSummary+Diff Listを持つExcel（`{指番}_{モジュール}_{サイド}_all.xlsx`）を1グループ1ファイルで出力するフォルダ |

## ディレクトリ構成

```
Ledger-merger/
├── app.py                                # View層: Streamlit UI
├── utils/
│   ├── __init__.py
│   ├── ledger_finder.py                  # Model層: フォルダ走査・台帳検出
│   ├── ledger_merger.py                  # Model層: 統合Excel（図形変更量詳細.xlsx）生成
│   ├── group_summary_builder.py          # Model層: 指番_モジュール_サイド別集計の生成
│   └── master_ledger_builder.py          # Model層: 統合図面管理台帳.xlsx（Master）の生成・蓄積
├── tests/
│   ├── unit/                             # 各モジュールのユニットテスト
│   ├── regression/
│   │   ├── bugfix/                       # 不具合再発防止テスト
│   │   ├── spec/                         # 仕様確認テスト
│   │   └── README.md                     # 不具合ID/受入条件 ⇄ テストの対応表
│   └── fixtures/
│       └── dxf_diff_manager_output/      # 実データの一部（コミット済み、テスト用）
├── requirements.txt
├── README.md
└── TECHNICAL.md
```

| モジュール | 責務 | Streamlit依存 |
|---|---|---|
| `app.py` | 入力受付（ZIP・統合図面管理台帳.xlsx）、結果表示、ダウンロード | あり |
| `utils/ledger_finder.py` | フォルダ再帰走査、台帳ファイルの検出・読み込み | なし |
| `utils/ledger_merger.py` | `図形変更量詳細.xlsx`（bytes）の生成 | なし |
| `utils/group_summary_builder.py` | `指番_モジュール_サイド別集計` 各ファイル（bytes）の生成 | なし |
| `utils/master_ledger_builder.py` | `統合図面管理台帳.xlsx`（bytes）の生成・前回分とのマージ | なし |

Model層（`utils/`）は Streamlit に依存しない純粋関数として実装しており、
`tests/unit/` から UI なしで直接呼び出してテストできる。

## アーキテクチャ（データフロー）

```
[ZIPアップロード（複数可）]        [統合図面管理台帳.xlsx アップロード（必須）]
        │                                     │
        ▼                                     ▼
一時ディレクトリに展開                  read_master_rows(bytes)
(tempfile.TemporaryDirectory)         (utils/master_ledger_builder.py)
        │                                     │
        ▼                                     │
find_ledger_files(root_dir)                   │
(utils/ledger_finder.py)                      │
        │                                     │
        ▼                                     │
list[LedgerEntry] + folders_without_ledger     │
        │                                     │
        ├──────────────┬──────────────────────┤
        ▼              ▼                      ▼
build_merged_workbook  build_group_workbooks   build_master_workbook(entries,
(ledger_merger.py)     (group_summary_builder.py)  previous_master_rows)
        │              │                      (master_ledger_builder.py)
        ▼              ▼                      ▼
図形変更量詳細.xlsx     {指番}_{モジュール}_    統合図面管理台帳.xlsx
                       {サイド}_all.xlsx（複数）
        │              │                      │
        └──────────────┴──────────────────────┘
                        ▼
              zipfile へまとめる（app.py）
                        ▼
              統合図面台帳.zip
                        ▼
              st.download_button（ボタン1つ）
```

### `utils/ledger_finder.py`

- `find_ledger_files(root_dir) -> (list[LedgerEntry], list[str])`
  - `os.walk(root_dir)` で再帰的に走査し、**「自分より深い階層にxlsxを含むフォルダが
    無い、xlsxを直下に持つフォルダ」を「DXF-diff-manager 出力フォルダ」として扱う**
    （サブフォルダの有無そのものでは判定しない）。DXF-diff-manager の出力フォルダには
    元DXFを格納する `dxf図面` 等の非xlsxサブフォルダが付随する場合があるが、そのサブ
    フォルダ自体はxlsxを持たないため出力フォルダの深さ判定に影響しない。一方、ZIP展開
    時のラッパーフォルダ（直下にもxlsxがあるが、より深い階層の各出力フォルダにもxlsx
    がある場合）は中間ラッパーとして評価対象から除外する
    （2026-07-28 の不具合修正。旧仕様「サブフォルダを持たない葉フォルダのみ」は、
    DXF-diff-manager 出力に `dxf図面` サブフォルダが追加されたことで前提が崩れていた。
    詳細は `tests/regression/bugfix/test_output_folder_detection.py`）。
  - 各対象フォルダ内の `.xlsx` のうち、`~$` 始まりの一時ロックファイルと
    `NON_LEDGER_FILENAMES = {"diff_labels.xlsx", "unchanged_labels.xlsx"}`
    （DXF-diff-manager が出力する台帳以外の固定ファイル）を除外したものを台帳候補とする。
  - 候補ファイルを `openpyxl.load_workbook(path, data_only=True, read_only=True)` で開き、
    シート名に `Diff List` と `Summary` の両方が存在し、`Diff List` のヘッダーが
    `DIFF_LIST_HEADERS`（12列）と一致し、`Summary` シートに `SUMMARY_LABELS`（9項目、
    Ledger-merger自身の統合Excel列名）が **`_SOURCE_LABEL_ALIASES` 経由で** すべて
    存在するものだけを有効な台帳と判定する（ファイル名は不問）。開けない/壊れている
    ファイルや条件を満たさないファイルは単に `None` を返して読み飛ばす。
  - **`_SOURCE_LABEL_ALIASES`**: DXF-diff-manager の `Summary` シート側の実際の
    ラベル文言（例:「削除図形 総数」）と、Ledger-merger 自身の `SUMMARY_LABELS`
    （例:「削除図形数 合計」）の対応表。DXF-diff-manager 側でラベル文言が変更されても
    Ledger-merger の統合Excel列名は変えない設計。「総図形数 合計」「入力図面総数」は
    DXF-diff-manager のペアリング方式（Type A: `アップロード図面...` / Type B・C:
    `流用先図面...`）によって文言が変わるため複数エイリアスを許容する。旧バージョンの
    Ledger-merger が使っていた旧ラベル文言（削除図形数 合計 等）との互換性は無い
    （2026-07-28、ユーザー判断により意図的に非対応）。
  - **`OPTIONAL_SUMMARY_LABELS`**（2026-08追加）: `完全新規図面数`・`新規作成率 [%]`
    （DXF-diff-manager が2026-08にSummaryシートへ追加した2指標）。`_SOURCE_LABEL_ALIASES`
    には他のラベルと同様に含めて解決するが、**`SUMMARY_LABELS`（台帳判定の必須9項目）
    には含めない**——必須にすると、この2指標を持たない旧バージョンの台帳がすべて無効
    判定されてしまうため。存在すれば `summary_values` に取り込まれ、存在しなければ
    キー自体が無い（呼び出し側は `.get(key, 0)` 等で欠損を0として扱う）。
  - `Diff List` の各データ行のうち、**`Total Entities` が空欄（`None`）の行は除外**する。
    これは実際には差分抽出されていない図番ペアの関係記録（親子マスター管理用、
    `Summary` シートの集計にも含まれない）であり、統合 Diff List には含めない。
    除外後に1行も残らない場合は無効な台帳として扱う（`folders_without_ledger` に積む）。
  - 台帳と判定したファイルの**直接の親フォルダ名**を `LedgerEntry.package_name` とする。
  - 対象フォルダ内に有効な台帳が1件も見つからない場合、そのフォルダの**ベース名のみ**
    （パスやファイル名の詳細は含めない）を `folders_without_ledger` に積む（重複除去・出現順）。
    xlsxを一切含まないフォルダ（`dxf図面` 等）は判定対象にすらならないため、
    ここには積まれない。
  - 戻り値の `entries` は `package_name` でソート済み。

- `reconcile_missing_folders(entries, folders_without_ledger) -> list[str]`
  - `find_ledger_files()` は1入力ソース（1ZIP分の展開先ディレクトリ）ごとに独立して判定する。
    複数ZIPを集約する `app.py` 側で、同じフォルダ名が異なるZIPに存在し、片方には有効な
    台帳があり、もう片方には無い場合、単純な `list.extend()` の集約結果は
    「統合成功」と「台帳が見つからなかったフォルダ」の両方に同じ名前が矛盾して現れてしまう。
    この関数は `entries` に存在する `package_name` を `folders_without_ledger` から除外し、
    かつ複数ソースに渡る重複名も1つにまとめる（出現順維持）。**複数ZIPを集約する箇所では
    必ずこの関数を経由すること**（`app.py` の `if run:` ブロック末尾で呼んでいる）。

- `LedgerEntry`（dataclass）
  - `package_name: str` — DXF-diff-manager の出力フォルダ名
  - `source_path: str` — 元ファイルの絶対パス
  - `diff_list_rows: list` — `Diff List` のデータ行（ヘッダーと `Total Entities` が
    空欄の行を除く、12列のタプル）。常に `summary_values["差分抽出ペア数"]` と同数になる
    とは限らない（下記「Total Entities 空欄行の除外の正当性の根拠」の「注意」参照）。
  - `summary_values: dict` — `SUMMARY_LABELS`（Ledger-merger自身の列名）をキーとした
    `Summary` シートの値（`_SOURCE_LABEL_ALIASES` でエイリアス解決済み）

### `utils/ledger_merger.py`

- `build_merged_workbook(entries: list[LedgerEntry]) -> bytes`（出力ファイル: `図形変更量詳細.xlsx`）
  - `openpyxl.Workbook()` を直接操作してセル単位で書き込む
    （pandas + xlsxwriter のDataFrame経由ではなく、セルごとのフォント色制御が必要なため素のopenpyxlを採用）。
  - シート名は `Diff List` の1シートのみ。
  - 出力列（21列、`OUTPUT_HEADERS`）:
    `Sashiban, Module, Side, Child, Parent, Relation, Title, Subtitle, Recorded Date, Note, Deleted Entities, Added Entities, Diff Entities, Unchanged Entities, Total Entities, 削除図形数 合計, 追加図形数 合計, 変更図形数 合計, 図形総数 合計, 図形変更率 [%], Diff Package`
    （2026-07-31、ユーザー要望により `Child` の前に `Sashiban`/`Module`/`Side` を追加し、
    `Diff Package` を先頭から最終列へ移動。列位置は決め打ちではなく
    `OUTPUT_HEADERS.index(...)` で解決するため、列順を変えてもロジック側の修正は不要）。
    （2026-08、ユーザー要望により図面統計系4列〈変更なし図形数 合計・入力図面総数・
    差分抽出ペア数・流用率 [%]〉を削除しエンティティ統計5項目のみに絞った。あわせて
    「差分図形数 合計」→「変更図形数 合計」、「総図形数 合計」→「図形総数 合計」に
    表示名を変更。この5項目の対応表は `_MERGED_SUMMARY_COLUMNS`（正規キー→表示列名の
    タプル）として `ledger_merger.py` 側に定義し、`ledger_finder.SUMMARY_LABELS`
    〈台帳判定用の必須9項目〉とは分離した）。
  - `Sashiban`/`Module`/`Side` は `utils/group_summary_builder.parse_sashiban_module_side()`
    で、台帳ファイル名（`entry.source_path`）を主・Diff Package名（`entry.package_name`）を
    従として逆算する。どちらの命名規則にも一致しない場合は3列とも空欄（`None`）
    （`find_entries_with_unresolved_sashiban()` で検出でき、`app.py` に警告表示される）。
  - 各 `LedgerEntry` ブロック内:
    - `Diff Package`列（最終列、`DIFF_PACKAGE_COL`）: ブロック最初の行のみ黒字（`FF000000`）、以降は薄いグレー（`FFA6A6A6`）。
    - `Sashiban`/`Module`/`Side`〜`Total Entities`列: 元の `Diff List` 行をそのまま転記。`Recorded Date` 列のみ `number_format = "YYYY-MM-DD HH:MM:SS"`。
      `Deleted/Added/Diff/Unchanged/Total Entities`（`ENTITY_COLS`）は `#,##0` ＋中央揃い
      （`'n/a'` と数値が混在するため、表示位置を揃える目的。2026-07-28追加）。
    - Summary由来5項目列: ブロックの最初の行のみ値を記入。カウント系は `#,##0`、
      `図形変更率 [%]` は `0.00%` の `number_format` を適用（`PERCENT_LABELS`）。中央揃い。2行目以降は空欄。
  - ヘッダー行（1行目）は太字＋中央揃い、`freeze_panes = "A2"`。

### `utils/group_summary_builder.py`

指番・モジュール・サイドは**台帳ファイル名**（`{指番}_{モジュール}_{サイド}[_-suffix].xlsx`。
DXF-diff-manager `model/master_ledger.py` の `MASTER_FILENAME_PATTERN` と同一規則）を
**主**として決定する。Diff Package名（出力フォルダ名。ZIPダウンロード時にユーザーが
自由編集できるテキスト欄に由来し、モジュール/サイドが欠落しうる）は**従**とし、
**リビジョン番号のみ**フォルダ名から取得する（台帳ファイル名の末尾サフィックスは
リビジョン以外の自由文字列でありうるため）。フォルダ名の命名規則は
`dxf_diff_results_Type{A/B/C}_{指番}_{モジュール}_{サイド}_{リビジョン}`
（2026-07-28以前の手動命名 `dxf_diff_results_Pair{A/B/C}_...` にも対応。2026-08以降、
末尾の `_{リビジョン}` は省略される場合がある）。

- `parse_sashiban_module_side(package_name, filename=None) -> (sashiban, module, side) | (None, None, None)`
  - `filename`（`entry.source_path`）が `LEDGER_FILENAME_PATTERN` に一致すればそこから
    決定する（主）。一致しない、または `filename` 省略時は `package_name` を
    `SASHIBAN_MODULE_SIDE_PATTERN`（指番=`[A-Z]{2}\d{2}-\d{4}-\d`、モジュール/サイド=
    `[A-Z0-9]{4|3}` または `"na"`）で判定する（従）。どちらにも一致しない場合は
    `(None, None, None)`（`utils/ledger_merger.py`・`utils/master_ledger_builder.py` の
    `Sashiban`/`Module`/`Side`・`Work Master`・`Summary` 生成で使用）。
- `parse_group_and_revision(package_name, filename=None) -> (group_key, revision) | None`
  - `parse_sashiban_module_side()` と同じ優先順位で指番/モジュール/サイドを決定し
    `group_key`（`指番_モジュール_サイド`）を組み立てる。**レビジョンは常に
    `package_name` からのみ取得**（`SASHIBAN_MODULE_SIDE_PATTERN` の任意リビジョン群、
    無ければ `GROUP_REVISION_PATTERN` の緩いフォールバック、それも無ければ `None`）。
  - 指番/モジュール/サイドがどちらの規則にも一致しない場合のみ、`package_name` を
    `GROUP_REVISION_PATTERN`（groupを1文字列として緩く取り出す。リビジョン必須）で
    判定する最終フォールバックがある。これにより、指番書式に一致しない過去の
    手動命名フォルダ（例: `..._OK_01`）の解釈は変えない。
- `find_entries_with_unresolved_sashiban(entries) -> list[LedgerEntry]`
  - `parse_sashiban_module_side(entry.package_name, entry.source_path)` が失敗する
    エントリ（台帳ファイル名のミスタイプ等）を返す。`app.py` の警告表示に使う
    （2026-08-03、指番が黙って除外される不具合の再発防止として追加）。
- `aggregate_input_drawing_totals_by_sashiban(entries) -> dict[sashiban, int|float]`
  - `entries` を `group_entries()` でグルーピングし、各グループの「アップロード図面総数」
    TOTAL値（`{group_key}_all.xlsx` の Summaryシートと同じ計算）を指番ごとに合算する。
    命名規則に一致しないグループは対象外。値の欠損（`summary_values` にキーが無い等）は
    0として扱う（`_total_value`/`_revision_value` を経由せず独自に集計しているのは、
    テスト用の簡易フィクスチャ〈`summary_values={}`〉でも安全に動作させるため）。
    `utils/master_ledger_builder.compute_summary_rows()` の「指番図面総数」列に使用。
- `group_entries(entries) -> dict[group_key, list[(revision, LedgerEntry)]]`
  - **まず `package_name`（同一フォルダ）でグルーピングし、`Diff List` 内の最大
    `Recorded Date` が最も新しい候補を1つに絞ってから**、その勝者の指番/モジュール/
    サイド・レビジョンを決定する（差分抽出のやり直しで古い実行結果がフォルダに
    残っていた場合の取り違え防止。2026-07-28に実データで確認したケース: 図番抽出に
    失敗した古い実行結果〈`na_na` ファイル名〉と、後で成功した新しい実行結果が同じ
    フォルダに混在していた）。**この「フォルダ内で1つに絞ってから指番を決める」順序は
    重要**——指番決定を先にすると、同一フォルダ内の候補がファイル名の違いで別グループに
    分裂し、取り違え防止が機能しなくなる（2026-08-03、ファイル名優先化の際に一度
    この回帰を作り、`tests/regression/bugfix/test_sashiban_filename_fallback.py` で
    再発防止）。レビジョン昇順にソートする（リビジョン省略形〈`revision` が `None`〉の
    エントリは末尾に並べる。2026-08）。
- `aggregate_diff_list_by_child(revision_entries) -> list[tuple]`
  - グループ内の全レビジョンの `diff_list_rows` を `Child` ごとに集計する。
    `Deleted/Added/Diff/Unchanged/Total Entities` の5列は**全レビジョンにわたって単純合計**
    する（`'n/a'` の値は数値でないためスキップし、ある `Child` の全出現が `'n/a'` の列は
    そのまま `'n/a'` として残す）。非数値列（`Parent`/`Relation`/`Title`/`Subtitle`/
    `Recorded Date`/`Note`）は、`Recorded Date` が最も新しい行の値を採用する。
- `build_group_workbook(group_key, revision_entries) -> bytes`
  - `Summary` シート: `TOTAL` 列 + レビジョン列（`"01"`, `"02"`, ...）。**グループ内の
    全エントリがリビジョン省略形（`revision` が全件 `None`）の場合はレビジョン別の列を
    出さず `TOTAL` 列のみとする**（2026-08。一部だけ省略形が混在する場合〈通常は発生
    しないが命名規則の移行期に起こりうる〉は、見出しを `"-"` として列自体は残す）。
    行構成は DXF-diff-manager の Summary シート表記そのまま（削除図形 総数 等、Type A
    表記のみ対応。「既知の制約」参照）。11項目のうちカウント系8項目はレビジョンごとの
    値をそのまま横に並べ、`TOTAL` 列は単純合計。`図形変更率 [%]`・`流用率 [%]`・
    `新規作成率 [%]`（2026-08追加）の3項目のみ `TOTAL` 列は単純合計ではなく
    `TOTAL(分子)/TOTAL(分母)` で再計算する（ユーザー提供の実際の参照ファイルと
    完全一致することを `tests/regression/spec/test_group_summary_export.py` で検証済み。
    ただしこの参照ファイルは2026-08の2指標追加より前のものであり、対応する行の期待値は
    旧形式台帳の既定値である0を追加している）。「完全新規図面数」「新規作成率 [%]」を
    持たない旧バージョンの台帳（`entry.summary_values` にキーが無い）は
    `_revision_value()` の既定値により0として扱われる（例外にはならない）。
  - `Diff List` シート: `Diff Package` 列・Summary合計列は含めない（12列のみ）。
    `Deleted/Added/Diff/Unchanged/Total Entities` 列は `#,##0` ＋中央揃い。ヘッダー行も
    中央揃い。
- `build_group_workbooks(entries) -> dict[filename, bytes]`
  - 検出できた全グループについて `{group_key}_all.xlsx` を生成する。

### `utils/master_ledger_builder.py`

`統合図面管理台帳.xlsx` は `Master`→`Work Master`→`Summary` の3シート構成
（2026-07-31、`Work Master`・`Summary` を追加）。

#### Master（従来からのシート）

- `extract_unique_child_parent_rows(entries) -> dict[(child, parent), tuple]`
  - 全エントリの `diff_list_rows` から `(Child, Parent)` ペアでユニーク化する。
    **同じペアが複数エントリにまたがる場合は `Recorded Date` が最も新しい行を採用する**
    （2026-08、それまでの先勝ち〈`setdefault`〉から変更。同一 Child-Parent ペアが
    複数のDXF-diff-manager出力フォルダに異なる実行時刻で記録される実データケース
    〈例: 図面Aの新旧比較が異なる指番フォルダ双方の差分抽出対象になり、それぞれの
    実行時刻が異なる〉を確認したための修正。`_recorded_date_or_min()` により
    `Recorded Date` が `datetime` でない場合は `datetime.min` として扱う）。
- `read_master_rows(file_bytes) -> dict[(child, parent), tuple] | None`
  - アップロードされた `統合図面管理台帳.xlsx` の `Master` シートを読み込む。
    シート名・ヘッダー（`DIFF_LIST_HEADERS` と一致）が想定と異なる場合は `None`
    （`app.py` 側で警告表示し、今回分のみで作成する）。

#### Work Master（指番ごとのChild-Parentユニーク化）

- `WORK_MASTER_HEADERS`: `("Sashiban",) + (DIFF_LIST_HEADERS から Relation を除いたもの)`
  = `Sashiban, Child, Parent, Title, Subtitle, Recorded Date, Note, Deleted/Added/Diff/
  Unchanged/Total Entities`（12列）。
- `_extract_unique_work_master_entries(entries) -> dict[(sashiban, child, parent), (sashiban, diff_list_row)]`
  - `parse_sashiban_module_side(entry.package_name, entry.source_path)` で解決できた
    エントリのみ対象（台帳ファイル名・出力フォルダ名のどちらからも指番を逆算できない
    エントリは除外。`find_entries_with_unresolved_sashiban()` で検出可能）。
    `(sashiban, Child, Parent)` でユニーク化し、
    `extract_unique_child_parent_rows` と同じ規則で `Recorded Date` が最も新しい行を
    採用する。`diff_list_row` は `DIFF_LIST_HEADERS` 12列（`Relation` を含む）のまま
    保持する内部専用の中間表現——`extract_unique_work_master_rows()`（Work Master出力用に
    `Relation` を除いた形へ変換）と `compute_summary_rows()`（`Relation` を用いて
    完全新規図面数・差分ペア総数を算出）の両方から共有される（2026-08新設）。
- `extract_unique_work_master_rows(entries) -> dict[(sashiban, child, parent), tuple]`
  - `_extract_unique_work_master_entries()` の結果から `Relation` 列を除いた
    `WORK_MASTER_HEADERS` 12列の辞書に変換する公開関数。シグネチャ・戻り値の形状は
    2026-08の変更前後で変わらない。
- `read_work_master_rows(file_bytes) -> dict[(sashiban, child, parent), tuple] | None`
  - `read_master_rows` と同様のパターンで `Work Master` シートを読み込む。シートが
    存在しない（旧バージョンで作成したファイル等）・ヘッダー不一致の場合は `None`
    （`app.py` 側で今回分のみのフォールバックに使う。Masterと異なりこの場合は
    警告を出さない——旧バージョンからの移行時に毎回警告が出るのを避けるため）。

`build_master_workbook()` は Master・Work Masterともに、`_merge_by_recorded_date()`
ヘルパーで前回分（アップロードされた辞書）と今回分（`extract_unique_*`の戻り値）を
マージする：**同じキーが両方に存在する場合は `Recorded Date` が新しい方（同日時なら
今回分）を採用**し、前回のみに存在するキーはそのまま保持、今回のみに存在するキーは
追加する（2026-08、それまでの「同じキーは常に今回データで上書き」から変更。古い実行
結果を含むZIPを後から統合した際に、既に蓄積済みの新しいデータを誤って古いデータで
上書きしてしまう実データケースを防ぐための修正）。

#### Summary（指番ごとの実行時点スナップショットの追記ログ）

Master/Work Masterとは異なり、**キー単位のマージ・上書きを行わない**。今回アップロード
したZIPのデータのみから算出した指番ごとの集計行を、アップロードされた前回分の末尾に
単純追記する（同じ指番の行が実行回数分だけ増えていく）。同一指番の推移を実行日時ごとに
追えるようにするための設計（2026-07-31、ユーザーとの確認により決定。詳細は
`tests/regression/README.md` 参照）。

- `SUMMARY_HEADERS`: `指番, 削除図形総数, 追加図形総数, 変更図形総数, 図形総数,
  図形変更率 [%], 差分ペア総数, 完全新規図面数, 指番図面総数, 流用率 [%],
  新規作成率 [%], 日付`（12列、列名はユーザー指定によりMaster/Work Masterと異なり
  日本語。2026-08、`完全新規図面数`・`新規作成率 [%]` をDXF-diff-manager自身の
  Summaryシートと同じ相対位置〈差分ペア総数の直下・流用率[%]の直下〉に追加。
  10列→12列）。
- `BRAND_NEW_RELATION = "完全新規図面"`（DXF-diff-manager `model/master_ledger.py` の
  `save_master_to_bytes()` の `brand_new_mask` と同じ完全一致判定。`"完全新規図面-changed"`
  は含まない）。
- `compute_summary_rows(entries, run_timestamp) -> list[tuple]`（指番昇順）
  - **今回のZIP入力（`entries`）のみ**から算出する（Work Masterの累積データは使わない。
    「指番図面総数」がその性質上、今回入力のみからしか算出できない〈後述〉ため、
    他の列もスコープを合わせている）。
  - `_extract_unique_work_master_entries(entries)` の結果（`Relation` を含む中間形式）を
    指番ごとにグルーピングし、以下を算出:
    - `削除図形総数`/`追加図形総数`/`図形総数`: それぞれ `Deleted`/`Added`/`Total
      Entities` 列の合計（非数値〈`'n/a'` 等〉は0として扱う。`_numeric_sum()`）。
      **完全新規図面の行も含めたまま合計する**（DXF-diff-manager 自身の集計〈`master_df`
      全体からの合計〉と範囲を揃えるため。完全新規図面の `Deleted`/`Diff`/`Unchanged`
      は元々 `'n/a'` のため自動的に除外され、`Added`/`Total` は数値としてそのまま
      合算される）。
    - `変更図形総数` = `削除図形総数` + `追加図形総数`。
    - `図形変更率 [%]` = `変更図形総数` / `図形総数`（`図形総数`が0なら0.0）。
    - `完全新規図面数`（2026-08追加） = `Relation == BRAND_NEW_RELATION` の行の
      `Child` ユニーク数。
    - `差分ペア総数` = `Relation != BRAND_NEW_RELATION` の行に限定した、その指番の
      ユニーク `(Child, Parent)` ペア数（**2026-08、完全新規図面を除外するよう変更**。
      それまでは完全新規図面の行もペア数に含めていた。DXF-diff-manager自身の
      「差分抽出ペア数」〈`status == 'complete'` のペア数、完全新規図面を含まない〉と
      定義を揃えるための変更）。
    - `指番図面総数` = `aggregate_input_drawing_totals_by_sashiban(entries)`
      （`utils/group_summary_builder.py`）の値。**`指番_モジュール_サイド別集計`
      フォルダの `*_all.xlsx` と同じ計算を経由するため、この列は今回のZIP入力のみ
      からしか算出できない**（`指番_モジュール_サイド別集計` 自体が毎回フレッシュに
      生成される仕様のため）。
    - `流用率 [%]` = `差分ペア総数` / `指番図面総数`（`指番図面総数`が0なら0.0）。
    - `新規作成率 [%]`（2026-08追加） = `完全新規図面数` / `指番図面総数`
      （`指番図面総数`が0なら0.0）。
  - `run_timestamp` は `build_master_workbook()` 呼び出し時刻（`datetime.now()`）。
    同一実行内の全指番行で共通。
- `read_summary_rows(file_bytes) -> list[tuple]`
  - アップロードされた `統合図面管理台帳.xlsx` の `Summary` シートの全行をそのまま
    読み込む（`Master`/`Work Master` と異なりキー付き辞書ではなく単純なリスト。
    ユニーク化しないため）。シートが存在しない・ヘッダー不一致の場合は空リスト
    `[]` を返す（`None` ではなく、警告なしで「今回分のみ追記」にフォールバックする
    設計のため）。

#### build_master_workbook

- `build_master_workbook(entries, previous_master_rows=None, previous_work_master_rows=None, previous_summary_rows=None) -> bytes`
  - Master: `extract_unique_child_parent_rows(entries)` と `previous_master_rows` を
    `_merge_by_recorded_date()` でマージ（同じキーは `Recorded Date` が新しい方）、
    `Child` 昇順でシート `Master` に書き込む。
  - Work Master: `extract_unique_work_master_rows(entries)` と `previous_work_master_rows`
    を同様に `_merge_by_recorded_date()` でマージ、`(Sashiban, Child)` 昇順でシート
    `Work Master` に書き込む。
  - Summary: `compute_summary_rows(entries, ...)` の結果を `previous_summary_rows`
    （リスト）の**末尾に追記**、シート `Summary` に書き込む（マージ・ソートなし、
    アップロードされた行 → 今回の行、の順）。
  - 3シートとも `Deleted/Added/Diff/Unchanged/Total Entities` 列（Summaryは
    `削除/追加/変更図形総数`・`図形総数`・`差分ペア総数`・`完全新規図面数`・
    `指番図面総数`）は `#,##0` ＋中央揃い、`図形変更率 [%]`・`流用率 [%]`・
    `新規作成率 [%]` は `0.00%` ＋中央揃い、ヘッダー行も中央揃い。

### `app.py`

- 入力は2つ:
  1. `st.file_uploader(accept_multiple_files=True, type=["zip"])` — DXF-diff-manager
     出力フォルダ群のZIP（複数可）。各ZIPは `tempfile.TemporaryDirectory()` に展開し
     `find_ledger_files` に渡す。key は `f"zip_uploader_{zip_uploader_version}"`
     （`st.session_state["zip_uploader_version"]`）。「新規統合の実行」時にこの
     バージョンをインクリメントしてウィジェットを再生成し、選択済みファイルを
     クリアする（2026-07-29）。
  2. `st.file_uploader(type=["xlsx"])` — 前回ダウンロードした `統合図面管理台帳.xlsx`
     （`master_upload`）。`統合図面管理台帳.xlsx` に蓄積機能を持たせた2026-07-28の変更で
     必須化した（任意にすると「アップロードし忘れて履歴が失われる」事故が起きうるため）。
     ただし `st.session_state["use_last_master"]` が `True` の間はこのアップローダー
     自体を表示せず、直近の統合成功時に保存した `st.session_state["master_bytes"]`
     を `master_bytes_override` として自動的に使用する（再アップロード不要。
     「新規統合の実行」ボタンで有効化。2026-07-29）。
- 「統合実行」ボタンは `disabled=not has_input`
  （`has_input = bool(zip_files) and (master_upload is not None or
  master_bytes_override is not None)`）に加え、`type` を
  `"secondary" if "final_zip_bytes" in st.session_state else "primary"` で動的に
  切り替える（2026-07-29。統合成功でダウンロードボタンが有効になった時点で白背景に
  戻すため。`streamlit` スキルの「状態に応じたボタンの色分け」を参照）。統合成功時
  （`else` 分岐の末尾）で `st.rerun()` を呼び、同一run内での描画順序の制約
  （ボタンは処理より前に描画される）を回避して即座に反映する。台帳0件の失敗分岐は
  従来どおり `st.rerun()` を呼ばない（エラーメッセージがフラッシュ的に消えるのを
  避けるため、意図的に非対称）。
- ZIPファイル1件単位で `st.empty()` + `progress(ratio, text=...)` の
  進捗バーを表示し、完了後 `placeholder.empty()` で消す。
- 「統合実行」ボタン押下時に全ZIPから `LedgerEntry` と `folders_without_ledger`
  を集約し、1件以上あれば `build_merged_workbook` / `build_master_workbook` /
  `build_group_workbooks` をそれぞれ実行し、`zipfile.ZipFile` で単一ZIP
  `統合図面台帳.zip`（ファイル名固定）にまとめて `st.session_state` に保存する。
  `master_bytes`（統合図面管理台帳.xlsxのbytes）も同時に `st.session_state` に保存し、
  「新規統合の実行」時の自動使用元を兼ねる。ダウンロードボタンは1つ
  （「統合台帳をダウンロード」）のみ。
- アップロードされた（または自動使用された）`統合図面管理台帳.xlsx` から
  `read_master_rows`/`read_work_master_rows`/`read_summary_rows` をそれぞれ読み込む。
  `read_master_rows` が `None` の場合のみ `st.warning()` を表示（`Master` が読めない
  ことをファイル不正の代表的なシグナルとして扱う）。`read_work_master_rows` が
  `None`（Work Masterシートが無い旧バージョン等）・`read_summary_rows` が空リスト
  （Summaryシートが無い等）の場合は、警告なしにそれぞれ今回分のみで作成・追記する
  （旧バージョンからの移行時に警告が重複して出るのを避けるため）。
- ZIP自体が開けない場合（`zipfile.BadZipFile`）は `st.warning()` で即時表示する
  （フォルダ名一覧とは別枠。ZIPはフォルダではないため）。
- 結果（成功件数・台帳が無いフォルダ名一覧・ダウンロードボタン）は `st.session_state` を介して
  描画するため、再実行（rerun）後も表示が保持される。フォルダ名一覧は
  `st.expander("⚠️ 台帳ファイルが見つからなかったフォルダ（N件）")` の中に
  **フォルダ名のみ**（パスやファイル名は含めない）で表示する。
- ダウンロードボタン（`st.download_button`）も「統合実行」ボタンと同様、
  `type = "secondary" if st.session_state.get("downloaded_once") else "primary"`
  で動的に色分けする。戻り値が `True` の run で `st.session_state["downloaded_once"]
  = True` を立てたうえで `st.rerun()` し、同一run内での描画順序の制約（ボタンは
  クリック処理より前に描画される）を回避して即座に白背景へ反映する（2026-07-29、
  「統合実行」ボタンの色分け修正時に見落としていた箇所。ダウンロード自体は
  ブラウザ側で即時開始されるため `st.rerun()` で中断されない）。`downloaded_once`
  が `True` の間、「新規統合の実行」ボタン（`type="primary"`）を表示する。押すと
  `use_last_master=True`・`zip_uploader_version` を +1 にした上で、結果表示系の
  キー（`final_zip_bytes`/`merged_count`/`merged_missing_folders`/
  `merged_unresolved_sashiban`/`group_summary_count`/`downloaded_once`）のみを pop し、`master_bytes` は
  保持したまま `st.rerun()` する（自動使用の入力元として次回に持ち越すため）。
- `use_last_master` モード中に「別のファイルをアップロードし直す」ボタンを押すと
  `use_last_master=False` にして `st.rerun()` し、通常の手動アップロードUIに戻る
  （ユーザー確認により追加したエスケープハッチ。2026-07-29）。
- `.streamlit/config.toml` でテーマ（`primaryColor = "#0365C0"` 等、DXF-diff-manager と同一）を指定。
  単一アクションのツールのため `st.subheader("Step N: ...")` のような連番見出しは使わない
  （セクション名のみの見出しに留める）。ボタンの `width` は既定（`'content'`）のまま
  指定しない（`streamlit` スキルの「ボタンの幅は既定のまま使う」を参照。2026-07-28、
  従来 `width='stretch'` を指定していたダウンロードボタンから撤去）。

### セッション状態

| キー | 内容 |
|---|---|
| `final_zip_bytes` | `統合図面台帳.zip` のバイト列 |
| `master_bytes` | 直近の統合成功時に生成した `統合図面管理台帳.xlsx` のバイト列。「新規統合の実行」時の自動使用元を兼ねる（2026-07-29） |
| `merged_count` | 統合した Diff Package 数 |
| `merged_missing_folders` | 台帳ファイルが見つからなかったフォルダ名（ベース名のみ）の一覧 |
| `merged_unresolved_sashiban` | 台帳は見つかったが指番を特定できなかったエントリの一覧（`"{package_name} / {ファイル名}"` 形式。2026-08-03） |
| `group_summary_count` | `指番_モジュール_サイド別集計` に生成されたファイル数 |
| `downloaded_once` | ダウンロードボタンが押されたかどうか。`True` の間だけ「新規統合の実行」ボタンを表示（2026-07-29） |
| `use_last_master` | `True` の間、統合図面管理台帳.xlsxのアップロード欄を隠し `master_bytes` を自動使用（2026-07-29） |
| `zip_uploader_version` | ZIPアップローダーの `key` に使うバージョン番号。「新規統合の実行」時に+1してウィジェットを再生成しクリアする（2026-07-29） |

## macOS の ZIP（Finder/ditto）への対応

macOS の Finder「圧縮」や `ditto` コマンドで ZIP 化すると、各フォルダをミラーする
`__MACOSX/<同名フォルダ>/` が作られ、その中にリソースフォークファイル
（`._図面親子管理台帳.xlsx` 等、本物と同じ拡張子）が入る。このミラーフォルダは本物と
同じベース名を持つため、対策をしないと「台帳はあるが無効なファイルしかないフォルダ」
として誤って `folders_without_ledger` に積まれてしまう
（実際に起きた不具合: 29フォルダすべてが正しく処理されているにもかかわらず、
同じ29フォルダ名が「台帳ファイルが見つからなかったフォルダ」にも表示された）。

`find_ledger_files`（`utils/ledger_finder.py`）では `os.walk` の `dirnames` から
`__MACOSX` を都度除外し、そのサブツリーを走査しないようにしている。加えて
`._` で始まるファイルも候補から除外している（`__MACOSX` を経由しない単独の
リソースフォークファイルへの保険）。`tests/unit/test_ledger_merger.py` の
`test_macosx_mirror_folder_not_reported_as_missing` で回帰確認している。

### Windows で作成した ZIP は対象外（検証済み）

`__MACOSX` ミラーフォルダは macOS（Finder/`ditto`）特有の挙動であり、Windows の ZIP機能
（エクスプローラー「送る→圧縮(zip形式)フォルダー」、`Compress-Archive`、7-Zip、WinRAR等）
には存在しないため、同種の誤検出は発生しない。

Windows には別の既知の問題（エクスプローラー標準のZIP機能がファイル名にUTF-8フラグを
付けず、システムのコードページ＝日本語WindowsならCP932でエンコードすることがあり、
Linux上の `zipfile`（CP437既定）で展開すると日本語ファイル名が文字化けする）があるが、
これを実際に再現して検証した結果、**本アプリには実害がない**ことを確認済み:
- `.xlsx` 拡張子はASCIIのため文字化けの影響を受けず、台帳候補の判定は機能する
- ファイル内容（Excelバイナリ）自体は文字化けと無関係で `openpyxl` は正常に読める
- `Diff Package` 名はファイル名ではなく**親フォルダ名**（ASCII）から取得しているため
  影響を受けない

そのため、Windows由来のZIPに対する追加の文字コード対策コードは実装していない。

## 入力方式に関する制約

Streamlit Cloud にはサーバー側のローカルファイルシステムがなく、ブラウザの
`st.file_uploader` はフォルダ構造（相対パス）を保持できないため、対象フォルダを
ZIP 化してアップロードする方式に統一している（ローカルパス直接指定は当初実装したが、
入力方式が2つ並ぶとユーザーが混乱するため廃止した）。`統合図面管理台帳.xlsx` は
単一ファイルのため通常の `st.file_uploader(type=["xlsx"])` で受け付ける。

## 「Total Entities 空欄行の除外」の正当性の根拠

`utils/ledger_finder.py` が `Total Entities` 空欄行を除外する判断は、DXF-diff-manager側の
ソースコード（`DXF-diff-manager/model/master_ledger.py` の `update_parent_child_master()` /
`save_master_to_bytes()`）まで遡って裏付けを取った:

- `update_parent_child_master()` は新規行追加・既存行更新のいずれでも、エンティティ数
  （Deleted/Added/Diff/Unchanged/Total Entities）は `if entity_counts:`（実際に差分抽出が
  成功したペアのみ `entity_counts` を保持）の場合だけ設定する。Title/Subtitle/Relation/
  Recorded Date はこれと無関係に設定されるため、**「Total Entities が空欄」＝「このペアは
  実際には差分抽出されていない（親子マスター管理のための関係記録のみ）」が DXF-diff-manager
  自身の設計として保証されている**。
- `save_master_to_bytes()` の Summary 集計（削除図形数 合計 等）は
  `pd.to_numeric(master_df[col], errors='coerce').sum(skipna=True)` で計算されるため、
  空欄行・`'n/a'`行は自動的に合計から除外される。

**注意（厳密な保証ではない点）**: 2026-07-09時点の実データ（29フォルダ）では
「除外後の行数 ＝ Summary の『差分抽出ペア数』」が全件一致していたが、これは
DXF-diff-manager の実装から推測される強い相関であり、`差分抽出ペア数`
（`pairs` リストから算出）と `Total Entities` 非空行数（`master_df` から算出）は
別々のデータソースに基づくため、常に一致する保証ではない（2026-07-28に取得した別の
実データでは、`差分抽出ペア数` より `Total Entities` 非空行数の方が多いフォルダが
複数見つかっている——ペアリング済みだが実際の比較処理で例外的に一部だけ成功した等の
ケースと推測される。`utils/ledger_finder.py` 自体の判定は `Total Entities is None`
のみに基づいており、この潜在的なズレの影響を受けない）。

## 複数ZIP集約時のフォルダ名の矛盾防止

複数ZIPをアップロードした際、`find_ledger_files()` は1ZIPごとに独立して呼ばれるため、
同じフォルダ名が異なるZIPに存在し、片方には有効な台帳があり、もう片方には無い場合、
単純に結果を `extend()` するだけでは「統合成功」と「台帳が見つからなかったフォルダ」の
両方に同じ名前が矛盾して表示されてしまう（実際に発生を確認し修正済み）。
`reconcile_missing_folders()`（`utils/ledger_finder.py`）で、`entries` に存在する
`package_name` を `folders_without_ledger` から除外し、複数ソースに渡る重複名も除去する。
`app.py` の `if run:` ブロック末尾、`build_merged_workbook` 呼び出しの前に必ず通すこと。

## 既知の制約

- 同名の `Diff Package`（出力フォルダ名）が複数の ZIP に存在し、**両方とも**有効な台帳を
  持つ場合は、自動的な重複排除・名前の disambiguation は行わない。両方とも別ブロックとして
  統合される（実データにも `... 2` のような重複名フォルダが存在するため、意図的にこの挙動
  とした）。片方だけ有効な場合は上記の `reconcile_missing_folders()` で矛盾を解消する。
- `Diff List` のヘッダー列構成が変わった場合は `DIFF_LIST_HEADERS`
  （`utils/ledger_finder.py`）の更新が必要。`Summary` シートのラベル文言が変わった場合は
  `_SOURCE_LABEL_ALIASES` にエイリアスを追加する（`SUMMARY_LABELS` 自体・Ledger-merger
  側の統合Excel列名は変更しない設計）。
- `utils/group_summary_builder.py` の `Summary` シート行ラベルは DXF-diff-manager の
  Type A（`アップロード図面...`）表記に固定している。Type B/C（`流用先図面...`）の
  グループが混在する場合、値は正しく計算されるが、行ラベルの表記はType Aのまま表示される
  （既知の制約。現状の実データはすべてType Aのため未対応）。
- `utils/group_summary_builder.py` のグルーピングは、DXF-diff-manager のZIPダウンロード
  ファイル名の命名規則（`dxf_diff_results_Type{A/B/C}_{指番}_{モジュール}_{サイド}_
  {リビジョン}`〈リビジョンは2026-08以降省略可〉、または旧手動命名の `Pair{A/B/C}`）に
  一致するフォルダ名のみを対象とする。一致しないフォルダは `指番_モジュール_サイド別集計`
  の対象外となるが、`図形変更量詳細.xlsx`・`統合図面管理台帳.xlsx` には引き続き含まれる。
- 台帳が見つからない理由（ファイル形式不一致・読み込み失敗・本当に存在しない等）は
  区別せず、フォルダ名のみで一括して報告する（意図的な簡略化）。
- `統合図面管理台帳.xlsx` の蓄積はユーザーが毎回正しいファイルを再アップロードすることに
  依存する（サーバー側に状態を持たない設計のため）。誤って古いバージョンや無関係な
  ファイルをアップロードした場合、`read_master_rows()` がヘッダー不一致で `None` を返し
  今回分のみで作成されるため、履歴を静かに失うことはないが、**間違った履歴を正として
  上書きしてしまう可能性はチェックできない**。
- `Summary` シートはキー単位のマージを行わず単純追記するため、**同じ統合図面台帳.zipを
  誤って複数回アップロード→ダウンロードすると、同じ内容の行が重複して増え続ける**
  （Master/Work Masterのような重複排除は働かない。意図的な設計——実行履歴として
  すべて残すことを優先しているため）。
- `Summary` の「指番図面総数」・「流用率 [%]」・「新規作成率 [%]」は今回アップロードした
  ZIPのみから算出する（`指番_モジュール_サイド別集計` フォルダ自体が毎回フレッシュに
  生成される仕様のため）。他の列（削除/追加/変更図形総数・図形総数・差分ペア総数・
  完全新規図面数）も同じスコープに揃えており、Work Masterの累積データからは算出しない。
  そのため、複数回に分けて同じ指番のZIPをアップロードした場合、Summaryの各行は
  「その回にアップロードした分だけ」の値になる（Work Master自体は正しく累積される）。
- **`SUMMARY_HEADERS` は2026-08に10列から12列（`完全新規図面数`・`新規作成率 [%]` 追加）に
  変わったため、旧バージョンで出力した `統合図面管理台帳.xlsx`（10列の `Summary`）を
  アップロードすると、`read_summary_rows()` がヘッダー不一致で空リストを返し、
  `Summary` シートの過去の履歴行は引き継がれない**（警告は出ず、今回分のみで新規作成
  される。`Master`/`Work Master` はヘッダー変更が無いため蓄積に影響しない。移行時に
  一度だけ発生する既知の非互換——ユーザー承認済み）。

## 依存パッケージ

- `streamlit>=1.40.0`
- `openpyxl>=3.1.0`

## システム要件

- Python 3.10+
- Streamlit Cloud 実行可（ローカルファイルシステム不要、ZIPアップロードで対応）

## トラブルシューティング

[README.md](README.md) の「よくある問題」を参照。

---
最終更新: 2026-08-03
