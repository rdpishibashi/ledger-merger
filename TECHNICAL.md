# Ledger-merger 技術仕様

## 概要

DXF-diff-manager の出力フォルダ群（ZIP化してアップロード）を再帰的に走査し、各台帳の
`Diff List` シートを集約する Streamlit アプリ。統合実行のたびに、単一ZIP
`統合図面台帳.zip` として以下の3種類を出力する。

| 出力 | 内容 |
|---|---|
| `図形変更量詳細.xlsx` | 全 Diff Package を1シートに集約した統合Excel（旧名 `統合_図面親子管理台帳_*.xlsx`）。毎回フレッシュに生成される |
| `統合図面管理台帳.xlsx` | `Child`-`Parent` ペア単位でユニーク化した `Master` シートのみのExcel。前回分をアップロードすると同じペアは上書き、それ以外は保持して蓄積する |
| `指番_モジュール_サイド別集計/` | DXF-diff-manager のZIPダウンロードファイル名の命名規則（`dxf_diff_results_Type{A/B/C}_{指番}_{モジュール}_{サイド}_{リビジョン}`）から検出したグループ単位で、レビジョン横断のSummary+Diff Listを持つExcel（`{指番}_{モジュール}_{サイド}_all.xlsx`）を1グループ1ファイルで出力するフォルダ |

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
  - 出力列（22列、`OUTPUT_HEADERS`）:
    `Diff Package, Child, Parent, Relation, Title, Subtitle, Recorded Date, Note, Deleted Entities, Added Entities, Diff Entities, Unchanged Entities, Total Entities, 削除図形数 合計, 追加図形数 合計, 差分図形数 合計, 変更なし図形数 合計, 総図形数 合計, 図形変更率 [%], 入力図面総数, 差分抽出ペア数, 流用率 [%]`
  - 各 `LedgerEntry` ブロック内:
    - A列（`Diff Package`）: ブロック最初の行のみ黒字（`FF000000`）、以降は薄いグレー（`FFA6A6A6`）。
    - B〜M列: 元の `Diff List` 行をそのまま転記。`Recorded Date` 列のみ `number_format = "YYYY-MM-DD HH:MM:SS"`。
      `Deleted/Added/Diff/Unchanged/Total Entities`（`ENTITY_COLS`）は `#,##0` ＋中央揃い
      （`'n/a'` と数値が混在するため、表示位置を揃える目的。2026-07-28追加）。
    - N〜V列（Summaryの9項目）: ブロックの最初の行のみ値を記入。カウント系は `#,##0`、
      `図形変更率 [%]` と `流用率 [%]` は `0.00%` の `number_format` を適用。中央揃い。2行目以降は空欄。
  - ヘッダー行（1行目）は太字＋中央揃い、`freeze_panes = "A2"`。

### `utils/group_summary_builder.py`

DXF-diff-manager のZIPダウンロードファイル名の命名規則
（`dxf_diff_results_Type{A/B/C}_{指番}_{モジュール}_{サイド}_{リビジョン}`。2026-07-28以前の
手動命名 `dxf_diff_results_Pair{A/B/C}_...` にも対応）に依存する。

- `parse_group_and_revision(package_name) -> (group_key, revision) | None`
  - フォルダ名から正規表現でグループキー（`指番_モジュール_サイド`）とレビジョンを取り出す。
    一致しないフォルダ（この命名規則に従っていない過去データ等）は対象外（`None`）。
- `group_entries(entries) -> dict[group_key, list[(revision, LedgerEntry)]]`
  - `package_name` でグルーピングし、レビジョン昇順にソートする。
  - **同一フォルダ（同一 `package_name`）に複数の有効な台帳がある場合、`Diff List` 内の
    最大 `Recorded Date` が最も新しいものだけを採用する**（差分抽出のやり直しで古い
    実行結果がフォルダに残っていた場合の取り違え防止。2026-07-28に実データで確認した
    ケース: 図番抽出に失敗した古い実行結果〈`na_na` ファイル名〉と、後で成功した新しい
    実行結果が同じフォルダに混在していた）。
- `aggregate_diff_list_by_child(revision_entries) -> list[tuple]`
  - グループ内の全レビジョンの `diff_list_rows` を `Child` ごとに集計する。
    `Deleted/Added/Diff/Unchanged/Total Entities` の5列は**全レビジョンにわたって単純合計**
    する（`'n/a'` の値は数値でないためスキップし、ある `Child` の全出現が `'n/a'` の列は
    そのまま `'n/a'` として残す）。非数値列（`Parent`/`Relation`/`Title`/`Subtitle`/
    `Recorded Date`/`Note`）は、`Recorded Date` が最も新しい行の値を採用する。
- `build_group_workbook(group_key, revision_entries) -> bytes`
  - `Summary` シート: `TOTAL` 列 + レビジョン列（`"01"`, `"02"`, ...）。行構成は
    DXF-diff-manager の Summary シート表記そのまま（削除図形 総数 等、Type A 表記のみ
    対応。「既知の制約」参照）。カウント系9項目のうち7項目はレビジョンごとの値を
    そのまま横に並べ、`TOTAL` 列は単純合計。`図形変更率 [%]`・`流用率 [%]` の2項目のみ
    `TOTAL` 列は単純合計ではなく `TOTAL(分子)/TOTAL(分母)` で再計算する（ユーザー提供の
    実際の参照ファイルと完全一致することを
    `tests/regression/spec/test_group_summary_export.py` で検証済み）。
  - `Diff List` シート: `Diff Package` 列・Summary9項目の合計列は含めない（12列のみ）。
    `Deleted/Added/Diff/Unchanged/Total Entities` 列は `#,##0` ＋中央揃い。ヘッダー行も
    中央揃い。
- `build_group_workbooks(entries) -> dict[filename, bytes]`
  - 検出できた全グループについて `{group_key}_all.xlsx` を生成する。

### `utils/master_ledger_builder.py`

- `extract_unique_child_parent_rows(entries) -> dict[(child, parent), tuple]`
  - 全エントリの `diff_list_rows` から `(Child, Parent)` ペアでユニーク化する
    （同じペアが複数エントリにまたがる場合はどれか1件を採用。データはどれも同じはず
    という前提）。
- `read_master_rows(file_bytes) -> dict[(child, parent), tuple] | None`
  - アップロードされた `統合図面管理台帳.xlsx` の `Master` シートを読み込む。
    シート名・ヘッダー（`DIFF_LIST_HEADERS` と一致）が想定と異なる場合は `None`
    （`app.py` 側で警告表示し、今回分のみで作成する）。
- `build_master_workbook(entries, previous_master_rows=None) -> bytes`
  - 今回のユニーク化データと `previous_master_rows`（アップロードされた前回分、無ければ
    `None`）をマージする。**同じ `(Child, Parent)` は今回のデータで上書き**、前回のみに
    存在するペアはそのまま保持する（蓄積）。`Child` 昇順でソートし、単一シート `Master`
    （`DIFF_LIST_HEADERS` 12列）として書き込む。`Deleted/Added/Diff/Unchanged/Total
    Entities` 列は `#,##0` ＋中央揃い、ヘッダー行も中央揃い。

### `app.py`

- 入力は2つ、どちらも必須:
  1. `st.file_uploader(accept_multiple_files=True, type=["zip"])` — DXF-diff-manager
     出力フォルダ群のZIP（複数可）。各ZIPは `tempfile.TemporaryDirectory()` に展開し
     `find_ledger_files` に渡す。
  2. `st.file_uploader(type=["xlsx"])` — 前回ダウンロードした `統合図面管理台帳.xlsx`
     （`master_upload`）。`統合図面管理台帳.xlsx` に蓄積機能を持たせた2026-07-28の変更で
     必須化した（任意にすると「アップロードし忘れて履歴が失われる」事故が起きうるため）。
- 「統合実行」ボタンは `type="primary"` + `disabled=not has_input`
  （`has_input = bool(zip_files) and master_upload is not None`）。両方揃うまでグレー
  アウトし、色を手動で変更するコードは書かない（テーマの `primaryColor` に委ねる。
  `streamlit` スキルの「UIテーマ・ボタン・セクション構成」を参照）。
- ZIPファイル1件単位で `st.empty()` + `progress(ratio, text=...)` の
  進捗バーを表示し、完了後 `placeholder.empty()` で消す。
- 「統合実行」ボタン押下時に全ZIPから `LedgerEntry` と `folders_without_ledger`
  を集約し、1件以上あれば `build_merged_workbook` / `build_master_workbook` /
  `build_group_workbooks` をそれぞれ実行し、`zipfile.ZipFile` で単一ZIP
  `統合図面台帳.zip`（ファイル名固定）にまとめて `st.session_state` に保存する。
  ダウンロードボタンは1つ（「統合台帳をダウンロード」）のみ。
- アップロードされた `統合図面管理台帳.xlsx` が読み込めない場合（`read_master_rows`
  が `None`）は `st.warning()` で表示し、今回分のデータのみで `Master` を作成する。
- ZIP自体が開けない場合（`zipfile.BadZipFile`）は `st.warning()` で即時表示する
  （フォルダ名一覧とは別枠。ZIPはフォルダではないため）。
- 結果（成功件数・台帳が無いフォルダ名一覧・ダウンロードボタン）は `st.session_state` を介して
  描画するため、再実行（rerun）後も表示が保持される。フォルダ名一覧は
  `st.expander("⚠️ 台帳ファイルが見つからなかったフォルダ（N件）")` の中に
  **フォルダ名のみ**（パスやファイル名は含めない）で表示する。
- `.streamlit/config.toml` でテーマ（`primaryColor = "#0365C0"` 等、DXF-diff-manager と同一）を指定。
  単一アクションのツールのため `st.subheader("Step N: ...")` のような連番見出しは使わない
  （セクション名のみの見出しに留める）。ボタンの `width` は既定（`'content'`）のまま
  指定しない（`streamlit` スキルの「ボタンの幅は既定のまま使う」を参照。2026-07-28、
  従来 `width='stretch'` を指定していたダウンロードボタンから撤去）。

### セッション状態

| キー | 内容 |
|---|---|
| `final_zip_bytes` | `統合図面台帳.zip` のバイト列 |
| `merged_count` | 統合した Diff Package 数 |
| `merged_missing_folders` | 台帳ファイルが見つからなかったフォルダ名（ベース名のみ）の一覧 |
| `group_summary_count` | `指番_モジュール_サイド別集計` に生成されたファイル数 |

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
  {リビジョン}`、または旧手動命名の `Pair{A/B/C}`）に一致するフォルダ名のみを対象とする。
  一致しないフォルダは `指番_モジュール_サイド別集計` の対象外となるが、`図形変更量詳細.xlsx`
  ・`統合図面管理台帳.xlsx` には引き続き含まれる。
- 台帳が見つからない理由（ファイル形式不一致・読み込み失敗・本当に存在しない等）は
  区別せず、フォルダ名のみで一括して報告する（意図的な簡略化）。
- `統合図面管理台帳.xlsx` の蓄積はユーザーが毎回正しいファイルを再アップロードすることに
  依存する（サーバー側に状態を持たない設計のため）。誤って古いバージョンや無関係な
  ファイルをアップロードした場合、`read_master_rows()` がヘッダー不一致で `None` を返し
  今回分のみで作成されるため、履歴を静かに失うことはないが、**間違った履歴を正として
  上書きしてしまう可能性はチェックできない**。

## 依存パッケージ

- `streamlit>=1.40.0`
- `openpyxl>=3.1.0`

## システム要件

- Python 3.10+
- Streamlit Cloud 実行可（ローカルファイルシステム不要、ZIPアップロードで対応）

## トラブルシューティング

[README.md](README.md) の「よくある問題」を参照。

---
最終更新: 2026-07-28
