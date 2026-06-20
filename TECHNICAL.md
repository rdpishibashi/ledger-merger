# Ledger-merger 技術仕様

## 概要

DXF-diff-manager の出力フォルダ群（各フォルダに台帳ファイル `図面親子管理台帳.xlsx` を含む）
を再帰的に走査し、各台帳の `Diff List` シートを1つの統合 Excel に集約する Streamlit アプリ。

## ディレクトリ構成

```
Ledger-merger/
├── app.py                          # View層: Streamlit UI
├── utils/
│   ├── __init__.py
│   ├── ledger_finder.py            # Model層: フォルダ走査・台帳検出
│   └── ledger_merger.py            # Model層: 統合Excel生成
├── tests/
│   └── unit/
│       └── test_ledger_merger.py   # 実データを使った回帰テスト
├── requirements.txt
├── README.md
└── TECHNICAL.md
```

| モジュール | 責務 | Streamlit依存 |
|---|---|---|
| `app.py` | 入力受付（ZIP）、結果表示、ダウンロード | あり |
| `utils/ledger_finder.py` | フォルダ再帰走査、台帳ファイルの検出・読み込み | なし |
| `utils/ledger_merger.py` | 統合 Excel（bytes）の生成 | なし |

Model層（`utils/`）は Streamlit に依存しない純粋関数として実装しており、
`tests/unit/` から UI なしで直接呼び出してテストできる。

## アーキテクチャ（データフロー）

```
[ZIPアップロード] → 一時ディレクトリに展開 → find_ledger_files(root_dir)
                          (tempfile.TemporaryDirectory)   (utils/ledger_finder.py)
                                                                  ↓
                                          list[LedgerEntry] + folders_without_ledger
                                                                  ↓
                                                  build_merged_workbook(entries)
                                                      (utils/ledger_merger.py)
                                                                  ↓
                                                      統合済み Excel (bytes)
                                                                  ↓
                                                      st.download_button
```

### `utils/ledger_finder.py`

- `find_ledger_files(root_dir) -> (list[LedgerEntry], list[str])`
  - `os.walk(root_dir)` で再帰的に走査するが、**サブフォルダを持つ中間（ラッパー）フォルダは
    評価対象から除外し、サブフォルダを持たない葉フォルダのみ**を「DXF-diff-manager 出力フォルダ」
    として扱う（実データでは各出力フォルダがサブフォルダを持たないため）。
  - 各葉フォルダ内の `.xlsx` のうち、`~$` 始まりの一時ロックファイルと
    `NON_LEDGER_FILENAMES = {"diff_labels.xlsx", "unchanged_labels.xlsx"}`
    （DXF-diff-manager が出力する台帳以外の固定ファイル）を除外したものを台帳候補とする。
  - 候補ファイルを `openpyxl.load_workbook(path, data_only=True, read_only=True)` で開き、
    シート名に `Diff List` と `Summary` の両方が存在し、`Diff List` のヘッダーが
    `DIFF_LIST_HEADERS`（12列）と一致し、`Summary` シートに `SUMMARY_LABELS`（9項目）が
    すべて存在するものだけを有効な台帳と判定する（ファイル名は不問）。開けない/壊れている
    ファイルや条件を満たさないファイルは単に `None` を返して読み飛ばす。
  - `Diff List` の各データ行のうち、**`Total Entities` が空欄（`None`）の行は除外**する。
    これは実際には差分抽出されていない図番ペアの関係記録（親子マスター管理用、
    `Summary` シートの集計にも含まれない）であり、統合 Diff List には含めない。
    除外後に1行も残らない場合は無効な台帳として扱う（`folders_without_ledger` に積む）。
  - 台帳と判定したファイルの**直接の親フォルダ名**を `LedgerEntry.package_name` とする。
  - 葉フォルダ内に有効な台帳が1件も見つからない場合、そのフォルダの**ベース名のみ**
    （パスやファイル名の詳細は含めない）を `folders_without_ledger` に積む（重複除去・出現順）。
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
    空欄の行を除く、12列のタプル）。常に `summary_values["差分抽出ペア数"]` と同数になる。
  - `summary_values: dict` — `SUMMARY_LABELS` をキーとした `Summary` シートの値

### `utils/ledger_merger.py`

- `build_merged_workbook(entries: list[LedgerEntry]) -> bytes`
  - `openpyxl.Workbook()` を直接操作してセル単位で書き込む
    （pandas + xlsxwriter のDataFrame経由ではなく、セルごとのフォント色制御が必要なため素のopenpyxlを採用）。
  - シート名は `Diff List` の1シートのみ。
  - 出力列（22列、`OUTPUT_HEADERS`）:
    `Diff Package, Child, Parent, Relation, Title, Subtitle, Recorded Date, Note, Deleted Entities, Added Entities, Diff Entities, Unchanged Entities, Total Entities, 削除図形数 合計, 追加図形数 合計, 差分図形数 合計, 変更なし図形数 合計, 総図形数 合計, 図形変更率 [%], 入力図面総数, 差分抽出ペア数, 流用率 [%]`
  - 各 `LedgerEntry` ブロック内:
    - A列（`Diff Package`）: ブロック最初の行のみ黒字（`FF000000`）、以降は薄いグレー（`FFA6A6A6`）。
    - B〜M列: 元の `Diff List` 行をそのまま転記。`Recorded Date` 列のみ `number_format = "YYYY-MM-DD HH:MM:SS"` を設定。
    - N〜V列（Summaryの9項目）: ブロックの最初の行のみ値を記入。カウント系は `#,##0`、
      `図形変更率 [%]` と `流用率 [%]` は `0.00%` の `number_format` を適用。2行目以降は空欄。
  - ヘッダー行は太字、`freeze_panes = "A2"`。

### `app.py`

- 入力は `st.file_uploader(accept_multiple_files=True, type=["zip"])` によるZIPアップロードのみ
  （ローカルパス直接指定は廃止。入力方式が2つ並ぶとユーザーが混乱するため、ZIP一本化に統一した）。
  各ZIPは `tempfile.TemporaryDirectory()` に展開し `find_ledger_files` に渡す。
- 「統合実行」ボタンは `type="primary"` + `disabled=not has_input`
  （`has_input = bool(zip_files)`）。入力が無い間はグレーアウトし、
  色を手動で変更するコードは書かない（テーマの `primaryColor` に委ねる。
  `streamlit` スキルの「UIテーマ・ボタン・セクション構成」を参照）。
- ZIPファイル1件単位で `st.empty()` + `progress(ratio, text=...)` の
  進捗バーを表示し、完了後 `placeholder.empty()` で消す。
- 「統合実行」ボタン押下時に全ZIPから `LedgerEntry` と `folders_without_ledger`
  （台帳が見つからなかったフォルダ名）を集約し、1件以上あれば `build_merged_workbook` を
  実行して `st.session_state` に結果を保存する。
- ZIP自体が開けない場合（`zipfile.BadZipFile`）は `st.warning()` で即時表示する
  （フォルダ名一覧とは別枠。ZIPはフォルダではないため）。
- 結果（成功件数・台帳が無いフォルダ名一覧・ダウンロードボタン）は `st.session_state` を介して
  描画するため、再実行（rerun）後も表示が保持される。フォルダ名一覧は
  `st.expander("⚠️ 台帳ファイルが見つからなかったフォルダ（N件）")` の中に
  **フォルダ名のみ**（パスやファイル名は含めない）で表示する。
- `.streamlit/config.toml` でテーマ（`primaryColor = "#0365C0"` 等、DXF-diff-manager と同一）を指定。
  単一アクションのツールのため `st.subheader("Step N: ...")` のような連番見出しは使わない
  （セクション名のみの見出しに留める）。

### セッション状態

| キー | 内容 |
|---|---|
| `merged_bytes` | 統合済み Excel のバイト列 |
| `merged_filename` | ダウンロード時のファイル名 |
| `merged_count` | 統合した Diff Package 数 |
| `merged_missing_folders` | 台帳ファイルが見つからなかったフォルダ名（ベース名のみ）の一覧 |

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
入力方式が2つ並ぶとユーザーが混乱するため廃止した）。

## 「Total Entities 空欄行の除外」の正当性の根拠

`utils/ledger_finder.py` が `Total Entities` 空欄行を除外する判断は、DXF-diff-manager側の
ソースコード（`DXF-diff-manager/app.py` の `update_parent_child_master()` /
`save_master_to_bytes()`）まで遡って裏付けを取った:

- `update_parent_child_master()` は新規行追加・既存行更新のいずれでも、エンティティ数
  （Deleted/Added/Diff/Unchanged/Total Entities）は `if entity_counts:`（実際に差分抽出が
  成功したペアのみ `entity_counts` を保持）の場合だけ設定する。Title/Subtitle/Relation/
  Recorded Date はこれと無関係に設定されるため、**「Total Entities が空欄」＝「このペアは
  実際には差分抽出されていない（親子マスター管理のための関係記録のみ）」が DXF-diff-manager
  自身の設計として保証されている**。
- `save_master_to_bytes()` の Summary 集計（削除図形数 合計 等）は
  `master_df[col].sum(skipna=True)` で計算されるため、空欄行は自動的に合計から除外される。
  つまり Summary の値は、もともと「Total Entities が空欄でない行」だけを集計したものである。

このため、`utils/ledger_finder.py` の除外ロジックは DXF-diff-manager の出力データの設計と
一致している。実データ（29フォルダ）に対して以下を全件検証済み（`tests/unit/test_ledger_merger.py`）:
- 件数: 除外後の行数 ＝ Summary の「差分抽出ペア数」（全29フォルダで完全一致）
- 数値: 除外後の各エンティティ列の合計 ＝ Summary の対応する合計値（全29フォルダ・5項目で完全一致）
- 順序・値: 除外後に残る行は、元の Diff List シートの該当行を順序・値ともに完全に保持

**注意（厳密な保証ではない点）**: 「件数一致」は DXF-diff-manager の現在の実装から推測される
強い相関であり、全29フォルダで実証済みだが、`差分抽出ペア数`（`pairs` リストから算出）と
`Total Entities` 非空行数（`master_df` から算出）は別々のデータソースに基づくため、
将来 DXF-diff-manager 側で「ペアリング済みだが実際の比較処理は例外で失敗した」ケースが
発生すると、理論上はこの一致が崩れる可能性がある。`utils/ledger_finder.py` 自体の判定は
`Total Entities is None` のみに基づいており、この潜在的なズレの影響を受けない
（Summary の「差分抽出ペア数」を直接の判定基準には使っていない）。

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
- `Diff List` のヘッダー列構成、`Summary` シートのラベル名が変わった場合は検出条件
  （`DIFF_LIST_HEADERS` / `SUMMARY_LABELS`、`utils/ledger_finder.py`）の更新が必要。
- 出力フォルダの判定は「サブフォルダを持たない葉フォルダ」という前提に依存している。
  DXF-diff-manager の出力構造が将来サブフォルダを持つようになった場合は
  `find_ledger_files`（`utils/ledger_finder.py`）の判定ロジックの見直しが必要。
- 台帳が見つからない理由（ファイル形式不一致・読み込み失敗・本当に存在しない等）は
  区別せず、フォルダ名のみで一括して報告する（意図的な簡略化）。

## 依存パッケージ

- `streamlit>=1.40.0`
- `openpyxl>=3.1.0`

## システム要件

- Python 3.10+
- Streamlit Cloud 実行可（ローカルファイルシステム不要、ZIPアップロードで対応）

## トラブルシューティング

[README.md](README.md) の「よくある問題」を参照。

---
最終更新: 2026-06-20
