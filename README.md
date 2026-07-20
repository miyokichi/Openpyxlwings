# openpyxlwings

`openpyxlwings` は、Excel ファイルを高速に読み取り、既存ブックの画像・図形・グラフなどを壊しにくい方法で書き込むための Python ライブラリです。

- 読み取りには `openpyxl` を使うため、Microsoft Excel を起動しません。
- 書き込みには `xlwings` を使い、Microsoft Excel 本体を通して保存します。
- 読み取りと書き込みは、どちらも `ExcelWorkbook` から操作できます。
- 帳票テンプレートへの値の流し込み、罫線表の編集、定型フォーマットからのデータ抽出に対応しています。

## 目次

- [必要環境](#必要環境)
- [インストール](#インストール)
- [クイックスタート](#クイックスタート)
- [設計と安全性](#設計と安全性)
- [基本操作](#基本操作)
  - [読み取り](#読み取り)
  - [便利関数](#便利関数)
  - [書き込み](#書き込み)
- [書き込み指示をまとめる（WritePlan）](#書き込み指示をまとめるwriteplan)
- [使用例：帳票テンプレートへの書き込み](#使用例帳票テンプレートへの書き込み)
- [高度な表操作](#高度な表操作)
  - [罫線表を検出・編集する](#罫線表を検出編集する)
  - [Excelフォーマットから表を抽出する](#excelフォーマットから表を抽出する)
- [リファレンス](#リファレンス)
  - [API 一覧](#api-一覧)
  - [CLI](#cli)
  - [旧APIからの移行](#旧apiからの移行)
- [開発](#開発)
- [ライセンス](#ライセンス)

## 必要環境

- Python 3.12 以上
- 読み取りのみ: Microsoft Excel は不要
- 書き込み: `xlwings` と Microsoft Excel が必要

主に、Microsoft Excel がインストールされた Windows 環境を想定しています。

## インストール

このリポジトリをローカルへインストールします。

```bash
pip install .
```

開発環境を作る場合は、dev 依存関係も同期します。

```bash
uv sync --dev
```

wheel を直接インストールする場合:

```bash
pip install dist/openpyxlwings-0.1.0-py3-none-any.whl
```

## クイックスタート

読み取りと書き込みは、同じ `ExcelWorkbook` から行えます。

```python
from openpyxlwings import ExcelWorkbook

with ExcelWorkbook("report.xlsx", visible=False) as book:
    rows = book.read_range("Data", "A1:D20")
    book.write_values("Summary", "B2", "更新済み")

print(rows)
```

読み取りだけなら Excel は起動しません。最初の書き込み時にライブラリ専用の Excel インスタンスが起動し、`with` ブロックを正常に抜けると保存して閉じます。途中で例外が発生した場合は保存しません。

### サンプルファイル

動作確認には [samples/openpyxlwings_sample.xlsx](samples/openpyxlwings_sample.xlsx) を利用できます。

| シート名 | 内容 |
| --- | --- |
| `QuickReadWrite` | 基本的な読み取り・書き込み |
| `BorderedTable` | 罫線表の検出・編集 |
| `BrokenBorder` | 内部罫線が欠けた表の読み取り |

サンプルファイルを再生成するには、次を実行します。

```bash
uv run python scripts/create_sample_workbook.py
```

## 設計と安全性

`ExcelWorkbook` は、処理に応じて内部の実装を切り替えます。

| 処理 | 使用ライブラリ | 動作 |
| --- | --- | --- |
| 読み取り | `openpyxl` | Excel を起動せず、値を高速に読む |
| 書き込み | `xlwings` | Excel 本体で値を書き、既存オブジェクトを保ったまま保存する |

書き込み時は、ユーザーが既に開いている Excel ではなく、原則として新しい専用インスタンスを作成します。対象ファイルが別の Excel で開かれて読み取り専用になった場合は、裏側から無理に保存せずエラーにします。

> [!NOTE]
> この設計は、画像、図形、グラフ、印刷設定、マクロなどを含むテンプレートへ値を書き込む用途を想定しています。

## 基本操作

セル位置は `"B2"` のようなアドレス、または1始まりの行番号・列番号で指定できます。

```python
from openpyxlwings import ExcelWorkbook

with ExcelWorkbook("report.xlsx") as book:
    value = book.read_cell_at("Sheet1", row=2, column=2)  # B2
    book.write_values_at("Sheet1", row=3, column=2, values="完了")  # B3
```

### 読み取り

#### 範囲を読み取る

```python
from openpyxlwings import ExcelWorkbook

with ExcelWorkbook("report.xlsx") as book:
    values = book.read_range("Sheet1", "A1:D5")

print(values)
```

戻り値は二次元リストです。

```python
[
    ["Name", "Score", "Department", "Date"],
    ["Alice", 95, "Sales", "2026-06-01"],
    ["Bob", 88, "Marketing", "2026-06-02"],
]
```

#### 1つのセルを読み取る

```python
from openpyxlwings import ExcelWorkbook

with ExcelWorkbook("report.xlsx") as book:
    title = book.read_cell("Summary", "A1")

print(title)
```

#### 行番号・列番号で1つのセルを読み取る

```python
from openpyxlwings import ExcelWorkbook

with ExcelWorkbook("report.xlsx") as book:
    score = book.read_cell_at("Data", row=2, column=2)

print(score)
```

#### 行番号・列番号で範囲を読み取る

```python
from openpyxlwings import ExcelWorkbook

with ExcelWorkbook("report.xlsx") as book:
    rows = book.read_range_at(
        "Data",
        start_row=1,
        start_column=1,
        end_row=10,
        end_column=4,
    )

print(rows)
```

#### シート全体を読み取る

```python
from openpyxlwings import ExcelWorkbook

with ExcelWorkbook("report.xlsx") as book:
    rows = book.read_sheet("Data")

for row in rows:
    print(row)
```

`read_sheet()` は、末尾の空行や空列を取り除いた二次元リストを返します。

#### シート名を取得する

```python
from openpyxlwings import ExcelWorkbook

with ExcelWorkbook("report.xlsx") as book:
    print(book.sheet_names())
```

### 便利関数

短い処理なら、クラスを明示せずに便利関数も使えます。

```python
from openpyxlwings import read_range, read_sheet, sheet_names

values = read_range("report.xlsx", "Sheet1", "A1:D5")
rows = read_sheet("report.xlsx", "Data")
names = sheet_names("report.xlsx")
```

行番号・列番号で指定する便利関数もあります。

```python
from openpyxlwings import read_cell_at, read_range_at

value = read_cell_at("report.xlsx", "Sheet1", row=2, column=2)
rows = read_range_at(
    "report.xlsx",
    "Sheet1",
    start_row=1,
    start_column=1,
    end_row=10,
    end_column=4,
)
```

### 書き込み

#### 1つのセルに書き込む

```python
from openpyxlwings import ExcelWorkbook

with ExcelWorkbook("report.xlsx", visible=False) as book:
    book.write_values("Summary", "B2", "完了")
```

#### 行番号・列番号で書き込む

```python
from openpyxlwings import ExcelWorkbook

with ExcelWorkbook("report.xlsx", visible=False) as book:
    book.write_values_at("Summary", row=2, column=2, values="完了")
```

#### 表形式のデータを書き込む

```python
from openpyxlwings import ExcelWorkbook

data = [
    ["Name", "Score"],
    ["Alice", 95],
    ["Bob", 88],
    ["Charlie", 91],
]

with ExcelWorkbook("report.xlsx") as book:
    book.write_values("Data", "A1", data)
```

#### 複数箇所にまとめて書き込む

同じファイルに複数回書き込む場合も、`ExcelWorkbook` を1つ使えば大丈夫です。

```python
from openpyxlwings import ExcelWorkbook

with ExcelWorkbook("report.xlsx", visible=False) as book:
    book.write_values("Summary", "B2", "更新済み")
    book.write_values("Summary", "B3", "2026-06-11")
    book.write_values("Data", "A1", [
        ["Product", "Amount"],
        ["A", 1200],
        ["B", 3400],
    ])
```

`with` ブロックを正常に抜けると、自動で保存して閉じます。
途中で例外が発生した場合は、保存せずに閉じます。

#### 値だけを消してから書き込む

`clear_contents()` は、指定範囲の値や数式だけを消します。
セルの書式、画像、図形、グラフは削除しません。

```python
from openpyxlwings import ExcelWorkbook

with ExcelWorkbook("template.xlsx") as book:
    book.clear_contents("Data", "A2:F1000")
    book.write_values("Data", "A2", [
        ["Alice", 95, "Sales"],
        ["Bob", 88, "Marketing"],
    ])
```

行番号・列番号で範囲を消す場合は `clear_contents_at()` を使います。

```python
from openpyxlwings import ExcelWorkbook

with ExcelWorkbook("template.xlsx") as book:
    book.clear_contents_at(
        "Data",
        start_row=2,
        start_column=1,
        end_row=1000,
        end_column=6,
    )
```

#### 便利関数で書き込む

一度だけ書き込むなら `write_values()` も使えます。

```python
from openpyxlwings import write_values

write_values("report.xlsx", "Summary", "B2", "完了")
```

行番号・列番号で書き込む便利関数:

```python
from openpyxlwings import write_values_at

write_values_at("report.xlsx", "Summary", row=2, column=2, values="完了")
```

## 書き込み指示をまとめる（WritePlan）

`WritePlan` を使うと、書き込み内容を `with` ブロックの外で先に組み立てておき、`with` ブロック内の好きなタイミングで `book.apply()` を呼んでまとめて実行できます。

```python
from openpyxlwings import ExcelWorkbook, WritePlan

# with の外で書き込み指示を組み立ててキャッシュしておく
plan = WritePlan()
plan.write_values("Summary", "B2", "更新済み")
plan.write_values("Summary", "B3", "2026-06-11")
plan.clear_contents("Data", "A2:F1000")
plan.write_values("Data", "A2", [
    ["Alice", 95, "Sales"],
    ["Bob", 88, "Marketing"],
])

# with の中の特定のタイミングで、貯めた指示をまとめて実行する
with ExcelWorkbook("report.xlsx", visible=False) as book:
    current = book.read_cell("Summary", "A1")   # 読み取りはこれまで通り使える
    book.apply(plan)
```

`WritePlan` は Excel も `xlwings` も使わない、書き込み指示を貯めるだけのオブジェクトです。
`book.apply()` を呼ぶまで Excel は起動しません。
そのため、`apply()` を呼ぶより前であれば、`with` ブロック内でも `openpyxl` による読み取りを続けて使えます。

メソッドは `ExcelWorkbook` の書き込みメソッドと同じ名前・引数です。
`self` を返すので、続けて書く（メソッドチェーン）こともできます。

```python
plan = (
    WritePlan()
    .write_values_at("Data", row=2, column=1, values=[[1, 2], [3, 4]])
    .clear_contents_at("Data", start_row=10, start_column=1, end_row=100, end_column=6)
)
```

`WritePlan` で使えるメソッド:

| メソッド | 内容 |
| --- | --- |
| `plan.write_values(sheet, cell, values, expand=False)` | セルまたは範囲への書き込みを予約する |
| `plan.write_values_at(sheet, row, column, values, expand=False)` | 行番号・列番号での書き込みを予約する |
| `plan.clear_contents(sheet, address)` | 指定範囲のクリアを予約する |
| `plan.clear_contents_at(sheet, start_row, start_column, end_row, end_column)` | 行番号・列番号での範囲クリアを予約する |
| `plan.add_bordered_table(table)` | 罫線テーブル（全体・部分どちらも）の編集内容（スナップショット）を予約する |
| `plan.clear()` | 予約した指示をすべて取り消す |
| `len(plan)` | 予約済みの指示の数を返す |

予約した指示は `for op in plan:` で確認できます。
`apply()` はプランを消費しないため、同じ `WritePlan` を複数のブックへ適用することもできます。

```python
plan = WritePlan().write_values("Summary", "B2", "確定")

for path in ["report_a.xlsx", "report_b.xlsx"]:
    with ExcelWorkbook(path, visible=False) as book:
        book.apply(plan)
```

行番号・列番号を受け取るメソッド（`write_values_at` / `clear_contents_at`）は、予約した時点で値の検証を行います。
不正な行番号・列番号を渡すと、`apply()` を待たずにその場で `ValueError` になります。

予約に渡した値は、**予約した時点でスナップショット（コピー）**されます。
数値や文字列はもちろん、リストや二次元データを渡したあとに元の変数を書き換えても、予約済みの内容は変わりません。
そのため遅延評価を意識せず、その場の値をそのまま積んでいけます。

```python
plan = WritePlan()
row = [1, 2, 3]
plan.write_values("Sheet", "A1", row)
row[0] = 99            # 予約済みの内容は [1, 2, 3] のまま
```

### 罫線テーブルの編集を予約する

`get_bordered_table()` で取得した罫線テーブルの編集も `WritePlan` に積めます。
罫線テーブルの検出は `openpyxl` による読み取りなので Excel は起動しません。
編集（`set_value` / `add_row` など）はすべてメモリ上の操作なので、`apply()` を呼ぶまで Excel への書き込みは発生しません。

検出・編集は Excel を起動しないので、`with` を使わずに1つの `book` で書けます。
`book.apply(plan)` で初めて Excel が起動するため、最後に `book.close()` で Excel を閉じてください。

```python
from openpyxlwings import ExcelWorkbook, WritePlan

plan = WritePlan()
book = ExcelWorkbook("report.xlsx")

# 検出と編集（ここまで Excel は起動しない・すべてメモリ上）
table = book.get_bordered_table("Report", row=5, column=3, header_rows=1, header_columns=1)
table.set_value(row=2, column=2, value=99)
table.add_row([100, 200, 300], row_headers=["新規"])
plan.add_bordered_table(table)        # ← 編集内容をスナップショットして予約
# ... 途中で plan.write_values(...) などを直感的に追加 ...

book.apply(plan)   # ← ここで初めて Excel を1回起動し、まとめて書き込み・保存
book.close()       # ← 起動した Excel を閉じる（apply 後は必須）
```

`book.close()` は「起動した Excel を終了してファイルを解放する」ためのものです。
保存自体は `apply()`（`save=True` 既定）で済んでいるので、`close()` の主目的は Excel プロセスの後始末です。
`apply()` を呼ぶ前なら Excel は起動していないため、検出・編集だけで終える場合は `close()` を省略しても実害はありません。

`with` を使わない場合は自動クリーンアップが効かないので、`apply()` 前後で例外が起きても確実に Excel を閉じたいときは try/finally が安全です。

```python
from openpyxlwings import ExcelWorkbook, WritePlan

plan = WritePlan()
book = ExcelWorkbook("report.xlsx")
try:
    table = book.get_bordered_table("Report", row=5, column=3, header_rows=1, header_columns=1)
    table.set_value(row=2, column=2, value=99)
    table.add_row([100, 200, 300], row_headers=["新規"])
    plan.add_bordered_table(table)
    book.apply(plan)
finally:
    book.close()
```

`plan.add_bordered_table(table)` は、予約した時点のテーブルの値・範囲・行/列挿入を**スナップショット（コピー）**します。
そのため、予約した後に同じ `table` をさらに編集しても、予約済みの内容は変わりません。

`table.save()` が「Excel へ反映＋保存」までを即座に行うのに対し、`plan.add_bordered_table()` は反映も保存も `book.apply()` のタイミングまで遅延します。

`book.apply(plan, save=False)` のように `save=False` を渡すと、保存せずに書き込みだけ行います。
その場合でも、`with` ブロックを正常に抜けるときに自動で保存されます。

## 使用例：帳票テンプレートへの書き込み

`template.xlsx` にロゴ画像、図形、印刷設定、数式、グラフが入っている想定です。
`input.xlsx` から値を読み取り、テンプレートの指定範囲だけを書き換えます。

```python
from openpyxlwings import ExcelWorkbook

source_path = "input.xlsx"
template_path = "template.xlsx"

with ExcelWorkbook(source_path) as source:
    rows = source.read_range("Data", "A2:D100")

with ExcelWorkbook(template_path, visible=False) as template:
    template.clear_contents("Report", "A5:D200")
    template.write_values("Report", "A5", rows)
    template.write_values("Report", "B2", "月次レポート")
```

この処理では、読み取り側は Excel を起動しません。
書き込み側だけ Excel を起動しますが、ユーザーが既に開いている Excel ではなく、ライブラリ専用のインスタンスで処理します。

## 高度な表操作

### 罫線表を検出・編集する

Excel の「テーブル機能」ではなく、普通のセル範囲に作られた表も扱えます。
**表の左上セル** を指定すると、そこから右・下方向に表の範囲を探索します。

```python
from openpyxlwings import ExcelWorkbook

with ExcelWorkbook("report.xlsx", visible=False) as book:
    table = book.get_bordered_table(
        "Report",
        row=5,      # 表の左上セル
        column=3,
        header_rows=1,
        header_columns=1,
    )

    print(table.range)
    print(table.column_headers)
    print(table.row_headers)
    print(table.data)
```

`header_rows` は列見出しの行数、`header_columns` は行見出しの列数です。
どちらも複数指定できます。

`table.data` は見出しを除いた本体を **列ごとのリスト**（列方向優先）で返します。

`table.row_headers` は本文行ごと、`table.column_headers` は本文列ごとの見出しを返します。
見出しが1段（`header_columns` または `header_rows` が1）なら `["East", "West"]` のように
**値のリスト**、複数段なら `[("2026", "Sales"), ("2027", "Sales")]` のように
**タプルのリスト**になります。

`get_bordered_table()` は表の探し方が2通りあり、引数の組み合わせで切り替えます。

| 引数 | 内容 |
| --- | --- |
| `row` / `column` | 表の左上セルの位置で探す（両方必須）。`header_values` とは併用不可 |
| `header_values` | 見出し行の値で探す。`row`/`column` とは併用不可 |
| `value_header_contains` | 値領域の列見出しに含まれる文字列（部分一致）。`columns="all"` では必須、`"selected"` では任意 |
| `columns` | `"all"`（既定・表全体）か `"selected"`（指定列のみの部分テーブル） |
| `header_rows` | 見出しの行数。見出し指定のときは「表の何行目が見出しか」の意味にもなる |
| `header_columns` | 行見出しの列数（左上セル指定のときのみ。見出し指定では自動決定） |
| `match_case` | 見出し比較で大文字小文字を区別するか（既定は区別しない） |

見出しで探す2つのモード（`columns="all"` / `"selected"`）の詳しい動きは後述します。

#### 罫線がなくても読める（検出のしくみ）

表の範囲探索は「セルに **値があるか、何かしらの罫線があるか**」だけを見ます。
左上セル指定でも見出し指定でも同じ探索を使うので、次のような表がすべて同じ書き方で読めます。

- 罫線がまったくない、値だけの表
- 値が入っていない、罫線だけの表（これから値を流し込むテンプレートの枠など）
- 罫線が一部欠けている表・内側の格子線がない表
- 結合セルを含む表（結合範囲は左上セルにだけ値が入ります）

```python
from openpyxlwings import ExcelWorkbook

with ExcelWorkbook("report.xlsx") as book:
    # 罫線が欠けていても、値だけでも、左上セルさえ指すだけで読める
    table = book.get_bordered_table("Report", row=5, column=3, header_columns=1)

    print(table.range)
    print(table.data)
```

範囲の探索は次のルールで止まります（終了条件）。

- 右方向: 現在の行範囲に「値も罫線もないセルしかない列」が現れたらそこで終了
- 下方向: 現在の列範囲に「値も罫線もないセルしかない行」が現れたらそこで終了
- ただし、**次の行が上罫線だけ**（＝表の下端の閉じ線が下のセル側に引かれているケース）や
  **次の列が左罫線だけ**（＝右端の閉じ線）の場合は表の外とみなします
- 右・下への拡張は互いに影響するため、どちらにも広がらなくなるまで繰り返します
  （L字型にデータが伸びていても外接矩形まで広がります）

見出し指定の場合は、一致した見出しセルの `header_rows - 1` 行上を表の先頭行として、
同じルールで左・右・下に範囲を広げます。

つまり「値も罫線もない空の行・列で表が囲まれている」ことが唯一の前提です。
表の中に完全に空の行（値も罫線もない行）があると、そこで表が終わったと判断されるので注意してください。

#### 表の中身を変更する

```python
from openpyxlwings import ExcelWorkbook

with ExcelWorkbook("report.xlsx", visible=False) as book:
    table = book.get_bordered_table("Report", row=5, column=3, header_rows=1, header_columns=1)

    table.set_value(row=1, column=2, value="売上")
    table.set_body_value(row=2, column=3, value=1200)

    table.save()
```

`set_value()` は見出しを含む表全体の位置で指定します。
`set_body_value()` は見出しを除いた本文部分の位置で指定します。
どちらも 1 始まりです。

`table.save()` は元のファイルへ書き戻しますが、`table.save("out.xlsx")` のように
**パスを渡すと元ファイルを変更せず別ファイルへ保存**します（Excel の「名前を付けて保存」）。
元ファイルを残したままテンプレートを別名で書き出したいときに便利です。
なお別名保存すると、それ以降このセッションの保存先は新しいファイルに切り替わるため、
続けて `table.save()`（パス省略）を呼ぶと同じ新ファイルへ保存されます。

行見出しで本文行を探し、本文の値を読み取ったり、行全体を差し替えたりできます。

```python
from openpyxlwings import ExcelWorkbook

with ExcelWorkbook("report.xlsx", visible=False) as book:
    table = book.get_bordered_table("Report", row=5, column=3, header_rows=1, header_columns=1)

    row = table.find_body_row("東日本")
    print(row)

    values = table.get_body_row_by_header("東日本")
    print(values)  # 行見出し列を除いた本文値

    table.set_body_row_by_header("東日本", [1200, 980, 760])

    table.save()
```

行見出し列が複数ある場合は、`("東日本", "法人")` のようにすべての行見出し値を指定します。
同じ行見出しに複数行が一致する場合は、誤更新を避けるためエラーになります。
`get_body_row_by_header()` は新しいリストを返すため、戻り値を書き換えてもテーブル自体は変わりません。

`WritePlan` を使う場合は、行を書き換えた後のテーブルを通常どおり予約します。
予約時点でテーブル全体がコピーされるため、その後で同じテーブルを変更しても予約済みの値は変わりません。

```python
from openpyxlwings import ExcelWorkbook, WritePlan

plan = WritePlan()
with ExcelWorkbook("report.xlsx") as book:
    table = book.get_bordered_table(
        "Report", row=5, column=3, header_rows=1, header_columns=1
    )
    table.set_body_row_by_header("東日本", [1200, 980, 760])
    plan.add_bordered_table(table)
    book.apply(plan)
```

#### 行や列を追加する

```python
from openpyxlwings import ExcelWorkbook

with ExcelWorkbook("report.xlsx", visible=False) as book:
    table = book.get_bordered_table("Report", row=5, column=3, header_rows=1, header_columns=1)

    table.add_row([100, 200, 300], row_headers=["新規"])
    table.add_column([400, 500, 600], column_headers=["計画"])

    table.save()
```

`add_row()` は本文行を追加します。
`add_column()` は本文列を追加します。
保存時には Excel 上で行や列を挿入するため、表の下や右にある既存セルを上書きせず、押し出して表を広げます。

#### 見出しを追加する

```python
from openpyxlwings import ExcelWorkbook

with ExcelWorkbook("report.xlsx", visible=False) as book:
    table = book.get_bordered_table("Report", row=5, column=3, header_rows=1, header_columns=1)

    table.add_header_row(["区分", "4月", "5月", "6月"])
    table.add_header_column(["分類", "東日本", "西日本"])

    table.save()
```

追加された行や列には罫線を引き直します。
v1では罫線のみを整え、塗りつぶしやフォント、表示形式の完全コピーは行いません。

#### 見出し行の値から罫線表を探す

セル位置がわからなくても、`header_values` を指定すると見出しの値から表を探せます。
次のように、左側に行見出し列があり、右側に `amount` のような値列が複数続く表を想定しています。

```text
header1      header2       amount   amount
header_col1  header2_col1
header_col2  header2_col2
header_col3  header2_col3
```

`amount` 列が何列あるかわからない場合でも、`value_header_contains="amount"` を指定すると、最初に `amount` を含む列から右側を値領域として扱います。
その左側は行見出し列になります。

```python
from openpyxlwings import ExcelWorkbook

with ExcelWorkbook("report.xlsx") as book:
    table = book.get_bordered_table(
        "Sheet1",
        header_values=["header1", "header2"],
        value_header_contains="amount",
    )

    print(table.row_headers)
    print(table.column_headers)
    print(table.data)
```

この例では、`header1` と `header2` が行見出し列、`amount` を含む列が値領域の列見出しになります。
文字比較はデフォルトで大文字小文字を区別しません（`match_case=True` で区別）。

このモード（`columns="all"`、既定）では次の条件があります。

- `header_values` は表の見出し行に **左端から連続して・この順番で** 並んでいる必要があります
- `header_values` は行見出し列を **すべて** カバーしている必要があります（足りないとエラー）
- `value_header_contains` は必須で、値領域の始まりを決めます
- 見出しが表の1行目にない場合は `header_rows=2` のように見出し行の位置を指定します

#### 列見出しの一部だけ指定して取得する（部分テーブル）

`columns="selected"` を指定すると、**必要な列だけ** を指定してその列のデータだけを
取得できます。戻り値は同じ `BorderTable` ですが `partial=True` の **部分テーブル** になり、
Excelへ書き戻すときは取得した列・追加した行/列にだけ書き込み、
取得していない既存列は変更しません。

```text
header1      header2       amount   amount
header_col1  header2_col1  100      200
header_col2  header2_col2  300      400
header_col3  header2_col3  500      600
```

```python
from openpyxlwings import ExcelWorkbook

with ExcelWorkbook("report.xlsx") as book:
    table = book.get_bordered_table(
        "Sheet1",
        header_values=["header1"],          # 完全一致（リストで複数指定可・順序保持）
        value_header_contains="amount",      # 部分一致（含む列をすべて取得・任意）
        columns="selected",
    )

    print(table.column_headers)   # ['amount', 'amount']（値列の見出し。header2 は含まれない）
    print(table.row_headers)      # ['header_col1', 'header_col2', 'header_col3']
    print(table.data)             # [[100, 300, 500], [200, 400, 600]]（値列を列ごとのリストで）
    print(table.source_columns)   # 各列の元のExcel列番号。メモリ上で追加した列は None

    # 部分テーブルとして編集できる
    table.add_row([700, 800], row_headers=["header_col4"])
    table.add_column([1, 2, 3, 4], column_headers=["ratio"])

    table.save()  # 取得列＋追加分のみ書き戻し
```

`columns="all"` と違い、`header_values` は表の見出し行の **どこにあってもよく**、
連続している必要もありません。指定した列が行見出し列（`row_headers`）になり、
`value_header_contains` に一致した列が本文（`data`）になります。

各列のデータは表の先頭行から読み取り、見出しより下は
**値がある / 上罫線がある / 下罫線がある** いずれかなら継続、いずれもなければ終了します。
列ごとに長さが異なる場合は `None` で埋めて矩形化します。

書き戻し時の動きは全体テーブルと少し異なります。

- 途中に挿入した行は Excel 上でもその位置に行挿入されます
- 追加した列は、仮想的な挿入位置にかかわらず表の **右端** に追加されます
- 表の見出しより上の行（タイトル行など）は書き換えません

`WritePlan` への予約は全体・部分どちらも `plan.add_bordered_table(table)` です。

#### 部分テーブルでも行見出しで本文行を探す

`find_body_row()` / `get_body_row_by_header()` / `set_body_row_by_header()` は部分テーブルでも同じように使えます。
行見出しとして扱われるのは `header_values` で指定した列（完全一致で選んだ列）で、
`value_header_contains` で選んだ列は本文（値）列として扱われます。

```python
from openpyxlwings import ExcelWorkbook

with ExcelWorkbook("report.xlsx", visible=False) as book:
    table = book.get_bordered_table(
        "Sheet1",
        header_values=["header1"],
        value_header_contains="amount",
        columns="selected",
    )

    row = table.find_body_row("header_col1")
    print(row)

    values = table.get_body_row_by_header("header_col1")
    print(values)  # 選択された本文列の値だけを返す

    table.set_body_row_by_header("header_col1", [1200, 980])

    table.save()
```

#### 行の値で列を絞り込んで部分テーブルにする

取得済みのテーブルから、**特定の行の値を条件に列を絞り込んだ部分テーブル**を作れます。
判定に使う行は `find_body_row()` と同じく行見出しで指定します。

```text
metric   prodA   prodB   prodC   prodD
sales    120     80      200     50
rate     0.9     0.6     1.1     0.4    ← この行の値で列を選ぶ
```

```python
from openpyxlwings import ExcelWorkbook

with ExcelWorkbook("report.xlsx", visible=False) as book:
    table = book.get_bordered_table("Metrics", row=2, column=2, header_columns=1)

    # rate 行が 0.8 以上の列だけを持つ部分テーブル
    subset = table.select_columns_by_row("rate", lambda v: v is not None and v >= 0.8)

    print(subset.column_headers)   # ['prodA', 'prodC']
    print(subset.source_columns)   # 各列の元のExcel列番号

    subset.set_body_row_by_header("sales", [150, 210])
    subset.save()                  # 絞り込んだ列だけ書き戻し
```

- `condition` には **呼び出し可能オブジェクト**（セルの生の値を受け取って真偽を返す）か、
  **プレーンな値**（見出し比較と同じ正規化＝前後空白除去・大文字小文字無視で完全一致。
  `match_case=True` で大小区別）を渡せます
- 行見出し列（先頭 `header_columns` 列）は常に保持されるので、絞り込み後も
  `find_body_row()` / `get_body_row_by_header()` / `set_body_row_by_header()` がそのまま使えます
- 1列も条件に合わない場合はエラーになります
- 元になるテーブルは全体・部分どちらでもよく、部分テーブルから更に絞り込むこともできます
- 値はコピーされるため、絞り込み後のテーブルを編集しても元のテーブルは変わりません。
  書き戻し位置がずれないよう、**検出直後（または保存直後）のテーブルから派生させてください**
  （元テーブルに未保存の行・列追加がある状態で派生させない）

#### 見出しの値で行・列を絞り込む

`select_columns_by_row()` が「本文行の値」で列を選ぶのに対し、**見出しの値そのもの**で
絞り込む2つのメソッドもあります。どちらも lambda（または完全一致するプレーン値）で判定します。

- `select_columns_by_column_header(condition)` … **列ヘッダー値**で **列** を絞る
- `select_rows_by_row_header(condition)` … **行ヘッダー値**で **行** を絞る

```text
metric   prodA   prodB   prodC   prodD
sales    120     80      200     50
rate     0.9     0.6     1.1     0.4
flag     OK      NG      ok
```

```python
from openpyxlwings import ExcelWorkbook

with ExcelWorkbook("report.xlsx", visible=False) as book:
    table = book.get_bordered_table("Metrics", row=2, column=2, header_columns=1)

    # 列ヘッダーが "prodA"/"prodC" の列だけ（列を絞る）
    cols = table.select_columns_by_column_header(lambda h: h in ("prodA", "prodC"))
    print(cols.column_headers)     # ['prodA', 'prodC']

    # 行ヘッダーが "sales"/"flag" の行だけ（行を絞る）
    rows = table.select_rows_by_row_header(lambda h: h in ("sales", "flag"))
    print(rows.row_headers)        # ['sales', 'flag']

    rows.add_row([1, 2, 3, 4], row_headers=["cost"])   # 末尾に行追加
    rows.add_column([9, 8, 9], column_headers=["extra"])  # 右端に列追加
    rows.save()   # 絞り込んだ行だけ書き戻し（除外行はそのまま温存）
```

- `condition` はセル値（複数ヘッダー行/列のときはタプル）を受け取る呼び出し可能オブジェクトか、
  正規化して完全一致するプレーン値（`match_case=True` で大小区別）
- 列版は行見出し列を、行版は列見出し行を常に保持します。1つも一致しなければエラー
- **行版（行サブセット）** は、保持した行を元のExcel行へ書き戻し、除外した行はそのまま残します。
  追加した行は表の末尾、追加した列は右端に書き込まれます。`add_row`/`add_column`/`save` に対応
- 列サブセットと行サブセットは相互に連鎖できません（片方に絞った後、もう一方では絞れない）

#### 表の検出条件と、見つからないときのチェックリスト

どちらの指定方法でも、表の範囲は同じ「値または罫線があるセルをたどる」領域探索で
決まります（終了条件は前述）。罫線の欠け・罫線なし・値なし・結合セルをすべて許容し、
前提は「値も罫線もない空の行・列で表が囲まれていること」だけです。

見出し指定（`header_values`）では、範囲の決め方と一致条件が次のようになります。

- 一致セルの `header_rows - 1` 行上を表の先頭行とみなし、そこから左右・下に範囲を広げます
- 見出しの比較は「前後の空白を除去した文字列の完全一致」です。既定では大文字小文字を
  区別しませんが、**見出し内部の空白・全角半角・セル内改行の違いは不一致** になります
- 一致セルは候補として順に試され、成立しない候補（見出し行の下に本文行がない、
  値領域が見つからない等）は読み飛ばされます

見出し指定で見つからないときは、(1) 見出しの文字列が完全一致しているか、
(2) `header_rows` が実際の見出し位置と合っているか、(3) 見出し行の下に本文行が
あるか、の順に確認するのが早道です。
表の位置がわかっているなら、左上セル指定に切り替えるのが確実です。

なお、表のすぐ隣（間に空の行・列がない位置）にメモ書きや別の表があると、
領域探索がそれらを取り込んで範囲を広く判定します。表の周囲は1行・1列以上
空けてください。

`header_values` を複数指定した場合は、`("header_col1", "header2_col1")` のように
すべての行見出し値を指定します。同じ行見出しに複数行が一致する場合は、誤更新を
避けるためエラーになります。

### Excelフォーマットから表を抽出する

別のExcelファイルを「表のフォーマット定義」として使い、同じ構造を持つ表を対象ブックから自動検索できます。
フォーマット定義は1シートにつき1パターンで、シート名がパターン名になります。

サンプル:

```text
samples/extraction_format.xlsx
samples/extraction_input.xlsx
```

#### フォーマットExcelを作る

`extraction_format.xlsx` の `amount_table` シートには、次のような2行のひな形が入っています。

```text
header1             header2             {{columns[].header | contains("amount")}}
{{rows[].header1}}  {{rows[].header2}}  {{rows[].amounts[]:float}}
```

固定文字とプレースホルダーを、実際の表と同じセル配置で記述します。

| 書式 | 意味 |
| --- | --- |
| `{{title}}` | 単一セルを抽出する |
| `{{rows[].name}}` | 下方向へ繰り返す値を抽出する |
| `{{columns[].header}}` | 右方向へ繰り返す値を抽出する |
| `{{rows[].amounts[]}}` | 可変行・可変列の交点を二次元的に抽出する |
| `:int`, `:float`, `:date` など | 値を指定型へ変換・検証する |
| `contains("amount")` | セルに指定文字列が含まれることを要求する |
| `equals("value")` | セルが指定文字列と完全一致することを要求する |

固定文字は対象Excelと完全一致する必要があります。
空のテンプレートセルは照合対象外です。

#### `{{title}}` で単一セルを取得する

繰り返しではない単一セルは、単純な名前のプレースホルダーで取得できます。
例えば、フォーマットExcelの `report_info` シートを次のように作ります。

```text
report_title  {{title}}
report_date   {{report_date:date}}
```

対象Excelに次の表があるとします。

```text
report_title  月次売上レポート
report_date   2026-06-24
```

```python
from openpyxlwings import ExcelFormat, ExcelWorkbook

formats = ExcelFormat.load("samples/extraction_format.xlsx")
pattern = formats["report_info"]

with ExcelWorkbook("samples/extraction_input.xlsx") as book:
    matches = book.extract(pattern)

report = matches[0]
print(report.data)
```

結果:

```python
{
    "title": "月次売上レポート",
    "report_date": date(2026, 6, 24),
}
```

数式セルの場合は、計算済みの値が `data["title"]`、数式文字列が `formulas["title"]` に入ります。

#### 対象Excelから抽出する

```python
from openpyxlwings import ExcelFormat, ExcelWorkbook

formats = ExcelFormat.load("samples/extraction_format.xlsx")
pattern = formats["amount_table"]

with ExcelWorkbook("samples/extraction_input.xlsx") as book:
    matches = book.extract(pattern)

for match in matches:
    print(match.sheet)
    print(match.range)
    print(match.data)
```

一致する表が複数ある場合は、シート位置・セル位置の順にすべて返します。

抽出結果の例:

```python
{
    "columns": [
        {"header": "amount"},
        {"header": "amount forecast"},
        {"header": "amount final"},
    ],
    "rows": [
        {
            "header1": "header_col1",
            "header2": "detail1",
            "amounts": [100.0, 120.0, 140.0],
        },
    ],
}
```

`ExtractedMatch` は以下の情報を持ちます。

| 属性 | 内容 |
| --- | --- |
| `sheet` | 一致したシート名 |
| `range` | 一致したセル範囲 |
| `data` | 計算済み値を意味付きの辞書・リストに変換した結果 |
| `formulas` | `data` と同じ構造で保持する数式文字列。通常セルは `None` |
| `source_cells` | `data` と同じ構造で保持する元セルの位置 |

Excelに数式の計算結果が保存されていない場合、`data` の値は `None` でも `formulas` には `=SUM(...)` などの式が入ります。

#### 検索対象を絞る

```python
with ExcelWorkbook("input.xlsx") as book:
    matches = book.extract(
        pattern,
        sheets=["Sheet1", "Sheet2"],
        ranges={"Sheet1": "A1:Z200"},
    )
```

`sheets` でシートを限定できます。
`ranges` だけを指定した場合は、辞書に含まれるシートと範囲だけを検索します。

v1では、可変行のプロトタイプをフォーマット表の最終行、可変列のプロトタイプを最終列に配置してください。
YAMLフォーマットと入れ子になった複数階層の繰り返しは、現在未対応です。

## リファレンス

### API 一覧

```python
from openpyxlwings import ExcelWorkbook, WritePlan
```

| API | 内容 |
| --- | --- |
| `ExcelWorkbook(path)` | 読み書き用の中心クラス |
| `book.sheet_names()` | シート名一覧を返す |
| `book.read_cell(sheet, cell)` | 1セルの値を返す |
| `book.read_cell_at(sheet, row, column)` | 行番号・列番号で1セルの値を返す |
| `book.read_range(sheet, address)` | 指定範囲を二次元リストで返す |
| `book.read_range_at(sheet, start_row, start_column, end_row, end_column)` | 行番号・列番号で指定範囲を返す |
| `book.read_sheet(sheet=None)` | シート全体を二次元リストで返す |
| `book.write_values(sheet, cell, values)` | セルまたは範囲に値を書き込む |
| `book.write_values_at(sheet, row, column, values)` | 行番号・列番号で値を書き込む |
| `book.clear_contents(sheet, address)` | 指定範囲の値や数式だけを消す |
| `book.clear_contents_at(sheet, start_row, start_column, end_row, end_column)` | 行番号・列番号で指定範囲の値や数式だけを消す |
| `WritePlan()` | 書き込み指示を貯めるオブジェクトを作る |
| `book.apply(plan, save=True)` | `WritePlan` に貯めた書き込みをまとめて実行する |
| `book.get_bordered_table(sheet, *, row=None, column=None, header_values=None, value_header_contains=None, columns="all", header_rows=1, header_columns=0, match_case=False)` | 表を取得する。`row`/`column`（左上セル指定）か `header_values`（見出し指定）のどちらかで表を探す。どちらも罫線不要のゆるい領域探索で、`columns="selected"` で指定列のみの部分テーブルを返す |
| `ExcelFormat.load(path)` | Excelフォーマットブックを読み込む |
| `book.extract(pattern, sheets=None, ranges=None)` | フォーマットに一致する表をすべて抽出する |
| `table.find_body_row(row_header)` | 行見出しに一致する本文行の1始まり位置を返す |
| `table.get_body_row_by_header(row_header)` | 行見出しに一致する行の本文値を新しいリストで返す |
| `table.set_body_row_by_header(row_header, values)` | 行見出しに一致する行の本文値を置き換える。`WritePlan` では変更後に `plan.add_bordered_table(table)` で予約する |
| `table.select_columns_by_row(row_header, condition, *, match_case=False)` | 指定行の値で列を絞った部分テーブルを返す |
| `table.select_columns_by_column_header(condition, *, match_case=False)` | 列ヘッダー値で列を絞った部分テーブルを返す |
| `table.select_rows_by_row_header(condition, *, match_case=False)` | 行ヘッダー値で行を絞った部分テーブルを返す（除外行は保存時も温存） |
| `table.save(path=None)` | 編集した罫線テーブルを書き戻す。`path` 指定で元ファイルを変更せず別ファイルへ保存 |
| `book.save(path=None)` | 明示的に保存する。`path` 指定で元ファイルを変更せず別ファイルへ保存 |
| `book.close(save=True)` | 開いている内部セッションを閉じる |

### 互換用の名前

以前のコード向けに、以下の名前も残しています。
ただし、新しく書くコードでは `ExcelWorkbook` を使うのがおすすめです。

```python
from openpyxlwings import ExcelReader, ExcelWriter

assert ExcelReader is ExcelWorkbook
assert ExcelWriter is ExcelWorkbook
```

### 旧APIからの移行

罫線テーブルの取得APIは `get_bordered_table()` 1本に統合しました。
旧APIと旧クラスは削除済みです。次の対応で書き換えてください。

| 旧 | 新 |
| --- | --- |
| `book.get_bordered_table(sheet, row, column, ...)` | `book.get_bordered_table(sheet, row=row, column=column, ...)`（キーワード引数に変更。**セルは表内の任意位置ではなく左上を指す仕様に変更**され、罫線条件なしで領域を探索します） |
| `book.get_bordered_table_by_header(sheet, header_values, value_header_contains=..., header_row=...)` | `book.get_bordered_table(sheet, header_values=..., value_header_contains=..., header_rows=...)` |
| `book.get_bordered_table_by_columns(sheet, header_values, ...)` | `book.get_bordered_table(sheet, header_values=..., columns="selected", ...)` |
| `SelectedColumnsTable` | `BorderTable`（`partial=True` の部分テーブル） |
| `plan.add_selected_columns_table(table)` | `plan.add_bordered_table(table)` |
| `require_inner_borders=...` | 廃止。検出が罫線を前提としなくなったため、引数ごと削除されました |

部分テーブル（旧 `SelectedColumnsTable`）はプロパティやメソッドの形も全体テーブルに揃えたため、以下が変わっています。

| 項目 | 旧 `SelectedColumnsTable` | 新 `BorderTable`（`columns="selected"`） |
| --- | --- | --- |
| `data` | 選択した全列（行見出し列を含む）の本文 | 行見出し列を **除いた** 本文（列ごとのリスト） |
| `column_headers` | 全列の見出しのフラットなリスト | 本文列のみの見出し行のリスト（`[[...]]`） |
| `columns` | `_SelectedColumn` のリスト | 見出しを含む値の列リスト。元のExcel列は `source_columns` |
| `row_count` | 本文の行数 | 見出しを含むグリッドの行数 |
| `add_row(values)` | 全列分の値を1つのリストで渡す | 本文列分の `values` と `row_headers=` に分けて渡す |
| `add_column(values, header=...)` | 見出しはスカラーの `header=` | 見出しは `column_headers=[...]`（リスト） |
| `set_value(row, column, ...)` | 本文基準の座標 | 見出しを含む表全体の座標（本文基準は `set_body_value()`） |
| 途中への行挿入 | 保存時は末尾にまとめて挿入 | 保存時もその位置に行挿入 |

### CLI

シート名を表示:

```bash
openpyxlwings sheets report.xlsx
```

指定範囲を JSON で表示:

```bash
openpyxlwings read report.xlsx Sheet1 A1:D20
```

## 開発

このパッケージは `src/` レイアウトです。

```text
src/
  openpyxlwings/
    __init__.py
    workbook.py
    reader.py
    writer.py
    exceptions.py
    cli.py
tests/
  test_reader.py
  test_api.py
```

ローカル確認:

```bash
uv run pytest
uv build
```

## ライセンス

MIT License です。詳細は `LICENSE` を参照してください。
