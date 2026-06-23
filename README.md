openpyxlwings
=============

`openpyxlwings` は、Excel ファイルを「高速に読み取り」「既存ブックを壊しにくく書き込み」するための Python ライブラリです。

中心になるクラスは `ExcelWorkbook` です。
読み取りと書き込みで別々のクラスを使う必要はありません。

- 読み取りは `openpyxl` を使います。
- 書き込みは `xlwings` を使い、Excel 本体経由で保存します。
- 書き込み時は、既にユーザーが開いている Excel 画面を操作せず、ライブラリ専用の Excel インスタンスを作って処理します。
- 画像、図形、グラフ、印刷設定、マクロなどを含むテンプレートファイルへの値流し込みを想定しています。

このライブラリが向いている場面
------------------------------

- Excel ファイルから大量の値を速く読み取りたい
- 画像や図形入りの帳票テンプレートに値を書き込みたい
- `openpyxl` で保存したときに画像や図形が消える問題を避けたい
- 読み取りは軽く、書き込みだけ Excel 本体に任せたい
- ユーザーが開いている Excel 作業画面に干渉したくない

仕組み
------

`ExcelWorkbook` の中で、読み取り用と書き込み用の処理を内部的に分けています。

| 処理 | 内部で使うもの | 目的 |
| --- | --- | --- |
| 読み取り | `openpyxl` | Excel を起動せず、高速に値を読む |
| 書き込み | `xlwings` | Excel 本体で開いて値を書き、既存オブジェクトを壊しにくく保存する |

`openpyxl` は読み取りには便利ですが、複雑な Excel ファイルを保存すると Excel 独自のオブジェクトに影響が出ることがあります。
そのため、このライブラリでは `openpyxl` を読み取り専用に使い、書き込みは `xlwings` 経由で Excel に任せます。

書き込み時の安全方針
--------------------

書き込みでは、既に開いている Excel アプリケーションを探して使うのではなく、原則として新しい Excel インスタンスを作成します。
これにより、ユーザーが手作業で開いている Excel ウィンドウやブックを誤って操作しにくくしています。

また、対象ファイルが既に Excel で開かれていて読み取り専用になった場合は、無理に保存せずエラーにします。
これは、ユーザーが開いているブックに裏側から勝手に書き込まないためです。

インストール
------------

このリポジトリをローカルでインストールする場合:

```bash
pip install .
```

uv で開発環境を作る場合:

```bash
uv sync --dev
```

テスト:

```bash
uv run pytest
```

配布用ファイルの作成:

```bash
uv build
```

wheel を直接インストールする場合:

```bash
pip install dist/openpyxlwings-0.1.0-py3-none-any.whl
```

サンプルExcel
-------------

動作確認用のサンプルファイルを用意しています。

```text
samples/openpyxlwings_sample.xlsx
```

含まれるシート:

| シート名 | 内容 |
| --- | --- |
| `QuickReadWrite` | 通常の読み取り・書き込みAPIを試すための表 |
| `BorderedTable` | 罫線テーブル検出・編集を試すための表 |
| `BrokenBorder` | 内部罫線が欠けている例外確認用の表 |

サンプルファイルを再生成する場合:

```bash
uv run python scripts/create_sample_workbook.py
```

罫線テーブルの読み取り例:

```python
from openpyxlwings import ExcelWorkbook

with ExcelWorkbook("samples/openpyxlwings_sample.xlsx") as book:
    table = book.get_bordered_table(
        "BorderedTable",
        row=5,
        column=3,
        header_rows=2,
        header_columns=1,
    )

    print(table.range)
    print(table.column_headers)
    print(table.row_headers)
    print(table.data)
```

注意点
------

読み取りだけなら Microsoft Excel は不要です。

書き込みには `xlwings` と Microsoft Excel が必要です。
このライブラリは、特に Windows 上で Excel がインストールされている環境を想定しています。

基本的な使い方
--------------

### 1つのクラスで読み書きする

```python
from openpyxlwings import ExcelWorkbook

with ExcelWorkbook("report.xlsx") as book:
    rows = book.read_range("Data", "A1:D20")
    book.write_values("Summary", "B2", "更新済み")
```

`ExcelWorkbook` は、読み取り時には `openpyxl` を使い、書き込み時には `xlwings` を使います。
利用者側で reader / writer を切り替える必要はありません。

### 行番号・列番号でアクセスする

`"B2"` のような Excel 形式のアドレスではなく、行番号と列番号でもアクセスできます。
行番号・列番号は Excel と同じく 1 始まりです。

```python
from openpyxlwings import ExcelWorkbook

with ExcelWorkbook("report.xlsx") as book:
    value = book.read_cell_at("Sheet1", row=2, column=2)
    book.write_values_at("Sheet1", row=3, column=2, values="完了")

print(value)
```

この例では、`row=2, column=2` が `B2`、`row=3, column=2` が `B3` に対応します。

読み取り
--------

### 範囲を読み取る

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

### 1つのセルを読み取る

```python
from openpyxlwings import ExcelWorkbook

with ExcelWorkbook("report.xlsx") as book:
    title = book.read_cell("Summary", "A1")

print(title)
```

### 行番号・列番号で1つのセルを読み取る

```python
from openpyxlwings import ExcelWorkbook

with ExcelWorkbook("report.xlsx") as book:
    score = book.read_cell_at("Data", row=2, column=2)

print(score)
```

### 行番号・列番号で範囲を読み取る

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

### シート全体を読み取る

```python
from openpyxlwings import ExcelWorkbook

with ExcelWorkbook("report.xlsx") as book:
    rows = book.read_sheet("Data")

for row in rows:
    print(row)
```

`read_sheet()` は、末尾の空行や空列を取り除いた二次元リストを返します。

### シート名を取得する

```python
from openpyxlwings import ExcelWorkbook

with ExcelWorkbook("report.xlsx") as book:
    print(book.sheet_names())
```

便利関数
--------

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

書き込み
--------

### 1つのセルに書き込む

```python
from openpyxlwings import ExcelWorkbook

with ExcelWorkbook("report.xlsx", visible=False) as book:
    book.write_values("Summary", "B2", "完了")
```

### 行番号・列番号で書き込む

```python
from openpyxlwings import ExcelWorkbook

with ExcelWorkbook("report.xlsx", visible=False) as book:
    book.write_values_at("Summary", row=2, column=2, values="完了")
```

### 表形式のデータを書き込む

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

### 複数箇所にまとめて書き込む

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

### 値だけを消してから書き込む

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

### 便利関数で書き込む

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

帳票テンプレートに値を流し込む例
------------------------------

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

罫線で区切られた表を編集する
----------------------------

Excel の「テーブル機能」ではなく、普通のセル範囲に罫線が引かれている表も扱えます。
起点セルを指定すると、そのセルを含む罫線テーブルを検出します。

```python
from openpyxlwings import ExcelWorkbook

with ExcelWorkbook("report.xlsx", visible=False) as book:
    table = book.get_bordered_table(
        "Report",
        row=5,
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

### 表の中身を変更する

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

### 行や列を追加する

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

### 見出しを追加する

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

### 見出し行の値から罫線テーブルを探す

次のように、左側に行見出し列があり、右側に `amount` のような値列が複数続く表も取得できます。

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
    table = book.get_bordered_table_by_header(
        "Sheet1",
        header_values=["header1", "header2"],
        value_header_contains="amount",
    )

    print(table.row_headers)
    print(table.column_headers)
    print(table.data)
```

この例では、`header1` と `header2` が行見出し列、`amount` を含む列が値領域の列見出しになります。
文字比較はデフォルトで大文字小文字を区別しません。

Excelフォーマットから表を抽出する
----------------------------------

別のExcelファイルを「表のフォーマット定義」として使い、同じ構造を持つ表を対象ブックから自動検索できます。
フォーマット定義は1シートにつき1パターンで、シート名がパターン名になります。

サンプル:

```text
samples/extraction_format.xlsx
samples/extraction_input.xlsx
```

### フォーマットExcelを作る

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

### `{{title}}` で単一セルを取得する

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

### 対象Excelから抽出する

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

### 検索対象を絞る

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

API 一覧
--------

```python
from openpyxlwings import ExcelWorkbook
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
| `book.get_bordered_table(sheet, row, column, header_rows=1, header_columns=0)` | 起点セルを含む罫線テーブルを取得する |
| `book.get_bordered_table_by_header(sheet, header_values, value_header_contains=...)` | 見出し行の値と値列見出しの文字列から罫線テーブルを取得する |
| `ExcelFormat.load(path)` | Excelフォーマットブックを読み込む |
| `book.extract(pattern, sheets=None, ranges=None)` | フォーマットに一致する表をすべて抽出する |
| `book.save(path=None)` | 明示的に保存する |
| `book.close(save=True)` | 開いている内部セッションを閉じる |

互換用の名前
------------

以前のコード向けに、以下の名前も残しています。
ただし、新しく書くコードでは `ExcelWorkbook` を使うのがおすすめです。

```python
from openpyxlwings import ExcelReader, ExcelWriter

assert ExcelReader is ExcelWorkbook
assert ExcelWriter is ExcelWorkbook
```

CLI
---

シート名を表示:

```bash
openpyxlwings sheets report.xlsx
```

指定範囲を JSON で表示:

```bash
openpyxlwings read report.xlsx Sheet1 A1:D20
```

開発メモ
--------

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

ライセンス
----------

MIT License です。詳細は `LICENSE` を参照してください。
