## Introduction

ExcelCompare is a command line tool (coming soon API) to diff Excel / Open document (ods) (Open office, Libre office) spreadsheets.

It uses the [Apache POI](http://poi.apache.org) library to read Excel files.
And the [OdfToolkit] (http://incubator.apache.org/odftoolkit) library to read Open document (ods) files.

This software is distributed under the [MIT](http://www.opensource.org/licenses/MIT) license.

## Features / Limitations

* Identifies extra cells / sheets in addition to common cells.
* Prints all diffs & extra cells on stdout.
* Process exits with 0 if workbooks match, 1 otherwise.
* Works with xls, xlsx, xlsm, ods. You may compare any of these with each other.
* Compares only cell "contents". Formatting, macros are currently ignored.
* Using --ignore1 & --ignore2 (both optional) you may tell the diff to skip any number of sheets / rows / columns / cells.
* Other flags to control diffing (see below for description of these): --diff_numeric_precision, --diff_ignore_formulas, --diff_format.

Report bugs / issues / requests [here](https://github.com/na-ka-na/ExcelCompare/issues)

## Installation

### Prerequisites

* Requires Java 1.8 or higher.
* Assumes Java is added to PATH (to check open a cmd and run java -version)
* No other platform specific requirements
* A shell script and a bat script are packaged

Just [download](https://github.com/na-ka-na/ExcelCompare/releases) the zip file.

Extract it anywhere (and optionally you add the extracted folder to PATH).

### macOS

[Homebrew](http://brew.sh/) makes it easy to install ExcelCompare:

    $ brew update
    $ brew install excel-compare

## Usage

    $ excel_cmp <diff-flags> <file1> <file2> [--ignore1 <sheet-ignore-spec> ..] [--ignore2 <sheet-ignore-spec> ..]

Notes:

* --ignore1 (file1) and --ignore2 (file2) are independent of each other
* Give one and only one &lt;sheet-ignore-spec> per sheet
* File path is assumed relative to current directory unless full path is provided
* file1 and/or file2 can be '/dev/null' or '\\.\NUL' (on windows) which is treated as empty file. This is useful for using ExcelCompare for git diff.

### Diff flags

* --diff_numeric_precision: by default numbers are diffed with double precision, to change that specify this flag as --diff_numeric_precision=0.0001
* --diff_ignore_formulas: by default for cells with formula, formula is compared instead of the evaluated value. Use this flag to compare evaluated value instead
* --diff_format: by default output is in 'excel_cmp' format, use --diff_format=unified to output in Unified Diff format instead

### Sheet Ignore Spec
    <sheet-name>:<row-ignore-spec>:<column-ignore-spec>:<cell-ignore-spec>

* Leaving &lt;sheet-name> blank corresponds to this spec applying to all sheets, for example ::A will ignore column A in all sheets
* To ignore whole sheet, just provide &lt;sheet-name>
* Any cell satisfying any ignore spec in the sheet (row, col, or cell) will be ignored in diff
* You may provide only &lt;cell-ignore-spec> as - &lt;sheet-name>:::&lt;cell-ignore-spec>

### Row Ignore Spec
    <comma sep list of row or row-range>

* Row numbers begin from 1
* Range of rows may be provide as: 1-10
* Rows and ranges may be mixed as: 1-10,12,20-30 etc.

### Column Ignore Spec
    <comma sep list of column or column-range>

* Similar to Row Ignore Spec
* Columns are letters starting with A

### Cell Ignore Spec
    <comma sep list of cell or cell-range>

* Similar to Row Ignore Spec
* Cells are in usual Excel notation A1 D10
* Range may be provided as A1-D10

### Examples

* Diff all cells

        excel_cmp 1.xlsx 2.xlsx
    
* Ignore Sheet1 in 1.xlsx
    
        excel_cmp 1.xlsx 2.xlsx --ignore1 Sheet1
 
* Ignore Sheet1 in both

        excel_cmp 1.xlsx 2.xlsx --ignore1 Sheet1 --ignore2 Sheet1

* Ignore column A in both 

        excel_cmp 1.xlsx 2.xlsx --ignore1 Sheet1::A --ignore2 Sheet1::A
        
* Ignore column A across all sheets in both 

        excel_cmp 1.xlsx 2.xlsx --ignore1 ::A --ignore2 ::A

* Ignore columns A,D and rows 1-5, 20-25

        excel_cmp 1.xlsx 2.xlsx --ignore1 Sheet1:1-5,20-25:A,D --ignore2 Sheet1:1-5,20-25:A,D

* Ignore columns A,D and rows 1-5, 20-25 and cells F6,H8

        excel_cmp 1.xlsx 2.xlsx --ignore1 Sheet1:1-5,20-25:A,D:F6,H8 --ignore2 Sheet1:1-5,20-25:A,D:F6,H8

* Ignore column A in Sheet1 and column B in Sheet2

        excel_cmp 1.xlsx 2.xlsx --ignore1 Sheet1::A Sheet2::B --ignore2 Sheet1::A Sheet2::B

* Ignore cells A1-B10 in Sheet2 of both files

        excel_cmp 1.xlsx 2.xlsx --ignore1 Sheet2:::A1-B10 --ignore2 Sheet2:::A1-B10

* Ignore column A in all sheets of both files

        excel_cmp 1.xlsx 2.xlsx --ignore1 ::A --ignore2 ::A

## Native ("excel_cmp") Output format
* Each diff or extra cell is reported per line as follows

        DIFF  Cell at      <Cell> => <Value1> v/s <Value2>
        EXTRA Cell in <WB> <Cell> => <Value>

* Then a summary

        ----------------- DIFF -------------------
        Sheets: [<Set of sheets with diffs>]
        Rows: [<Set of rows with diffs>]
        Cols: [<Set of columns with diffs>]
        ----------------- EXTRA WB1 -------------------
        Sheets: [<Set of extra sheets in WB1>]
        Rows: [<Set of extra rows in WB1>]
        Cols: [<Set of extra columns in WB1>]
        ----------------- EXTRA WB2 -------------------
        Sheets: [<Set of extra sheets in WB2>]
        Rows: [<Set of extra rows in WB2>]
        Cols: [<Set of extra columns in WB2>]
        -----------------------------------------

* Then one line

        Excel files <file1> and <file2> differ|match

### Examples

* Diffs in cells and extra cells

```
> excel_cmp xxx.xlsx yyy.xlsx
DIFF  Cell at     Sheet1!A1 => 'a' v/s 'aa'
EXTRA Cell in WB1 Sheet1!B1 => 'cc'
DIFF  Cell at     Sheet1!D4 => '4.0' v/s '14.0'
EXTRA Cell in WB2 Sheet1!J10 => 'j'
EXTRA Cell in WB1 Sheet1!K11 => 'k'
EXTRA Cell in WB1 Sheet2!A1 => 'abc'
EXTRA Cell in WB2 Sheet3!A1 => 'haha'
----------------- DIFF -------------------
Sheets: [Sheet1]
Rows: [1, 4]
Cols: [A, D]
----------------- EXTRA WB1 -------------------
Sheets: [Sheet1, Sheet2]
Rows: [1, 11]
Cols: [B, K, A]
----------------- EXTRA WB2 -------------------
Sheets: [Sheet1, Sheet3]
Rows: [10, 1]
Cols: [J, A]
-----------------------------------------
Excel files xxx.xlsx and yyy.xlsx differ
```

* Only extra cells

```
excel_cmp xxx.xlsx yyy.xlsx --ignore1 Sheet1 --ignore2 Sheet1
EXTRA Cell in WB1 Sheet2!A1 => 'abc'
EXTRA Cell in WB2 Sheet3!A1 => 'haha'
----------------- DIFF -------------------
Sheets: []
Rows: []
Cols: []
----------------- EXTRA WB1 -------------------
Sheets: [Sheet2]
Rows: [1]
Cols: [A]
----------------- EXTRA WB2 -------------------
Sheets: [Sheet3]
Rows: [1]
Cols: [A]
-----------------------------------------
Excel files xxx.xlsx and yyy.xlsx differ
```

* No diff

```
excel_cmp xxx.xlsx yyy.xlsx --ignore1 Sheet1 Sheet2 Sheet3 --ignore2 Sheet1 Sheet2 Sheet3
----------------- DIFF -------------------
Sheets: []
Rows: []
Cols: []
----------------- EXTRA WB1 -------------------
Sheets: []
Rows: []
Cols: []
----------------- EXTRA WB2 -------------------
Sheets: []
Rows: []
Cols: []
-----------------------------------------
Excel files xxx.xlsx and yyy.xlsx match
```

## Unified Diff output format
* Diffs are reported in the "unified diff" style, with no surrounding context (_i.e._, a la `diff -U0`).
* Each sheet containing a diff or an extra cell begins with a header as follows:
		--<FileName1>!<SheetName>
		++<FileName2>!<SheetName>
* Each row containing a diff or an extra cell begins with a line that identifies a contiguous series of cells in the row as follows:
		@@ <Row><ColumnM>,<Row><ColumnN> <Row><ColumnM>,<Row><ColumnN>  @@
* Each diff or extra cell is reported as follows:
		-<ColumnMValue1>
		-...
		-<ColumnNValue1>
		+<ColumnMValue2>
		+...
		+<ColumnNValue2>
* If there are multiple series of diff or extra cells, the row header and cell data will be repeated, with the column numbers idetifying the start and end of each series.
* There is no summary, and if there are no diffs and no extra cells, the output is empty.

### Examples

* Diffs in cells and extra cells

```diff
> excel_cmp --diff-format=unified xxx.xlsx yyy.xlsx
--- xxx.xlsx!Sheet1
+++ yyy.xlsx!Sheet1
@@ A1,B1 A1,B1 @@
-a
-cc
+aa
+
@@ D4 D4 @@
-4.0
+14.0
@@ J10 J10 @@
-
+j
@@ K11 K11 @@
-k
+
--- xxx.xlsx!Sheet2
+++ yyy.xlsx!Sheet2
@@ A1 A1 @@
-abc
+
--- xxx.xlsx!Sheet3
+++ yyy.xlsx!Sheet3
@@ A1 A1 @@
-
+haha
```

* Only extra cells

```diff
> excel_cmp --diff-format=unified xxx.xlsx yyy.xlsx --ignore1 Sheet1 --ignore2 Sheet1
--- xxx.xlsx!Sheet2
+++ yyy.xlsx!Sheet2
@@ A1 A1 @@
-abc
+
--- xxx.xlsx!Sheet3
+++ yyy.xlsx!Sheet3
@@ A1 A1 @@
-
+haha
```

* No diff

```diff
> excel_cmp --diff-format=unified xxx.xlsx yyy.xlsx --ignore1 Sheet1 Sheet2 Sheet3 --ignore2 Sheet1 Sheet2 Sheet3
```
