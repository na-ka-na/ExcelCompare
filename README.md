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

Report bugs / issues / requests [here](https://github.com/na-ka-na/ExcelCompare/issues)

## Installation

### Prerequisites

* Requires Java 1.6 or higher.
* Assumes Java is added to PATH (to check open a cmd and run java -version)
* No other platform specific requirements
* A shell script and a bat script are packaged

Just [download](https://github.com/na-ka-na/ExcelCompare/releases) the zip file.

Extract it anywhere (and optionally you add the bin folder to PATH).

## Usage

    $ excel_cmp <file1> <file2> [--ignore1 <sheet-ignore-spec> ..] [--ignore2 <sheet-ignore-spec> ..]

Notes:

* --ignore1 (file1) and --ignore2 (file2) are independent of each other
* Give one and only one &lt;sheet-ignore-spec> per sheet
* File path is assumed relative to current directory unless full path is provided

### Sheet Ignore Spec
    <sheet-name>:<row-ignore-spec>:<column-ignore-spec>:<cell-ignore-spec>
    
* Everything except &lt;sheet-name> is optional
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
        
* Ignore column A in all spreadsheets in both

        excel_cmp 1.xlsx 2.xlsx --ignore1 ::A --ignore2 ::A       

* Ignore columns A,D and rows 1-5, 20-25

        excel_cmp 1.xlsx 2.xlsx --ignore1 Sheet1:1-5,20-25:A,D --ignore2 Sheet1:1-5,20-25:A,D

* Ignore columns A,D and rows 1-5, 20-25 and cells F6,H8

        excel_cmp 1.xlsx 2.xlsx --ignore1 Sheet1:1-5,20-25:A,D:F6,H8 --ignore2 Sheet1:1-5,20-25:A,D:F6,H8

* Ignore column A in Sheet1 and column B in Sheet2

        excel_cmp 1.xlsx 2.xlsx --ignore1 Sheet1::A Sheet2::B --ignore2 Sheet1::A Sheet2::B

* Ignore cells A1-B10 in Sheet2 of both files

        excel_cmp 1.xlsx 2.xlsx --ignore1 Sheet2:::A1-B10 --ignore2 Sheet2:::A1-B10


## Output format
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
<pre>
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
</pre>

* Only extra cells
<pre>
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
</pre>


* No diff
<pre>
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
</pre>
