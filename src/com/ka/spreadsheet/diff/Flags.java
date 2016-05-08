package com.ka.spreadsheet.diff;

import java.io.File;

public class Flags {

  private static final String DEBUG_FLAG = "--debug";
  // double value, default null
  private static final String DIFF_NUMERIC_PRECISION_FLAG = "--diff_numeric_precision";

  public static boolean DEBUG = false;
  public static Double DIFF_NUMERIC_PRECISION = null;
  public static File WORKBOOK1 = null;
  public static File WORKBOOK2 = null;
  public static WorkbookIgnores WORKBOOK_IGNORES1 = null;
  public static WorkbookIgnores WORKBOOK_IGNORES2 = null;

  public static boolean parseFlags(String[] args) {
    int idx = findFlag(DEBUG_FLAG, args);
    if (idx != -1) {
      DEBUG = true;
      args = removeFlag(idx, args);
    }
    idx = findFlag(DIFF_NUMERIC_PRECISION_FLAG, args);
    if (idx != -1) {
      DIFF_NUMERIC_PRECISION = parseDoubleFlagValue(idx, args);
      args = removeFlag(idx, args);
    }
    if (args.length < 2) {
      System.out.println(usage());
      return false;
    }
    WORKBOOK1 = new File(args[0]);
    WORKBOOK2 = new File(args[1]);
    WORKBOOK_IGNORES1 = WorkbookIgnores.parseWorkbookIgnores(args, "--ignore1");
    WORKBOOK_IGNORES2 = WorkbookIgnores.parseWorkbookIgnores(args, "--ignore2");
    return true;
  }

  private static int findFlag(String flag, String[] args) {
    for (int i = 0; i < args.length; i++) {
      if (args[i].startsWith(flag))
        return i;
    }
    return -1;
  }

  private static double parseDoubleFlagValue(int flagIdx, String[] args) {
    String flag = args[flagIdx];
    return Double.parseDouble(flag.substring(flag.indexOf("=") + 1, flag.length()));
  }

  private static String[] removeFlag(int flagIdx, String[] args) {
    String[] args1 = new String[args.length - 1];
    for (int i = 0; i < flagIdx; i++)
      args1[i] = args[i];
    for (int i = flagIdx + 1; i < args.length; i++)
      args1[i - 1] = args[i];
    return args1;
  }

  private static String usage() {
    return "Usage> excel_cmp <diff-flags> <file1> <file2> [--ignore1 <sheet-ignore-spec> <sheet-ignore-spec> ..] [--ignore2 <sheet-ignore-spec> <sheet-ignore-spec> ..]"
        + "\n"
        + "\n"
        + "Notes: * Prints all diffs & extra cells on stdout"
        + "\n"
        + "       * Process exits with 0 if workbooks match, 1 otherwise"
        + "\n"
        + "       * Works with both xls, xlsx, xlsm, ods. You may compare any of xls, xlsx, xlsm, ods with each other"
        + "\n"
        + "       * Compares only cell \"contents\". Formatting, macros are not diffed. Although there is rudimentary support for recognizing if an xlsm file contains macros"
        + "\n"
        + "       * Using --ignore1 & --ignore2 (optional) you may tell the diff to ignore cells"
        + "\n"
        + "       * Give one and only one <sheet-ignore-spec> for a sheet"
        + "\n"
        + "\n"
        + "\n"
        + "Diff flags"
        + "\n"
        + "       * --diff_numeric_precision: by default numbers are diffed with double precision, to change that specify this flag as --diff_numeric_precision=0.0001"
        + "\n"
        + "\n"
        + "Sheet Ignore Spec:  <sheet-name>:<row-ignore-spec>:<column-ignore-spec>:<cell-ignore-spec>"
        + "\n"
        + "                    * Leaving <sheet-name> blank corresponds to this spec applying to all sheets"
        + "\n"
        + "                      for example ::A will ignore column A in all sheets"
        + "\n"
        + "                    * To ignore whole sheet, just provide <sheet-name>"
        + "\n"
        + "                    * Any cell satisfying any ignore spec in the sheet (row, col, or cell) will be ignored in diff"
        + "\n"
        + "                    * You may provide only <cell-ignore-spec> as - <sheet-name>:::<cell-ignore-spec>"
        + "\n"
        + "\n"
        + "Row Ignore Spec:    <comma sep list of row or row-range>"
        + "\n"
        + "                    * Row numbers begin from 1"
        + "\n"
        + "                    * Range of rows may be provide as: 1-10"
        + "\n"
        + "                    * Rows and ranges may be mixed as: 1-10,12,20-30 etc."
        + "\n"
        + "\n"
        + "Column Ignore Spec: <comma sep list of column or column-range>"
        + "\n"
        + "                    * Similar to Row Ignore Spec"
        + "\n"
        + "                    * Columns are letters starting with A"
        + "\n"
        + "\n"
        + "Cell Ignore Spec:   <comma sep list of cell or cell-range>"
        + "\n"
        + "                    * Similar to Row Ignore Spec"
        + "\n"
        + "                    * Cells are in usual Excel notation A1 D10"
        + "\n"
        + "                    * Range may be provided as A1-D10"
        + "\n"
        + "\n"
        + "Example command line: "
        + "\n"
        + "excel_cmp 1.xlsx 2.xlsx --ignore1 Sheet1:::A1,B1,J10,K11,D4 Sheet2:::A1 --ignore2 Sheet1:::A1,D4,J10 Sheet3:::A1"
        + "\n";
  }
}
