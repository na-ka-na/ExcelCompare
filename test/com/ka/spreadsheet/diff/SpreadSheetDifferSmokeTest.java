package com.ka.spreadsheet.diff;

import static com.ka.spreadsheet.diff.TestUtils.assertTrue;
import static com.ka.spreadsheet.diff.TestUtils.verifyFileContentsSame;

import java.io.File;
import java.io.PrintStream;

import javax.annotation.Nullable;

public class SpreadSheetDifferSmokeTest {

  private static final File TEMP_DIR = new File("test/resources");

  private static final boolean isWindows = "\\".equals(System.getProperty("file.separator"));

  public static void main(String[] args) throws Exception {
    testDiff(
        "Identical xlsx files",
        new String[] {"test/resources/ss1.xlsx", "test/resources/ss1.xlsx"},
        resultFile("test/resources/ss1_xlsx_ss1_xlsx.out"),
        null);
    testDiff(
        "Diff xlsx files",
        new String[] {"test/resources/ss1.xlsx", "test/resources/ss2.xlsx"},
        resultFile("test/resources/ss1_xlsx_ss2_xlsx.out"),
        null);
    testDiff(
        "Diff ods files",
        new String[] {"test/resources/ss1.ods", "test/resources/ss2.ods"},
        resultFile("test/resources/ss1_ods_ss2_ods.out"),
        null);
    testDiff(
        "Diff xlsx and ods",
        new String[] {"test/resources/ss3.xlsx", "test/resources/ss3.ods"},
        resultFile("test/resources/ss3_xlsx_ss3_ods.out"),
        null);
    testDiff(
        "Diff ods and xlsx",
        new String[] {"test/resources/ss3.ods", "test/resources/ss3.xlsx"},
        resultFile("test/resources/ss3_ods_ss3_xlsx.out"),
        null);
    testDiff(
        "Missing file",
        new String[] {"test/resources/missingfile", "test/resources/ss1.xlsx"},
        null,
        resultFile("test/resources/missing_file.err"));
    testDiff(
        "Bad file",
        new String[] {"test/resources/badfile.txt", "test/resources/ss1.xlsx"},
        null,
        resultFile("test/resources/bad_file.err"));
    testDiff(
        "Numeric and formula xls xlsx",
        new String[] {"test/resources/numeric_and_formula.xls",
            "test/resources/numeric_and_formula.xlsx"},
        resultFile("test/resources/numeric_and_formula.xls.xlsx.out"),
        null);
    testDiff(
        "Numeric and formula xls odf",
        new String[] {"test/resources/numeric_and_formula.xls",
            "test/resources/numeric_and_formula.ods"},
        resultFile("test/resources/numeric_and_formula.xls.ods.out"),
        null);
    testDiff(
        "Numeric and formula odf xlsx with flag",
        new String[] {"--diff_ignore_formulas",
                      "test/resources/numeric_and_formula.ods",
                      "test/resources/numeric_and_formula.xlsx"},
        resultFile("test/resources/numeric_and_formula_ignoreformulaflag.ods.xlsx.out"),
        null);
    testDiff(
        "Nullable Sheet",
        new String[] {"test/resources/MultiSheet.xls", "test/resources/MultiSheet.xls",
            "--ignore1", "::B", "--ignore2", "::B"},
        resultFile("test/resources/nullableSheet_xls.out"),
        null);
    testDiff(
        "Ignore single cell",
        new String[] {"test/resources/ss3.xlsx", "test/resources/ss3.ods",
            "--ignore1", "Sheet1:2:B", "--ignore2", "Sheet1:2:B"},
        resultFile("test/resources/ss3_xlsx_ss3_ignore2B_ods.out"),
        null);
    testDiff(
        "Macro diff",
        new String[] {"test/resources/ss_with_macro.xlsm",
            "test/resources/ss_without_macro.xlsx"},
        resultFile("test/resources/macro_diff.out"),
        null);
    testDiff(
        "Numeric precision diff without flag",
        new String[] {"test/resources/ss1_numeric_precision.xlsx",
            "test/resources/ss2_numeric_precision.xlsx"},
        resultFile("test/resources/numeric_precision_diff.out"),
        null);
    testDiff(
        "Numeric precision diff with flag",
        new String[] {"--diff_numeric_precision=0.0000001",
            "test/resources/ss1_numeric_precision.xlsx",
            "test/resources/ss2_numeric_precision.xlsx"},
        resultFile("test/resources/numeric_precision_no_diff.out"),
        null);
    if (!isWindows) {
      testDiff(
          "File1 is /dev/null",
          new String[] {"test/resources/ss1.xlsx", "/dev/null"},
          resultFile("test/resources/ss1_xlsx_dev_null.out"),
          null);
      testDiff(
          "File2 is /dev/null",
          new String[] {"/dev/null", "test/resources/ss1.xlsx"},
          resultFile("test/resources/dev_null_ss1_xlsx.out"),
          null);
    }
    testDiff(
        "With without formula with flag",
        new String[] {"--diff_ignore_formulas",
                      "test/resources/ss_without_formula.xlsx",
                      "test/resources/ss_with_formula.xlsx"},
        resultFile("test/resources/ss_with_without_formula_ignoreformulaflag.out"),
        null);
    testDiff(
        "With without formula without flag",
        new String[] {"test/resources/ss_without_formula.xlsx",
                      "test/resources/ss_with_formula.xlsx"},
        resultFile("test/resources/ss_with_without_formula.out"),
        null);
    testDiff(
        "Identical xlsx files with Unified Diff output format",
        new String[] {"--diff_format=unified",
                      "test/resources/ss1.xlsx",
                      "test/resources/ss1.xlsx"},
        resultFile("test/resources/ss1_xlsx_ss1_xlsx_udiff.out"),
        null);
    testDiff(
        "Diff xlsx files with Unified Diff output format",
        new String[] {"--diff_format=unified",
                      "test/resources/ss1.xlsx",
                      "test/resources/ss2.xlsx"},
        resultFile("test/resources/ss1_xlsx_ss2_xlsx_udiff.out"),
        null);
     testDiff(
        "Diff ods files with Unified Diff output format",
        new String[] {"--diff_format=unified",
                      "test/resources/ss1.ods", "test/resources/ss2.ods"},
        resultFile("test/resources/ss1_ods_ss2_ods_udiff.out"),
        null);
    testDiff(
        "Diff xlsx and ods with Unified Diff output format",
        new String[] {"--diff_format=unified",
                      "test/resources/ss3.xlsx", "test/resources/ss3.ods"},
        resultFile("test/resources/ss3_xlsx_ss3_ods_udiff.out"),
        null);
    testDiff(
        "Diff ods and xlsx with Unified Diff output format",
        new String[] {"--diff_format=unified",
                      "test/resources/ss3.ods", "test/resources/ss3.xlsx"},
        resultFile("test/resources/ss3_ods_ss3_xlsx_udiff.out"),
        null);
    testDiff(
        "Numeric and formula xls xlsx with Unified Diff output format",
        new String[] {"--diff_format=unified",
                      "test/resources/numeric_and_formula.xls",
            "test/resources/numeric_and_formula.xlsx"},
        resultFile("test/resources/numeric_and_formula.xls.xlsx_udiff.out"),
        null);
    testDiff(
        "Numeric and formula xls odf with Unified Diff output format",
        new String[] {"--diff_format=unified",
                      "test/resources/numeric_and_formula.xls",
            "test/resources/numeric_and_formula.ods"},
        resultFile("test/resources/numeric_and_formula.xls.ods_udiff.out"),
        null);
    testDiff(
        "Numeric and formula odf xlsx with flag with Unified Diff output format",
        new String[] {"--diff_format=unified",
                      "--diff_ignore_formulas",
                      "test/resources/numeric_and_formula.ods",
                      "test/resources/numeric_and_formula.xlsx"},
        resultFile("test/resources/numeric_and_formula_ignoreformulaflag.ods.xlsx_udiff.out"),
        null);
    testDiff(
        "Ignore single cell with Unified Diff output format",
        new String[] {"--diff_format=unified",
                      "test/resources/ss3.xlsx", "test/resources/ss3.ods",
            "--ignore1", "Sheet1:2:B", "--ignore2", "Sheet1:2:B"},
        resultFile("test/resources/ss3_xlsx_ss3_ignore2B_ods_udiff.out"),
        null);
    testDiff(
        "With without formula with flag with Unified Diff output format",
        new String[] {"--diff_format=unified",
                      "--diff_ignore_formulas",
                      "test/resources/ss_without_formula.xlsx",
                      "test/resources/ss_with_formula.xlsx"},
        resultFile("test/resources/ss_with_without_formula_ignoreformulaflag_udiff.out"),
        null);
    testDiff(
        "With without formula without flag with Unified Diff output format",
        new String[] {"--diff_format=unified",
                      "test/resources/ss_without_formula.xlsx",
                      "test/resources/ss_with_formula.xlsx"},
        resultFile("test/resources/ss_with_without_formula_udiff.out"),
        null);
    System.err.println("All tests pass");
  }

  private static File resultFile(String resultFile) {
    return new File(isWindows ? ("win_" + resultFile) : resultFile);
  }

  public static void testDiff(String testName, String[] args, @Nullable File expectedOutFile,
      @Nullable File expectedErrFile) throws Exception {
    System.err.print(testName + "... ");
    PrintStream oldOut = System.out;
    PrintStream oldErr = System.err;
    File outFile = File.createTempFile("testOutput", "out", TEMP_DIR);
    File errFile = File.createTempFile("testOutput", "err", TEMP_DIR);
    outFile.deleteOnExit();
    errFile.deleteOnExit();
    boolean testCompleted = false;
    PrintStream out = null;
    PrintStream err = null;
    try {
      out = new PrintStream(outFile);
      try {
        err = new PrintStream(errFile);
        try {
          System.setOut(out);
          System.setErr(err);
          SpreadSheetDiffer.doDiff(args);
          testCompleted = true;
        } finally {
          System.setOut(oldOut);
          System.setErr(oldErr);
        }
      } finally {
        if (err != null)
          err.close();
      }

    } finally {
      if (out != null)
        out.close();
    }
    assertTrue(testCompleted);
    verifyFileContentsSame(errFile, expectedErrFile);
    verifyFileContentsSame(outFile, expectedOutFile);
    System.err.println("passed");
  }
}
