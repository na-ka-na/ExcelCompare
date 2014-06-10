package com.ka.spreadsheet.diff;

import static com.ka.spreadsheet.diff.TestUtils.assertTrue;
import static com.ka.spreadsheet.diff.TestUtils.verifyFileContentsSame;

import java.io.File;
import java.io.PrintStream;

import javax.annotation.Nullable;

public class SpreadSheetDifferSmokeTest {
	
	private static final File TEMP_DIR = new File("test/resources");

	public static void main(String[] args) throws Exception {
		testDiff(
			"Identical xlsx files",
			new String[]{"test/resources/ss1.xlsx", "test/resources/ss1.xlsx"},
			new File("test/resources/ss1_xlsx_ss1_xlsx.out"),
			null);
		testDiff(
			"Diff xlsx files",
			new String[]{"test/resources/ss1.xlsx", "test/resources/ss2.xlsx"},
			new File("test/resources/ss1_xlsx_ss2_xlsx.out"),
			null);
		testDiff(
			"Diff ods files",
			new String[]{"test/resources/ss1.ods", "test/resources/ss2.ods"},
			new File("test/resources/ss1_ods_ss2_ods.out"),
			null);
		testDiff(
			"Diff xlsx and ods",
			new String[]{"test/resources/ss3.xlsx", "test/resources/ss3.ods"},
			new File("test/resources/ss3_xlsx_ss3_ods.out"),
			null);
		testDiff(
			"Diff ods and xlsx",
			new String[]{"test/resources/ss3.ods", "test/resources/ss3.xlsx"},
			new File("test/resources/ss3_ods_ss3_xlsx.out"),
			null);
		testDiff(
			"Missing file",
			new String[]{"test/resources/missingfile", "test/resources/ss1.xlsx"},
			null,
			new File("test/resources/missing_file.err"));
		testDiff(
			"Bad file",
			new String[]{"test/resources/badfile.txt", "test/resources/ss1.xlsx"},
			null,
			new File("test/resources/bad_file.err"));
		testDiff(
			"Numeric and formula xls xlsx",
			new String[]{"test/resources/numeric_and_formula.xls", "test/resources/numeric_and_formula.xlsx"},
			new File("test/resources/numeric_and_formula.xls.xlsx.out"),
			null);
		testDiff(
			"Numeric and formula xls odf",
			new String[]{"test/resources/numeric_and_formula.xls", "test/resources/numeric_and_formula.ods"},
			new File("test/resources/numeric_and_formula.xls.ods.out"),
			null);
		testDiff(
			"Nullable Sheet",
			new String[]{"test/resources/Multisheet.xls", "test/resources/Multisheet.xls", "--ignore1", "::B", "--ignore2", "::B"},
			new File("test/resources/nullableSheet_xls.out"),
			null);
		System.out.println("All tests pass");
	}
	
	public static void testDiff(String testName, String[] args,
		@Nullable File expectedOutFile, @Nullable File expectedErrFile) throws Exception {
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
				if(err != null)	err.close();
			}
			
		} finally {
			if(out != null)	out.close();
		}
		assertTrue(testCompleted);
		verifyFileContentsSame(errFile, expectedErrFile);
		verifyFileContentsSame(outFile, expectedOutFile);
		System.out.println(testName + " passed");
	}
}
