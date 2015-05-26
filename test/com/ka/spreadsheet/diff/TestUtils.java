package com.ka.spreadsheet.diff;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileReader;
import java.io.IOException;
import java.util.LinkedList;

import javax.annotation.Nullable;

public class TestUtils {

	/**
	 * Treat nulls as empty files.
	 */
	public static void verifyFileContentsSame(@Nullable File actual, @Nullable File expected) throws IOException {
		LinkedList<String> actualLines = actual == null
			? new LinkedList<String>()
			: readFileIntoLines(actual);
		LinkedList<String> expectedLines = expected == null
			? new LinkedList<String>()
			: readFileIntoLines(expected);
		for (int lineNum = 1;; lineNum++) {
			String actualLine = actualLines.poll();
			String expectedLine = expectedLines.poll();
			if ((actualLine == null) && (expectedLine == null)) {
				break;
			}
			assertEquals("Line " + lineNum + " differs", actualLine, expectedLine);
		}
	}

	public static LinkedList<String> readFileIntoLines(File file) throws IOException {
		LinkedList<String> lines = new LinkedList<String>();
		BufferedReader reader = new BufferedReader(new FileReader(file));
		try {
			String line = null;
			while ((line = reader.readLine()) != null) {
				lines.add(line);
			}
		}
		finally	{
			if (reader != null)	reader.close();
		}
		return lines;
	}

	public static void assertEquals(Object actual, Object expected) {
		assertEquals("assertEquals failed", actual, expected);
	}

	public static void assertEquals(String messagePrefix, Object actual, Object expected) {
		if (((actual == null && expected != null) || (actual != null && expected == null)) || !actual.equals(expected)) {
			throw new AssertionError(
				messagePrefix
				+ "\nactual: " + actual
				+ "\nexpected: " + expected);
		}
	}

	public static void assertTrue(boolean expected) {
		assertEquals("assertTrue failed", true, expected);
	}

	public static void assertFalse(boolean expected) {
		assertEquals("assertFalse failed", false, expected);
	}

	public static void fail() {
		throw new AssertionError();
	}
}
