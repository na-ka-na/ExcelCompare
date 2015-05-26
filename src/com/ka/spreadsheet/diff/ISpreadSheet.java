package com.ka.spreadsheet.diff;

import java.util.Iterator;

/**
 * All indexes are zero based
 */
public interface ISpreadSheet {

	Iterator<ISheet> getSheetIterator();
}

interface ISheet {

	String getName();

	int getSheetIndex();

	Iterator<IRow> getRowIterator();
}

interface IRow {

	int getRowIndex();

	Iterator<ICell> getCellIterator();
}

interface ICell {

	int getRowIndex();

	int getColumnIndex();

	String getStringValue();
}
