package com.ka.spreadsheet.diff;

import java.util.Iterator;

import javax.annotation.Nullable;

/**
 * All indexes are zero based
 */
public interface ISpreadSheet {

	Iterator<ISheet> getSheetIterator();

	@Nullable
	Boolean hasMacro();
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
