package com.ka.spreadsheet.diff;

import java.io.File;

/**
 * Contract:
 *	*	Caller will call init() before calling any other methods.
 *	*	Caller will process the worksheets of the spreadsheets in order
 *		by sheet index.  Sheets may not be processed in name order.
 *	*	Caller will process each worksheet in its entirety before processing
 *		the next worksheet.
 *	*	Caller will process the rows of a worksheet in number order from the
 *		top (e.g., 1, 2, ...).
 *	*	Caller will process each row of a worksheet in its entirety before
 *		processing the next row.
 *	*	Caller will process each cell of a row in alphabetic order from the
 *		left (e.g., B1, B2, B3, ...).
 *	*	Caller will call reportExtraCell() to report "extra" cells in a row 
 *		after calling reportDiffCell() to report any differences in the row,
 *		and will not call reportDiffCell() for the same row afterwards.  Extra
 *		cells will always be at the end of the row (e.g., if File1_Sheet1
 *		contains A1-A3, and File2_Sheet1 contains A1-A5, then
 *		File2_Sheet1!A4-A5 are extra cells).
 *	*	Caller will call reportExtraCell() to report "extra" rows in a
 *		worksheet after calling reportDiffCell() to report any differences in
 *		last row, and will not call reportDiffCell() for the same worksheet
 *		afterwards.  Extra will always be at the bottom of the worksheet (e.g.,
 *		if File1_Sheet1 contains A1-B3, and File2_Sheet1 contains A1-C3, then
 *		File2_Sheet1!C1-C3 are extra cells).
 *	*	Caller will call reportWorkbooksDiffer() will only be called once all
 *		differing and extra cells and any other differences have been reported.
 *	*	Caller will call finish() after calling all other methods, and will
 *		not call any other methods afterwards.
 */
public interface SpreadSheetDiffCallback {

  void init(String file1, String file2);

  void finish();

  void reportDiffCell(CellPos c1, CellPos c2);

  void reportExtraCell(boolean inFirstSpreadSheet, CellPos c);

  void reportMacroOnlyIn(boolean inFirstSpreadSheet);

  void reportWorkbooksDiffer(boolean differ);
}
