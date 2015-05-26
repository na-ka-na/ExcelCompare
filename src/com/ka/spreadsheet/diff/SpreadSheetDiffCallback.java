package com.ka.spreadsheet.diff;

import java.io.File;

public interface SpreadSheetDiffCallback {

    void reportDiffCell(CellPos c1, CellPos c2);

    void reportExtraCell(boolean inFirstSpreadSheet, CellPos c);

    void reportWorkbooksDiffer(boolean differ, File file1, File file2);
}
