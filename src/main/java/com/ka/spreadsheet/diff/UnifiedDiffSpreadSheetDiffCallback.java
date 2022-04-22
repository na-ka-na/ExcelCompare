package com.ka.spreadsheet.diff;

import java.io.File;
import java.util.ArrayList;
import java.util.List;

import static com.ka.spreadsheet.diff.SpreadSheetUtils.CELL_INTERNAL_TO_USER;

// Output format, inspired by the traditional Unified Diff format (but using cell references instead of line numbers):
//      +--- column 1
//      V
//        --- file_1_name!sheet_1_name
//        +++ file_2_name!sheet_1_name
//        @@ cell_1_ref,cell_n_ref cell_1_ref,cell_n_ref @@
//        -old_cell_1_data
//        -...
//        -old_cell_n_data
//        +new_cell_1_data
//        +...
//        +new_cell_n_data
//        @@@ ...
//        ...
//        --- file_1_name!sheet_2_name
//        +++ file_2_name!sheet_2_name
//        ...
// Each cell data line is always present, even if the data is empty (thus just "-" or "+).
public class UnifiedDiffSpreadSheetDiffCallback extends SpreadSheetDiffCallbackBase {

  private final String lineSeparator = System.getProperty("line.separator");
  private String file1;
  private String file2;
  private DiffCell prevDiffCell = new DiffCell("", -2, -2, null, null);
  private List<DiffCell> currentCellBlock = new ArrayList<DiffCell>();

  @Override
  public void init(String file1, String file2) {
    super.init(file1, file2);
    this.file1 = file1;
    this.file2 = file2;
  }

  @Override
  public void finish() {
    super.finish();
    printAndEmptyCellBlock();
  }

  @Override
  public void reportMacroOnlyIn(boolean inFirstSpreadSheet) {
    super.reportMacroOnlyIn(inFirstSpreadSheet);
    System.out.println("Unified diff format does not support macros, however WB" + (inFirstSpreadSheet ? "1" : "2") + " contains at least one macro that is not in the other workbook.");
  }

  @Override
  public void reportExtraCell(boolean inFirstSpreadSheet, CellPos c) {
    super.reportExtraCell(inFirstSpreadSheet, c);
    accumulateAndMaybePrint(new DiffCell(
      c.getSheetName(),
      c.getRowIndex(),
      c.getColumnIndex(),
      (inFirstSpreadSheet ? c.getCellValue() : null),
      (!inFirstSpreadSheet ? c.getCellValue() : null)
    ));
  }

  @Override
  public void reportDiffCell(CellPos c1, CellPos c2) {
    super.reportDiffCell(c1, c2);
    accumulateAndMaybePrint(new DiffCell(
      c1.getSheetName(),
      c1.getRowIndex(),
      c1.getColumnIndex(),
      c1.getCellValue(),
      c2.getCellValue()
    ));
  }

  private void accumulateAndMaybePrint(DiffCell diffCell) {
    if (!isSameSheet(prevDiffCell, diffCell)) {
      printAndEmptyCellBlock();
      System.out.println("--- " + file1 + "!" + diffCell.sheetName);
      System.out.println("+++ " + file2 + "!" + diffCell.sheetName);
    }
    if (!isSameRow(prevDiffCell, diffCell)) {
      printAndEmptyCellBlock();
    }
    if (!isSameCellBlock(prevDiffCell, diffCell)) {
      printAndEmptyCellBlock();
    }
    currentCellBlock.add(diffCell);
    prevDiffCell = diffCell;
  }

  private boolean isSameSheet(DiffCell c1, DiffCell c2) {
    return c1.sheetName.equals(c2.sheetName);
  }

  private boolean isSameRow(DiffCell c1, DiffCell c2) {
    return c1.rowIndex == c2.rowIndex;
  }

  private boolean isSameCellBlock(DiffCell c1, DiffCell c2) {
    return c1.colIndex == (c2.colIndex - 1);
  }

  // TODO: Make this handle multiple rows, maybe.  What would that output look like?  This might be a bad idea.
  private void printAndEmptyCellBlock() {
    if (currentCellBlock.size() > 0) {
      StringBuilder sheet1Lines = new StringBuilder();
      StringBuilder sheet2Lines = new StringBuilder();
      String cellRange = CELL_INTERNAL_TO_USER(currentCellBlock.get(0).rowIndex, currentCellBlock.get(0).colIndex);
      if (currentCellBlock.size() > 1) {
          cellRange = cellRange + "," +
            CELL_INTERNAL_TO_USER(currentCellBlock.get(0).rowIndex, currentCellBlock.get(currentCellBlock.size()-1).colIndex);
      }
      System.out.println("@@ -" + cellRange + " +" + cellRange + " @@");
      int prevCol = -1;
      for (DiffCell rowCell : currentCellBlock) {
        assert currentCellBlock.get(0).rowIndex == rowCell.rowIndex : "printAndEmptyCellBlock() only supports one row at a time.";
        assert prevCol == -1 || prevCol == (rowCell.colIndex -1 ) : "printAndEmptyCellBlock() only supports a contiguous range of cells at a time.";
        sheet1Lines.append("-");
        sheet2Lines.append("+");
        if (rowCell.c1Value != null) {
          sheet1Lines.append(rowCell.c1Value);
        }
        if (rowCell.c2Value != null) {
          sheet2Lines.append(rowCell.c2Value);
        }
        sheet1Lines.append(lineSeparator);
        sheet2Lines.append(lineSeparator);
      }
      System.out.print(sheet1Lines.toString());
      System.out.print(sheet2Lines.toString());
    }
    currentCellBlock = new ArrayList<DiffCell>();
  }

  private class DiffCell {
    String sheetName;
    int rowIndex;
    int colIndex;
    CellValue c1Value;
    CellValue c2Value;

    public DiffCell(String _sheetName, int _rowIndex, int _colIndex, CellValue _c1Value, CellValue _c2Value) {
      sheetName = _sheetName;
      rowIndex = _rowIndex;
      colIndex = _colIndex;
      c1Value = _c1Value;
      c2Value = _c2Value;
    }
  }
}
