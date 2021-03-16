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
public class UnifiedDiffSpreadSheetDiffCallback implements SpreadSheetDiffCallback {

  private final String lineSeparator = System.getProperty("line.separator");
  private List<DiffCell> diffCells = new ArrayList<DiffCell>();

  private String file1;
  private String file2;

  private CellPos previousCell = null;

  @Override
  public void init(String file1, String file2) {
    this.file1 = file1;
    this.file2 = file2;
  }

  @Override
  public void finish() {
  }

  @Override
  public void reportWorkbooksDiffer(boolean differ) {
    String sheetName = "";
    int row = -1;
    int col = -1;
    List<DiffCell> cellBlock = null;
    for (DiffCell diffCell : diffCells) {
     if (!diffCell.sheetName.equals(sheetName)) {
        // New sheet, finish the active row (if any) and start the new sheet.
        if (cellBlock != null) {
          printRow(cellBlock);
        }
        sheetName = diffCell.sheetName;
        System.out.println("--- " + file1 + "!" + sheetName);
        System.out.println("+++ " + file2 + "!" + sheetName);
        cellBlock = null;
        row = -1;
      }
      if (diffCell.rowIndex != row) {
        // New row, finish the active row (if any) and start the new row.
        if (cellBlock != null) {
          printRow(cellBlock);
        }
        cellBlock = null;
        row = diffCell.rowIndex;
        col = -1;
      }
      if (diffCell.colIndex != col) {
        // New cell-block, finish the active cell-block (if any) and start a new cell-block.
        if (cellBlock != null) {
          printRow(cellBlock);
        }
        cellBlock = new ArrayList<DiffCell>();
      }
      // Add this cell to the active cell-block.
      cellBlock.add(diffCell);
      col = diffCell.colIndex + 1;
    }
    if (cellBlock != null) {
      printRow(cellBlock);
    }
  }

  @Override
  public void reportMacroOnlyIn(boolean inFirstSpreadSheet) {
    System.out.println("Unified diff format does not support macros, however WB" + (inFirstSpreadSheet ? "1" : "2") + " contains at least one macro that is not in the other workbook.");
  }

  @Override
  public void reportExtraCell(boolean inFirstSpreadSheet, CellPos c) {
    assert previousCell == null || c.compareCellPositions(previousCell) >= 0 :
      "Cell-ordering contract violated.  Previous=" + previousCell.getCellPosition()
      + ", current=" + c.getCellPosition();
    previousCell = c;
    diffCells.add(new DiffCell(
      c.getSheetName(),
      c.getRowIndex(),
      c.getColumnIndex(),
      (inFirstSpreadSheet ? c.getCellValue() : null),
      (!inFirstSpreadSheet ? c.getCellValue() : null)
    ));
  }

  @Override
  public void reportDiffCell(CellPos c1, CellPos c2) {
    assert (c1.getRowIndex() == c2.getRowIndex())
      && (c1.getColumnIndex() == c2.getColumnIndex()) : "Cells are not at the same position. Cell 1="
      + c1.getCellPosition() + ", cell 2=" + c2.getCellPosition();
    assert previousCell == null || c1.compareCellPositions(previousCell) >= 0 :
      "Cell-ordering contract violated.  Previous=" + previousCell.getCellPosition()
      + ", current=" + c1.getCellPosition();
    previousCell = c1;
    diffCells.add(new DiffCell(
      c1.getSheetName(),
      c1.getRowIndex(),
      c1.getColumnIndex(),
      c1.getCellValue(),
      c2.getCellValue()
    ));
  }

  // TODO: Make this handle multiple rows, maybe.  What would that output look like?  This might be a bad idea.
  private void printRow(List<DiffCell> rowCells) {
    StringBuilder sheet1Lines = new StringBuilder();
    StringBuilder sheet2Lines = new StringBuilder();
    String cellRange = CELL_INTERNAL_TO_USER(rowCells.get(0).rowIndex, rowCells.get(0).colIndex);
    if (rowCells.size() > 1) {
        cellRange = cellRange + "," +
          CELL_INTERNAL_TO_USER(rowCells.get(0).rowIndex, rowCells.get(rowCells.size()-1).colIndex);
    }
    System.out.println("@@ -" + cellRange + " +" + cellRange + " @@");
    int prevCol = -1;
    for (DiffCell rowCell : rowCells) {
        if (rowCells.get(0).rowIndex != rowCell.rowIndex) {
            throw new RuntimeException("printRow() only supports one row at a time.");
        }
        if (prevCol != -1 && prevCol != (rowCell.colIndex -1 )) {
            throw new RuntimeException("printRow() only supports a contiguous range of cells at a time.");
        }
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

  private class DiffCell implements Comparable<DiffCell> {
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
    
    public int compareTo(DiffCell other) {
        int sheetComparison = sheetName.compareTo(other.sheetName);
        return sheetComparison != 0 ? sheetComparison
                : rowIndex != other.rowIndex ? rowIndex - other.rowIndex
                : colIndex - other.colIndex;
    }
  }
}
