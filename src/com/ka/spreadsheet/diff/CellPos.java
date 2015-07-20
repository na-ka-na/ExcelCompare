package com.ka.spreadsheet.diff;

import static com.ka.spreadsheet.diff.SpreadSheetUtils.CELL_INTERNAL_TO_USER;
import static com.ka.spreadsheet.diff.SpreadSheetUtils.COL_INTERNAL_TO_USER;
import static com.ka.spreadsheet.diff.SpreadSheetUtils.ROW_INTERNAL_TO_USER;

public class CellPos {

  private final ISheet sheet;
  private final ICell cell;

  public CellPos(ISheet sheet, ICell cell) {
    this.sheet = sheet;
    this.cell = cell;
  }

  public String getSheetName() {
    return sheet.getName();
  }

  public int getRow() {
    return ROW_INTERNAL_TO_USER(cell.getRowIndex());
  }

  public String getColumn() {
    return COL_INTERNAL_TO_USER(cell.getColumnIndex());
  }

  public String getCellPosition() {
    return getSheetName() + "!" + CELL_INTERNAL_TO_USER(cell.getRowIndex(), cell.getColumnIndex());
  }

  @Override
  public String toString() {
    return getCellPosition() + " => " + cell.toString();
  }

  public Object getCellValue() {
    return cell.getValue();
  }

  /**
   * Compare positions in {Sheet, Row, Column} order.
   */
  public int compareCellPositions(CellPos o) {
    int c = sheet.getSheetIndex() - o.sheet.getSheetIndex();
    if (c == 0) {
      c = cell.getRowIndex() - o.cell.getRowIndex();
      if (c == 0) {
        c = cell.getColumnIndex() - o.cell.getColumnIndex();
      }
    }
    return c;
  }
}
