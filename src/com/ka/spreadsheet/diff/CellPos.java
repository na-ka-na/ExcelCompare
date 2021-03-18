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

  public int getRowIndex() {
    return cell.getRowIndex();
  }

  public String getColumn() {
    return COL_INTERNAL_TO_USER(cell.getColumnIndex());
  }

  public int getColumnIndex() {
    return cell.getColumnIndex();
  }

  public String getCellPosition() {
    return getSheetName() + "!" + CELL_INTERNAL_TO_USER(cell.getRowIndex(), cell.getColumnIndex());
  }

  @Override
  public String toString() {
    return getCellPosition() + " => " + cell.toString();
  }

  public CellValue getCellValue() {
    try {
      return cell.getValue();
    } catch (Exception e) {
      throw new RuntimeException("Error reading Cell at " + getCellPosition() + ": " + e.getMessage(), e);
    }
  }

  /**
   * Compare positions in {Sheet, Row, Column} order.
   * Returns -/0/+ as in Comparable.compareTo().
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
