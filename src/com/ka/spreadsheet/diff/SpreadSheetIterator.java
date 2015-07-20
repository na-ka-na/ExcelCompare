package com.ka.spreadsheet.diff;

import java.util.Iterator;

public class SpreadSheetIterator {

  private final Iterator<ISheet> sheetIterator;
  private ISheet sheet;
  private SheetIgnores currSheetIgnores;
  private Iterator<IRow> rows;
  private Iterator<ICell> cells;

  private boolean seenNext;
  private ICell nextCell;
  private final WorkbookIgnores workbookIgnores;

  SpreadSheetIterator(ISpreadSheet spreadSheet, WorkbookIgnores workbookIgnores) {
    this.workbookIgnores = workbookIgnores;
    this.sheetIterator = spreadSheet.getSheetIterator();
  }

  boolean hasNext() {
    if (!seenNext) {
      nextCell = null;
      seenNext = true;
      while (nextCell == null) {
        if ((cells != null) && (cells.hasNext())) {
          ICell cell = cells.next();
          if (!ignoreCol(cell) && !ignoreCell(cell)) {
            nextCell = cell;
          }
        } else if ((rows != null) && (rows.hasNext())) {
          IRow row = rows.next();
          if (!ignoreRow(row)) {
            cells = row.getCellIterator();
          }
        } else if (sheetIterator.hasNext()) {
          sheet = sheetIterator.next();
          currSheetIgnores = workbookIgnores.fetchSheetIgnores(sheet.getName());
          if (!ignoreSheet()) {
            rows = sheet.getRowIterator();
          }
        } else {
          break;
        }
      }
    }
    return nextCell != null;
  }

  private boolean ignoreSheet() {
    return (currSheetIgnores != null) && currSheetIgnores.isWholeSheetIgnored();
  }

  private boolean ignoreRow(IRow row) {
    return (currSheetIgnores != null) && (currSheetIgnores.isRowIgnored(row.getRowIndex()));
  }

  private boolean ignoreCol(ICell cell) {
    return (currSheetIgnores != null) && (currSheetIgnores.isColIgnored(cell.getColumnIndex()));
  }

  private boolean ignoreCell(ICell cell) {
    return (currSheetIgnores != null)
        && (currSheetIgnores.isCellIgnored(cell.getRowIndex(), cell.getColumnIndex()));
  }

  public CellPos next() {
    seenNext = false;
    return new CellPos(sheet, nextCell);
  }
}
