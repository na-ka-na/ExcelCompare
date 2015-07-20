package com.ka.spreadsheet.diff;

import java.util.Iterator;

import javax.annotation.Nullable;

import org.odftoolkit.simple.SpreadsheetDocument;
import org.odftoolkit.simple.table.Cell;
import org.odftoolkit.simple.table.Row;
import org.odftoolkit.simple.table.Table;

public class SpreadSheetOdf implements ISpreadSheet {

  private final SpreadsheetDocument spreadsheetDocument;

  public SpreadSheetOdf(SpreadsheetDocument spreadsheetDocument) {
    this.spreadsheetDocument = spreadsheetDocument;
  }

  @Override
  public Iterator<ISheet> getSheetIterator() {
    return new Iterator<ISheet>() {

      private int currSheetIdx = 0;

      @Override
      public boolean hasNext() {
        return currSheetIdx < spreadsheetDocument.getSheetCount();
      }

      @Override
      public ISheet next() {
        Table table = spreadsheetDocument.getSheetByIndex(currSheetIdx);
        SheetOdf sheetOdf = new SheetOdf(table, currSheetIdx);
        currSheetIdx++;
        return sheetOdf;
      }

      @Override
      public void remove() {
        throw new UnsupportedOperationException();
      }
    };
  }

  @Override
  @Nullable
  public Boolean hasMacro() {
    return null;
  }
}


class SheetOdf implements ISheet {

  private final Table table;
  private final int sheetIdx;

  public SheetOdf(Table table, int sheetIdx) {
    this.table = table;
    this.sheetIdx = sheetIdx;
  }

  @Override
  public String getName() {
    return table.getTableName();
  }

  @Override
  public int getSheetIndex() {
    return sheetIdx;
  }

  @Override
  public Iterator<IRow> getRowIterator() {
    final Iterator<Row> rowIterator = table.getRowIterator();
    return new Iterator<IRow>() {

      @Override
      public boolean hasNext() {
        return rowIterator.hasNext();
      }

      @Override
      public IRow next() {
        return new RowOdf(rowIterator.next());
      }

      @Override
      public void remove() {
        throw new UnsupportedOperationException();
      }
    };
  }
}


class RowOdf implements IRow {

  private final Row row;

  public RowOdf(Row row) {
    this.row = row;
  }

  @Override
  public int getRowIndex() {
    return row.getRowIndex();
  }

  @Override
  public Iterator<ICell> getCellIterator() {
    final int numCells = row.getCellCount();
    return new Iterator<ICell>() {
      private int currCellIdx = 0;

      @Override
      public boolean hasNext() {
        return currCellIdx < numCells;
      }

      @Override
      public ICell next() {
        return new CellOdf(row.getCellByIndex(currCellIdx++));
      }

      @Override
      public void remove() {
        throw new UnsupportedOperationException();
      }
    };
  }
}


class CellOdf implements ICell {

  private final Cell cell;

  public CellOdf(Cell cell) {
    this.cell = cell;
  }

  @Override
  public int getRowIndex() {
    return cell.getRowIndex();
  }

  @Override
  public int getColumnIndex() {
    return cell.getColumnIndex();
  }

  @Override
  public String getStringValue() {
    String formula = cell.getFormula();
    if (formula != null) {
      return formula;
    }
    String valueType = cell.getValueType();
    if (valueType != null) {
      if (valueType.equals("float")) {
        return String.valueOf(cell.getDoubleValue());
      } else if (valueType.equals("boolean")) {
        return String.valueOf(cell.getBooleanValue());
      }
    }
    return cell.getStringValue();
  }
}
