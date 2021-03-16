package com.ka.spreadsheet.diff;

import java.io.File;
import java.util.LinkedHashSet;
import java.util.Set;

public class StdoutSpreadSheetDiffCallback implements SpreadSheetDiffCallback {

  private final Set<Object> sheets = new LinkedHashSet<Object>();
  private final Set<Object> rows = new LinkedHashSet<Object>();
  private final Set<Object> cols = new LinkedHashSet<Object>();
  private final Set<Object> macros = new LinkedHashSet<Object>();

  private final Set<Object> sheets1 = new LinkedHashSet<Object>();
  private final Set<Object> rows1 = new LinkedHashSet<Object>();
  private final Set<Object> cols1 = new LinkedHashSet<Object>();

  private final Set<Object> sheets2 = new LinkedHashSet<Object>();
  private final Set<Object> rows2 = new LinkedHashSet<Object>();
  private final Set<Object> cols2 = new LinkedHashSet<Object>();

  private final Set<Object> macros1 = new LinkedHashSet<Object>();
  private final Set<Object> macros2 = new LinkedHashSet<Object>();

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
    reportSummary("DIFF", sheets, rows, cols, macros);
    reportSummary("EXTRA WB1", sheets1, rows1, cols1, macros1);
    reportSummary("EXTRA WB2", sheets2, rows2, cols2, macros2);
    System.out.println("-----------------------------------------");
    System.out.println("Excel files " + file1 + " and " + file2 + " "
        + (differ ? "differ" : "match"));
  }

  @Override
  public void reportMacroOnlyIn(boolean inFirstSpreadSheet) {
    String name = "unknown";
    (inFirstSpreadSheet ? macros1 : macros2).add(name);
    System.out.println("EXTRA macro name: " + name + " found only in " + wb(inFirstSpreadSheet));
  }

  @Override
  public void reportExtraCell(boolean inFirstSpreadSheet, CellPos c) {
    assert previousCell == null || c.compareCellPositions(previousCell) >= 0 :
      "Cell-ordering contract violated.  Previous=" + previousCell.getCellPosition()
      + ", current=" + c.getCellPosition();
    previousCell = c;
    if (inFirstSpreadSheet) {
      sheets1.add(c.getSheetName());
      rows1.add(c.getRow());
      cols1.add(c.getColumn());
    } else {
      sheets2.add(c.getSheetName());
      rows2.add(c.getRow());
      cols2.add(c.getColumn());
    }
    System.out.println("EXTRA Cell in " + wb(inFirstSpreadSheet) + " " + c.getCellPosition()
        + " => '" + c.getCellValue() + "'");
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
    sheets.add(c1.getSheetName());
    rows.add(c1.getRow());
    cols.add(c1.getColumn());
    System.out.println("DIFF  Cell at     " + c1.getCellPosition() + " => '" + c1.getCellValue()
        + "' v/s '" + c2.getCellValue() + "'");
  }

  private void reportSummary(String what, Set<Object> sheets, Set<Object> rows, Set<Object> cols,
      Set<Object> macros) {
    System.out.println("----------------- " + what + " -------------------");
    System.out.println("Sheets: " + sheets);
    System.out.println("Rows: " + rows);
    System.out.println("Cols: " + cols);
    if (!macros.isEmpty()) {
      System.out.println("Macros: " + macros);
    }
  }

  private String wb(boolean inFirstSpreadSheet) {
    return inFirstSpreadSheet ? "WB1" : "WB2";
  }
}
