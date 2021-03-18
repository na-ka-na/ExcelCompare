package com.ka.spreadsheet.diff;

import java.io.File;

public abstract class SpreadSheetDiffCallbackBase implements SpreadSheetDiffCallback {

  private CellPos previousCell = null;

  @Override
  public void init(String file1, String file2) {
  }

  @Override
  public void finish() {
  }

  @Override
  public void reportDiffCell(CellPos c1, CellPos c2) {
    assert (c1.getRowIndex() == c2.getRowIndex())
      && (c1.getColumnIndex() == c2.getColumnIndex()) : "Cells are not at the same position. Cell 1="
      + c1.getCellPosition() + ", cell 2=" + c2.getCellPosition();
    assert previousCell == null || c1.compareCellPositions(previousCell) > 0 :
      "Cell-ordering contract violated.  Previous=" + previousCell.getCellPosition()
      + ", current=" + c1.getCellPosition();
    previousCell = c1;
  }

  @Override
  public void reportExtraCell(boolean inFirstSpreadSheet, CellPos c) {
    assert previousCell == null || c.compareCellPositions(previousCell) > 0 :
      "Cell-ordering contract violated.  Previous=" + previousCell.getCellPosition()
      + ", current=" + c.getCellPosition();
    previousCell = c;
  }

  @Override
  public void reportMacroOnlyIn(boolean inFirstSpreadSheet){
  }

  @Override
  public void reportWorkbooksDiffer(boolean differ){
  }
}
