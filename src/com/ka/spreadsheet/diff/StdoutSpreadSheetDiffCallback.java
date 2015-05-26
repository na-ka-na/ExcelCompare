package com.ka.spreadsheet.diff;

import java.io.File;
import java.util.LinkedHashSet;
import java.util.Set;

public class StdoutSpreadSheetDiffCallback implements SpreadSheetDiffCallback {

	private final Set<Object> sheets = new LinkedHashSet<Object>();
	private final Set<Object> rows = new LinkedHashSet<Object>();
	private final Set<Object> cols = new LinkedHashSet<Object>();

	private final Set<Object> sheets1 = new LinkedHashSet<Object>();
	private final Set<Object> rows1 = new LinkedHashSet<Object>();
	private final Set<Object> cols1 = new LinkedHashSet<Object>();

	private final Set<Object> sheets2 = new LinkedHashSet<Object>();
	private final Set<Object> rows2 = new LinkedHashSet<Object>();
	private final Set<Object> cols2 = new LinkedHashSet<Object>();

    @Override
    public void reportWorkbooksDiffer(boolean differ, File file1, File file2) {
    	reportSummary("DIFF", sheets, rows, cols);
        reportSummary("EXTRA WB1", sheets1, rows1, cols1);
        reportSummary("EXTRA WB2", sheets2, rows2, cols2);
        System.out.println("-----------------------------------------");
        System.out.println("Excel files " + file1 + " and " + file2 + " " + (differ ? "differ" : "match"));
    }

    @Override
    public void reportExtraCell(boolean inFirstSpreadSheet, CellPos c) {
        if (inFirstSpreadSheet){
            sheets1.add(c.getSheetName());
            rows1.add(c.getRow());
            cols1.add(c.getColumn());
        } else {
            sheets2.add(c.getSheetName());
            rows2.add(c.getRow());
            cols2.add(c.getColumn());
        }
        String wb = inFirstSpreadSheet ? "WB1" : "WB2";
        System.out.println("EXTRA Cell in " + wb + " " + c.getCellPosition() +" => '" + c.getStringValue() + "'");
    }

    @Override
    public void reportDiffCell(CellPos c1, CellPos c2) {
        sheets.add(c1.getSheetName());
        rows.add(c1.getRow());
        cols.add(c1.getColumn());
        System.out.println("DIFF  Cell at     " + c1.getCellPosition()+" => '"+ c1.getStringValue() +"' v/s '" + c2.getStringValue() + "'");
    }

    private void reportSummary(String what, Set<Object> sheets, Set<Object> rows, Set<Object> cols) {
        System.out.println("----------------- "+what+" -------------------");
        System.out.println("Sheets: " + sheets);
        System.out.println("Rows: " + rows);
        System.out.println("Cols: " + cols);
    }
}
