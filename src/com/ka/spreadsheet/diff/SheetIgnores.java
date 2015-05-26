package com.ka.spreadsheet.diff;

import static com.ka.spreadsheet.diff.SpreadSheetUtils.CELL_USER_TO_INTERNAL;
import static com.ka.spreadsheet.diff.SpreadSheetUtils.COL_USER_TO_INTERNAL;
import static com.ka.spreadsheet.diff.SpreadSheetUtils.ROW_USER_TO_INTERNAL;

import java.util.ArrayList;
import java.util.List;

public class SheetIgnores {

    private static class Range {

    	private final int[] low;
        private final int[] high;

        Range(int low, int high){
            this.low = new int[]{low};
            this.high = new int[]{high};
        }

        Range(int low1, int low2, int high1, int high2){
            this.low = new int[]{low1, low2};
            this.high = new int[]{high1, high2};
        }

        boolean lies_between(int ... x){
            for (int i=0; i<x.length; i++){
                if (!((low[i] <= x[i]) && (x[i] <= high[i])))
                    return false;
            }
            return true;
        }
    }

    private boolean completeIgnore;
    private boolean rowIgnoresPresent;
    private boolean colIgnoresPresent;
    private boolean cellIgnoresPresent;
    private String sheetName;
    private List<Range> rowIgnores;
    private List<Range> colIgnores;
    private List<Range> cellIgnores;

    public String sheetName(){
        return sheetName;
    }

    public boolean isWholeSheetIgnored(){
        return completeIgnore;
    }

    public boolean isRowIgnored(int row){
        return rowIgnoresPresent && isIgnored(rowIgnores, row);
    }

    public boolean isColIgnored(int col){
        return colIgnoresPresent && isIgnored(colIgnores, col);
    }

    public boolean isCellIgnored(int row, int col){
        return cellIgnoresPresent && isIgnored(cellIgnores, row, col);
    }

    public boolean isIgnored(List<Range> rngs, int ... pt){
        for (Range r : rngs){
            if (r.lies_between(pt))
                return true;
        }
        return false;
    }

    public static SheetIgnores newSheetIgnore(String val){
        return new SheetIgnores().parse(val);
    }

    // Assume val is not null & non-empty
    private SheetIgnores parse(String val){
        String[] parts = val.split(":");
        sheetName = parts[0];
        completeIgnore = parts.length == 1;
        if ((parts.length > 1) && (!parts[1].isEmpty())){
            rowIgnoresPresent = true;
            rowIgnores = formRowIgnores(parts[1]);
        }
        if ((parts.length > 2) && (!parts[2].isEmpty())){
            colIgnoresPresent = true;
            colIgnores = formColIgnores(parts[2]);
        }
        if ((parts.length > 3) && (!parts[3].isEmpty())){
            cellIgnoresPresent = true;
            cellIgnores = formCellIgnores(parts[3]);
        }
        if (parts.length > 4)
            throw new IllegalArgumentException("Illegal Sheet Ignores argument " + val);
        return this;
    }

    private static List<Range> formRowIgnores(String val){
        List<Range> ret = new ArrayList<Range>();
        if (val != null){
            for (String rng : val.split(",")){
                String[] rngs = rng.split("-");
                if(rngs.length == 1){ // Single row
                    int row = ROW_USER_TO_INTERNAL(Integer.parseInt(rngs[0]));
                    ret.add(new Range(row, row));
                } else if (rngs.length == 2){
                    int row1 = ROW_USER_TO_INTERNAL(Integer.parseInt(rngs[0]));
                    int row2 = ROW_USER_TO_INTERNAL(Integer.parseInt(rngs[1]));
                    ret.add(new Range(row1, row2));
                } else {
                    throw new IllegalArgumentException("Illegal row ignore specifier " + val);
                }
            }
        }
        return ret;
    }

    private static List<Range> formColIgnores(String val){
        List<Range> ret = new ArrayList<Range>();
        if (val != null){
            for (String rng : val.split(",")){
                String[] rngs = rng.split("-");
                if(rngs.length == 1){ // Single col
                    int col = COL_USER_TO_INTERNAL(rngs[0]);
                    ret.add(new Range(col, col));
                } else if (rngs.length == 2){
                    int col1 = COL_USER_TO_INTERNAL(rngs[0]);
                    int col2 = COL_USER_TO_INTERNAL(rngs[1]);
                    ret.add(new Range(col1, col2));
                } else {
                    throw new IllegalArgumentException("Illegal col ignore specifier " + val);
                }
            }
        }
        return ret;
    }

    private static List<Range> formCellIgnores(String val){
        List<Range> ret = new ArrayList<Range>();
        if (val != null){
            for (String rng : val.split(",")){
                String[] rngs = rng.split("-");
                if(rngs.length == 1){ // Single cell
                    int[] rowcol = CELL_USER_TO_INTERNAL(rngs[0]);
                    ret.add(new Range(rowcol[0], rowcol[1], rowcol[0], rowcol[1]));
                } else if (rngs.length == 2){
                    int[] rowcol1 = CELL_USER_TO_INTERNAL(rngs[0]);
                    int[] rowcol2 = CELL_USER_TO_INTERNAL(rngs[1]);
                    ret.add(new Range(rowcol1[0], rowcol1[1], rowcol2[0], rowcol2[1]));
                } else {
                    throw new IllegalArgumentException("Illegal cell ignore specifier " + val);
                }
            }
        }
        return ret;
    }
}
