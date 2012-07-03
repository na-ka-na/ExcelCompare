package com.ka.excelcmp;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Workbook;

import static com.ka.excelcmp.ExcelUtils.CELL_POI_TO_USER;
import static com.ka.excelcmp.ExcelUtils.COL_POI_TO_USER;
import static com.ka.excelcmp.ExcelUtils.ROW_POI_TO_USER;


public class CellPos implements Comparable<CellPos>{
    private final Workbook wb;
    private final int sheetIdx;
    private final Cell cell;
    
    public String sheet(){
        return wb.getSheetName(sheetIdx);
    }
    public int row(){
        return ROW_POI_TO_USER(cell.getRowIndex());
    }
    public String col(){
        return COL_POI_TO_USER(cell.getColumnIndex());
    }
    
    protected CellPos(Workbook wb, int sheetIdx, Cell cell){
        this.wb = wb;
        this.sheetIdx = sheetIdx;
        this.cell = cell;
    }
    
    public String cellPos(){
        return wb.getSheetName(sheetIdx)+"!"
                +CELL_POI_TO_USER(cell.getRowIndex(), cell.getColumnIndex());
    }
    
    public String toString(){
        return cellPos() +" => " + cell.toString();
    }
    
    public Object value(){
        return cell.toString();
    }
    
    @Override
    public int compareTo(CellPos o) {
        int c = sheetIdx - o.sheetIdx;
        if (c == 0){
            c = cell.getRowIndex() - o.cell.getRowIndex();
            if (c == 0){
                c = cell.getColumnIndex() - o.cell.getColumnIndex();
            }
        }
        return c;
    }
}