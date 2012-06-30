package com.ka.util;
import java.io.File;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;


public class ExcelCompare {

    static String usage(){
        return    "Usage> excel_cmp <file1> <file2> [--ignore1 <sheet-ignore-spec> <sheet-ignore-spec> ..] [--ignore2 <sheet-ignore-spec> <sheet-ignore-spec> ..]" + "\n"
                + "\n"
                + "Notes: * Prints all diffs & extra cells on stdout" + "\n"
                + "       * Process exits with 0 if workbooks match, 1 otherwise" + "\n"
                + "       * Works with both xls, xlsx. You may compare an xls with xlsx too" + "\n"
                + "       * Compares only cell \"contents\". Formatting, macros are not diffed" + "\n"
                + "       * Using --ignore1 & --ignore2 (optional) you may tell the diff to ignore cells" + "\n"
                + "       * Give one and only one <sheet-ignore-spec> for a sheet" + "\n"
                + "\n"
                + "Sheet Ignore Spec:  <sheet-name>:<row-ignore-spec>:<column-ignore-spec>:<cell-ignore-spec>" + "\n"
                + "                    * Everything except <sheet-name> is optional" + "\n"
                + "                    * To ignore whole sheet, just provide <sheet-name>" + "\n"
                + "                    * Any cell satisfying any ignore spec in the sheet (row, col, or cell) will be ignored in diff" + "\n"
                + "                    * You may provide only <cell-ignore-spec> as - <sheet-name>:::<cell-ignore-spec>" + "\n"
                + "\n"
                + "Row Ignore Spec:    <comma sep list of row or row-range>" + "\n"
                + "                    * Row numbers begin from 1" + "\n"
                + "                    * Range of rows may be provide as: 1-10" + "\n"
                + "                    * Rows and ranges may be mixed as: 1-10,12,20-30 etc." + "\n"
                + "\n"
                + "Column Ignore Spec: <comma sep list of column or column-range>" + "\n"
                + "                    * Similar to Row Ignore Spec" + "\n"
                + "                    * Columns are letters starting with A" + "\n"
                + "\n"
                + "Cell Ignore Spec:   <comma sep list of cell or cell-range>" + "\n"
                + "                    * Similar to Row Ignore Spec" + "\n"
                + "                    * Cells are in usual Excel notation A1 D10" + "\n"
                + "                    * Range may be provided as A1-D10" + "\n"
                + "\n"
                + "Example command line: " + "\n"
                + "excel_cmp 1.xlsx 2.xlsx --ignore1 Sheet1:::A1,B1,J10,K11,D4 Sheet2:::A1 --ignore2 Sheet1:::A1,D4,J10 Sheet3:::A1" + "\n"
                ;
    }
    
    /*
     * TODO: Provide API (callbacks)
     * TODO: Add tests
     * TODO: Better display of results
     */
    
    public static void main(String[] args) throws Exception{
        if ((args.length < 2)){
            System.out.println(usage());
            return;
        }
        File file1 = new File(args[0]);
        File file2 = new File(args[1]);
        
        Workbook wb1 = WorkbookFactory.create(file1);
        Workbook wb2 = WorkbookFactory.create(file2);
        
        Map<String,SheetIgnores> sheetIgnores1 = parseSheetIgnores(args, "--ignore1");
        Map<String,SheetIgnores> sheetIgnores2 = parseSheetIgnores(args, "--ignore2");

        WorkbookIterator wi1 = new WorkbookIterator(wb1, sheetIgnores1);
        WorkbookIterator wi2 = new WorkbookIterator(wb2, sheetIgnores2);
        
        boolean isDiff = false;
        CellPos c1 = null, c2 = null;
        while (true){
            if ((c1==null) && wi1.hasNext()) c1 = wi1.next();
            if ((c2==null) && wi2.hasNext()) c2 = wi2.next();
            
            if ((c1!=null) && (c2!=null)){
                int c = c1.compareTo(c2);
                if (c == 0){
                    if (!c1.value().equals(c2.value())){
                        isDiff = true;
                        reportDiff(c1, c2);
                    }
                    c1 = c2 = null;
                }
                else if (c < 0){
                    isDiff = true;
                    reportExtraCell("WB1", c1);
                    c1 = null;
                }
                else {
                    isDiff = true;
                    reportExtraCell("WB2", c2);
                    c2 = null;
                }
            } else {
                break;
            }
        }
        if ((c1!=null) && (c2==null)){
            do {
                isDiff = true;
                reportExtraCell("WB1", c1);
                c1 = wi1.hasNext() ? wi1.next() : null;
            } while (c1 != null);
        }
        else if ((c1==null) && (c2!=null)){
            do {
                isDiff = true;
                reportExtraCell("WB2", c2);
                c2 = wi2.hasNext() ? wi2.next() : null;
            } while (c2 != null);
        }
        if ((c1!=null) || (c2!=null)){
            throw new IllegalStateException("Something wrong");
        }
        
        if (isDiff){
            System.out.println("Excel files " + file1 +" and " + file2 +" differ");
            System.exit(1);
        } else {
            System.out.println("Excel files " + file1 +" and " + file2 +" match");
            System.exit(0);
        }
    }
    
    static void reportDiff(CellPos c1, CellPos c2){
        System.out.println("DIFF  Cell at     " + c1.cellPos()+" => '"+ c1.value() +"' v/s '" + c2.value() + "'");
    }
    
    static void reportExtraCell(String wb, CellPos c){
        System.out.println("EXTRA Cell in " + wb + " " + c.cellPos() +" => '" + c.value() + "'");
    }
    
    static Map<String,SheetIgnores> parseSheetIgnores(String[] args, String opt){
        int start = -1, end = -1;
        for (int i=0; i<args.length; i++){
            if (start == -1){
                if (opt.equals(args[i])){
                    start = i+1;
                }
            }
            else {
                if (args[i].startsWith("--")){
                    end = i;
                }
            }
        }
        if (end == -1) end = args.length;
        
        Map<String,SheetIgnores> ret = new HashMap<String,SheetIgnores>();
        if (start != -1){
            for (int i=start; i<end; i++){
                SheetIgnores s = new SheetIgnores().parse(args[i]);
                ret.put(s.sheetName, s);
            }
        }
        return ret;
    }
    
    
    static String convertToLetter(int col){
        if (col >= 26) {
            int mod26 = (col % 26);
            return convertToLetter(((col - mod26) / 26) - 1) + convertToLetter(mod26);
        } else {
            return String.valueOf((char)(col + 65));
        }
    }
    
    static int convertFromLetter(String col){
        int idx=0;
        for (int i=col.length()-1, exp=0; i>=0; i--,exp++){
            idx += Math.pow(26, exp) * (col.charAt(i)-'A'+1);
        }
        return idx - 1;
    }

    // POI_TO_USER
    static String COL_POI_TO_USER(int col){
        return convertToLetter(col);
    }
    static int ROW_POI_TO_USER(int row){
        return row+1;
    }
    static String CELL_POI_TO_USER(int row, int col){
        return COL_POI_TO_USER(col) + ROW_POI_TO_USER(row);
    }
    
    
    // USER_TO_POI
    static int COL_USER_TO_POI(String col){
        return convertFromLetter(col);
    }
    static int ROW_USER_TO_POI(int row){
        return row-1;
    }
    static final Pattern cellPat = Pattern.compile("([A-Z]+)(\\d+)");
    static int[] CELL_USER_TO_POI(String cell){
        Matcher matcher = cellPat.matcher(cell);
        if (!matcher.matches())
            throw new IllegalArgumentException("Illegal cell specifier " + cell);
        return new int[]{ROW_USER_TO_POI(Integer.parseInt(matcher.group(2))),
                         COL_USER_TO_POI(matcher.group(1))};
    }
}

class CellPos implements Comparable<CellPos>{
    Workbook wb;
    int sheetIdx;
    Cell cell;
    CellPos(Workbook wb, int sheetIdx, Cell cell){
        this.wb = wb;
        this.sheetIdx = sheetIdx;
        this.cell = cell;
    }
    String cellPos(){
        return wb.getSheetName(sheetIdx)+"!"
                +ExcelCompare.CELL_POI_TO_USER(cell.getRowIndex(), cell.getColumnIndex());
    }
    public String toString(){
        return cellPos() +" => " + cell.toString();
    }
    Object value(){
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

class WorkbookIterator{
    Workbook wb;
    Map<String,SheetIgnores> sheetIgnores;
    int currSheetIdx;
    SheetIgnores currSheetIgnores;
    Iterator<Row> rows;
    Iterator<Cell> cells;
    
    WorkbookIterator(Workbook wb, Map<String,SheetIgnores> sheetIgnores){
        this.wb = wb;
        this.sheetIgnores = sheetIgnores;
        this.currSheetIdx = -1;
        _seenNext = false;
    }
    
    boolean _seenNext;
    Cell _nextCell;
    
    boolean hasNext(){
        if (!_seenNext){
            _nextCell = null;
            _seenNext = true;
            while (_nextCell == null){
                if ((cells != null) && (cells.hasNext())){
                    Cell cell = cells.next();
                    if (!ignoreCol(cell) && !ignoreCell(cell)){
                        _nextCell = cell;
                    }
                }
                else if ((rows != null) && (rows.hasNext())){
                    Row row = rows.next();
                    if (!ignoreRow(row)){
                        cells = row.cellIterator();
                    }
                }
                else if (currSheetIdx < (wb.getNumberOfSheets()-1)){
                    currSheetIdx++;
                    Sheet sheet = wb.getSheetAt(currSheetIdx);
                    currSheetIgnores = sheetIgnores.get(wb.getSheetName(currSheetIdx));
                    if (!ignoreSheet()){
                        rows = sheet.rowIterator();
                    }
                }
                else{
                    break;
                }
            }
        }
        return _nextCell != null;
    }

    boolean ignoreSheet(){
        return (currSheetIgnores!=null) && currSheetIgnores.completeIgnore;
    }
    
    boolean ignoreRow(Row row){
        return (currSheetIgnores!=null) && (currSheetIgnores.isRowIgnored(row.getRowNum()));
    }
    
    boolean ignoreCol(Cell cell){
        return (currSheetIgnores!=null) && (currSheetIgnores.isColIgnored(cell.getColumnIndex()));
    }
    
    boolean ignoreCell(Cell cell){
        return (currSheetIgnores!=null) && (currSheetIgnores.isCellIgnored(cell.getRowIndex(), cell.getColumnIndex()));
    }
    
    CellPos next(){
        _seenNext = false;
        return new CellPos(wb, currSheetIdx, _nextCell);
    }
}

class SheetIgnores{
    boolean completeIgnore;
    boolean rowIgnoresPresent;
    boolean colIgnoresPresent;
    boolean cellIgnoresPresent;
    String sheetName;
    List<Range> rowIgnores;
    List<Range> colIgnores;
    List<Range> cellIgnores;
    
    boolean isRowIgnored(int row){
        return rowIgnoresPresent && isIgnored(rowIgnores, row);
    }
    
    boolean isColIgnored(int col){
        return colIgnoresPresent && isIgnored(colIgnores, col);
    }
    
    boolean isCellIgnored(int row, int col){
        return cellIgnoresPresent && isIgnored(cellIgnores, row, col);
    }
    
    boolean isIgnored(List<Range> rngs, int ... pt){
        for (Range r : rngs){
            if (r.lies_between(pt))
                return true;
        }
        return false;
    }
    
    // Assume val is not null & non-empty
    SheetIgnores parse(String val){
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
    
    static List<Range> formRowIgnores(String val){
        List<Range> ret = new ArrayList<Range>();
        if (val != null){
            for (String rng : val.split(",")){
                String[] rngs = rng.split("-");
                if(rngs.length == 1){ // Single row
                    int row = ExcelCompare.ROW_USER_TO_POI(Integer.parseInt(rngs[0]));
                    ret.add(new Range(row, row));
                } else if (rngs.length == 2){
                    int row1 = ExcelCompare.ROW_USER_TO_POI(Integer.parseInt(rngs[0]));
                    int row2 = ExcelCompare.ROW_USER_TO_POI(Integer.parseInt(rngs[1]));
                    ret.add(new Range(row1, row2));
                } else {
                    throw new IllegalArgumentException("Illegal row ignore specifier " + val);
                }
            }
        }
        return ret;
    }
    
    static List<Range> formColIgnores(String val){
        List<Range> ret = new ArrayList<Range>();
        if (val != null){
            for (String rng : val.split(",")){
                String[] rngs = rng.split("-");
                if(rngs.length == 1){ // Single col
                    int col = ExcelCompare.COL_USER_TO_POI(rngs[0]);
                    ret.add(new Range(col, col));
                } else if (rngs.length == 2){
                    int col1 = ExcelCompare.COL_USER_TO_POI(rngs[0]);
                    int col2 = ExcelCompare.COL_USER_TO_POI(rngs[1]);
                    ret.add(new Range(col1, col2));
                } else {
                    throw new IllegalArgumentException("Illegal col ignore specifier " + val);
                }
            }
        }
        return ret;
    }
    
    static List<Range> formCellIgnores(String val){
        List<Range> ret = new ArrayList<Range>();
        if (val != null){
            for (String rng : val.split(",")){
                String[] rngs = rng.split("-");
                if(rngs.length == 1){ // Single cell
                    int[] rowcol = ExcelCompare.CELL_USER_TO_POI(rngs[0]);
                    ret.add(new Range(rowcol[0], rowcol[1], rowcol[0], rowcol[1]));
                } else if (rngs.length == 2){
                    int[] rowcol1 = ExcelCompare.CELL_USER_TO_POI(rngs[0]);
                    int[] rowcol2 = ExcelCompare.CELL_USER_TO_POI(rngs[1]);
                    ret.add(new Range(rowcol1[0], rowcol1[1], rowcol2[0], rowcol2[1]));
                } else {
                    throw new IllegalArgumentException("Illegal cell ignore specifier " + val);
                }
            }
        }
        return ret;
    }
}

class Range{
    final int[] low, high;
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
