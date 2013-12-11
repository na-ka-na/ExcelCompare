package com.ka.excelcmp;
import java.io.File;
import java.util.HashMap;
import java.util.LinkedHashSet;
import java.util.Map;
import java.util.Set;

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
        final File file1 = new File(args[0]);
        final File file2 = new File(args[1]);
        
        Workbook wb1 = WorkbookFactory.create(file1);
        Workbook wb2 = WorkbookFactory.create(file2);
        SpreadSheetExcel sse1 = new SpreadSheetExcel(wb1);
        SpreadSheetExcel sse2 = new SpreadSheetExcel(wb2);
        
        Map<String,SheetIgnores> sheetIgnores1 = parseSheetIgnores(args, "--ignore1");
        Map<String,SheetIgnores> sheetIgnores2 = parseSheetIgnores(args, "--ignore2");

        SpreadSheetIterator wi1 = new SpreadSheetIterator(sse1, sheetIgnores1);
        SpreadSheetIterator wi2 = new SpreadSheetIterator(sse2, sheetIgnores2);
        
        DiffReportCallback call = new DiffReportCallback() {
            Set<Object> sheets = new LinkedHashSet<Object>();
            Set<Object> rows = new LinkedHashSet<Object>();
            Set<Object> cols = new LinkedHashSet<Object>();
            
            Set<Object> sheets1 = new LinkedHashSet<Object>();
            Set<Object> rows1 = new LinkedHashSet<Object>();
            Set<Object> cols1 = new LinkedHashSet<Object>();
            
            Set<Object> sheets2 = new LinkedHashSet<Object>();
            Set<Object> rows2 = new LinkedHashSet<Object>();
            Set<Object> cols2 = new LinkedHashSet<Object>();
            
            @Override
            public void reportWorkbooksDiffer(boolean differ) {
                reportSummary();
                System.out.println("Excel files " + file1 + " and " + file2 + " " + (differ ? "differ" : "match"));
            }
            
            @Override
            public void reportExtraCell(boolean firstWb, CellPos c) {
                if (firstWb){
                    sheets1.add(c.getSheetName());
                    rows1.add(c.getRow());
                    cols1.add(c.getColumn());
                } else {
                    sheets2.add(c.getSheetName());
                    rows2.add(c.getRow());
                    cols2.add(c.getColumn());
                }
                String wb = firstWb ? "WB1" : "WB2";
                System.out.println("EXTRA Cell in " + wb + " " + c.getCellPosition() +" => '" + c.getStringValue() + "'");
            }
            
            @Override
            public void reportDiffCell(CellPos c1, CellPos c2) {
                sheets.add(c1.getSheetName());
                rows.add(c1.getRow());
                cols.add(c1.getColumn());
                System.out.println("DIFF  Cell at     " + c1.getCellPosition()+" => '"+ c1.getStringValue() +"' v/s '" + c2.getStringValue() + "'");
            }
            
            private void reportSummary(){
                reportS("DIFF", sheets, rows, cols);
                reportS("EXTRA WB1", sheets1, rows1, cols1);
                reportS("EXTRA WB2", sheets2, rows2, cols2);
                System.out.println("-----------------------------------------");
            }
            
            @SuppressWarnings("hiding")
            private void reportS(String what, Set<Object> sheets, Set<Object> rows, Set<Object> cols) {
                System.out.println("----------------- "+what+" -------------------");
                System.out.println("Sheets: " + sheets);
                System.out.println("Rows: " + rows);
                System.out.println("Cols: " + cols);
            }
        };
        
        boolean differ = doDiff(wi1, wi2, call);
        
        System.exit(differ ? 1 : 0);
    }
    
    private static boolean doDiff(SpreadSheetIterator wi1, SpreadSheetIterator wi2, DiffReportCallback call){
        boolean isDiff = false;
        CellPos c1 = null, c2 = null;
        while (true){
            if ((c1==null) && wi1.hasNext()) c1 = wi1.next();
            if ((c2==null) && wi2.hasNext()) c2 = wi2.next();
            
            if ((c1!=null) && (c2!=null)){
                int c = c1.compareTo(c2);
                if (c == 0){
                    if (!c1.getStringValue().equals(c2.getStringValue())){
                        isDiff = true;
                        call.reportDiffCell(c1, c2);
                    }
                    c1 = c2 = null;
                }
                else if (c < 0){
                    isDiff = true;
                    call.reportExtraCell(true, c1);
                    c1 = null;
                }
                else {
                    isDiff = true;
                    call.reportExtraCell(false, c2);
                    c2 = null;
                }
            } else {
                break;
            }
        }
        if ((c1!=null) && (c2==null)){
            do {
                isDiff = true;
                call.reportExtraCell(true, c1);
                c1 = wi1.hasNext() ? wi1.next() : null;
            } while (c1 != null);
        }
        else if ((c1==null) && (c2!=null)){
            do {
                isDiff = true;
                call.reportExtraCell(false, c2);
                c2 = wi2.hasNext() ? wi2.next() : null;
            } while (c2 != null);
        }
        if ((c1!=null) || (c2!=null)){
            throw new IllegalStateException("Something wrong");
        }
        
        call.reportWorkbooksDiffer(isDiff);
        
        return isDiff;
    }
    
    private static Map<String,SheetIgnores> parseSheetIgnores(String[] args, String opt){
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
                SheetIgnores s = SheetIgnores.newSheetIgnore(args[i]);
                ret.put(s.sheetName(), s);
            }
        }
        return ret;
    }
}
