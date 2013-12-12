package com.ka.spreadsheet.diff;
import java.io.File;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.odftoolkit.simple.SpreadsheetDocument;


public class SpreadSheetDiffer {

    static String usage(){
        return    "Usage> excel_cmp <file1> <file2> [--ignore1 <sheet-ignore-spec> <sheet-ignore-spec> ..] [--ignore2 <sheet-ignore-spec> <sheet-ignore-spec> ..]" + "\n"
                + "\n"
                + "Notes: * Prints all diffs & extra cells on stdout" + "\n"
                + "       * Process exits with 0 if workbooks match, 1 otherwise" + "\n"
                + "       * Works with both xls, xlsx, ods. You may compare any of xls, xlsx, ods with each other" + "\n"
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
    
    public static void main(String[] args) {
    	int ret = doDiff(args);
    	System.exit(ret);
    }
    
    public static int doDiff(String[] args) {
    	int ret = -1;
    	try {
    		ret = doDiff(args, new StdoutSpreadSheetDiffCallback());
    	} catch (Exception e) {
    		//e.printStackTrace(System.err);
    		System.err.println("Diff failed: " + e.getMessage());
    	}
    	return ret;
    }
    
    public static int doDiff(String[] args, SpreadSheetDiffCallback diffCallback) throws Exception {
        if ((args.length < 2)){
            System.out.println(usage());
            return -1;
        }
        final File file1 = new File(args[0]);
        final File file2 = new File(args[1]);
        
        if (!verifyFile(file1) || !verifyFile(file2)) {
        	return -1;
        }
        
        ISpreadSheet ss1 = loadSpreadSheet(file1);
        ISpreadSheet ss2 = loadSpreadSheet(file2);
        
        Map<String,SheetIgnores> sheetIgnores1 = parseSheetIgnores(args, "--ignore1");
        Map<String,SheetIgnores> sheetIgnores2 = parseSheetIgnores(args, "--ignore2");

        SpreadSheetIterator ssi1 = new SpreadSheetIterator(ss1, sheetIgnores1);
        SpreadSheetIterator ssi2 = new SpreadSheetIterator(ss2, sheetIgnores2);
        
        boolean isDiff = false;
        CellPos c1 = null, c2 = null;
        while (true){
            if ((c1==null) && ssi1.hasNext()) c1 = ssi1.next();
            if ((c2==null) && ssi2.hasNext()) c2 = ssi2.next();
            
            if ((c1!=null) && (c2!=null)){
                int c = c1.compareTo(c2);
                if (c == 0){
                    if (!c1.getStringValue().equals(c2.getStringValue())){
                        isDiff = true;
                        diffCallback.reportDiffCell(c1, c2);
                    }
                    c1 = c2 = null;
                }
                else if (c < 0){
                    isDiff = true;
                    diffCallback.reportExtraCell(true, c1);
                    c1 = null;
                }
                else {
                    isDiff = true;
                    diffCallback.reportExtraCell(false, c2);
                    c2 = null;
                }
            } else {
                break;
            }
        }
        if ((c1!=null) && (c2==null)){
            do {
                isDiff = true;
                diffCallback.reportExtraCell(true, c1);
                c1 = ssi1.hasNext() ? ssi1.next() : null;
            } while (c1 != null);
        }
        else if ((c1==null) && (c2!=null)){
            do {
                isDiff = true;
                diffCallback.reportExtraCell(false, c2);
                c2 = ssi2.hasNext() ? ssi2.next() : null;
            } while (c2 != null);
        }
        if ((c1!=null) || (c2!=null)){
            throw new IllegalStateException("Something wrong");
        }
        
        diffCallback.reportWorkbooksDiffer(isDiff, file1, file2);
        
        return isDiff ? 1 : 0;
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
    
    private static boolean verifyFile(File file) {
    	if (!file.exists()) {
    		System.err.println("File: " + file + " does not exist.");
    		return false;
    	}
    	if (!file.canRead()) {
    		System.err.println("File: " + file + " not readable.");
    		return false;
    	}
    	if (!file.isFile()) {
    		System.err.println("File: " + file + " is not a file.");
    		return false;
    	}
    	return true;
    }
    
    private static ISpreadSheet loadSpreadSheet(File file) throws Exception {
    	// assume file is excel by default
    	Exception excelReadException = null;
    	try {
    		Workbook workbook = WorkbookFactory.create(file);
    		return new SpreadSheetExcel(workbook);
    	} catch (Exception e) {
    		excelReadException = e;
    	}
    	Exception odfReadException = null;
    	try {
    		SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.loadDocument(file);
    		return new SpreadSheetOdf(spreadsheetDocument);
    	} catch (Exception e) {
    		odfReadException = e;
    	}
    	if (file.getName().matches(".*\\.ods.*")) {
    		throw new RuntimeException("Failed to read as ods file: " + file, odfReadException);
    	} else {
    		throw new RuntimeException("Failed to read as excel file: " + file, excelReadException);
    	}
    }
}
