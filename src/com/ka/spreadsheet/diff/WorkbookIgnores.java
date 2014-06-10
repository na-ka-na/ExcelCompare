package com.ka.spreadsheet.diff;

import java.util.HashMap;
import java.util.Map;

public class WorkbookIgnores {
	private Map<String,SheetIgnores> ignore;
	
	public WorkbookIgnores(String[] args, String opt) {
		ignore = parseSheetIgnores(args, opt);
	}
	
	public SheetIgnores fetchSheetIgnores(String sheetName) {
		return ignore.get(sheetName) != null? ignore.get(sheetName) : (ignore.keySet().contains("") ? ignore.get("") : null);
	}
	
	
	private Map<String,SheetIgnores> parseSheetIgnores(String[] args, String opt){
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
