package com.ka.spreadsheet.diff;

import java.util.HashMap;
import java.util.Map;

import javax.annotation.Nullable;

public class WorkbookIgnores {
	private Map<String,SheetIgnores> ignore;
	
	/**
	 * Create container for workbook ignores based on given command line args that follow
	 * the opt argument.
	 * @param args array of command line parameters
	 * @param opt name of matching command line argument
	 */	
	public WorkbookIgnores(String[] args, String opt) {
		ignore = parseSheetIgnores(args, opt);
	}
	
	/**
	 * Create container for workbook ignores based on given command line args that follow
	 * the opt argument. Append additional ignores from workbookIgnores
	 * @param args array of command line parameters
	 * @param opt name of matching command line argument
	 * @param workbookIgnores used for additional ignores
	 */
	public WorkbookIgnores(String[] args, String opt, WorkbookIgnores workbookIgnores)
	{
		this(args, opt);
		ignore.putAll(workbookIgnores.ignore);
	}

	public @Nullable SheetIgnores fetchSheetIgnores(String sheetName) {
		SheetIgnores ignoredByName = ignore.get(sheetName);
		SheetIgnores ignoredAll = ignore.get("");
		return ignoredByName != null ? ignoredByName : ((ignoredAll != null) ? ignoredAll : null );
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
