package com.ka.spreadsheet.diff;

import java.util.HashMap;
import java.util.Map;

import javax.annotation.Nullable;

public class WorkbookIgnores {
  private Map<String, SheetIgnores> ignores;

  public WorkbookIgnores(Map<String, SheetIgnores> ignores) {
    this.ignores = ignores;
  }

  public @Nullable SheetIgnores fetchSheetIgnores(String sheetName) {
    SheetIgnores ignoredByName = ignores.get(sheetName);
    SheetIgnores ignoredAll = ignores.get("");
    return ignoredByName != null ? ignoredByName : ((ignoredAll != null) ? ignoredAll : null);
  }

  public static WorkbookIgnores parseWorkbookIgnores(String[] args, String opt) {
    int start = -1, end = -1;
    for (int i = 0; i < args.length; i++) {
      if (start == -1) {
        if (opt.equals(args[i])) {
          start = i + 1;
        }
      } else {
        if (args[i].startsWith("--")) {
          end = i;
        }
      }
    }
    if (end == -1)
      end = args.length;

    Map<String, SheetIgnores> ret = new HashMap<String, SheetIgnores>();
    if (start != -1) {
      for (int i = start; i < end; i++) {
        SheetIgnores s = SheetIgnores.newSheetIgnore(args[i]);
        ret.put(s.sheetName(), s);
      }
    }
    return new WorkbookIgnores(ret);
  }
}
