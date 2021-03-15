package com.ka.spreadsheet.diff;

import static com.ka.spreadsheet.diff.Flags.WORKBOOK1;
import static com.ka.spreadsheet.diff.Flags.WORKBOOK2;

import java.io.File;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.odftoolkit.simple.SpreadsheetDocument;


public class SpreadSheetDiffer {

  public static void main(String[] args) {
    int ret = doDiff(args);
    System.exit(ret);
  }

  public static int doDiff(String[] args) {
    int ret = -1;
    try {
      if (Flags.parseFlags(args)) {
        SpreadSheetDiffCallback formatter;
        switch (Flags.DIFF_FORMAT) {
          case EXCEL_CMP:
            formatter = new StdoutSpreadSheetDiffCallback();
            break;
          case UNIFIED:
            formatter = new UnifiedDiffSpreadSheetDiffCallback();
            break;
          default:
            throw new IllegalArgumentException("Unknown diff formatter");
        }
        ret = doDiff(formatter);
      }
    } catch (Exception e) {
      if (Flags.DEBUG) {
        e.printStackTrace(System.err);
      } else {
        System.err.println("Diff failed: " + e.getMessage());
      }
    }
    return ret;
  }

  public static int doDiff(SpreadSheetDiffCallback diffCallback) throws Exception {
    if (!verifyFile(WORKBOOK1) || !verifyFile(WORKBOOK2)) {
      return -1;
    }

    ISpreadSheet ss1 = isDevNull(WORKBOOK1) ? emptySpreadSheet() : loadSpreadSheet(WORKBOOK1);
    ISpreadSheet ss2 = isDevNull(WORKBOOK2) ? emptySpreadSheet() : loadSpreadSheet(WORKBOOK2);

    ISpreadSheetIterator ssi1 = isDevNull(WORKBOOK1) ?
        emptySpreadSheetIterator() : new SpreadSheetIterator(ss1, Flags.WORKBOOK_IGNORES1);
    ISpreadSheetIterator ssi2 = isDevNull(WORKBOOK2) ?
        emptySpreadSheetIterator() : new SpreadSheetIterator(ss2, Flags.WORKBOOK_IGNORES2);

    boolean isDiff = false;
    CellPos c1 = null, c2 = null;
    while (true) {
      if ((c1 == null) && ssi1.hasNext())
        c1 = ssi1.next();
      if ((c2 == null) && ssi2.hasNext())
        c2 = ssi2.next();

      if ((c1 != null) && (c2 != null)) {
        int c = c1.compareCellPositions(c2);
        if (c == 0) {
          if (!c1.getCellValue().compare(c2.getCellValue())) {
            isDiff = true;
            diffCallback.reportDiffCell(c1, c2);
          }
          c1 = c2 = null;
        } else if (c < 0) {
          isDiff = true;
          diffCallback.reportExtraCell(true, c1);
          c1 = null;
        } else {
          isDiff = true;
          diffCallback.reportExtraCell(false, c2);
          c2 = null;
        }
      } else {
        break;
      }
    }
    if ((c1 != null) && (c2 == null)) {
      do {
        isDiff = true;
        diffCallback.reportExtraCell(true, c1);
        c1 = ssi1.hasNext() ? ssi1.next() : null;
      } while (c1 != null);
    } else if ((c1 == null) && (c2 != null)) {
      do {
        isDiff = true;
        diffCallback.reportExtraCell(false, c2);
        c2 = ssi2.hasNext() ? ssi2.next() : null;
      } while (c2 != null);
    }
    if ((c1 != null) || (c2 != null)) {
      throw new IllegalStateException("Something wrong");
    }

    Boolean hasMacro1 = ss1.hasMacro();
    Boolean hasMacro2 = ss2.hasMacro();
    if ((hasMacro1 != null) && (hasMacro2 != null) && (hasMacro1 != hasMacro2)) {
      isDiff = true;
      diffCallback.reportMacroOnlyIn(hasMacro1);
    }

    diffCallback.reportWorkbooksDiffer(isDiff, WORKBOOK1, WORKBOOK2);

    return isDiff ? 1 : 0;
  }

  private static boolean isDevNull(File file) {
    return "/dev/null".equals(file.getAbsolutePath())
        || "\\\\.\\NUL".equals(file.getAbsolutePath());
  }

  private static boolean verifyFile(File file) {
    if (isDevNull(file)) {
      return true;
    }
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

  private static ISpreadSheet emptySpreadSheet() {
    return new ISpreadSheet() {
      @Override
      public Boolean hasMacro() {
        return false;
      }
      @Override
      public Iterator<ISheet> getSheetIterator() {
        return new Iterator<ISheet>() {
          @Override
          public boolean hasNext() {
            return false;
          }
          @Override
          public ISheet next() {
            throw new IllegalStateException();
          }
          @Override
          public void remove() {
            throw new IllegalStateException();
          }
        };
      }
    };
  }

  private static ISpreadSheetIterator emptySpreadSheetIterator() {
    return new ISpreadSheetIterator() {
      @Override
      public boolean hasNext() {
        return false;
      }
      @Override
      public CellPos next() {
        throw new IllegalStateException();
      }
    };
  }
}
