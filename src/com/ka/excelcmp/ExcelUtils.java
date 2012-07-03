package com.ka.excelcmp;

import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class ExcelUtils {

    public static String convertToLetter(int col){
        if (col >= 26) {
            int mod26 = (col % 26);
            return convertToLetter(((col - mod26) / 26) - 1) + convertToLetter(mod26);
        } else {
            return String.valueOf((char)(col + 65));
        }
    }
    
    public static int convertFromLetter(String col){
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
