package com.ka.spreadsheet.diff;

import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class SpreadSheetUtils {

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

    // INTERNAL_TO_USER
    public static String COL_INTERNAL_TO_USER(int col){
        return convertToLetter(col);
    }
    
    public static int ROW_INTERNAL_TO_USER(int row){
        return row+1;
    }
    
    public static String CELL_INTERNAL_TO_USER(int row, int col){
        return COL_INTERNAL_TO_USER(col) + ROW_INTERNAL_TO_USER(row);
    }
    
    
    // USER_TO_INTERNAL
    public static int COL_USER_TO_INTERNAL(String col){
        return convertFromLetter(col);
    }
    
    public static int ROW_USER_TO_INTERNAL(int row){
        return row-1;
    }
    
    public static final Pattern cellPat = Pattern.compile("([A-Z]+)(\\d+)");
    
    public static int[] CELL_USER_TO_INTERNAL(String cell){
        Matcher matcher = cellPat.matcher(cell);
        if (!matcher.matches())
            throw new IllegalArgumentException("Illegal cell specifier " + cell);
        return new int[]{ROW_USER_TO_INTERNAL(Integer.parseInt(matcher.group(2))),
                         COL_USER_TO_INTERNAL(matcher.group(1))};
    }
}
