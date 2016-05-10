package com.ka.spreadsheet.diff;

public class CellValue {

  private final boolean hasFormula;
  private final String formula;
  private final Object value;

  public CellValue(boolean hasFormula, String formula, Object value) {
    this.hasFormula = hasFormula;
    this.formula = formula;
    this.value = value;
  }

  @Override
  public String toString() {
    if (hasFormula && !Flags.DIFF_IGNORE_FORMULAS) {
      return String.valueOf(formula);
    } else {
      return String.valueOf(value);
    }
  }

  public boolean compare(CellValue other) {
    if (!Flags.DIFF_IGNORE_FORMULAS) {
      if (hasFormula ^ other.hasFormula) {
        return false;
      } else if (hasFormula) {
        if (formula == null) {
          return other.formula == null;
        } else {
          return formula.equals(other.formula);
        }
      }
    }
    if (value == null) {
      return other.value == null;
    } else if (other.value == null) {
      return false;
    } else { // both not null
      if (value.equals(other.value)) {
        return true;
      }
      if ((Flags.DIFF_NUMERIC_PRECISION != null)
          && (value instanceof Double) && (other.value instanceof Double)
          && (Math.abs((Double) value - (Double) other.value) < Flags.DIFF_NUMERIC_PRECISION)) {
        return true;
      }
      return false;
    }
  }
}
