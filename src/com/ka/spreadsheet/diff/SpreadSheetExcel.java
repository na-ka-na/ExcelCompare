package com.ka.spreadsheet.diff;

import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class SpreadSheetExcel implements ISpreadSheet {

	private Workbook workbook;
	
	public SpreadSheetExcel(Workbook workbook) {
		this.workbook = workbook;
	}
	
	@Override
	public Iterator<ISheet> getSheetIterator() {
		return new Iterator<ISheet>() {
			
			private int currSheetIdx = 0;

			@Override
			public boolean hasNext() {
				return currSheetIdx < workbook.getNumberOfSheets();
			}

			@Override
			public ISheet next() {
				Sheet sheet = workbook.getSheetAt(currSheetIdx);
				SheetExcel sheetExcel = new SheetExcel(sheet, currSheetIdx);
				currSheetIdx++;
				return sheetExcel;
			}

			@Override
			public void remove() {
				throw new UnsupportedOperationException();
			}
		};
	}
}

class SheetExcel implements ISheet {

	private Sheet sheet;
	private int sheetIdx;
	
	public SheetExcel(Sheet sheet, int sheetIdx) {
		this.sheet = sheet;
		this.sheetIdx = sheetIdx;
	}
	
	@Override
	public String getName() {
		return sheet.getSheetName();
	}
	
	@Override
	public int getSheetIndex() {
		return sheetIdx;
	}

	@Override
	public Iterator<IRow> getRowIterator() {
		final Iterator<Row> rowIterator = sheet.rowIterator();
		return new Iterator<IRow>() {
			
			@Override
			public boolean hasNext() {
				return rowIterator.hasNext();
			}

			@Override
			public IRow next() {
				return new RowExcel(rowIterator.next());
			}

			@Override
			public void remove() {
				throw new UnsupportedOperationException();
			}
		};
	}
}

class RowExcel implements IRow {
	
	private Row row;
	
	public RowExcel(Row row) {
		this.row = row;
	}

	@Override
	public int getRowIndex() {
		return row.getRowNum();
	}

	@Override
	public Iterator<ICell> getCellIterator() {
		final Iterator<Cell> cellIterator = row.cellIterator();
		return new Iterator<ICell>() {
			
			@Override
			public boolean hasNext() {
				return cellIterator.hasNext();
			}

			@Override
			public ICell next() {
				return new CellExcel(cellIterator.next());
			}

			@Override
			public void remove() {
				throw new UnsupportedOperationException();
			}
		};
	}
}

class CellExcel implements ICell {

	private Cell cell;
	
	public CellExcel(Cell cell) {
		this.cell = cell;
	}

	@Override
	public int getRowIndex() {
		return cell.getRowIndex();
	}

	@Override
	public int getColumnIndex() {
		return cell.getColumnIndex();
	}

	@Override
	public String getStringValue() {
		return cell.getStringCellValue();
	}
}
