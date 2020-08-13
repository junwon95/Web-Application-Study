package org.springframework.samples.petclinic.admin;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.List;

public abstract class ExcelManager {

	Workbook wb;

	final int START_ROW = 3;

	final int START_COL = 3;

	ByteArrayInputStream dataToExcel(List<?> table) {
		try {
			this.wb = new XSSFWorkbook();
			ByteArrayOutputStream out = new ByteArrayOutputStream();
			// create sheets
			this.makeSheets(table);
			// write
			wb.write(out);
			return new ByteArrayInputStream(out.toByteArray());
		}
		catch (IOException e) {
			throw new RuntimeException("fail to import data to Excel file: " + e.getMessage());
		}
	}

	public void makeSheet(List<?> table, int referenceId) {
		String sheetName = makeSheetName(table, referenceId);
		Sheet sheet = wb.createSheet(sheetName);
		String[] fields = getFields();

		// makeHeader
		makeHeader(sheet, fields);

		// writeData
		writeData(sheet, table);
	}

	public String makeSheetName(List<?> table, int referenceId) {
		// get class name
		String className = table.get(0).getClass().getName();
		// remove class prefix
		String sheetName = className.substring(className.lastIndexOf('.') + 1) + "s";
		// get sheet's number
		if (referenceId != 0)
			sheetName = className.concat(Integer.toString(referenceId));

		return sheetName;
	}

	public void makeHeader(Sheet sheet, String[] fields) {

		final int COL_SIZE = fields.length;

		Row tableNameRow = sheet.createRow(START_ROW);
		for (int i = START_COL; i < START_COL + COL_SIZE; i++) {
			Cell cell = tableNameRow.createCell(i);
			cell.setCellStyle(style("HEADER"));
		}
		Cell tableNameCell = tableNameRow.getCell(START_COL);
		tableNameCell.setCellValue(sheet.getSheetName());
		sheet.addMergedRegion(new CellRangeAddress(START_ROW, START_ROW, START_COL, START_COL + COL_SIZE - 1));

		Row fieldNamesRow = sheet.createRow(START_ROW + 1);
		for (int i = START_COL; i < START_COL + COL_SIZE; i++) {
			Cell cell = fieldNamesRow.createCell(i);
			cell.setCellValue(fields[i - START_COL]);
			cell.setCellStyle(style("HEADER"));
		}
	}

	// Cell styles
	public CellStyle style(String type) {
		CellStyle style = wb.createCellStyle();
		switch (type) {
		case "HEADER":
			style.setBorderBottom(BorderStyle.MEDIUM);
			style.setBorderLeft(BorderStyle.MEDIUM);
			style.setBorderRight(BorderStyle.MEDIUM);
			style.setBorderTop(BorderStyle.MEDIUM);
			style.setFillForegroundColor(IndexedColors.LEMON_CHIFFON.getIndex());
			style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
			break;
		case "DATA":
			style.setBorderBottom(BorderStyle.DASHED);
			style.setLocked(false);
			break;
		case "DATA_LEFT":
			style.setBorderBottom(BorderStyle.DASHED);
			style.setBorderLeft(BorderStyle.MEDIUM);
			break;
		case "DATA_RIGHT":
			style.setBorderBottom(BorderStyle.DASHED);
			style.setBorderRight(BorderStyle.MEDIUM);
			style.setLocked(false);
			break;
		case "LINK":
			Font font = wb.createFont();
			font.setColor(IndexedColors.BLUE.getIndex());
			style.setFont(font);
			break;
		case "FOOTER":
			style.setBorderBottom(BorderStyle.MEDIUM);
		}
		return style;
	}

	public abstract void makeSheets(List<?> table);

	public abstract void writeData(Sheet sheet, List<?> Table);

	public abstract String[] getFields();

}
