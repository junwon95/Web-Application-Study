package org.springframework.samples.petclinic.admin;

import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.List;

public abstract class ExcelManager {

	static Workbook wb;

	int START_ROW = 3;

	int START_COL = 3;

	ByteArrayInputStream dataToExcel(List<?> table) throws IOException {
		try (ByteArrayOutputStream out = new ByteArrayOutputStream();) {
			wb = new XSSFWorkbook();
			// create sheets
			this.makeSheets(table, 0);
			// write
			wb.write(out);
			return new ByteArrayInputStream(out.toByteArray());
		}
		catch (IOException e) {
			throw new RuntimeException("fail to import data to Excel file: " + e.getMessage());
		}
		finally {
			wb.close();
		}
	}

	public void makeSheet(List<?> table, int referenceId) {
		String sheetName = makeSheetName(table, referenceId);
		Sheet sheet = wb.createSheet(sheetName);
		String[] fields = getFields();

		// make Header
		makeHeader(sheet, fields);

		// write Data
		writeData(sheet, table);

		// make hyperlink
		for (int i = START_ROW + 2; i < START_ROW + 2 + table.size(); i++) {
			Cell hyperCell = sheet.getRow(i).createCell(START_COL + fields.length);
			setHyperCell(hyperCell, table.get(i - START_ROW - 2));
		}
	}

	public String makeSheetName(List<?> table, int referenceId) {
		// get class name
		String className = table.get(0).getClass().getName();
		// remove class prefix
		String sheetName = className.substring(className.lastIndexOf('.') + 1) + "s";
		// get sheet's number
		if (referenceId != 0)
			sheetName = sheetName.concat(Integer.toString(referenceId));

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

	public void setCell(Cell cell, String value, CellStyle style) {
		cell.setCellValue(value);
		cell.setCellStyle(style);
	}

	public void writeData(Sheet sheet, List<?> table) {

		final int ROW_LEN = table.size();
		int COL_LEN = getFields().length;

		for (int i = START_ROW + 2; i < START_ROW + 2 + ROW_LEN; i++) {
			String[] fieldValues = getFieldValues(table.get(i - START_ROW - 2));

			Row row = sheet.createRow(i);

			Cell cell = row.createCell(START_COL);
			setCell(cell, fieldValues[0], style("DATA_LEFT"));

			for (int j = START_COL + 1; j < START_COL + COL_LEN - 1; j++) {
				cell = row.createCell(j);
				setCell(cell, fieldValues[j - START_COL], style("DATA"));
			}

			cell = row.createCell(START_COL + COL_LEN - 1);
			setCell(cell, fieldValues[COL_LEN - 1], style("DATA_RIGHT"));

		}
		// auto sizer
		for (int i = START_COL; i <= START_COL + COL_LEN; i++)
			sheet.autoSizeColumn(i);
	}

	public void setHyperLink(Cell cell, String sheetName) {
		Hyperlink hyperlink = wb.getCreationHelper().createHyperlink(HyperlinkType.DOCUMENT);
		hyperlink.setAddress(sheetName + "!B2");
		cell.setHyperlink(hyperlink);
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

	public abstract <T> String[] getFieldValues(T entity);

	public abstract void makeSheets(List<?> table, int referenceId);

	public abstract String[] getFields();

	public abstract <T> void setHyperCell(Cell cell, T entity);

}
