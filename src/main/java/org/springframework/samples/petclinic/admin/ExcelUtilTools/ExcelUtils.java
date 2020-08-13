package org.springframework.samples.petclinic.admin.ExcelUtilTools;

import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.samples.petclinic.admin.Administer;
import org.springframework.samples.petclinic.admin.ReferencedBy;

import java.beans.IntrospectionException;
import java.beans.PropertyDescriptor;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;

import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.List;

public class ExcelUtils {

	private final Workbook wb;

	private final ByteArrayOutputStream out;

	private final int MAIN = 0;

	final int START_ROW = 2;

	final int START_COL = 2;

	int ROW_SIZE;

	int COL_SIZE;

	ExcelUtils() {
		this.wb = new XSSFWorkbook();
		this.out = new ByteArrayOutputStream();
	}

	public ByteArrayInputStream dataToExcel(List<?> table) throws IntrospectionException, InvocationTargetException {
		try {
			// create sheets
			makeSheet(table, MAIN);
			// write
			wb.write(out);
			return new ByteArrayInputStream(out.toByteArray());
		}
		catch (IOException | IllegalAccessException e) {
			throw new RuntimeException("fail to import data to Excel file: " + e.getMessage());
		}
	}

	public <T> void makeSheet(List<T> table, int referenceId)
			throws IllegalAccessException, IntrospectionException, InvocationTargetException {
		// get properties of table
		List<Field> fields = getAdministerFields(table);

		ROW_SIZE = table.size() + 2;
		COL_SIZE = fields.size();
		Cell[][] cells = new Cell[ROW_SIZE][COL_SIZE];

		String sheetName = makeSheetName(table, referenceId);

		// create sheet
		Sheet sheet = wb.createSheet(sheetName);
		// format sheet
		formatSheet(sheet, fields, cells);
		// add data to cells
		writeData(table, fields, cells);

		int rowIdx = START_ROW + 2;
		for (T t : table) {
			List<?> subTable = getSubTable(table, t);
			if (subTable.isEmpty())
				continue;

			// create hyperlink
			String subTableName = makeSheetName(subTable, getId(t));
			Cell hyperlinkCell = sheet.createRow(rowIdx++).createCell(START_COL + COL_SIZE);
			Hyperlink hyperlink = wb.getCreationHelper().createHyperlink(HyperlinkType.DOCUMENT);
			hyperlink.setAddress(subTableName + "!B2");
			hyperlinkCell.setHyperlink(hyperlink);

			// recursive call
			makeSheet(subTable, getId(t));
		}

	}

	public String makeSheetName(List<?> table, int referenceId) {
		// get class name
		String className = table.get(0).getClass().getName();
		// remove class prefix
		String sheetName = className.substring(className.lastIndexOf('.') + 1) + "s";
		// get sheet's number
		if (referenceId != MAIN)
			sheetName = className.concat(Integer.toString(referenceId));

		return sheetName;
	}

	public static int getId(Object object)
			throws IntrospectionException, InvocationTargetException, IllegalAccessException {
		// get class of object
		Class clazz = object.getClass();
		PropertyDescriptor pd = new PropertyDescriptor("id", clazz);
		Method method = pd.getReadMethod();

		return (int) method.invoke(object);
	}

	public List<Field> getAdministerFields(List<?> table) {

		Class clazz = table.get(0).getClass();
		List<Field> fields = new ArrayList<>();

		while (true) {
			List<Field> tmp = new ArrayList<>();
			for (Field field : clazz.getDeclaredFields()) {
				if (field.isAnnotationPresent(Administer.class)) {
					tmp.add(field);
				}
			}
			// add to front to maintain order of fields
			fields.addAll(0, tmp);
			//
			if (fields.get(0).getName().equals("id"))
				break;
			else
				clazz = clazz.getSuperclass();
		}

		return fields;
	}

	public void formatSheet(Sheet sheet, List<Field> fields, Cell[][] cells) {
		int rowIdx = START_ROW;
		int colIdx = START_COL;
		// HEADER region
		Row tableNameRow = sheet.createRow(rowIdx++);
		for (int i = 0; i < COL_SIZE; i++) {
			cells[0][i] = tableNameRow.createCell(colIdx++);
			addCellStyle(cells[0][i], "HEADER");
		}
		cells[0][0].setCellValue(sheet.getSheetName());
		sheet.addMergedRegion(new CellRangeAddress(START_ROW, START_ROW, START_COL, colIdx));

		Row fieldNamesRow = sheet.createRow(rowIdx++);
		colIdx = START_COL;
		for (int i = 0; i < COL_SIZE; i++) {
			cells[1][i] = fieldNamesRow.createCell(colIdx);
			cells[1][i].setCellValue(fields.get(i).getName());
			addCellStyle(cells[1][i], "HEADER");
		}

		// DATA region
		for (int i = 2; i < ROW_SIZE; i++) {
			for (int j = 0; j < COL_SIZE; j++) {
				cells[i][j] = sheet.createRow(rowIdx).createCell(colIdx++);
				addCellStyle(cells[i][j], "BOTTOM_DASHED");
				if (rowIdx == START_ROW + ROW_SIZE - 1)
					addCellStyle(cells[i][j], "BOTTOM_BORDER");
				if (colIdx == START_COL)
					addCellStyle(cells[i][j], "LEFT_BORDER");
				else if (colIdx == START_COL + COL_SIZE - 1)
					addCellStyle(cells[i][j], "RIGHT_BORDER");
			}
			rowIdx++;
			colIdx = START_COL;
		}

	}

	public void writeData(List<?> table, List<Field> fields, Cell[][] cells)
			throws IllegalAccessException, IntrospectionException, InvocationTargetException {
		for (int i = 2; i < ROW_SIZE; i++) {
			for (int j = 0; j < COL_SIZE; j++) {
				String value = getFieldValue(table.get(i - 2), fields.get(j));
				cells[i][j].setCellValue(value);
			}
		}
	}

	public <T> List<?> getSubTable(List<T> table, T t) throws IllegalAccessException {

		Class clazz = table.get(0).getClass();

		for (Field field : clazz.getDeclaredFields()) {
			if (field.isAnnotationPresent(ReferencedBy.class)) {
				return (List<?>) field.get(t);
			}
		}
		return new ArrayList<>();
	}

	public static String getFieldValue(Object object, Field field)
			throws IntrospectionException, InvocationTargetException, IllegalAccessException {
		// get class of object
		String fieldName = field.getName();
		Class clazz = object.getClass();
		PropertyDescriptor pd = new PropertyDescriptor(fieldName, clazz);
		Method method = pd.getReadMethod();

		return method.invoke(object).toString();
	}

	// Cell styles
	public void addCellStyle(Cell cell, String type) {
		CellStyle style = cell.getCellStyle();
		Font font = wb.createFont();
		switch (type) {
		case "HEADER":
			style.setBorderTop(BorderStyle.MEDIUM);
			style.setBorderLeft(BorderStyle.MEDIUM);
			style.setBorderRight(BorderStyle.MEDIUM);
			style.setBorderBottom(BorderStyle.MEDIUM);
			style.setFillForegroundColor(IndexedColors.LEMON_CHIFFON.getIndex());
			style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
			break;
		case "BOTTOM_DASHED":
			style.setBorderBottom(BorderStyle.DASHED);
			break;
		case "TOP_BORDER":
			style.setBorderTop(BorderStyle.MEDIUM);
			break;
		case "LEFT_BORDER":
			style.setBorderLeft(BorderStyle.MEDIUM);
			break;
		case "RIGHT_BORDER":
			style.setBorderRight(BorderStyle.MEDIUM);
			break;
		case "BOTTOM_BORDER":
			style.setBorderBottom(BorderStyle.MEDIUM);
			break;
		case "BLUE_FONT":
			font.setColor(IndexedColors.BLUE.getIndex());
			style.setFont(font);
			break;
		case "UNLOCKED":
			style.setLocked(false);
		}
	}

}
