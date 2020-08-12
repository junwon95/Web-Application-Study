package org.springframework.samples.petclinic.admin.ExcelUtilTools;

import java.beans.IntrospectionException;
import java.beans.PropertyDescriptor;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.samples.petclinic.admin.Administer;
import org.springframework.samples.petclinic.owner.Owner;
import org.springframework.samples.petclinic.owner.Pet;
import org.springframework.web.multipart.MultipartFile;

public class ExcelManager {

	static final String APPLICATION_NAME = "org.springframework.samples.petclinic";

	public static String TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

	// attributes of Owner relation
	static String SHEET = "AdministrationData";

	public static boolean hasExcelFormat(MultipartFile file) {
		return TYPE.equals(file.getContentType());
	}

	// UPLOAD
	public static List<Owner> excelToOwners(InputStream is) {
		try {
			Workbook workbook = new XSSFWorkbook(is);

			Sheet sheet = workbook.getSheet(SHEET);
			Iterator<Row> rows = sheet.iterator();

			List<Owner> owners = new ArrayList<Owner>();

			int rowNumber = 0;
			while (rows.hasNext()) {
				Row currentRow = rows.next();

				// skip header
				if (rowNumber == 0) {
					rowNumber++;
					continue;
				}

				Iterator<Cell> cellsInRow = currentRow.iterator();

				Owner owner = new Owner();

				int cellIdx = 0;
				while (cellsInRow.hasNext()) {
					Cell currentCell = cellsInRow.next();
					// get value from each cell of row and set to table
					switch (cellIdx) {
					case 0:
						owner.setId((int) Math.round(currentCell.getNumericCellValue()));
						break;
					case 1:
						owner.setFirstName(currentCell.getStringCellValue());
						break;
					case 2:
						owner.setLastName(currentCell.getStringCellValue());
						break;
					case 3:
						owner.setAddress(currentCell.getStringCellValue());
						break;
					case 4:
						owner.setCity(currentCell.getStringCellValue());
						break;
					case 5:
						owner.setTelephone(currentCell.getStringCellValue());
						break;
					default:
						break;
					}
					cellIdx++;
				}
				owners.add(owner);
			}
			workbook.close();
			return owners;
		}
		catch (IOException e) {
			throw new RuntimeException("fail to parse Excel file: " + e.getMessage());
		}
	}

	static String[] ENTITY_NAMES = { "owner.Owner", "owner.Pet" };

	// ALL DATA DOWNLOAD
	public static ByteArrayInputStream dataToExcel(List<Owner> owners) {
		try (Workbook workbook = new XSSFWorkbook(); ByteArrayOutputStream out = new ByteArrayOutputStream()) {

			// SHEET
			Sheet sheet = workbook.createSheet("dataBaseInstance");
			// freeze header rows
			sheet.createFreezePane(3, 3);

			// HEADER
			List<List<String>> allFields = getAdministerFields(ENTITY_NAMES);
			headerFormatter(sheet, allFields, workbook);

			// DATA
			dataFormatter(sheet, owners, allFields, workbook);

			// autosize columns to fit text
			int cnt = 0;
			for (List<String> fields : allFields) {
				for (String f : fields) {
					sheet.autoSizeColumn(cnt++);
				}
			}

			// write
			workbook.write(out);
			return new ByteArrayInputStream(out.toByteArray());
		}
		catch (IOException | ClassNotFoundException | IntrospectionException | InvocationTargetException
				| IllegalAccessException e) {
			throw new RuntimeException("fail to import data to Excel file: " + e.getMessage());
		}
	}

	/**
	 * @param ENTITY_NAMES
	 * @return allFields returns all fields annotated with @Administer including the
	 * fields of parent classes. searches recursively until field "id" is found.
	 * @throws ClassNotFoundException
	 *
	 */
	public static List<List<String>> getAdministerFields(String[] ENTITY_NAMES) throws ClassNotFoundException {
		List<List<String>> allFields = new ArrayList<>();

		for (String name : ENTITY_NAMES) {
			List<String> fields = new ArrayList<>();

			Class clazz = Class.forName(APPLICATION_NAME + '.' + name);

			while (true) {
				List<String> tmp = new ArrayList<>();
				for (Field field : clazz.getDeclaredFields()) {
					if (field.isAnnotationPresent(Administer.class)) {
						tmp.add(field.getName());
					}
				}
				// add to front to maintain order of fields
				fields.addAll(0, tmp);

				//
				if (fields.get(0).equals("id"))
					break;
				else
					clazz = clazz.getSuperclass();
			}

			allFields.add(fields);
		}

		return allFields;
	}

	public static void headerFormatter(Sheet sheet, List<List<String>> allFields, Workbook workbook) {

		Row tableRow = sheet.createRow(1);
		Row attributeRow = sheet.createRow(2);
		int firstCol, lastCol = 1;
		int entityIdx = 0;
		for (List<String> fields : allFields) {
			// name attribute cells
			firstCol = lastCol;
			for (String f : fields) {
				Cell cell = tableRow.createCell(lastCol);
				cell.setCellStyle(style(0, workbook));

				cell = attributeRow.createCell(lastCol++);
				cell.setCellValue(f);
				cell.setCellStyle(style(0, workbook));
			}
			// name and merge table header cells
			String str = ENTITY_NAMES[entityIdx++];
			CellUtil.createCell(tableRow, firstCol, str.substring(str.indexOf('.') + 1));
			sheet.addMergedRegion(new CellRangeAddress(1, 1, firstCol, lastCol - 1));
		}

	}

	public static void dataFormatter(Sheet sheet, List<Owner> owners, List<List<String>> allFields, Workbook workbook)
			throws IntrospectionException, InvocationTargetException, IllegalAccessException {

		int entityIdx = 0;
		int rowIdx = 3;
		int colIdx = 1;

		for (Owner owner : owners) {
			int startRow = rowIdx;

			Row row = sheet.createRow(rowIdx);
			colIdx = makeCells(row, colIdx, owner, allFields.get(entityIdx), workbook);

			entityIdx++;
			for (Pet pet : owner.getPets()) {
				makeCells(row, colIdx, pet, allFields.get(entityIdx), workbook);
				row = sheet.createRow(++rowIdx);
			}
			if (owner.getPets().size() == 0) {
				CellUtil.createCell(row, colIdx, "", style(1, workbook));
				for (int i = 1; i < allFields.get(entityIdx).size() - 1; i++) {
					CellUtil.createCell(row, colIdx + i, "", style(2, workbook));
				}
				CellUtil.createCell(row, colIdx + allFields.get(entityIdx).size() - 1, "", style(3, workbook));
			}
			entityIdx--;

			// if more than one pet
			if (rowIdx - 1 > startRow) {
				for (int i = startRow + 1; i < rowIdx; i++) {
					row = sheet.getRow(i);
					row.createCell(1).setCellStyle(style(1, workbook));
					for (int j = 2; j < colIdx - 1; j++) {
						row.createCell(j).setCellStyle(style(2, workbook));
					}
					row.createCell(colIdx - 1).setCellStyle(style(3, workbook));
				}
				for (int i = 1; i < colIdx; i++)
					sheet.addMergedRegion(new CellRangeAddress(startRow, rowIdx - 1, i, i));
			}

			colIdx = 1;
		}

	}

	public static int makeCells(Row row, int colIdx, Object object, List<String> fields, Workbook workbook)
			throws IntrospectionException, InvocationTargetException, IllegalAccessException {

		Class clazz = object.getClass();

		Cell cell = row.createCell(colIdx);
		for (String field : fields) {
			PropertyDescriptor pd = new PropertyDescriptor(field, clazz);
			Method method = pd.getReadMethod();

			cell = row.createCell(colIdx++);
			cell.setCellStyle(style(2, workbook));

			if (field.equals("id")) {
				cell.setCellValue((int) method.invoke(object));
				cell.setCellStyle(style(1, workbook));
			}
			else {
				cell.setCellValue(method.invoke(object).toString());
			}
		}
		cell.setCellStyle(style(3, workbook));

		return colIdx;
	}

	public static CellStyle style(int type, Workbook workbook) {
		CellStyle style = workbook.createCellStyle();
		switch (type) {
		case 0:
			style.setBorderBottom(BorderStyle.MEDIUM);
			style.setBorderLeft(BorderStyle.MEDIUM);
			style.setBorderRight(BorderStyle.MEDIUM);
			style.setBorderTop(BorderStyle.MEDIUM);
			break;
		case 1:
			style.setBorderBottom(BorderStyle.DASHED);
			style.setBorderLeft(BorderStyle.MEDIUM);
			style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
			break;
		case 2:
			style.setBorderBottom(BorderStyle.DASHED);
			style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
			break;
		case 3:
			style.setBorderBottom(BorderStyle.DASHED);
			style.setBorderRight(BorderStyle.MEDIUM);
			style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
			break;
		}
		return style;
	}

}
