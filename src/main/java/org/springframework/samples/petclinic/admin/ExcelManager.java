package org.springframework.samples.petclinic.admin;

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
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.hibernate.jdbc.Work;
import org.springframework.samples.petclinic.owner.Owner;
import org.springframework.samples.petclinic.owner.Pet;
import org.springframework.web.multipart.MultipartFile;

public class ExcelManager {

	static final String APPLICATION_NAME = "org.springframework.samples.petclinic";
	public static String TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

	// attributes of Owner relation
	static String[] HEADERs = { "id", "firstName", "lastName", "address", "city", "telephone" };
	static String SHEET = "AdministrationData";

	public static boolean hasExcelFormat(MultipartFile file) {
		if (!TYPE.equals(file.getContentType())) {
			return false;
		}
		return true;
	}

	// DOWNLOAD
	public static ByteArrayInputStream ownersToExcel(List<Owner> owners) {

		try (Workbook workbook = new XSSFWorkbook(); ByteArrayOutputStream out = new ByteArrayOutputStream();) {
			Sheet sheet = workbook.createSheet(SHEET);
			sheet.setAutoFilter(new CellRangeAddress(0, 0, 0, HEADERs.length));
			sheet.createFreezePane(0, 1);

			// Header
			Row headerRow = sheet.createRow(0);

			for (int col = 0; col < HEADERs.length; col++) {
				Cell cell = headerRow.createCell(col);
				cell.setCellValue(HEADERs[col]);
			}

			// Data
			CellStyle unlockedCellStyle = workbook.createCellStyle();
			unlockedCellStyle.setLocked(false);

			int rowIdx = 1;
			for (Owner owner : owners) {
				Row row = sheet.createRow(rowIdx++);

				row.createCell(0).setCellValue(owner.getId());
				row.createCell(1).setCellValue(owner.getFirstName());
				row.createCell(2).setCellValue(owner.getLastName());
				row.createCell(3).setCellValue(owner.getAddress());
				row.createCell(4).setCellValue(owner.getCity());
				row.createCell(5).setCellValue(owner.getTelephone());

				row.setRowStyle(unlockedCellStyle);
			}

			workbook.write(out);
			return new ByteArrayInputStream(out.toByteArray());
		}
		catch (IOException e) {
			throw new RuntimeException("fail to import data to Excel file: " + e.getMessage());
		}
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


	// ------------ Current TODO

	static String[] ENTITY_NAMES = {"owner.Owner", "owner.Pet"};

	// ALL DATA DOWNLOAD
	public static ByteArrayInputStream dataToExcel(List<Owner> owners) {
		try (Workbook workbook = new XSSFWorkbook(); ByteArrayOutputStream out = new ByteArrayOutputStream()) {

			// SHEET
			Sheet sheet = workbook.createSheet("dataBaseInstance");
			// freeze header rows
			sheet.createFreezePane(0, 2);

			// HEADER
			// retrieve field names of each entity
			List<List<String>> allFields = getAdministerFields(ENTITY_NAMES);
			headerFormatter(sheet, allFields, workbook);

			// TODO: Data
			dataFormatter(sheet, owners, allFields, workbook);

			// add filter
//			sheet.setAutoFilter(new CellRangeAddress(1, 1, 0, columnLength));

			workbook.write(out);
			return new ByteArrayInputStream(out.toByteArray());
		}
		catch (IOException | ClassNotFoundException | IntrospectionException | InvocationTargetException | IllegalAccessException e) {
			throw new RuntimeException("fail to import data to Excel file: " + e.getMessage());
		}
	}


	/**
	 *
	 * @param ENTITY_NAMES
	 * @return allFields
	 * returns all fields annotated with @Administer including the fields of parent classes.
	 * searches recursively until field "id" is found.
	 *
	 * @throws ClassNotFoundException
	 *
	 */
	public static List<List<String>> getAdministerFields(String[] ENTITY_NAMES) throws ClassNotFoundException {
		List<List<String>> allFields = new ArrayList<>();

		for (String name : ENTITY_NAMES){
			List<String> fields = new ArrayList<>();

			Class clazz = Class.forName(APPLICATION_NAME + '.' + name);

			while(true){
				List<String> tmp = new ArrayList<>();
				for (Field field : clazz.getDeclaredFields()) {
					if (field.isAnnotationPresent(Administer.class)) {
						tmp.add(field.getName());
					}
				}
				// add to front to maintain order of fields
				fields.addAll(0,tmp);

				//
				if(fields.get(0).equals("id")) break;
				else clazz = clazz.getSuperclass();
			}

			allFields.add(fields);
		}

		return allFields;
	}

	public static void headerFormatter(Sheet sheet, List<List<String>> allFields, Workbook workbook) {

		Row tableRow = sheet.createRow(0);
		Row attributeRow = sheet.createRow(1);
		int firstCol, lastCol = 0;
		int entityIdx = 0;
		for (List<String> fields : allFields) {
			// name attribute cells
			firstCol = lastCol;
			for(String f : fields){
				Cell cell = attributeRow.createCell(lastCol++);
				cell.setCellValue(f);
				cellStyler(cell, workbook);
			}
			// name and merge table header cells
			Cell cell = tableRow.createCell(firstCol);
			cell.setCellValue(ENTITY_NAMES[entityIdx++]);
			sheet.addMergedRegion(new CellRangeAddress(0,0, firstCol, lastCol-1));
			cellStyler(cell, workbook);
		}
	}

	public static void dataFormatter(Sheet sheet, List<Owner> owners, List<List<String>> allFields, Workbook workbook)
		throws IntrospectionException, InvocationTargetException, IllegalAccessException {

		int entityIdx = 0;
		int rowIdx = 2;
		int colIdx = 0;

		for(Owner owner : owners){
			int startRow = rowIdx;
			Row row = sheet.createRow(rowIdx);
			colIdx = makeCells(row, colIdx, owner, allFields.get(entityIdx), workbook);

			entityIdx++;
			for(Pet pet : owner.getPets()){
				makeCells(row, colIdx, pet, allFields.get(entityIdx), workbook);
				row = sheet.createRow(++rowIdx);
			}
			entityIdx--;

			if(rowIdx-1 > startRow){
				for(int i = 0; i < colIdx; i++)
					sheet.addMergedRegion(new CellRangeAddress(startRow,rowIdx-1, i, i));
			}

			colIdx = 0;
		}
	}

	public static int makeCells(Row row, int colIdx, Object object, List<String> fields, Workbook workbook)
		throws IntrospectionException, InvocationTargetException, IllegalAccessException {

		Class clazz = object.getClass();

		for(String field : fields){
			PropertyDescriptor pd = new PropertyDescriptor(field, clazz);
			Method method = pd.getReadMethod();

			Cell cell = row.createCell(colIdx++);
			if (field.equals("id")){
				cell.setCellValue((int)method.invoke(object));
			}
			else{
				cell.setCellValue(method.invoke(object).toString());
			}
		}

		return colIdx;
	}

	public static void cellStyler(Cell cell, Workbook workbook){
		CellStyle style = workbook.createCellStyle();
		style.setBorderBottom(BorderStyle.MEDIUM);
		style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
		style.setBorderLeft(BorderStyle.MEDIUM);
		style.setLeftBorderColor(IndexedColors.BLACK.getIndex());
		style.setBorderRight(BorderStyle.MEDIUM);
		style.setRightBorderColor(IndexedColors.BLACK.getIndex());
		style.setBorderTop(BorderStyle.MEDIUM);
		style.setTopBorderColor(IndexedColors.BLACK.getIndex());
		cell.setCellStyle(style);
	}

}
