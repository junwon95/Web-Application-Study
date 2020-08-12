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

import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.samples.petclinic.admin.Administer;
import org.springframework.samples.petclinic.admin.LinkedEntityGetter;
import org.springframework.samples.petclinic.admin.ReferencedBy;
import org.springframework.samples.petclinic.dto.ExcelDto;
import org.springframework.samples.petclinic.owner.Owner;

public class ExcelManager2 {

	// ALL DATA DOWNLOAD
	public static ByteArrayInputStream dataToExcel(List entity) {
		try (Workbook workbook = new XSSFWorkbook(); ByteArrayOutputStream out = new ByteArrayOutputStream()) {

			sheetFormatter(workbook, entity, 0);
			// write
			workbook.write(out);
			return new ByteArrayInputStream(out.toByteArray());
		}
		catch (IOException | IntrospectionException | InvocationTargetException | IllegalAccessException e) {
			throw new RuntimeException("fail to import data to Excel file: " + e.getMessage());
		}
	}

	public static List<String> getAdministerFields(Class clazz) {

		List<String> fields = new ArrayList<>();

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

		return fields;

	}

	public static void sheetFormatter(Workbook workbook, List<Object> entity, int referenceId)
			throws IntrospectionException, InvocationTargetException, IllegalAccessException {
		// get class of object
		if (entity.get(0) == null)
			return;
		Class clazz = entity.get(0).getClass();
		String entityName = clazz.getName().substring(clazz.getName().lastIndexOf('.') + 1);
		List<String> fields = getAdministerFields(clazz);

		// create main sheet
		String sheetName = entityName;
		if (referenceId != 0) {
			sheetName = sheetName.concat(Integer.toString(referenceId));
		}
		Sheet sheet = workbook.createSheet(sheetName);
		sheet.createFreezePane(2, 3);

		// relative coordinates of table
		int ROW_SPACE = 1;
		final int COL_SPACE = 1;

		// row, col size of table
		final int entityRowSize = entity.size();
		final int entityColSize = fields.size();

		// create home button
		if (referenceId != 0) {
			Cell homeButton = sheet.createRow(ROW_SPACE + entityRowSize + 2).createCell(COL_SPACE);
			Hyperlink home = workbook.getCreationHelper().createHyperlink(HyperlinkType.DOCUMENT);
			home.setAddress("Owner" + "!B2");
			homeButton.setHyperlink(home);
			homeButton.setCellStyle(style("HEADER", workbook));
			homeButton.setCellValue("BACK");
		}

		// format HEADER
		Row tableNameRow = sheet.createRow(ROW_SPACE++);
		for (int i = 0; i < entityColSize; i++) {
			Cell headerCell = tableNameRow.createCell(COL_SPACE + i);
			headerCell.setCellValue(entityName);
			headerCell.setCellStyle(style("HEADER", workbook));
		}
		sheet.addMergedRegion(new CellRangeAddress(1, 1, 1, entityColSize));

		Row fieldNamesRow = sheet.createRow(ROW_SPACE++);
		for (int i = 0; i < entityColSize; i++) {
			Cell fieldCell = fieldNamesRow.createCell(COL_SPACE + i);
			fieldCell.setCellValue(fields.get(i));
			fieldCell.setCellStyle(style("HEADER", workbook));
		}

		// format DATA region
		for (int i = 0; i < entityRowSize; i++) {
			Row row = sheet.createRow(ROW_SPACE + i);

			Object object = entity.get(i);

			for (int j = 0; j < entityColSize; j++) {
				Cell cell;
				if (j == 0) {
					cell = row.createCell(COL_SPACE);
					cell.setCellValue(getId(object));
					cell.setCellStyle(style("DATA_LEFT", workbook));
					continue;
				}
				cell = row.createCell(COL_SPACE + j);
				cell.setCellValue(getFieldValue(object, fields.get(j)));
				if (j < entityColSize - 1)
					cell.setCellStyle(style("DATA", workbook));
				else
					cell.setCellStyle(style("DATA_RIGHT", workbook));

				// auto size setting
				sheet.autoSizeColumn(COL_SPACE + j);
			}

			Cell hyperlinkCell = row.createCell(COL_SPACE + entityColSize);
			// get @ReferencedBy field name and create new sheet using sheetFormatter()
			// recursively
			String referencedField = getReferencedField(clazz);

			if (!referencedField.isEmpty()) {
				String referencedFieldName = referencedField.substring(0, 1).toUpperCase()
						+ referencedField.substring(1, referencedField.length() - 1);

				// get linked entity
				List<Object> linkedEntity = getLinkedEntity(clazz, object);
				if (linkedEntity.size() == 0)
					continue;
				// set value of hyperlink cell
				hyperlinkCell.setCellValue("see " + referencedField);
				hyperlinkCell.setCellStyle(style("LINK", workbook));

				// recursive call of sheetFormatter
				sheetFormatter(workbook, linkedEntity, getId(object));

				// create hyperlink
				Hyperlink hyperlink = workbook.getCreationHelper().createHyperlink(HyperlinkType.DOCUMENT);
				hyperlink.setAddress(referencedFieldName + getId(object) + "!B2");
				hyperlinkCell.setHyperlink(hyperlink);
			}

		}
		// lock sheet
		sheet.protectSheet("123");
	}

	public static int getId(Object object)
			throws IntrospectionException, InvocationTargetException, IllegalAccessException {
		// get class of object
		Class clazz = object.getClass();
		PropertyDescriptor pd = new PropertyDescriptor("id", clazz);
		Method method = pd.getReadMethod();

		return (int) method.invoke(object);
	}

	public static String getFieldValue(Object object, String field)
			throws IntrospectionException, InvocationTargetException, IllegalAccessException {
		// get class of object
		Class clazz = object.getClass();
		PropertyDescriptor pd = new PropertyDescriptor(field, clazz);
		Method method = pd.getReadMethod();

		return method.invoke(object).toString();
	}

	public static String getReferencedField(Class clazz) {

		for (Field field : clazz.getDeclaredFields()) {
			if (field.isAnnotationPresent(ReferencedBy.class)) {
				return field.getName();
			}
		}
		return "";
	}

	public static List<Object> getLinkedEntity(Class clazz, Object object)
			throws InvocationTargetException, IllegalAccessException {

		// find and return linked entity field
		for (Method method : clazz.getDeclaredMethods()) {
			if (method.isAnnotationPresent(LinkedEntityGetter.class)) {
				List<Object> list = (List) method.invoke(object);
				return list;
			}
		}
		return null;
	}

	// ALL DATA UPLOAD
	public static ExcelDto excelToDB(InputStream is) {
		try {
			ExcelDto excelDto = new ExcelDto();
			Workbook workbook = new XSSFWorkbook(is);

			Sheet sheet = workbook.getSheet("Owner");

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

			excelDto.setOwners(owners);
			return excelDto;
		}
		catch (IOException e) {
			throw new RuntimeException("fail to parse Excel file: " + e.getMessage());
		}
	}

	// Cell styles
	public static CellStyle style(String type, Workbook workbook) {
		CellStyle style = workbook.createCellStyle();
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
			Font font = workbook.createFont();
			font.setColor(IndexedColors.BLUE.getIndex());
			style.setFont(font);
			break;
		case "FOOTER":
			style.setBorderBottom(BorderStyle.MEDIUM);
		}
		return style;
	}

}
