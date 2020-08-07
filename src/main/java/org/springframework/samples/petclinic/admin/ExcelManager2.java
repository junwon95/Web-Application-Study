package org.springframework.samples.petclinic.admin;

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

import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.multipart.MultipartFile;

public class ExcelManager2 {

	public static String TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
	static final String APPLICATION_NAME = "org.springframework.samples.petclinic";

	public static boolean hasExcelFormat(MultipartFile file) {
		return TYPE.equals(file.getContentType());
	}

	// ALL DATA DOWNLOAD
	public static ByteArrayInputStream dataToExcel(List entity) {
		try (Workbook workbook = new XSSFWorkbook(); ByteArrayOutputStream out = new ByteArrayOutputStream()) {

			sheetFormatter(workbook, entity);
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

		return fields;

	}

	public static void sheetFormatter(Workbook workbook, List<Object> entity) throws IntrospectionException, InvocationTargetException, IllegalAccessException {

		// get class of object
		Class clazz = entity.get(0).getClass();
		String entityName = clazz.getName().substring(clazz.getName().lastIndexOf('.')+1);
		List<String> fields = getAdministerFields(clazz);

		// create main sheet
		Sheet sheet = workbook.createSheet(entityName);

		// relative coordinates of table
		int ROW_SPACE = 1;
		final int COL_SPACE = 1;

		// row, col size of table
		final int entityRowSize = entity.size();
		final int entityColSize = fields.size();

		// format HEADER
		Row tableNameRow = sheet.createRow(ROW_SPACE++);
		Cell headerCell = tableNameRow.createCell(COL_SPACE);
		headerCell.setCellValue(entityName);
		sheet.addMergedRegion(new CellRangeAddress(1,1, 1, fields.size()));

		Row fieldNamesRow = sheet.createRow(ROW_SPACE++);
		for(int i = 0; i < entityColSize; i++){
			Cell fieldCell = fieldNamesRow.createCell(COL_SPACE+i);
			fieldCell.setCellValue(fields.get(i));
		}

		// format DATA region
		for(int i = 0; i < entityRowSize; i++){
			Row row = sheet.createRow(ROW_SPACE+i);
			Object object = entity.get(i);
			for(int j = 0; j < entityColSize; j++){
				if(j == 0){
					row.createCell(COL_SPACE).setCellValue(getId(object));
					continue;
				}
				Cell cell = row.createCell(COL_SPACE+j);
				cell.setCellValue(getFieldValue(object, fields.get(j)));
				sheet.autoSizeColumn(COL_SPACE+j);
			}
			Cell hyperlinkCell = row.createCell(COL_SPACE+entityColSize);

			//get @ReferencedBy field name and create new sheet using sheetFormatter() recursively
			String str = getReferencedField(clazz);
			if(str.isEmpty()) continue;
			else{
				String referencedFieldName = str.substring(0,1).toUpperCase() + str.substring(1,str.length()-1);
				workbook.createSheet(referencedFieldName + getId(object));

				// create hyperlink
				hyperlinkCell.setCellValue(getReferencedField(clazz));
				Hyperlink hyperlink = workbook.getCreationHelper().createHyperlink(HyperlinkType.DOCUMENT);
				hyperlink.setAddress(referencedFieldName + getId(object) + "!A1");
				headerCell.setHyperlink(hyperlink);

				// recursive call of sheetFormatter
				List<Object> linkedEntity = getLinkedEntity(clazz, object);
				sheetFormatter(workbook, linkedEntity);
			}

		}

	}

	public static int getId(Object object) throws IntrospectionException, InvocationTargetException, IllegalAccessException {
		// get class of object
		Class clazz = object.getClass();
		PropertyDescriptor pd = new PropertyDescriptor("id", clazz);
		Method method = pd.getReadMethod();

		return (int)method.invoke(object);
	}

	public static String getFieldValue(Object object, String field) throws IntrospectionException, InvocationTargetException, IllegalAccessException {
		// get class of object
		Class clazz = object.getClass();
		PropertyDescriptor pd = new PropertyDescriptor(field, clazz);
		Method method = pd.getReadMethod();

		return (String)method.invoke(object);
	}

	public static String getReferencedField(Class clazz){

		for (Field field : clazz.getDeclaredFields()) {
			if (field.isAnnotationPresent(ReferencedBy.class)) {
				return field.getName();
			}
		}
		return "";
	}

	public static List<Object> getLinkedEntity(Class clazz, Object object) throws IntrospectionException, InvocationTargetException, IllegalAccessException {

		for(Method method : clazz.getDeclaredMethods()){
			if(method.isAnnotationPresent(LinkedEntityGetter.class)) {
				return (List<Object>) method.invoke(object);
			}
		}
		return null;
	}


}
