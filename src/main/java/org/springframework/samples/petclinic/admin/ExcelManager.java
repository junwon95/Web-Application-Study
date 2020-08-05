package org.springframework.samples.petclinic.admin;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.samples.petclinic.owner.Owner;
import org.springframework.samples.petclinic.owner.OwnerRepository;
import org.springframework.web.multipart.MultipartFile;

public class ExcelManager {

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

	static String[] ENTITY_NAMES = {"Owner", "Pet"};
	static String[] OWNER_FIELDS = {"id", "firstName", "lastName", "address", "city", "telephone"};
	static String[] PET_FIELDS = {"name", "birth_date", "type"};
	// ALL DATA DOWNLOAD
	public static ByteArrayInputStream dataToExcel() {
		try (Workbook workbook = new XSSFWorkbook(); ByteArrayOutputStream out = new ByteArrayOutputStream();) {

			// SHEET
			Sheet sheet = workbook.createSheet("dataBaseInstance");
			// freeze header rows
			sheet.createFreezePane(0, 2);

			// HEADER
			// retrieve field names of each entity
			List<String[]> allFields = getFields();
			headerFormatter(sheet, allFields);

			// TODO: Data
			dataFormatter(sheet);

			// add filter
//			sheet.setAutoFilter(new CellRangeAddress(1, 1, 0, columnLength));

			workbook.write(out);
			return new ByteArrayInputStream(out.toByteArray());
		}
		catch (IOException | ClassNotFoundException e) {
			throw new RuntimeException("fail to import data to Excel file: " + e.getMessage());
		}
	}

	public static List<String[]> getFields() {
		// get list consisting of list of field names
		List<String[]> attributes = new ArrayList<>();
		attributes.add(OWNER_FIELDS);
		attributes.add(PET_FIELDS);

		return attributes;
	}

	public static void headerFormatter(Sheet sheet, List<String[]> allFields) {

		Row tableRow = sheet.createRow(0);
		Row attributeRow = sheet.createRow(1);
		int firstCol, lastCol = 0;
		int attributeIdx = 0;
		for (String[] fields : allFields) {
			// name attribute cells
			firstCol = lastCol;
			for(String f : fields){
				Cell cell = attributeRow.createCell(lastCol++);
				cell.setCellValue(f);
			}
			// name and merge table header cells
			Cell cell = tableRow.createCell(firstCol);
			cell.setCellValue(ENTITY_NAMES[attributeIdx++]);
			sheet.addMergedRegion(new CellRangeAddress(0,0, firstCol, lastCol-1));
		}
	}

	public static void dataFormatter(Sheet sheet){

	}
}
