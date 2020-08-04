package org.springframework.samples.petclinic.admin;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.samples.petclinic.owner.Owner;
import org.springframework.web.multipart.MultipartFile;

public class ExcelManager {
	public static String TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

	// attributes of Owner relation
	static String[] HEADERs = { "firstName", "lastName", "Address", "City", "Telephone" };
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

			// Header
			Row headerRow = sheet.createRow(0);

			for (int col = 0; col < HEADERs.length; col++) {
				Cell cell = headerRow.createCell(col);
				cell.setCellValue(HEADERs[col]);
			}

			int rowIdx = 1;
			for (Owner owner : owners) {
				Row row = sheet.createRow(rowIdx++);

				row.createCell(0).setCellValue(owner.getId());
				row.createCell(1).setCellValue(owner.getFirstName());
				row.createCell(2).setCellValue(owner.getLastName());
				row.createCell(3).setCellValue(owner.getAddress());
				row.createCell(4).setCellValue(owner.getCity());
				row.createCell(5).setCellValue(owner.getTelephone());
			}

			workbook.write(out);
			return new ByteArrayInputStream(out.toByteArray());
		} catch (IOException e) {
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
		} catch (IOException e) {
			throw new RuntimeException("fail to parse Excel file: " + e.getMessage());
		}
	}
}
