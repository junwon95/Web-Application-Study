package org.springframework.samples.petclinic.admin;

import org.apache.poi.ss.usermodel.*;
import org.springframework.samples.petclinic.owner.Pet;

import java.util.List;

public class PetExcelManager extends ExcelManager {

	private static final String[] FIELDS = { "id", "name", "birthDate", "type" };

	@Override
	public void makeSheets(List<?> table) {
		makeSheet(table, 0);
	}

	@Override
	public void writeData(Sheet sheet, List<?> table) {
		final int FIELDS = 4;
		int rowIdx = START_ROW + 2;
		for (Object t : table) {
			Pet pet = (Pet) t;

			int colIdx = START_COL;
			Row row = sheet.createRow(rowIdx++);

			Cell cell = row.createCell(colIdx++);
			cell.setCellValue(pet.getId());
			cell.setCellStyle(style("DATA_LEFT"));

			cell = row.createCell(colIdx++);
			cell.setCellValue(pet.getName());
			cell.setCellStyle(style("DATA"));

			cell = row.createCell(colIdx++);
			cell.setCellValue(pet.getBirthDate().toString());
			cell.setCellStyle(style("DATA"));

			cell = row.createCell(colIdx);
			cell.setCellValue(pet.getType().toString());
			cell.setCellStyle(style("DATA_RIGHT"));
		}
		for (int i = START_COL; i < START_COL + FIELDS; i++)
			sheet.autoSizeColumn(i);
	}

	@Override
	public String[] getFields() {
		return FIELDS;
	}

	@Override
	public CellStyle style(String type) {
		CellStyle style = super.style(type);

		return style;
	}

}
