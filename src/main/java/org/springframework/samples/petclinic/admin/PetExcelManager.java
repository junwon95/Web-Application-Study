package org.springframework.samples.petclinic.admin;

import org.apache.poi.ss.usermodel.*;
import org.springframework.samples.petclinic.owner.Pet;

import java.util.List;

public class PetExcelManager extends ExcelManager {

<<<<<<< HEAD
	private static final String[] FIELDS = { "id", "name", "birthDate", "type" };
=======
	public PetExcelManager() {
	}

	public PetExcelManager(int row, int col) {
		super.START_ROW += row;
		super.START_COL += col;
	}
>>>>>>> da2b995c2562f1c23f6300846db9542b6a0797fb

	@Override
	public void makeSheets(List<?> table, int referenceId) {
		makeSheet(table, referenceId);
	}

	@Override
<<<<<<< HEAD
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
=======
	public <T> String[] getFieldValues(T entity) {
		String[] fieldValues = new String[4];
		Pet pet = (Pet) entity;

		fieldValues[0] = pet.getId().toString();
		fieldValues[1] = pet.getName();
		fieldValues[2] = pet.getBirthDate().toString();
		fieldValues[3] = pet.getType().toString();

		return fieldValues;
>>>>>>> da2b995c2562f1c23f6300846db9542b6a0797fb
	}

	@Override
	public String[] getFields() {
		return new String[] { "id", "name", "birthDate", "type" };
	}

	@Override
	public <T> void setHyperCell(Cell cell, T entity) {
		Pet pet = (Pet) entity;
		if (pet.getVisits().isEmpty())
			return;
		setCell(cell, "see visits", style("LINK"));
		setHyperLink(cell, "Visits" + pet.getId());
	}

	@Override
	public CellStyle style(String type) {
		CellStyle style = super.style(type);

		return style;
	}

}
