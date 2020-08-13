package org.springframework.samples.petclinic.admin;

import org.apache.poi.ss.usermodel.*;
import org.springframework.samples.petclinic.owner.Owner;
import org.springframework.samples.petclinic.owner.Pet;

import java.util.List;

public class OwnerExcelManager extends ExcelManager {
<<<<<<< HEAD

	private static final String[] FIELDS = { "id", "firstName", "lastName", "Address", "city", "telephone" };
=======
>>>>>>> da2b995c2562f1c23f6300846db9542b6a0797fb

	@Override
	public void makeSheets(List<?> table, int referenceId) {
		makeSheet(table, referenceId);

		for (Object entity : table) {
			Owner owner = (Owner) entity;
			if (owner.getPets().isEmpty())
				continue;
			List<Pet> pets = owner.getPets();
			PetExcelManager petExcelManager = new PetExcelManager(10, 10);
			petExcelManager.makeSheets(pets, owner.getId());
		}
	}

	@Override
<<<<<<< HEAD
	public void writeData(Sheet sheet, List<?> table) {
		final int FIELDS = 6;
		int rowIdx = START_ROW + 2;
		for (Object t : table) {
			Owner owner = (Owner) t;

			int colIdx = START_COL;
			Row row = sheet.createRow(rowIdx++);

			Cell cell = row.createCell(colIdx++);
			cell.setCellValue(owner.getId());
			cell.setCellStyle(style("DATA_LEFT"));

			cell = row.createCell(colIdx++);
			cell.setCellValue(owner.getFirstName());
			cell.setCellStyle(style("DATA"));

			cell = row.createCell(colIdx++);
			cell.setCellValue(owner.getLastName());
			cell.setCellStyle(style("DATA"));

			cell = row.createCell(colIdx++);
			cell.setCellValue(owner.getAddress());
			cell.setCellStyle(style("DATA"));

			cell = row.createCell(colIdx++);
			cell.setCellValue(owner.getCity());
			cell.setCellStyle(style("DATA"));

			cell = row.createCell(colIdx);
			cell.setCellValue(owner.getTelephone());
			cell.setCellStyle(style("DATA_RIGHT"));
		}
		for (int i = START_COL; i < START_COL + FIELDS; i++)
			sheet.autoSizeColumn(i);
=======
	public <T> String[] getFieldValues(T entity) {
		String fieldValues[] = new String[6];
		Owner owner = (Owner) entity;

		fieldValues[0] = owner.getId().toString();
		fieldValues[1] = owner.getFirstName();
		fieldValues[2] = owner.getLastName();
		fieldValues[3] = owner.getAddress();
		fieldValues[4] = owner.getCity();
		fieldValues[5] = owner.getTelephone();

		return fieldValues;
>>>>>>> da2b995c2562f1c23f6300846db9542b6a0797fb
	}

	@Override
	public String[] getFields() {
		return new String[] { "id", "firstName", "lastName", "Address", "city", "telephone" };
	}

	@Override
	public CellStyle style(String type) {
		CellStyle style = super.style(type);

		return style;
	}

<<<<<<< HEAD
=======
	@Override
	public <T> void setHyperCell(Cell cell, T entity) {
		Owner owner = (Owner) entity;
		if (owner.getPets().isEmpty())
			return;
		setCell(cell, "see pets", style("LINK"));
		setHyperLink(cell, "Pets" + owner.getId());
	}

>>>>>>> da2b995c2562f1c23f6300846db9542b6a0797fb
}
