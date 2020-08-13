package org.springframework.samples.petclinic.admin;

import org.apache.poi.ss.usermodel.*;
import org.springframework.samples.petclinic.owner.Pet;

import java.util.List;

public class PetExcelManager extends ExcelManager {

	public PetExcelManager() {
	}

	public PetExcelManager(int row, int col) {
		super.START_ROW += row;
		super.START_COL += col;
	}

	@Override
	public void makeSheets(List<?> table, int referenceId) {
		makeSheet(table, referenceId);
	}

	@Override
	public <T> String[] getFieldValues(T entity) {
		String[] fieldValues = new String[4];
		Pet pet = (Pet) entity;

		fieldValues[0] = pet.getId().toString();
		fieldValues[1] = pet.getName();
		fieldValues[2] = pet.getBirthDate().toString();
		fieldValues[3] = pet.getType().toString();

		return fieldValues;
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
