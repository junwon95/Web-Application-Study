package org.springframework.samples.petclinic.admin;

import org.apache.poi.ss.usermodel.*;
import org.springframework.samples.petclinic.owner.Owner;
import org.springframework.samples.petclinic.owner.Pet;

import java.util.List;

public class OwnerExcelManager extends ExcelManager {

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

	@Override
	public <T> void setHyperCell(Cell cell, T entity) {
		Owner owner = (Owner) entity;
		if (owner.getPets().isEmpty())
			return;
		setCell(cell, "see pets", style("LINK"));
		setHyperLink(cell, "Pets" + owner.getId(), excelCoordinates());
	}

}
