package org.springframework.samples.petclinic.admin;

import org.apache.poi.ss.usermodel.*;
import org.springframework.samples.petclinic.owner.Owner;

import java.util.List;

public class OwnerExcelManager extends ExcelManager {

	private static final String[] FIELDS = { "id", "firstName", "lastName", "Address", "city", "telephone" };

	@Override
	public void makeSheets(List<?> table) {
		makeSheet(table, 0);
	}

	@Override
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
