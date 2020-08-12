package org.springframework.samples.petclinic.admin;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

// abstract class ex
public abstract class ExcelDownloader {

	public Workbook createWorkbook() {
		Workbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = sheet();
		return workbook;
	}

	public abstract XSSFSheet sheet();

}
