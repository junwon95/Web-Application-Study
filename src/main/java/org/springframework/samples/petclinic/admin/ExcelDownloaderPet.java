package org.springframework.samples.petclinic.admin;

import org.apache.poi.xssf.usermodel.XSSFSheet;

public class ExcelDownloaderPet extends ExcelDownloader {

	@Override
	public XSSFSheet sheet() {
		return null;
	}

}
