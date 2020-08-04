package org.springframework.samples.petclinic.admin;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.InputStreamResource;
import org.springframework.core.io.Resource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.samples.petclinic.owner.OwnerRepository;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.multipart.MultipartFile;

import java.util.Map;

@Controller
class AdminController {

	@Autowired
	ExcelService excelService;

	@GetMapping("/admin/administer")
	public String showAdministrationData(Map<String, Object> model) {
		model.put("owners", excelService.ownerRepository.findAll());
		return "admin/administer";
	}

	@GetMapping("/admin/download")
	public ResponseEntity<Resource> getFile() {
		String filename = "ownerData.xlsx";
		InputStreamResource file = new InputStreamResource(excelService.load());

		return ResponseEntity.ok()
			.header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=" + filename)
			.contentType(MediaType.parseMediaType("application/vnd.ms-excel"))
			.body(file);
	}

	// TODO: PASS excel file as param
	@PostMapping("/admin/upload")
	public String uploadFile(@RequestParam("file") MultipartFile file) {
		String message = "";

		if (ExcelManager.hasExcelFormat(file)) {
			excelService.save(file);
			return "admin/administer";
		}
		// no file err
		return "";
	}


}
