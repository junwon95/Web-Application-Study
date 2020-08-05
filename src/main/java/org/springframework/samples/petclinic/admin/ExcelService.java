package org.springframework.samples.petclinic.admin;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.samples.petclinic.owner.Owner;
import org.springframework.samples.petclinic.owner.OwnerRepository;
import org.springframework.samples.petclinic.owner.PetRepository;
import org.springframework.samples.petclinic.visit.VisitRepository;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.util.List;

@Service
public class ExcelService {

	@Autowired
	OwnerRepository ownerRepository;

	public void save(MultipartFile file) {
		try {
			List<Owner> owners = ExcelManager.excelToOwners(file.getInputStream());
			ownerRepository.saveAll(owners);
		}
		catch (IOException e) {
			throw new RuntimeException("fail to store excel data: " + e.getMessage());
		}
	}

	public ByteArrayInputStream load() {
		List<Owner> owners = ownerRepository.findAll();

//		ByteArrayInputStream in = ExcelManager.ownersToExcel(owners);
		ByteArrayInputStream in = ExcelManager.dataToExcel();

		return in;
	}

	public List<Owner> getAllOwners() {
		return ownerRepository.findAll();
	}

}
