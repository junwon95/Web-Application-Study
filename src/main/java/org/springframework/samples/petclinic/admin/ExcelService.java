package org.springframework.samples.petclinic.admin;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.samples.petclinic.dto.ExcelDto;
import org.springframework.samples.petclinic.owner.Owner;
import org.springframework.samples.petclinic.owner.OwnerRepository;
import org.springframework.samples.petclinic.owner.Pet;
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

	@Autowired
	PetRepository petRepository;

	@Autowired
	VisitRepository visitRepository;

	public void save(MultipartFile file) {
		try {
			ExcelDto dto = ExcelManager2.excelToDB(file.getInputStream());
			ownerRepository.saveAll(dto.getOwners());
			petRepository.saveAll(dto.getPets());
			visitRepository.saveAll(dto.getVisits());
		}
		catch (IOException e) {
			throw new RuntimeException("fail to store excel data: " + e.getMessage());
		}
	}

	public ByteArrayInputStream load() {
		List<Owner> owners = ownerRepository.findAll();
		for (Owner owner : owners) {
			for (Pet pet : owner.getPets()) {
				pet.setVisitsInternal(visitRepository.findByPetId(pet.getId()));
			}
		}
		ByteArrayInputStream in = ExcelManager2.dataToExcel(owners);

		return in;
	}

}
