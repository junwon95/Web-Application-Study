package org.springframework.samples.petclinic.admin;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.samples.petclinic.owner.Owner;
import org.springframework.samples.petclinic.owner.OwnerRepository;
import org.springframework.samples.petclinic.owner.Pet;
import org.springframework.samples.petclinic.owner.PetRepository;
import org.springframework.samples.petclinic.visit.VisitRepository;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;
import org.springframework.web.multipart.MultipartFile;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.util.Collections;
import java.util.List;

@Service
public class ExcelService {

	@Autowired
	OwnerRepository ownerRepository;
	@Autowired
	VisitRepository visitRepository;

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
		for(Owner owner : owners){
			for(Pet pet : owner.getPets()){
				pet.setVisitsInternal(visitRepository.findByPetId(pet.getId()));
			}
		}
		ByteArrayInputStream in = ExcelManager2.dataToExcel(owners);

		return in;
	}


}
