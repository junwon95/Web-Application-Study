package org.springframework.samples.petclinic.admin;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.samples.petclinic.admin.ExcelUtilTools.ExcelManager2;
import org.springframework.samples.petclinic.dto.ExcelDto;
import org.springframework.samples.petclinic.owner.Owner;
import org.springframework.samples.petclinic.owner.OwnerRepository;
import org.springframework.samples.petclinic.owner.Pet;
import org.springframework.samples.petclinic.owner.PetRepository;
import org.springframework.samples.petclinic.visit.Treatment;
import org.springframework.samples.petclinic.visit.TreatmentRepository;
import org.springframework.samples.petclinic.visit.Visit;
import org.springframework.samples.petclinic.visit.VisitRepository;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.beans.IntrospectionException;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.util.ArrayList;
import java.util.List;

@Service
public class ExcelService {

	@Autowired
	OwnerRepository ownerRepository;

	@Autowired
	PetRepository petRepository;

	@Autowired
	VisitRepository visitRepository;

	@Autowired
	TreatmentRepository treatmentRepository;

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

	public ByteArrayInputStream load() throws IntrospectionException, InvocationTargetException {
		// Owner ver
		List<Owner> owners = ownerRepository.findAll();
		for (Owner owner : owners) {
			for (Pet pet : owner.getPets()) {
				pet.setVisitsInternal(visitRepository.findVisitsByPetId(pet.getId()));
				for (Visit visit : pet.getVisits()) {
					List<Treatment> treatments = new ArrayList<>();
					treatments.add(treatmentRepository.findByVisitId(visit.getId()));
					visit.setTreatments(treatments);
				}
			}
		}
		OwnerExcelManager ownerExcelManager = new OwnerExcelManager();
		ByteArrayInputStream in = ownerExcelManager.dataToExcel(owners);

		// recursive ver
		// ByteArrayInputStream in = ExcelManager2.dataToExcel(owners);

		// Pet ver
//		List<Pet> pets = petRepository.findAll();
//		for (Pet pet : pets) {
//			pet.setVisitsInternal(visitRepository.findVisitsByPetId(pet.getId()));
//			for (Visit visit : pet.getVisits()) {
//				List<Treatment> treatments = new ArrayList<>();
//				treatments.add(treatmentRepository.findByVisitId(visit.getId()));
//				visit.setTreatments(treatments);
//			}
//		}
//		PetExcelManager petExcelManager = new PetExcelManager();
//		ByteArrayInputStream in = petExcelManager.dataToExcel(pets);

		return in;
	}

}
