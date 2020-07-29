package org.springframework.samples.petclinic.visit;

import org.springframework.samples.petclinic.dto.TreatmentDto;
import org.springframework.samples.petclinic.vet.VetRepository;
import org.springframework.samples.petclinic.vet.Vet;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.PostMapping;

import java.util.Map;

@Controller
class TreatmentController {

	private final TreatmentRepository treatmentRepository;

	private final VisitRepository visitRepository;

	private final VetRepository vetRepository;

	TreatmentController(TreatmentRepository treatmentRepository, VisitRepository visitRepository,
			VetRepository vetRepository) {
		this.treatmentRepository = treatmentRepository;
		this.visitRepository = visitRepository;
		this.vetRepository = vetRepository;
	}

	@GetMapping("/owners/{ownerId}/visits/{visitId}")
	public String initProcessTreatmentForm(@PathVariable int ownerId, @PathVariable int visitId,
			Map<String, Object> model) {
		model.put("ownerId", ownerId);
		model.put("visitId", visitId);

		Treatment treatment = treatmentRepository.findByVisitId(visitId);

		// Treatment treatment = treatmentRepository.findByVisitId(visitId);
		if (treatment == null) { // check to see if method works ***************

			TreatmentDto treatmentDto = new TreatmentDto();
			model.put("treatment", treatmentDto);

			// put collection of all vets to model
			model.put("vets", vetRepository.findAll());

			// create new Treatment and redirect to input HTML page
			return "pets/createOrUpdateTreatmentForm";
		}
		Vet vet = treatment.getVet();
		model.put("treatment", treatment);
		model.put("Fname", vet.getFirstName());
		model.put("Lname", vet.getLastName());

		return "pets/TreatmentDetails";
	}

	@PostMapping("/treatment/{ownerId}/{visitId}")
	public String postProcessTreatmentForm(@PathVariable int ownerId, @PathVariable int visitId,
			TreatmentDto treatmentDto, Map<String, Object> model) {
		Treatment treatment = new Treatment();
		treatment.setVisit(visitRepository.getOne(visitId));
		treatment.setDescription(treatmentDto.getDescription());
		treatment.setPrescription(treatmentDto.getPrescription());
		treatment.setVet(vetRepository.findById(treatmentDto.getVetId()));

		treatmentRepository.save(treatment);
		return "redirect:/owners/" + ownerId;

	}

}
