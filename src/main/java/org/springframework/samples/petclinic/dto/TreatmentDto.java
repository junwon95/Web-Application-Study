package org.springframework.samples.petclinic.dto;

import org.springframework.samples.petclinic.vet.Vet;

import java.util.Collection;

public class TreatmentDto {

	private String description;

	private String prescription;

	private int visitId;

	private int vetId;

	public String getDescription() {
		return description;
	}

	public void setDescription(String description) {
		this.description = description;
	}

	public String getPrescription() {
		return prescription;
	}

	public void setPrescription(String prescription) {
		this.prescription = prescription;
	}

	public int getVisitId() {
		return visitId;
	}

	public void setVisitId(int visitId) {
		this.visitId = visitId;
	}

	public int getVetId() {
		return vetId;
	}

	public void setVetId(int vetId) {
		this.vetId = vetId;
	}

}
