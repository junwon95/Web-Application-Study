package org.springframework.samples.petclinic.dto;

import java.util.ArrayList;
import java.util.List;

public class PetDto {

	private Integer id;

	private String name;

	private List<VisitDto> visits = new ArrayList<>();

	public List<VisitDto> getVisits() {
		return visits;
	}

	public void setVisits(List<VisitDto> visits) {
		this.visits = visits;
	}

}
