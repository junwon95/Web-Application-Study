package org.springframework.samples.petclinic.dto;

import lombok.Getter;
import lombok.Setter;
import org.springframework.samples.petclinic.owner.Owner;
import org.springframework.samples.petclinic.owner.Pet;
import org.springframework.samples.petclinic.visit.Visit;

import java.util.List;

@Getter
@Setter
public class ExcelDto {

	private List<Owner> owners;

	private List<Pet> pets;

	private List<Visit> visits;

}
