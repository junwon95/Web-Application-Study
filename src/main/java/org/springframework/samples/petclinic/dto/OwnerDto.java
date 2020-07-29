package org.springframework.samples.petclinic.dto;

import java.util.ArrayList;
import java.util.List;

public class OwnerDto {

	private Integer id;

	private String firstName, lastName, address, city, telephone;

	public Integer getId() {
		return id;
	}

	public void setId(Integer id) {
		this.id = id;
	}

	public String getFirstName() {
		return firstName;
	}

	public void setFirstName(String firstName) {
		this.firstName = firstName;
	}

	public String getLastName() {
		return lastName;
	}

	public void setLastName(String lastName) {
		this.lastName = lastName;
	}

	public String getAddress() {
		return address;
	}

	public void setAddress(String address) {
		this.address = address;
	}

	public String getCity() {
		return city;
	}

	public void setCity(String city) {
		this.city = city;
	}

	public String getTelephone() {
		return telephone;
	}

	public void setTelephone(String telephone) {
		this.telephone = telephone;
	}

	private List<PetDto> pets = new ArrayList<>();

	public List<PetDto> getPets() {
		return pets;
	}

	public void setPets(List<PetDto> pets) {
		this.pets = pets;
	}

}
