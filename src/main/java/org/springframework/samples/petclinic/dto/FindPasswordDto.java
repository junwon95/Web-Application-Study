package org.springframework.samples.petclinic.dto;

import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class FindPasswordDto {

	private String username;

	private String email;

}
