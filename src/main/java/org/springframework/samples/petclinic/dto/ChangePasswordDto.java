package org.springframework.samples.petclinic.dto;

import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class ChangePasswordDto {
	String username;
	String password;
	String newPassword;
}
