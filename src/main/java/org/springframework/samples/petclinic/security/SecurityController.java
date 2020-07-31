package org.springframework.samples.petclinic.security;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.mail.SimpleMailMessage;
import org.springframework.mail.javamail.JavaMailSender;
import org.springframework.mail.javamail.JavaMailSenderImpl;
import org.springframework.samples.petclinic.dto.ChangePasswordDto;
import org.springframework.samples.petclinic.dto.FindPasswordDto;
import org.springframework.samples.petclinic.dto.VerificationDto;
import org.springframework.security.core.context.SecurityContextHolder;
import org.springframework.security.core.userdetails.UserDetails;
import org.springframework.security.crypto.bcrypt.BCryptPasswordEncoder;
import org.springframework.security.crypto.password.PasswordEncoder;
import org.springframework.stereotype.Controller;
import org.springframework.util.StringUtils;
import org.springframework.validation.BindingResult;
import org.springframework.web.bind.WebDataBinder;
import org.springframework.web.bind.annotation.*;

import javax.validation.Valid;
import java.util.Map;

@Controller
public class SecurityController {

	@Autowired
	UserService userService;

	@Autowired
	JavaMailSenderImpl mailSender;

	private final PasswordEncoder passwordEncoder;

	private final MemberRepository memberRepository;

	SecurityController(MemberRepository memberRepository, PasswordEncoder passwordEncoder) {
		this.memberRepository = memberRepository;
		this.passwordEncoder = passwordEncoder;
	}

	@InitBinder("member")
	public void initMemberBinder(WebDataBinder dataBinder) {
		dataBinder.setValidator(new MemberValidator());
	}

	@InitBinder("passwordDto")
	public void initPasswordBinder(WebDataBinder dataBinder) {
		dataBinder.setValidator(new PasswordValidator());
	}

	@GetMapping("/signin")
	public String loginPage(Map<String, Object> model) {
		return "security/login";
	}

	@GetMapping("/register")
	public String initRegisterForm(Map<String, Object> model) {
		model.put("member", new Member());
		return "security/register";
	}

	@PostMapping("/register")
	public String postRegisterForm(Map<String, Object> model, @Valid Member member, BindingResult result) {
		if (memberRepository.findByUsername(member.getUsername()) != null) {
			result.rejectValue("username", "duplicate", "already exists");
		}
		if (memberRepository.findByEmail(member.getEmail()) != null) {
			result.rejectValue("email", "duplicate", "already exists");
		}

		if (result.hasErrors()) {
			return "security/register";
		}
		else {
			member.setPassword(passwordEncoder.encode(member.getPassword()));
			EmailVerificationToken emailVerificationToken = new EmailVerificationToken(member);

			SimpleMailMessage email = new SimpleMailMessage();
			email.setTo(member.getEmail());
			email.setSubject("Verification Code");
			email.setText("code: " + emailVerificationToken.getValue());
			mailSender.send(email);

			VerificationDto verificationDto = new VerificationDto();
			verificationDto.setMember(member);
			model.put("verificationDto", verificationDto);

			return "security/verify";
		}
	}

	@PostMapping("/verify")
	public String emailVerifyForm(Map<String, Object> model, VerificationDto verificationDto, BindingResult result) {
		BCryptPasswordEncoder encoder = new BCryptPasswordEncoder();
		String verificationCode = verificationDto.getVerificationCode();
		Member member = verificationDto.getMember();

		if (encoder.matches(member.getEmail(), verificationCode)) {
			memberRepository.save(member);
			return "redirect:";
		}
		else {
			return "redirect:/verificationFailed";
		}
	}

	@GetMapping("/verificationFailed")
	public String verificationFailedForm(Map<String, Object> model) {
		model.put("message", "Verification Failed");
		return "/error";
	}

	@GetMapping("/findPassword")
	public String findPasswordForm(Map<String, Object> model) {
		model.put("findPasswordDto", new FindPasswordDto());
		return "security/findPassword";
	}

	@PostMapping("/findPassword")
	public String postFindPasswordForm(Map<String, Object> model, FindPasswordDto findPasswordDto,
			BindingResult result) {
		Member member = memberRepository.findByUsername(findPasswordDto.getUsername());
		if (member == null) {
			result.rejectValue("username", "notFound", "not found");
			result.rejectValue("email", "notVerified", "not verified");
		}
		else if (!member.getEmail().equals(findPasswordDto.getEmail())) {
			result.rejectValue("email", "notVerified", "not verified");
		}
		if (result.hasErrors()) {
			return "security/findPassword";
		}

		BCryptPasswordEncoder bCryptPasswordEncoder = new BCryptPasswordEncoder();
		String newPassword = bCryptPasswordEncoder.encode("0000");
		memberRepository.changePassword(newPassword, member.getId());

		SimpleMailMessage email = new SimpleMailMessage();
		email.setTo(member.getEmail());
		email.setSubject("Password Changed");
		email.setText("password is set to: 0000");
		mailSender.send(email);

		return "security/passwordSent";
	}

	@GetMapping("/accessDenied")
	public String accessDeniedForm(Map<String, Object> model) {
		model.put("message", "Access Denied");
		return "/error";
	}

	@GetMapping("/profile")
	public String profileForm(Map<String, Object> model) {
		Object principal = SecurityContextHolder.getContext().getAuthentication().getPrincipal();

		// petclinic member
		if (principal instanceof UserDetails) {
			String username = ((UserDetails) principal).getUsername();
			Member member = memberRepository.findByUsername(username);
			model.put("member", member);
			return "security/profile";
		}
		// OAUTH member
		else {
			// TODO: link OAUTH with Members
			// redirect to custon
			return "redirect:";
		}

	}

	@GetMapping("/changePassword")
	public String changePasswordForm(Map<String, Object> model, String username) {
		ChangePasswordDto changePasswordDto = new ChangePasswordDto();
		changePasswordDto.setUsername(username);

		model.put("passwordDto", changePasswordDto);
		return "security/changePassword";
	}

	@PostMapping("/changePassword")
	public String postChangePasswordForm(Map<String, Object> model,
			@Valid @ModelAttribute("passwordDto") ChangePasswordDto passwordDto, BindingResult result) {
		String username = passwordDto.getUsername();
		String password = passwordDto.getPassword();
		String newPassword = passwordDto.getNewPassword();

		Member member = memberRepository.findByUsername(username);
		BCryptPasswordEncoder bCryptPasswordEncoder = new BCryptPasswordEncoder();

		if (StringUtils.hasLength(password) && !bCryptPasswordEncoder.matches(password, member.getPassword())) {
			result.rejectValue("password", "notVerified", "wrong password");
		}

		if (result.hasErrors()) {
			return "security/changePassword";
		}

		memberRepository.changePassword(bCryptPasswordEncoder.encode(newPassword), member.getId());
		model.put("member", member);
		return "security/profile";
	}

}
