package org.springframework.samples.petclinic.security;

import lombok.AllArgsConstructor;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;
import org.springframework.mail.javamail.JavaMailSender;
import org.springframework.security.config.annotation.authentication.builders.AuthenticationManagerBuilder;
import org.springframework.security.config.annotation.web.builders.HttpSecurity;
import org.springframework.security.config.annotation.web.configuration.EnableWebSecurity;
import org.springframework.security.config.annotation.web.configuration.WebSecurityConfigurerAdapter;
import org.springframework.security.crypto.bcrypt.BCryptPasswordEncoder;
import org.springframework.security.crypto.password.PasswordEncoder;
import org.springframework.security.web.access.AccessDeniedHandler;

@Configuration
@EnableWebSecurity
@AllArgsConstructor
public class SecurityConfig extends WebSecurityConfigurerAdapter {

	@Bean
	public PasswordEncoder passwordEncoder() {
		return new BCryptPasswordEncoder();
	}

	@Override
	protected void configure(HttpSecurity http) throws Exception {
		// https configuration for deployment on Heroku
		http.requiresChannel().requestMatchers(r -> r.getHeader("X-Forwarded-Proto") != null).requiresSecure();
		// authorization settings
		http
				// use basic authorizations
				.httpBasic().and().authorizeRequests()
				// permit resources
				.antMatchers("/resources/**", "/webjars/**").permitAll()
				// permit pages (main page, search, see list of vets)
				.antMatchers("/*", "/owners/find", "/findPassword").permitAll()
				// set authentication corresponding to role ADMIN
				.antMatchers("/owners/*/edit", "/owners/*/delete").hasRole("ADMIN")
				// all other requests are authenticated
				.anyRequest().authenticated().and()
				// form login config
				.formLogin().loginPage("/signin").permitAll().and().exceptionHandling()
				.accessDeniedHandler(new CustomAccessDeniedHandler()).and().logout().logoutUrl("/logout")
				.invalidateHttpSession(true).and()
				// oauth2 logout config
				.oauth2Login().and().exceptionHandling().accessDeniedHandler(new CustomAccessDeniedHandler());
		// Only for development
		// http.csrf().disable();

	}

	@Override
	protected void configure(AuthenticationManagerBuilder auth) throws Exception {
		auth.userDetailsService(myUserService()).passwordEncoder(passwordEncoder());
	}

	@Bean
	public UserService myUserService() {
		return new UserService();
	}

	@Bean
	public AccessDeniedHandler accessDeniedHandler() {
		return new CustomAccessDeniedHandler();
	}

}
