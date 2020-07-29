package org.springframework.samples.petclinic.visit;

import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;
import org.springframework.data.repository.query.Param;
import org.springframework.transaction.annotation.Transactional;

public interface TreatmentRepository extends JpaRepository<Treatment, Integer> {

	// void save(Treatment treatment);
	//
	// @Query("SELECT t FROM Treatment t JOIN Visit v WHERE t.visitId = :visit_id")
	// @Transactional(readOnly = true)
	// Treatment findByVisitId(@Param("visit_id") int visit_id);

	Treatment findByVisitId(Integer visitId);

}
