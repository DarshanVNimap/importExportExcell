package com.importExportExcel.model;

import java.util.Date;

import jakarta.persistence.Entity;
import jakarta.persistence.Id;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

@Entity
@Data
@NoArgsConstructor
@AllArgsConstructor(staticName = "build")
public class OfficeSupplySale {
	
	@Id
	private Long id;
	private Date date;
	private String salesRep ; 
	private String shift;
	private Double costPrice;
	private	Double sellPrice;

	
}
