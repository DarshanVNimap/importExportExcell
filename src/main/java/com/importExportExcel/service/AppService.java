package com.importExportExcel.service;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.util.List;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import com.importExportExcel.helper.ExcellHelper;
import com.importExportExcel.model.OfficeSupplySale;
import com.importExportExcel.repository.OfficeSaleRepository;

@Service
public class AppService {

	@Autowired
	private OfficeSaleRepository repository;
	
	public void importFile(MultipartFile file) throws IOException {
		List<OfficeSupplySale> sales = ExcellHelper.extractDataToList(file.getInputStream());
		repository.saveAll(sales);
	}
	
	public ByteArrayInputStream exportToExcel(String fileName) throws IOException {
		return ExcellHelper.exportTOExcel(getList(), fileName);
	}
	
	
	public List<OfficeSupplySale> getList(){
		return repository.findAll();
	}
	
}
