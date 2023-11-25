package com.importExportExcel.controller;

import java.io.IOException;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.InputStreamResource;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import com.importExportExcel.helper.ExcellHelper;
import com.importExportExcel.service.AppService;

@RestController
@RequestMapping("/api")
public class AppController {
	
	@Autowired
	private AppService service;
	
	@PostMapping("/import")
	public ResponseEntity<?> importExcelSheet(@RequestParam("file") MultipartFile file) throws IOException{
	if(ExcellHelper.checkUploadedFileType(file)) {
		service.importFile(file);
	}
		return ResponseEntity
					.status(HttpStatus.OK)
					.body("Imported successfully");
	}

	@GetMapping("/export/{name}")
	public ResponseEntity<?> exportExcelSheet(@PathVariable("name") String name) throws IOException{		

		return ResponseEntity
						.status(HttpStatus.OK)
						.header("Content-Disposition", "attachment; filename="+name)
						.contentType(MediaType.parseMediaType("application/vnd.ms-excel"))
						.body(new InputStreamResource(service.exportToExcel(name)));
	}
	
	@GetMapping("/get")
	public ResponseEntity<?> getSaleReport(){
		return ResponseEntity.
						status(HttpStatus.OK)
						.body(service.getList());
	}
}
