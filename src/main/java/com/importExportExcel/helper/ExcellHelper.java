package com.importExportExcel.helper;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.multipart.MultipartFile;

import com.importExportExcel.model.OfficeSupplySale;

import lombok.extern.slf4j.Slf4j;

@Slf4j
public class ExcellHelper {
	
	static String[] SHEET_HEADING = {
		"Id" , "Saler" , "Shift" , "Cost Price" , "Sale Price" , "Date"
	};


	// check file type

	public static boolean checkUploadedFileType(MultipartFile file) {

		return file.getContentType().equals("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
	}

	// Convert sheet to list of data;

	public static List<OfficeSupplySale> extractDataToList(InputStream is) throws IOException {

		List<OfficeSupplySale> report = new ArrayList<>();

		try (XSSFWorkbook book = new XSSFWorkbook(is)) {
			//		log.info(book.getSheetName(0));
					XSSFSheet sheet = book.getSheet("Sheet1");
					Iterator<Row> rows = sheet.iterator();
			
					int rowNumber = 0;
					while (rows.hasNext()) {
						rows.next();
						if (rowNumber <= 1) {
							rowNumber++;
			
							log.info("Row number: " + rowNumber);
							continue;
						}
			
						Iterator<Cell> cells = rows.next().iterator();
						Integer cellNumber = 0;
			
						OfficeSupplySale saleReport = new OfficeSupplySale();
			
						saleReport.setId((long) rowNumber);
						while (cells.hasNext()) {
			
							Cell cell = cells.next();
			
							log.info("Row : " + rowNumber + " : " + cellNumber + " : " + cell.getCellType().toString());
							switch (cellNumber) {
							case 0:
								saleReport.setDate(new Date((long) cell.getNumericCellValue()));
								break;
							case 1:
								saleReport.setSalesRep(cell.getStringCellValue());
								break;
							case 2:
								saleReport.setShift(cell.getStringCellValue());
								break;
							case 3:
								saleReport.setCostPrice(cell.getNumericCellValue());
								break;
							case 4:
								saleReport.setSellPrice(cell.getNumericCellValue());
								break;
							}
							cellNumber++;
						}
			//			break;
						report.add(saleReport);
						rowNumber++;
			
					}
		}

return report;
	}

	
	// Export sheet
	
	public static ByteArrayInputStream exportTOExcel(List<OfficeSupplySale> list , String sheetName) throws IOException {
		
		XSSFWorkbook book = new XSSFWorkbook();
		ByteArrayOutputStream out = new ByteArrayOutputStream();
		
		XSSFSheet sheet =  book.createSheet(sheetName);
		
		 Row rowHeading =sheet.createRow(0);
		 
		 for(int i=0;i<SHEET_HEADING.length;i++) {
			 Cell cell = rowHeading.createCell(i);
			 cell.setCellValue(SHEET_HEADING[i]);
			 
		 }
		 
		 Integer rowNumber = 0;
			SimpleDateFormat dateFormat = new SimpleDateFormat("dd-MM-YYYY");
		 for(OfficeSupplySale report : list) {
			 
			 rowNumber++;
			 Row row = sheet.createRow(rowNumber);
			 
			 row.createCell(0).setCellValue(report.getId());
			 row.createCell(1).setCellValue(report.getSalesRep());
			 row.createCell(2).setCellValue(report.getShift());
			 row.createCell(3).setCellValue(report.getCostPrice());
			 row.createCell(4).setCellValue(report.getSellPrice());
			 row.createCell(5).setCellValue(dateFormat.format(report.getDate()));
			 
		 }
		 
		 book.write(out);
		book.close();
		out.close();
		return new ByteArrayInputStream(out.toByteArray());
		
	}
	
}
