package com.marcoscheuerer.read;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class UpdateExcel {

	public void readAndUpdate(String filePath) throws IOException {
		/*
		 * 1. Get workbook from inputStream
		 * 2. Get sheet
		 * 3. Get cell number for amount
		 * 4. Get actual amount and add + 10
		 * 5. create a new cell at the end of row
		 * 6. write the updated amount in new cell
		 * 7. write object in output stream
		 * 8. close stream
		 */
		
		// get the workbook
		File file = new File(filePath);
		
		FileInputStream fis = new FileInputStream(file);
		
		Workbook workbook = WorkbookFactory.create(fis);
		// get the first sheet
		Sheet sheet = workbook.getSheetAt(0);
		// get the header row
		Row headerRow = sheet.getRow(0);
		// iterate over the header
		Iterator<Cell> cellIterator = headerRow.cellIterator();
		int amountCellIndex = 0;
		boolean isAmountCellFound = false;
		
		while (cellIterator.hasNext()) {
			Cell cell = cellIterator.next();
			
			if (cell.getStringCellValue().equalsIgnoreCase("amount")) {
				amountCellIndex = cell.getColumnIndex();
				isAmountCellFound = true;
				break;
			}
			
		}
		
		if (!isAmountCellFound) {
			System.out.println("amount cell not found");
			return;
		} else {
			System.out.println("Cell: " + amountCellIndex);
			
			int updatedAmountCellIndex = headerRow.getLastCellNum();
			
			headerRow.createCell(updatedAmountCellIndex).setCellValue("Updated Amount");
			// writing the cell value
			Iterator<Row> rowIterator = sheet.rowIterator();
			
			while(rowIterator.hasNext()) {
				Row row = rowIterator.next();
				if (row.getRowNum() == 0) {
					continue;
				} else {
					Cell cell = row.getCell(amountCellIndex);
					
					if (cell != null) {
						double amount = cell.getNumericCellValue();
						Cell updateCell = row.createCell(updatedAmountCellIndex);
						updateCell.setCellValue(amount + 10);
					}
					
				}
			}
		}
		
		// write workbook on output stream
		FileOutputStream fos = new FileOutputStream(file);
		
		workbook.write(fos);
	
		workbook.close();
		fis.close();
		fos.close();
		
//		FileOutputStream fos = new FileOutputStream(file);
	}
	
}
