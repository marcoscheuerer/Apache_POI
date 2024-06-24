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

public class ReadExcel {

	/*
	 * 1. Read data from single row
	 * 2. Read data from entire sheet
	 */
	
	public void readExcelDataForSingleRow(String filePath) throws IOException {
		
		/*
		 * 1. Get workbook object from input stream
		 * 2. Get 0th sheet object from above workbook
		 * 3. Get 0th row object
		 * 4. get 0th cell object
		 */
		
		// get workbook object
		File file = new File(filePath);
		FileInputStream fis = new FileInputStream(file);
		// get workbook object from above input stream
		Workbook workbook = WorkbookFactory.create(fis);
		// get sheet from above workbook
		Sheet sheet = workbook.getSheetAt(0);
		// get row from above sheet
		Row row = sheet.getRow(0);
		// get cell from above row 
		Cell cell = row.getCell(0);
		System.out.println(cell.getStringCellValue());
		
		workbook.close();
		fis.close();
	}
	
	public void readExcelDataForEntireSheet(String filePath) throws IOException {
		/*
		 * 1. get workbook object from input stream
		 * 2. create sheet iterator from above workbook
		 * 3. create row iterator from above sheet
		 * 4. create cell iterator from above row
		 * 5. find the type of cell
		 * 6. get data from cell
		 */
		
		// get workbook object
		File file = new File(filePath);
		FileInputStream fis = new FileInputStream(file);
		
		// get workbook object from above input stream
		Workbook workbook = WorkbookFactory.create(fis);
		// iterate through the workbook
		Iterator<Sheet> sheetIterator = workbook.sheetIterator();
		while (sheetIterator.hasNext()) {
			Sheet sheet = sheetIterator.next();
			// row iterator
			Iterator<Row> rowIterator = sheet.rowIterator();
			
			while (rowIterator.hasNext()) {
				Row row = rowIterator.next();
				// cell iterator
				Iterator<Cell> cellIterator = row.cellIterator();
				
				while (cellIterator.hasNext()) {
					Cell cell = cellIterator.next();
					switch (cell.getCellType()) {
						case BLANK:
							System.out.print("");
							break;
						case BOOLEAN:
							System.out.print(cell.getBooleanCellValue());
							break;
						case ERROR:
							System.out.print("error");
							break;
						case NUMERIC:
							System.out.print(cell.getNumericCellValue());
							break;
						case STRING:
							System.out.print(cell.getStringCellValue());
							break;
						case FORMULA:
							System.out.print(cell.getCellFormula());
							break;
						case _NONE:
							System.out.print("none");
							break;
						default:
					}
					System.out.print("\t");
				}
				System.out.println();
			}
		}
	}
	
}
