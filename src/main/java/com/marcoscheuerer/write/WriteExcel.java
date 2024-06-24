package com.marcoscheuerer.write;

import java.awt.Color;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteExcel {

	/*
	 * 1. single data in excel sheet
	 * 2. multiple data row in excel sheet
	 */
	
	/**
	 * 
	 * @param filePath
	 * @throws IOException
	 */
	public void writeSingleCellData(String filePath) throws IOException {
		/*
		 * 1. create a workbook
		 * 2. create a sheet in above workbook
		 * 3. create a row in above sheet
		 * 4. create a cell in above row
		 * 5. set data inside cell
		 */
		
		Workbook workbook = new XSSFWorkbook();
		Sheet sheet = workbook.createSheet("firstSheet");
		Row row = sheet.createRow(0);
		Cell cell = row.createCell(0);
		cell.setCellValue("First Cell");
		
		// write workbook on output stream
		File file = new File(filePath);
		FileOutputStream fos = new FileOutputStream(file);
		workbook.write(fos);
		
		// close stream
		fos.close();
		workbook.close();
	}
	
	public void writeMultipleCellData(String filePath) throws IOException {
		
		Workbook workbook = new XSSFWorkbook();
		Sheet sheet = workbook.createSheet("MultipleCellSheet");
		
		int[][] dataArray = getRandomDataArray(5, 6);
		
		for (int i = 0; i < dataArray.length; i++) {
			Row row = sheet.createRow(i);
			
			for (int j = 0; j < dataArray[0].length; j++) {
				Cell cell = row.createCell(j);
				cell.setCellValue(dataArray[i][j]);
			}
		}
		
		File file = new File(filePath);
		FileOutputStream fos = new FileOutputStream(file);
		workbook.write(fos);
		
		workbook.close();
		fos.close();
	}

	
	public void writeSingleCellDataWithFontStyle(String filePath) throws IOException {
		
		Workbook workbook = new XSSFWorkbook();
		Sheet sheet = workbook.createSheet();
		Row row = sheet.createRow(0);
		
		/*
		 * 1. cellStyle Object
		 * 2. configure style and font options
		 * 3. pass cellStyle object to cell object
		 */
		
		// set style for right horizontal alignment
		Cell cell1 = row.createCell(1);
		CellStyle style = workbook.createCellStyle();
		style.setAlignment(HorizontalAlignment.RIGHT);
		cell1.setCellStyle(style);
		cell1.setCellValue("Horizontal Alignment");
		
		// set style for right horizontal alignment
		Cell cell2 = row.createCell(2);
		CellStyle style2 = workbook.createCellStyle();
		style2.setBorderBottom(BorderStyle.THIN);
		cell2.setCellStyle(style2);
		cell2.setCellValue("Border Cell");
		
		// set style for background color
		Cell cell3 = row.createCell(3);
		CellStyle style3 = workbook.createCellStyle();
		style3.setFillBackgroundColor(IndexedColors.YELLOW.getIndex());
		style3.setFillPattern(FillPatternType.BIG_SPOTS);
		cell3.setCellStyle(style3);
		cell3.setCellValue("BG Color");
		
		// set style for font family
		Cell cell4 = row.createCell(4);
		CellStyle style4 = workbook.createCellStyle();
		Font font = workbook.createFont();
		font.setBold(true);
		font.setFontName("Chilanka");
		style4.setFont(font);
		cell4.setCellStyle(style4);
		cell4.setCellValue("Font");
		
		// set style for text wrapping
		Cell cell5 = row.createCell(5);
		CellStyle style5 = workbook.createCellStyle();
		style5.setWrapText(true);
		cell5.setCellStyle(style5);
		cell5.setCellValue("This is longer text");
		
		// set style for shrink to fit
		Cell cell6 = row.createCell(6);
		CellStyle style6 = workbook.createCellStyle();
		style6.setShrinkToFit(true);
		cell6.setCellStyle(style6);
		cell6.setCellValue("This text must shrink to fit.");
		
		// set style for cell rotation
		Cell cell7 = row.createCell(7);
		CellStyle style7 = workbook.createCellStyle();
		style7.setRotation((short)45);
		cell7.setCellStyle(style7);
		cell7.setCellValue("rotate");
		
		File file = new File(filePath);
		FileOutputStream fos = new FileOutputStream(file);
		workbook.write(fos);
		
		workbook.close();
		fos.close();
	}
	
	
	public void createCustomerFile(String filePath) throws IOException {
		/*
		 * 1. create workbook
		 * 2. create sheet
		 * 3. create header row in above sheet
		 * 4. create re-usable styling method
		 * 5. create body with random data
		 * 6. write workbook object in output stream
		 */
		
		Workbook workbook = new XSSFWorkbook();
		Sheet sheet = workbook.createSheet("firstSheet");
		
		// create header row
		createHeaderRow(sheet, workbook);
		
		// create data body
		int rowCount = 15;
		int colCount = 4;
		
		String[] monthArray = { "May", "June", "July", "August" };
		int[] amountArray = { 220, -340, 1000, -3000, 7999 };
		
		for (int i = 1; i < rowCount; i++) {
			
			int randomIndexForMonth = (int)(Math.random() * monthArray.length);
			int randomIndexForAmount = (int)(Math.random() * amountArray.length);
			int amount = amountArray[randomIndexForAmount];
			Row row = sheet.createRow(i);
			
			for (int j = 0; j < colCount; j++) {
				Cell cell = row.createCell(j);
				
				if (j == 0) {
					cell.setCellStyle(getCellStyle(workbook, HorizontalAlignment.LEFT, null));
					cell.setCellValue(i);
				} else if (j == 1) {
					cell.setCellStyle(getCellStyle(workbook, HorizontalAlignment.LEFT, null));
					cell.setCellValue(monthArray[randomIndexForMonth]);
				} else if (j == 2) {
					cell.setCellStyle(getCellStyle(workbook, HorizontalAlignment.CENTER, null));
					cell.setCellValue(amount);
				} else if (j == 3) {
					if (amount > 0) {
						cell.setCellStyle(getCellStyle(workbook, HorizontalAlignment.RIGHT, IndexedColors.GREEN));
						cell.setCellValue("Credit");
					} else {
						cell.setCellStyle(getCellStyle(workbook, HorizontalAlignment.RIGHT, IndexedColors.RED));
						cell.setCellValue("Debit");
					}
				}
			}
		}
		
		/**
		 * 1. create row after all our data
		 * 2. set cell value as Total amount and set style
		 * 3. merge column1 and column2
		 */
		
		Row sumRow = sheet.createRow(rowCount);
		// create total cell
		Cell totalCell = sumRow.createCell(0);
		totalCell.setCellStyle(getCellStyle(workbook, HorizontalAlignment.RIGHT, IndexedColors.YELLOW));
		CellRangeAddress cellRange = new CellRangeAddress(rowCount, rowCount, 0, 1);
		sheet.addMergedRegion(cellRange);
		totalCell.setCellValue("Total");
		
		// create formula cell
		Cell formulaCell = sumRow.createCell(2);
		formulaCell.setCellFormula("SUM(C2:C" + rowCount + ")");
		formulaCell.setCellStyle(getCellStyle(workbook, HorizontalAlignment.RIGHT, null));
		
		// create hyperlink cell
		int hyperLinkRowCount = rowCount + 1;
		Row hyperLinkRow = sheet.createRow(hyperLinkRowCount);
		Cell hyperLinkCell = hyperLinkRow.createCell(0);
		/*
		 * 1. create object of creationHelper
		 * 2. create object of hyperlink
		 * 3. set hyperlink of cell
		 */
		CreationHelper createHelper = workbook.getCreationHelper();
		Hyperlink link = createHelper.createHyperlink(HyperlinkType.URL);
		link.setAddress("http://www.google.de");
		hyperLinkCell.setHyperlink(link);
		hyperLinkCell.setCellValue("MyHyperLink");
		hyperLinkCell.setCellStyle(getCellStyle(workbook, HorizontalAlignment.CENTER, null));
		// merge cells from 0 to 3
		CellRangeAddress cellRange2 = new CellRangeAddress(hyperLinkRowCount, hyperLinkRowCount, 0, 3);
		sheet.addMergedRegion(cellRange2);
		
		// write workbook on output stream
		File file = new File(filePath);
		FileOutputStream fos = new FileOutputStream(file);
		workbook.write(fos);
		
		// close stream
		workbook.close();
		fos.close();
	}
	
	
	private CellStyle getCellStyle(Workbook workbook, HorizontalAlignment alignment, IndexedColors backgroundColor) {
		
		/*
		 * 1. user specific alignment
		 * 2. set border for each cell
		 * 3. set font family
		 * 4. set specific background color
		 */
		BorderStyle borderStyle = BorderStyle.THIN;
		
		CellStyle cellStyle = workbook.createCellStyle();
		cellStyle.setAlignment(alignment);
		cellStyle.setBorderBottom(borderStyle);
		cellStyle.setBorderLeft(borderStyle);
		cellStyle.setBorderTop(borderStyle);
		cellStyle.setBorderRight(borderStyle);
		
		// set font
		Font font = workbook.createFont();
		font.setFontName("Arial");
		font.setColor(IndexedColors.BLACK.getIndex());
		
		if (backgroundColor != null) {
			cellStyle.setFillBackgroundColor(backgroundColor.getIndex());
			cellStyle.setFillPattern(FillPatternType.BIG_SPOTS);
			font.setColor(IndexedColors.WHITE.getIndex());
		}
		
		cellStyle.setFont(font);
		
		return cellStyle;
	}
	
	
	private Row createHeaderRow(Sheet sheet, Workbook workbook) {
		Row headerRow = sheet.createRow(0);
		// create header row style
		CellStyle headerStyle = workbook.createCellStyle();
		headerStyle.setAlignment(HorizontalAlignment.CENTER);
		headerStyle.setFillBackgroundColor(IndexedColors.YELLOW.getIndex());
		headerStyle.setFillPattern(FillPatternType.BIG_SPOTS);
		
		// set header font
		Font font = workbook.createFont();
		font.setFontName("Arial");
		font.setBold(true);
		font.setColor(IndexedColors.WHITE.getIndex());
		headerStyle.setFont(font);
		
		// create cells
		Cell cell0 = headerRow.createCell(0);
		cell0.setCellStyle(headerStyle);
		cell0.setCellValue("S. No.");
		
		Cell cell1 = headerRow.createCell(1);
		cell1.setCellStyle(headerStyle);
		cell1.setCellValue("Month");
		
		Cell cell2 = headerRow.createCell(2);
		cell2.setCellStyle(headerStyle);
		cell2.setCellValue("Amount");
		
		Cell cell3 = headerRow.createCell(3);
		cell3.setCellStyle(headerStyle);
		cell3.setCellValue("Credit/Debit");
		
		// return headerRow
		return headerRow;
	}

	
	private int[][] getRandomDataArray(int countRow, int countCol) {
		int[][] dataArray = new int[countRow][countCol];
		
		for (int i = 0; i < countRow; i++) {
			for (int j = 0; j < countCol; j++) {
				dataArray[i][j] = (int)(Math.random() * 1000);
			}
		}
		
		return dataArray;
	}
	
}
