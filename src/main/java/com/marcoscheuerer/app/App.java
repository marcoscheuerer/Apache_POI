package com.marcoscheuerer.app;

import java.io.IOException;

import com.marcoscheuerer.read.ReadExcel;
import com.marcoscheuerer.read.UpdateExcel;
import com.marcoscheuerer.write.WriteExcel;

public class App {
	
	public static void main(String[] args) throws IOException {
//		WriteExcel write = new WriteExcel();
		UpdateExcel read = new UpdateExcel();
		String path ="/home/marco/customer.xlsx";
//		write.writeSingleCellData(path);
//		write.writeMultipleCellData(path);
//		write.writeSingleCellDataWithFontStyle(path);
//		write.createCustomerFile("/home/marco/customer.xlsx");
//		read.readExcelDataForSingleRow(path);
//		read.readExcelDataForEntireSheet(path);
		read.readAndUpdate(path);
		
//		System.out.println("File created!");
	}
}
