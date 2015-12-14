package com.happylife.demo.poi;

import java.io.File;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CreateNewExcel {
	public static void main(String[] args) throws Exception {
		// Create Blank workbook
		XSSFWorkbook workbook = new XSSFWorkbook();
		//Create a blank worksheet
		XSSFSheet spreadsheet = workbook.createSheet("Sheet Name");
		// create the first row, 0 based
		XSSFRow row = spreadsheet.createRow(0);
		// create the first cell in the row, 0 based
		Cell cell = row.createCell(0);
        cell.setCellValue("hello world");
		// Create file system using specific name
		FileOutputStream out = new FileOutputStream(new File("d:/delete/createworkbook.xlsx"));
		// write operation workbook using file out object
		workbook.write(out);
		out.close();
		System.out.println("createworkbook.xlsx written successfully");
	}
}
