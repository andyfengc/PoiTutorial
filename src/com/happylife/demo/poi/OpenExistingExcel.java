package com.happylife.demo.poi;

import java.io.File;
import java.io.FileInputStream;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class OpenExistingExcel {
	public static void main(String args[]) throws Exception {
		File file = new File("d:/delete/createworkbook.xlsx");
		FileInputStream fis = new FileInputStream(file);
		if (file.isFile() && file.exists()) {
			// Get the workbook instance for XLSX file
			XSSFWorkbook workbook = new XSSFWorkbook(fis);
			System.out.println("openworkbook.xlsx file open successfully.");
			// get worksheet, 0 based
			XSSFSheet sheet = workbook.getSheetAt(0);
			// get rows
			Iterator<Row> it = sheet.iterator();
			while (it.hasNext()) {
				Row row = (XSSFRow) it.next();
				// get cells
				Iterator<Cell> cellIt = row.cellIterator();
				while (cellIt.hasNext()) {
					Cell cell = cellIt.next();
					switch (cell.getCellType()) {
					case Cell.CELL_TYPE_NUMERIC:
						System.out.print(cell.getNumericCellValue() + " \t\t ");
						break;
					case Cell.CELL_TYPE_STRING:
						System.out.print(cell.getStringCellValue() + " \t\t ");
						break;
					}
				}
			}

		} else {
			System.out.println("Error to open openworkbook.xlsx file.");
		}
	}
}
