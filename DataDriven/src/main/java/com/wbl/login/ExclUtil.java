package com.wbl.login;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Properties;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.WebDriver;

public class ExclUtil extends Inistilization {

	public static WebDriver driver;

	private static XSSFWorkbook excelWBook;
	private static XSSFSheet excelWSheet;
	private static XSSFCell cell;
	private static XSSFRow row;
	public static int rowNumber;
	public static int columnNumber;
    public static String testData = "C:\\Users\\kanpuri\\testData.xls";

	public static void setRowNumber(int pRowNumber) {
		rowNumber = pRowNumber;
	}
	public static int getRowNumber() {
		return rowNumber;
	}
	public static void setColumnNumber(int pColumnNumber) {
		columnNumber = pColumnNumber;
	}
	public static int getColumnNumber() {
		return columnNumber;
	}
	public static void setExcelFileSheet(String sheetName) throws IOException {
		File src = new File("testData");
		FileInputStream fil = new FileInputStream(src);
		excelWBook = new XSSFWorkbook(fil);
		excelWSheet = excelWBook.getSheet("sheet1");
	}
	public static String getCellData(int rowNum, int colNum ) {
		try {
			cell = excelWSheet.getRow(rowNum).getCell(colNum);
			DataFormatter formatter = new DataFormatter();
			String cellData = formatter.formatCellValue(cell);
			return cellData;
		} catch (Exception e) {
			throw (e);
		}
	}
	public static XSSFRow getRowData(int RowNum) {
		try {
			row = excelWSheet.getRow(RowNum);
			return row;
		} catch (Exception e) {
			throw (e);
		}
	}

}
