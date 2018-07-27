package com.main.test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadingExcelData {

	public String ExecuteFlag;
	public static String data = null;
	public static XSSFSheet ExcelWSheet;
	private static XSSFWorkbook ExcelWBook;
	private static XSSFCell Cell;
	static FileInputStream fis;
	static FileOutputStream fos;
	

	public static void setExcelFiletoSetStatus(int rowNumber, String resultData)
			throws Exception {
		String sPath = "D:\\SeleniumPractice\\ApiAutomationUsingReflection\\src\\dataEngine\\DataEngine.xlsx";
		fis = new FileInputStream(sPath);
		ExcelWBook = new XSSFWorkbook(fis);
		ExcelWSheet = ExcelWBook.getSheet("sheet1");
		XSSFRow row1 = ExcelWSheet.getRow(rowNumber);
		XSSFCell r5c5 = row1.createCell(4);
		r5c5.setCellValue(resultData);
		fis.close();
		fos = new FileOutputStream(
				new File(
						"D:\\SeleniumPractice\\ApiAutomationUsingReflection\\src\\dataEngine\\DataEngine.xlsx"));
		ExcelWBook.write(fos);
		fos.close();

	}
	
	public static void setExcelFiletoSendComment(int rowNumber, String comment)
			throws Exception {
		String sPath = "D:\\SeleniumPractice\\ApiAutomationUsingReflection\\src\\dataEngine\\DataEngine.xlsx";
		fis = new FileInputStream(sPath);
		ExcelWBook = new XSSFWorkbook(fis);
		ExcelWSheet = ExcelWBook.getSheet("sheet1");
		XSSFRow row1 = ExcelWSheet.getRow(rowNumber);
		XSSFCell r5c5 = row1.createCell(5);
		r5c5.setCellValue(comment);
		fis.close();
		fos = new FileOutputStream(
				new File(
						"D:\\SeleniumPractice\\ApiAutomationUsingReflection\\src\\dataEngine\\DataEngine.xlsx"));
		ExcelWBook.write(fos);
		fos.close();

	}

	public static String getCellData(int RowNum, int ColNum) throws Exception {
		Cell = ExcelWSheet.getRow(RowNum).getCell(ColNum);
		String CellData = Cell.getStringCellValue();
		System.out.println("CellData----" + CellData);
		return CellData;
	}

	public static void get() throws IOException {

		try {
			// setExcelFile("data");
			int length = ExcelWSheet.getLastRowNum();
			System.out.println("length..........." + length);
			for (int iRow = 1; iRow <= length; iRow++) {
				String ExecuteFlag = ExcelWSheet.getRow(iRow).getCell(3)
						.getStringCellValue();
				if (ExecuteFlag.equals("Yes")) {
					data = getCellData(iRow, 2);
					System.out.println("the data is " + data);
				}
			}
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

}
