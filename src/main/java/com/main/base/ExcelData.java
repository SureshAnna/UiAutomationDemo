package com.main.base;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * @author Swathi
 *
 */
public class ExcelData {

	public String ExecuteFlag;
	public static String data = null;
	public static XSSFSheet ExcelWSheet;
	private static XSSFWorkbook ExcelWBook;
	private static XSSFCell Cell;
	static FileInputStream fis;
	static FileOutputStream fos;

	/**
	 * This method is used to read the excel and write the status result
	 * pass/fail in excel.
	 * 
	 * @param rowNumber
	 * @param resultData
	 * @throws Exception
	 */
	public static void setExcelFiletoSetStatus(int rowNumber, String resultData)
			throws Exception {
		String sPath = "D:\\SeleniumPractice\\ApiAutomationUsingReflection\\src\\test\\java\\dataEngine\\DataEngine.xlsx";
		fis = new FileInputStream(sPath);
		ExcelWBook = new XSSFWorkbook(fis);
		ExcelWSheet = ExcelWBook.getSheet("sheet1");
		XSSFRow row1 = ExcelWSheet.getRow(rowNumber);
		XSSFCell r5c5 = row1.createCell(4);
		r5c5.setCellValue(resultData);
		fis.close();
		fos = new FileOutputStream(
				new File(
						"D:\\SeleniumPractice\\ApiAutomationUsingReflection\\src\\test\\java\\dataEngine\\DataEngine.xlsx"));
		ExcelWBook.write(fos);
		fos.close();

	}

	/**
	 * This method is used to read the excel and write the failure exception in
	 * comment cell in excel.
	 * 
	 * @param rowNumber
	 * @param comment
	 * @throws Exception
	 */
	public static void setExcelFiletoSendComment(int rowNumber, String comment)
			throws Exception {
		String sPath = "D:\\SeleniumPractice\\ApiAutomationUsingReflection\\src\\test\\java\\dataEngine\\DataEngine.xlsx";
		fis = new FileInputStream(sPath);
		ExcelWBook = new XSSFWorkbook(fis);
		ExcelWSheet = ExcelWBook.getSheet("sheet1");
		XSSFRow row1 = ExcelWSheet.getRow(rowNumber);
		XSSFCell r5c5 = row1.createCell(5);
		r5c5.setCellValue(comment);
		fis.close();
		fos = new FileOutputStream(
				new File(
						"D:\\SeleniumPractice\\ApiAutomationUsingReflection\\src\\test\\java\\dataEngine\\DataEngine.xlsx"));
		ExcelWBook.write(fos);
		fos.close();

	}

	/**
	 * it takes row number and column number from the excel
	 * 
	 * @param RowNum
	 * @param ColNum
	 * @return
	 * @throws Exception
	 */
	public static String getCellData(int RowNum, int ColNum) throws Exception {
		Cell = ExcelWSheet.getRow(RowNum).getCell(ColNum);
		String CellData = Cell.getStringCellValue();
		return CellData;
	}

}
