package com.main.base;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.util.concurrent.TimeUnit;

import org.apache.commons.io.FileUtils;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.Select;
import org.testng.ITestResult;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

import com.relevantcodes.extentreports.ExtentReports;
import com.relevantcodes.extentreports.ExtentTest;
import com.relevantcodes.extentreports.LogStatus;

/**
 * @author Swathi
 *
 */
public class TestBase {
	public static WebDriver driver;
	public ExtentReports extent;
	public ExtentTest logger;
	public static XSSFWorkbook wb;
	public static XSSFSheet sheet;

	public String fname, lname, email, pwd, cpwd;
	public FileInputStream fis;
	public int length;

	String sPath = "C:\\Users\\Suresh\\git\\UiAutomationDemo\\src\\test\\java\\dataEngine\\DataEngine.xlsx";

	/**
	 * This method is used to open the brwoser and redirect to newtour site
	 */

	@BeforeTest
	public void openBrowser() throws Exception {
		try {

			String workingDir = System.getProperty("user.dir");
			extent = new ExtentReports(workingDir
					+ "\\test-output\\ExtentReport.html");
			extent.addSystemInfo("Host Name", "Sample");
			extent.addSystemInfo("Environment", "Automation Testing");
			extent.addSystemInfo("User Name", "Swathi");
			extent.loadConfig(new File(System.getProperty("user.dir")
					+ "\\extent-config.xml"));

			System.setProperty("webdriver.chrome.driver",
					"E:\\Selenium_Training\\Selenium_Training\\src\\Assignmnets\\drivers\\chromedriver.exe");
			driver = new ChromeDriver();
			driver.get("http://newtours.demoaut.com/index.php");

			driver.manage().window().maximize();
			logger = extent.startTest("Test Name", "Description");
			logger.log(LogStatus.PASS, "opened application successfully");
		} catch (Exception e) {
			logger.log(LogStatus.FAIL, "opened application failed");

		}

	}

	/**
	 * This method is used to register the application and taking results
	 * dynamically
	 */

	@Test(enabled = true, priority = 1)
	public void register() throws Exception {
		String result = "FAIL";
		try {
			fis = new FileInputStream(
					"C:\\Users\\Suresh\\git\\UiAutomationDemo\\src\\test\\java\\dataEngine\\Registeration.xlsx");
			wb = new XSSFWorkbook(fis);
			sheet = wb.getSheet("Sheet1");
			length = sheet.getLastRowNum();
			for (int row = 1; row <= length; row++) {
				fname = sheet.getRow(row).getCell(0).getStringCellValue();
				lname = sheet.getRow(row).getCell(1).getStringCellValue();
				email = sheet.getRow(row).getCell(2).getStringCellValue();
				pwd = sheet.getRow(row).getCell(3).getStringCellValue();
				cpwd = sheet.getRow(row).getCell(4).getStringCellValue();

				driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
				driver.findElement(By.linkText("REGISTER")).click();

				driver.findElement(By.name("firstName")).sendKeys(fname);
				driver.findElement(By.name("lastName")).sendKeys(lname);
				driver.findElement(By.name("email")).sendKeys(email);
				driver.findElement(By.name("password")).sendKeys(pwd);
				driver.findElement(By.name("confirmPassword")).sendKeys(cpwd);
				driver.findElement(By.name("register")).click();

				if (LogStatus.PASS != null) {
					result = "PASS";
					ExcelData.setExcelFiletoSetStatus(1, result);

				}
			}
			logger = extent.startTest("register", "registeration successfully");
			System.out.println("Registeration successfully");
			logger.log(LogStatus.PASS, "registeration successfully");

		} catch (FileNotFoundException e) {
			String message = e.getMessage();
			ExcelData.setExcelFiletoSetStatus(1, result);
			ExcelData.setExcelFiletoSendComment(1, message);
			logger.log(LogStatus.FAIL, "registeration Failed");

		}
	}

	/**
	 * This method is used to sign to the newtour application with registered
	 * data.
	 */

	@Test(enabled = true, priority = 2)
	public void signIn() throws Exception {
		String result = "FAIL";
		try {
			driver.get("http://newtours.demoaut.com/mercurysignon.php");
			driver.findElement(By.name("userName")).sendKeys("TestUseremail1@Test.com");
			driver.findElement(By.name("password")).sendKeys("Sample123");
			driver.findElement(By.name("login")).click();
			System.out.println("Login succssfully");
			logger.log(LogStatus.PASS, "signin verified");
			if (LogStatus.PASS != null) {
				result = "PASS";
				ExcelData.setExcelFiletoSetStatus(2, result);

			}
		} catch (Exception e) {
			String message = e.getMessage();
			ExcelData.setExcelFiletoSetStatus(2, result);
			ExcelData.setExcelFiletoSendComment(2, message);
			logger.log(LogStatus.FAIL, "signin failed");
		}

	}

	/**
	 * This method is used to find the flights availability.
	 */
	@Test(enabled = true, priority = 3)
	public void flightFinder() throws Exception {
		String result = "FAIL";
		try {
			Select oSelect = new Select(
					driver.findElement(By.name("passCount")));
			oSelect.selectByVisibleText("1");
			System.out.println("find flights successfully");
			driver.findElement(By.name("findFlights")).click();
			logger.log(LogStatus.PASS,
					"checking flights available successfully");
			if (LogStatus.PASS != null) {
				result = "PASS";
				ExcelData.setExcelFiletoSetStatus(3, result);

			}
		} catch (Exception e) {
			String message = e.getMessage();
			logger.log(LogStatus.FAIL, "checking available failed");
			ExcelData.setExcelFiletoSetStatus(3, result);
			ExcelData.setExcelFiletoSendComment(3, message);
		}

	}

	/**
	 * This method is used to select the available flights.
	 */

	@Test(enabled = true, priority = 4)
	public void selectFlight() throws Exception {
		String result = "FAIL";
		try {
			driver.findElement(By.name("reserveFlights")).click();
			System.out.println("reserved flights succssfully");
			logger.log(LogStatus.PASS, "selected flight");
			if (LogStatus.PASS != null) {
				result = "PASS";
				ExcelData.setExcelFiletoSetStatus(4, result);

			}
		} catch (Exception e) {
			String message = e.getMessage();
			logger.log(LogStatus.FAIL, "selected flight failed");
			ExcelData.setExcelFiletoSetStatus(4, result);
			ExcelData.setExcelFiletoSendComment(4, message);
		}
	}

	/**
	 * This method is used to book the available flight.
	 */

	@Test(enabled = true, priority = 5)
	public void buyFlight() throws Exception {
		String result = "FAIL";
		try {
			driver.findElement(By.name("buyFlights")).click();
			System.out.println("booked flights succssfully");
			logger.log(LogStatus.PASS, "flights booked");
			if (LogStatus.PASS != null) {
				result = "PASS";
				ExcelData.setExcelFiletoSetStatus(5, result);

			}
		} catch (Exception e) {
			String message = e.getMessage();
			logger.log(LogStatus.FAIL, "flights booking failed");
			ExcelData.setExcelFiletoSetStatus(5, result);
			ExcelData.setExcelFiletoSendComment(5, message);
		}

	}

	/**
	 * This method is used to close the browser.
	 */

	@Test(enabled = true, priority = 6)
	public void closeBrowser() throws Exception {
		String result = "FAIL";
		try {
			driver.quit();
			logger.log(LogStatus.PASS, "closed application successfully");
			if (LogStatus.PASS != null) {
				result = "PASS";
				ExcelData.setExcelFiletoSetStatus(6, result);

			}
		} catch (Exception e) {
			String message = e.getMessage();
			logger.log(LogStatus.FAIL, "closed application failed");
			ExcelData.setExcelFiletoSetStatus(6, result);
			ExcelData.setExcelFiletoSendComment(6, message);
		}
	}

	/**
	 * This method is used to take the failed method screenshots.
	 */

	@AfterMethod
	public void screenShot(ITestResult result) {

		if (ITestResult.SUCCESS == result.getStatus()) {
		} else if (ITestResult.FAILURE == result.getStatus()) {

			try {
				TakesScreenshot screenshot = (TakesScreenshot) driver;
				File src = screenshot.getScreenshotAs(OutputType.FILE);
				FileUtils.copyFile(src,
						new File("D:\\Error\\" + result.getName() + ".png"));
				System.out.println("Successfully captured a screenshot");
				logger.log(LogStatus.PASS,
						"Test Case failed is " + result.getName());
			} catch (Exception e) {
				System.out.println("Exception while taking screenshot "
						+ e.getMessage());
				logger.log(LogStatus.FAIL,
						"Test Case failed is " + result.getName());
			}

		}
		extent.endTest(logger);
	}

	/**
	 * This method is used to close the extent report.
	 */

	@AfterTest
	public void afterTest() {
		extent.flush();

	}

}