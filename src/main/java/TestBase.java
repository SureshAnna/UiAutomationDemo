import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.util.Map;
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

public class TestBase {
	public static WebDriver driver;
	public ExtentReports extent;
	public ExtentTest logger;
	public static XSSFWorkbook wb;
	public static XSSFSheet sheet;

	public String fname, lname, email, pwd, cpwd;
	public FileInputStream fis;
	public int length;

	String sPath = "D:\\SeleniumPractice\\ApiAutomationUsingReflection\\src\\test\\java\\dataEngine\\DataEngine.xlsx";

	@BeforeTest
	public void openBrowser() throws Exception {
		try {

			String workingDir = System.getProperty("user.dir");
			extent = new ExtentReports(workingDir
					+ "\\test-output\\Reports\\workspace.html");
			extent.addSystemInfo("Host Name", "Sample");
			extent.addSystemInfo("Environment", "Automation Testing");
			extent.addSystemInfo("User Name", "Swathi");
			extent.loadConfig(new File(System.getProperty("user.dir")
					+ "\\extent-config.xml"));

			System.setProperty("webdriver.chrome.driver",
					"D:\\SeleniumPractice\\ApiAutomationUsingReflection\\Driver\\chromedriver.exe");
			driver = new ChromeDriver();
			driver.get("http://newtours.demoaut.com/index.php");

			driver.manage().window().maximize();
			logger = extent.startTest("Test Name", "Description");
			logger.log(LogStatus.PASS, "opened application successfully");
		} catch (Exception e) {
			logger.log(LogStatus.FAIL, "opened application failed");

		}

	}

	@Test(enabled = true, priority = 1)
	public void register() throws Exception {
		String result = "FAIL";
		try {
			fis = new FileInputStream(
					"D:\\SeleniumPractice\\ApiAutomationUsingReflection\\src\\test\\java\\dataEngine\\Registeration.xlsx");
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
				System.out.println("Registeration successfully");
				logger = extent.startTest("register",
						"registeration successfully");
				logger.log(LogStatus.PASS, "registeration successfully");
				if (LogStatus.PASS != null) {
					result = "PASS";
					ReadingExcelData.setExcelFiletoSetStatus(1, result);

				}
			}

		} catch (FileNotFoundException e) {
			String message = e.getMessage();
			ReadingExcelData.setExcelFiletoSetStatus(1, result);
			ReadingExcelData.setExcelFiletoSendComment(1, message);
			logger.log(LogStatus.FAIL, "registeration Failed");

		}
	}

	@Test(enabled = true, priority = 2)
	public void signIn() throws Exception {
		String result = "FAIL";
		try {
			driver.get("http://newtours.demoaut.com/mercurysignon.php");
			driver.findElement(By.name("userName")).sendKeys("TestUser_11");
			driver.findElement(By.name("password")).sendKeys("Test@1234");
			driver.findElement(By.name("login")).click();
			System.out.println("Login succssfully");
			logger.log(LogStatus.PASS, "signin verified");
			if (LogStatus.PASS != null) {
				result = "PASS";
				ReadingExcelData.setExcelFiletoSetStatus(2, result);

			}
		} catch (Exception e) {
			String message = e.getMessage();
			ReadingExcelData.setExcelFiletoSetStatus(2, result);
			ReadingExcelData.setExcelFiletoSendComment(2, message);
			logger.log(LogStatus.FAIL, "signin failed");
		}

	}

	@Test(enabled = true, priority = 3)
	public void flightFinder() throws Exception {
		String result = "FAIL";
		try {
			Select oSelect = new Select(
					driver.findElement(By.name("passCount")));
			oSelect.selectByVisibleText("1");
			System.out.println("find flights successfully");
			driver.findElement(By.name("findFlights")).click();
			logger.log(LogStatus.PASS, "checking available successfully");
			if (LogStatus.PASS != null) {
				result = "PASS";
				ReadingExcelData.setExcelFiletoSetStatus(3, result);

			}
		} catch (Exception e) {
			String message = e.getMessage();
			logger.log(LogStatus.FAIL, "checking available failed");
			ReadingExcelData.setExcelFiletoSetStatus(3, result);
			ReadingExcelData.setExcelFiletoSendComment(3, message);
		}

	}

	@Test(enabled = true, priority = 4)
	public void selectFlight() throws Exception {
		String result = "FAIL";
		try {
			driver.findElement(By.name("reserveFlights")).click();
			System.out.println("reserved flights succssfully");
			logger.log(LogStatus.PASS, "selected flight");
			if (LogStatus.PASS != null) {
				result = "PASS";
				ReadingExcelData.setExcelFiletoSetStatus(4, result);

			}
		} catch (Exception e) {
			String message = e.getMessage();
			logger.log(LogStatus.FAIL, "selected flight failed");
			ReadingExcelData.setExcelFiletoSetStatus(4, result);
			ReadingExcelData.setExcelFiletoSendComment(4, message);
		}
	}

	@Test(enabled = true, priority = 5)
	public void buyFlight() throws Exception {
		String result = "FAIL";
		try {
			driver.findElement(By.name("buyFlights")).click();
			System.out.println("booked flights succssfully");
			logger.log(LogStatus.PASS, "flights booked");
			if (LogStatus.PASS != null) {
				result = "PASS";
				ReadingExcelData.setExcelFiletoSetStatus(5, result);

			}
		} catch (Exception e) {
			String message = e.getMessage();
			logger.log(LogStatus.FAIL, "flights booking failed");
			ReadingExcelData.setExcelFiletoSetStatus(5, result);
			ReadingExcelData.setExcelFiletoSendComment(5, message);
		}

	}

	@Test(enabled = true, priority = 6)
	public void closeBrowser() throws Exception {
		String result = "FAIL";
		try {
			driver.quit();
			logger.log(LogStatus.PASS, "closed application successfully");
			if (LogStatus.PASS != null) {
				result = "PASS";
				ReadingExcelData.setExcelFiletoSetStatus(6, result);

			}
		} catch (Exception e) {
			String message = e.getMessage();
			logger.log(LogStatus.FAIL, "closed application failed");
			ReadingExcelData.setExcelFiletoSetStatus(6, result);
			ReadingExcelData.setExcelFiletoSendComment(6, message);
		}
	}

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

	@AfterTest
	public void afterTest() {
		extent.flush();

	}

}