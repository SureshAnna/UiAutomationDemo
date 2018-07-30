import java.io.FileInputStream;
import java.lang.reflect.Constructor;
import java.lang.reflect.Method;
import java.util.ArrayList;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.IAnnotationTransformer;
import org.testng.annotations.ITestAnnotation;

public class DemoApplication extends TestBase implements IAnnotationTransformer{

	public static XSSFWorkbook wb;
	public static XSSFSheet sheet;

	public String ExecuteFlag;
	public String TestCaseToByPass;

	public String ITestannotationMethod;
	public String AllMethods[];
	public ArrayList<String> ar;

	public void transform(ITestAnnotation annotation, Class testClass,
			Constructor testConstructor, Method testMethod) {

		ArrayList<String> ar = new ArrayList<String>();
		ar.add(testMethod.getName());

		try {
			FileInputStream fis = new FileInputStream(
					"D:\\SeleniumPractice\\ApiAutomationUsingReflection\\src\\test\\java\\dataEngine\\DataEngine.xlsx");
			wb = new XSSFWorkbook(fis);
			sheet = wb.getSheet("Sheet1");

			int length = sheet.getLastRowNum();
			for (int row = 1; row <= length; row++) {
				ExecuteFlag = sheet.getRow(row).getCell(3).getStringCellValue();
				if (ExecuteFlag.equals("No")) {
					String[] BypassedMethods = new String[length];
					for (int i = 0; i < BypassedMethods.length; i++) {
						BypassedMethods[i] = sheet.getRow(row).getCell(2)
								.getStringCellValue();
						for (String j : ar) {
							if (j.equals(BypassedMethods[i])) {
								annotation.setEnabled(false);
							}

						}

					}
				}

			} // for close brace

		} catch (Exception e) {

			System.out.println(e.getMessage());
		}

	}
}
