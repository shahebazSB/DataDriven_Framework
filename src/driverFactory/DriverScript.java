package driverFactory;

import java.io.File;

import org.apache.commons.io.FileUtils;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.testng.Reporter;
import org.testng.annotations.Test;

import com.relevantcodes.extentreports.ExtentReports;
import com.relevantcodes.extentreports.ExtentTest;
import com.relevantcodes.extentreports.LogStatus;

import commonFunctions.FunctionLibrary;
import config.AppUtil;
import utilities.ExcelFileUtil;

public class DriverScript extends AppUtil{
ExtentReports reports;
ExtentTest logger;
String inputpath="./FileInput/TestData.xlsx";
String outputpath ="./FileOutput/DataDrivenResults.xlsx";
@Test
public void startTest() throws Throwable
{
	//define path of html report
	reports = new ExtentReports("./Reports/Login.html");
	//create reference object for excel file util class
	ExcelFileUtil xl = new ExcelFileUtil(inputpath);
	//count no of rows in login Sheet
	int rc =xl.rowCount("Login");
	Reporter.log("No of rows are::"+rc,true);
	//iterate all rows in login sheet
	for(int i=1;i<=rc;i++)
	{
		logger = reports.startTest("Login Test");
		logger.assignAuthor("Ranga");
		//read username and password cells
		String username = xl.getCellData("Login", i, 0);
		String password = xl.getCellData("Login", i, 1);
		logger.log(LogStatus.INFO,username+"     "+password);
		//call login method and assign parametrs
		boolean res =FunctionLibrary.adminLogin(username, password);
		if(res)
		{
			//if res is true write as login success into results cell
			xl.setCellData("Login", i, 2, "Login success", outputpath);
			//if res is true write as pass into status cell
			xl.setCellData("Login", i, 3, "pass", outputpath);
			logger.log(LogStatus.PASS, "Valid Username and password");
		}
		else
		{
			//adding screen shot for fail test
			File screen =((TakesScreenshot)driver).getScreenshotAs(OutputType.FILE);
			//copy screen shot into local system
			FileUtils.copyFile(screen, new File("./Screenshot/Iteration/"+i+"Login.png"));
			//if res is false write as login Fail into results cell
			xl.setCellData("Login", i, 2, "Login Fail", outputpath);
			//if res is false write as Fail into status cell
			xl.setCellData("Login", i, 3, "Fail", outputpath);
			logger.log(LogStatus.FAIL, "Invalid username and password");
		}
		
		reports.endTest(logger);
		reports.flush();
		
	}
}
}








