package com.testing.ts.hooks;

import java.io.File;
import java.io.FilenameFilter;
import java.io.IOException;
import java.util.Hashtable;
import java.util.concurrent.TimeUnit;

import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.testng.SkipException;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.DataProvider;

import com.relevantcodes.extentreports.LogStatus;
import com.testing.ts.pages.HomePage;
import com.testing.ts.util.BaseClass;
import com.testing.ts.util.FilepathsUtility;
import com.testing.ts.util.Read_XLS;
import com.testing.ts.util.TestUtil;
import com.testing.ts.util.VideoRecorder;
import com.testing.ts.util.WordDocumentManager;

import atu.testrecorder.ATUTestRecorder;

public class TestNGHooks extends BaseClass {
	
	public ATUTestRecorder recorder;
	JavascriptExecutor js = null;
	public VideoRecorder videoRecorder	=	new VideoRecorder();

	// object creation
	public WebDriver driver;
	public BrowserOperations loginpage;
	public HomePage homewindow;

	// Default Variables
	protected static final String TEST_CASES 	= "TestCases";
	protected static final String PASS_FAIL_SKIP = "Pass/Fail/Skip";
	protected static final String DATA_TO_RUN 	= "DataToRun";
	protected static final String RUNMODE 		= "Runmode";
	private static final String RESULT 			= "Result";
	private static final String FAILED 			= "Failed";
	private static final String PASSED 			= "PASSED";
	private static final String SKIPPED 		= "Skipped";
	protected static  int dataRowNo = 3;
	boolean skip                   	= false;
	protected boolean pass          = false;
	int k                          	= 0;
	String fullName                	= "";
	protected String HomeWindow              	= "";
	protected String testCaseNo              	= "";
	String SheetName               	= null;
	Read_XLS      FilePath        	= null;
	String TestCaseName            	= null;
	String TestDataToRun[]         	= null;
	String ToRunColumnNameTestCase 				= null;
	String ToRunColumnNameTestData 				= null;
	protected int noOfPass 								= 0;
	protected int noOfFail 								= 0;
	protected int noOfSkip 								= 0;
	protected int totalNoOfTestcases 					= 0;

	@BeforeTest
	public void beforeTest() throws Exception {
		init();
		FilePath = TestCaseListExcel;
		TestCaseName = this.getClass().getSimpleName();
		SheetName = "TestCases";
		ToRunColumnNameTestCase = "Runmode";

		if (!FilepathsUtility.checkToRunUtility(FilePath, SheetName, ToRunColumnNameTestCase, TestCaseName)) {
			FilepathsUtility.WriteResultUtility(FilePath, SheetName, "Pass/Fail/Skip", TestCaseName, "SKIP");
			throw new SkipException(
					TestCaseName + "'s Runmode Flag Is 'N' Or Blank. So Skipping Execution Of " + TestCaseName);
		}
		CMS_CONFIG = initConfigurations(CONFIG_FILE_PATH);

	}
	
	public WebDriver isTestcaseExecutable(Hashtable<String, String> inputData, String browser) throws Exception {
		/*********************************
		 * Check Point
		 ***********************************************/

		k++;
		dataRowNo++;
		testCaseNo = "TC_" + k + "_";
		//recorder = new ATUTestRecorder(this.getClass().getSimpleName() + "_TC_" + k + "_" + getCurrentDate(), true);
		//recorder.start();
		
		videoRecorder.startRecording(this.getClass().getSimpleName());
		extentReportDetails("<b>" + this.getClass().getSimpleName() + "_RUN_" + k + "</b>",
				this.getClass().getSimpleName());
		if (!TestUtil.isExecutable(this.getClass().getSimpleName(), DEMOgetVal)
				|| (!inputData.get("Runmode").equals("Y")))
		{
			skip = true;
			APPLICATION_LOGS.debug("Skipping the Testcase as RunMode for Testcase" + k + " is not given as Y");
			extentTest.log(LogStatus.SKIP, "Skipping the Testcase as RunMode for Testcase" + k + " is not given as Y");
			ReportLogAs("Skipping the Testcase as RunMode for Testcase" + k + " is not given as Y", "PASS");
			DEMOgetVal.setCellData(this.getClass().getSimpleName(), RESULT, dataRowNo, "Skipped");
			noOfSkip++;
			throw new SkipException("Skipping the test");
		}

		/***********************************
		 * Driver Initialization
		 **********************************************/
		//killChromeDriverServer();
		driver=initDriver(browser);
		//js = (JavascriptExecutor) driver;
		
		loginpage = new BrowserOperations(driver);
		homewindow = loginpage.openUrl(CONFIG.getProperty("DEMO_URL"));
		
		//HomeWindow = driver.getWindowHandle();
		APPLICATION_LOGS.debug(("Logged in Successfully"));
		ReportLogAs("Logged in Successfully", "PASS");
		driver.manage().timeouts().implicitlyWait(15, TimeUnit.SECONDS);
		return driver;
	}
	

	@AfterTest
	public void close() throws Exception {

		extentTest.log(LogStatus.INFO, "APPLICATION CLOSED SUCCESSFULLY");
		ReportLogAs("APPLICATION CLOSED SUCCESSFULLY", "INFO");
		// close report.
		extent.endTest(extentTest);

		// writing everything to document.
		// extent.flush();

		flushExtentReport();
		driver.close();
		killIeDriverServer();
		suiteResults();
	}

	// Updates TestData Sheet of a Test case with status as
	// Passed/Failed/Skipped for Currently used
	// set of Test Data after every @Test method block is executed fully or
	// partially
	@AfterMethod
	public void setStatus() throws Exception {
		// recorder.stop();
		videoRecorder.stopRecording();
		if (!skip && !pass) {
			noOfFail++;
			DEMOgetVal.setCellData(this.getClass().getSimpleName(), RESULT, dataRowNo, "Failed");
		}

		if (!skip && pass) {
			noOfPass++;
			DEMOgetVal.setCellData(this.getClass().getSimpleName(), RESULT, dataRowNo, "Passed");
		}

		if (!skip) {
			//closeAllWindowsExcept(HomeWindow);
			//driver.switchTo().window(HomeWindow);
		}

		skip = false;
		pass = false;
	}

	// This Method reads data from TestData Sheet just Before @Test method
	// invocation
	// and returns a 2D array of HashTable Objects to the Test Method.......
	@DataProvider
	public Object[][] getDetails() {
		return TestUtil.getData(this.getClass().getSimpleName(), DEMOgetVal);
	}
	
	public void consolidateScreenshotsInWordDoc() {
		
		try {
			String screenshotsConsolidatedFolderPath = FilepathsUtility.getCurrentRunPath()
					+ FilepathsUtility.getFileSeparator()+ "ConsolidatedScreenshots";
			
			new File(screenshotsConsolidatedFolderPath).mkdir();
			
			String testcaseName = testCaseNo;
			System.out.println(testcaseName);
			WordDocumentManager documentManager = new WordDocumentManager(
					screenshotsConsolidatedFolderPath,testcaseName);

			String screenshotsFolderPath = System.getProperty("user.dir").replace("\\", "\\\\")+BaseClass.REPORTS+BaseClass.RUN+"\\screenshots";
			File screenshotsFolder = new File(screenshotsFolderPath);

			FilenameFilter filenameFilter = new FilenameFilter() {
				      public boolean accept(File dir, String fileName) {
					return fileName.contains(testcaseName);
				}
			};

			File[] screenshots = screenshotsFolder.listFiles(filenameFilter);
			if (screenshots != null && screenshots.length > 0) {
				documentManager.createDocument();
				for (File screenshot : screenshots) {
					documentManager.addPicture(screenshot);
				}
			}
			
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	//@AfterSuite//(alwaysRun=true)
 	public void suiteResults() throws Exception{
 		consolidateScreenshotsInWordDoc();
 		
 		init();
		FilePath = TestCaseListExcel;
		TestCaseName = this.getClass().getSimpleName();
		SheetName = "TestCases";
		ToRunColumnNameTestCase = "Runmode";
		// Name of column In Test Case Data sheets.
		ToRunColumnNameTestData = DATA_TO_RUN;

		APPLICATION_LOGS.debug("no of pass--------------"+noOfPass);
		APPLICATION_LOGS.debug("no of fail--------------"+noOfFail);
		APPLICATION_LOGS.debug("no of skip---------------"+noOfSkip);
		totalNoOfTestcases = noOfFail+noOfPass+noOfSkip;
		FilepathsUtility.WriteResultUtility(FilePath, SheetName,
				"NumberOfSkip", TestCaseName, String.valueOf(noOfSkip));
		FilepathsUtility.WriteResultUtility(FilePath, SheetName,
				"NumberOfPass", TestCaseName, String.valueOf(noOfPass));
		FilepathsUtility.WriteResultUtility(FilePath, SheetName,
				"NumberOfFail", TestCaseName, String.valueOf(noOfFail));
		FilepathsUtility.WriteResultUtility(FilePath, SheetName,
				"TotalTestcases", TestCaseName, String.valueOf(totalNoOfTestcases));
		
		if(noOfSkip==totalNoOfTestcases)
		{
			FilepathsUtility.WriteResultUtility(FilePath, SheetName,
					PASS_FAIL_SKIP, TestCaseName, SKIPPED);
			
		}
		else if(((noOfPass==totalNoOfTestcases)||(noOfSkip!=0))&&(noOfFail==0))
		{
			FilepathsUtility.WriteResultUtility(FilePath, SheetName,
					PASS_FAIL_SKIP, TestCaseName, PASSED);
			
		}
		else if(noOfFail!=0)
		{
			FilepathsUtility.WriteResultUtility(FilePath, SheetName,
					PASS_FAIL_SKIP, TestCaseName, FAILED);
		}
 	}

}
