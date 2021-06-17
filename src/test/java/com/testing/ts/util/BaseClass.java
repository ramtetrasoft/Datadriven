package com.testing.ts.util;

import java.awt.AWTException;
import java.awt.Robot;
import java.awt.Toolkit;
import java.awt.datatransfer.StringSelection;
import java.awt.event.KeyEvent;
import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.io.OutputStream;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.Iterator;
import java.util.List;
import java.util.Properties;
import java.util.Random;
import java.util.Set;
import java.util.concurrent.TimeUnit;

import org.apache.commons.io.FileUtils;
import org.apache.log4j.Logger;
import org.apache.log4j.PropertyConfigurator;
import org.openqa.selenium.Alert;
import org.openqa.selenium.By;
import org.openqa.selenium.ElementNotVisibleException;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.NoAlertPresentException;
import org.openqa.selenium.NoSuchElementException;
import org.openqa.selenium.NoSuchWindowException;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.StaleElementReferenceException;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.FluentWait;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.Wait;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.Assert;
import org.testng.Reporter;

import com.jacob.com.LibraryLoader;
import com.relevantcodes.extentreports.ExtentReports;
import com.relevantcodes.extentreports.ExtentTest;
import com.relevantcodes.extentreports.LogStatus;
import com.testing.ts.pages.HomePage;

import atu.testng.reports.ATUReports;
import atu.testng.reports.logging.LogAs;
import autoitx4java.AutoItX;
import io.github.bonigarcia.wdm.WebDriverManager;

public class BaseClass {
	public static WebDriver driver = null;
	public static Properties CONFIG = null;
	public static Properties INPUTCONFIG = null;

	protected Properties CMS_CONFIG = null;
	public static Logger APPLICATION_LOGS = null;
	public static FileOutputStream output = null;
	public Read_XLS TestCaseListExcel = null;
	private static final String SRC_COM_INTELLECT_CORE_TESTOUTPUT_SCREENSHOTS = "\\Screenshots\\";
	public static final String CONFIG_FILE_PATH = System.getProperty("user.dir") + "\\Properties\\config.properties";
	public static final String INPUT_FILE_PATH = System.getProperty("user.dir") + "\\Properties\\input.properties";
	public static final String REPORTS="\\Reports\\";
	public static final String currentDateTime=new SimpleDateFormat("yyyy-MMM-dd hh-mm-ss").format(new Date());
	public static final String RUN="\\RUN_"+currentDateTime+"\\";
	public static final String log4jConfigFile = System.getProperty("user.dir") + "\\Properties\\log4j.properties";
	public static Xls_Reader DEMOgetVal = new Xls_Reader(System.getProperty("user.dir") + "//Testdata//DataSheet.xls");
	public WebElement element;
	public HomePage homeWindow;
	public static ExtentReports extent = new ExtentReports(FilepathsUtility.getExtentReportPath(), true);

	public static ExtentTest extentTest = null;

	public static void extentReportDetails(String ScenarioName, String ScenarioDescription) {

		extent.loadConfig(new File(System.getProperty("user.dir") + "\\Properties\\extent-config.xml"));
		// Start the test using the ExtentTest class object.
		extentTest = extent.startTest(ScenarioName, ScenarioDescription);
	}

	public void flushExtentReport() {
		extent.flush();

		String html = FilepathsUtility.getExtentReportPath();
		StringBuilder contentBuilder = new StringBuilder();
		try {
			BufferedReader in = new BufferedReader(new FileReader(html));
			String str;
			while ((str = in.readLine()) != null) {
				contentBuilder.append(str);
			}
			in.close();
		} catch (IOException e) {
		}

		String content = contentBuilder.toString();
		content = content.replaceAll("<img class=", "</td><td><img class=");
		content = content.replaceAll("<th>Details</th>", "<th>Details</th>\r\n" + "<th>Screenshots</th>");

		FileWriter fWriter = null;
		BufferedWriter writer = null;
		try {
			fWriter = new FileWriter(html);
			writer = new BufferedWriter(fWriter);
			writer.write(content);
			writer.newLine(); // this is not actually needed for html files -
								// can make your code more readable though
			writer.close(); // make sure you close the writer object
		} catch (Exception e) {
			// catch any exceptions here
		}
	}

	public void init() {
		TestCaseListExcel = new Read_XLS(System.getProperty("user.dir") + "\\Testdata\\DataSheet.xls");

	}

	
	public WebElement fluentWait(final By locator)
	{
		Wait<WebDriver> wait = new FluentWait<WebDriver>(driver).withTimeout(10, TimeUnit.SECONDS)
				.pollingEvery(1, TimeUnit.SECONDS).ignoring(NoSuchElementException.class);

		return wait.until(new java.util.function.Function<WebDriver, WebElement>() {
		public WebElement apply(WebDriver driver) {
				APPLICATION_LOGS.debug("Trying");
				return driver.findElement(locator);
			}
		});
	}

	
	/**************************************************************************************************************/

	public static WebDriver initDriver() {
		try {
			if (CONFIG.getProperty("BROWSER").equals("IE")) {
				WebDriverManager.iedriver().setup();
				//System.setProperty("webdriver.ie.driver", System.getProperty("user.dir")+"//WebDriverServer//IEDriverServer.exe");
				driver = new InternetExplorerDriver();
				//System.setProperty("webdriver.ie.driver", System.getProperty("user.dir")+"//WebDriverServer//IEDriverServer.exe");
				APPLICATION_LOGS.debug("Opening IE");
				//driver.manage().deleteAllCookies();
				//Runtime.getRuntime().exec("RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 255");
				
				Thread.sleep(5000);
			}
		} catch (Exception e) {
			e.printStackTrace();
			APPLICATION_LOGS.debug("DRIVER INITIALIZATION FAILED");
			Assert.fail("DRIVER INITIALIZATION FAILED");
		}
		return driver;
	}
	public static WebDriver initDriver(String browser) {
		try {
			 
			if(browser.equals("CHROME"))
			{
				WebDriverManager.chromedriver().setup();
				//System.setProperty("webdriver.chrome.driver", System.getProperty("user.dir")+"//WebDriverServer//chromedriver.exe");
				//DesiredCapabilities dc = new DesiredCapabilities();
				driver = new ChromeDriver();
				//driver.manage().deleteAllCookies();
				driver.manage().window().maximize();
				Thread.sleep(5000);
				APPLICATION_LOGS.debug("Opening CHROME");
			}
			else
			{
			//	System.setProperty("webdriver.gecko.driver", System.getProperty("user.dir")+"//WebDriverServer//geckodriver.exe");
				WebDriverManager.firefoxdriver().setup();
				driver = new FirefoxDriver();
				//driver.manage().deleteAllCookies();
				driver.manage().window().maximize();
				Thread.sleep(5000);
				APPLICATION_LOGS.debug("Opening FireFox");
			}
			
		} catch (Exception e) {
			e.printStackTrace();
			APPLICATION_LOGS.debug("DRIVER INITIALIZATION FAILED");
			Assert.fail("DRIVER INITIALIZATION FAILED");
		}
		return driver;
	}

	public static Properties initConfigurations(String filePath) {
		if (CONFIG == null) {
			// Logging

			APPLICATION_LOGS = Logger.getLogger("rbi");
			PropertyConfigurator.configure(log4jConfigFile);
			APPLICATION_LOGS.info("Log4j FilePath: " + log4jConfigFile + " is now Configured..");
			// config.prop

			CONFIG = new Properties();
			try {
				FileInputStream fs = new FileInputStream(filePath);
				CONFIG.load(fs);
				APPLICATION_LOGS.info("Config Filepath: " + filePath + " is now Configured..");
			} catch (Exception e) {
				e.printStackTrace();
			}
		}
		return CONFIG;
	}

	public static Properties initInputConfigurations(String filePath) {
		if (INPUTCONFIG == null) {
			// Logging

			APPLICATION_LOGS = Logger.getLogger("bi");
			PropertyConfigurator.configure(log4jConfigFile);
			APPLICATION_LOGS.info("Log4j FilePath: " + log4jConfigFile + " is now Configured..");
			// config.prop

			INPUTCONFIG = new Properties();
			try {
				FileInputStream fs = new FileInputStream(filePath);
				INPUTCONFIG.load(fs);
				APPLICATION_LOGS.info("Input Config Filepath: " + filePath + " is now Configured..");
			} catch (Exception e) {
				e.printStackTrace();
			}
		}
		return INPUTCONFIG;
	}

	public static void setpropertyKeyValue(String key, String value) {
		OutputStream out = null;

		try {

			Properties props = new Properties();

			File f = new File(INPUT_FILE_PATH);
			if (f.exists()) {

				props.load(new FileReader(f));
				// Change your values here
				props.setProperty(key, value);
			} else {

				// Set default values?
				props.setProperty(key, value);

				f.createNewFile();
			}

			out = new FileOutputStream(f);
			props.store(out, "INPUT CONFIG FILE UPDATED");

			System.out.println(props.get(value));

		} catch (Exception e) {
			e.printStackTrace();
		} finally {

			if (out != null) {

				try {

					out.close();
				} catch (IOException ex) {

					System.out
							.println("IOException: Could not close myApp.properties output stream; " + ex.getMessage());
					ex.printStackTrace();
				}
			}
		}
	}

	public void openURL(String url) {
		try {
			driver.manage().timeouts().implicitlyWait(15, TimeUnit.SECONDS);
			driver.get(url);
			driver.manage().window().maximize();
			APPLICATION_LOGS.debug("URL Opened successfully");
			Reporter.log("URL Opened successfully");
		} catch (Exception e) {
			APPLICATION_LOGS.debug("URL FAILED");
			Reporter.log("URL FAILED");
		}
	}

	
	public boolean switchToTitle(String title) throws InterruptedException {
		boolean switched = false;
		try {
			Set<String> handles = driver.getWindowHandles();
			for (String windowHandle : handles) {
				driver.switchTo().window(windowHandle);
				APPLICATION_LOGS.debug("Title " + driver.getTitle());
				if (driver.getTitle().equalsIgnoreCase(title.trim()) || driver.getTitle().contains(title.trim())) {
					APPLICATION_LOGS.debug("Finally Switched to " + driver.getTitle());
					extentTest.log(LogStatus.PASS, "Switched to <b>" + driver.getTitle() + "</b> window sccessfully");
					Thread.sleep(500);
					switched = true;
					break;
				}
			}
		} catch (NoSuchWindowException e) {
			extentTest.log(LogStatus.FAIL, "<b>Switching to Window FAILED " + title + "</b>");
			Assert.fail("Switching to Window FAILED" + title);
		}
		return switched;
	}

	public boolean switchToTitlewithoutAssert(String title) throws InterruptedException {
		boolean switched = false;
		try {
			Set<String> handles = driver.getWindowHandles();
			for (String windowHandle : handles) {
				driver.switchTo().window(windowHandle);
				if (driver.getTitle().equalsIgnoreCase(title.trim()) || driver.getTitle().contains(title.trim())) {
					APPLICATION_LOGS.debug("Finally Switched to " + driver.getTitle());
					extentTest.log(LogStatus.PASS, "Switched to <b>" + driver.getTitle() + "</b> window sccessfully");
					Thread.sleep(500);
					switched = true;
					break;
				}
			}
		} catch (NoSuchWindowException e) {
		}
		return switched;
	}

	public void acceptModalDialogue() {
		try {
			Alert alert = driver.switchTo().alert();
			alert.accept();
			APPLICATION_LOGS.debug("Inside Try....Alert is Present................................");
			extentTest.log(LogStatus.PASS, "Inside Try....Alert is Present................................");
		} catch (NoAlertPresentException e1) {
			APPLICATION_LOGS.debug("No Modal Dialogue/Alert present");
		}
	}

	public void selectFromDDL(WebElement element, String inputData) {

		if (inputData != "") {
			try {
				driver.manage().timeouts().implicitlyWait(60, TimeUnit.SECONDS);
				Select ddl = new Select(element);
				ddl.selectByVisibleText(inputData);
				Thread.sleep(500);
				element.sendKeys(Keys.TAB);		
				ReportLogAs("Successfully selected " + inputData, "PASS");
				APPLICATION_LOGS.debug("Successfully selected " + inputData);
				extentTest.log(LogStatus.PASS, "Successfully selected <b>" + inputData + "</b>");
			} catch (NoSuchElementException e1) {
				ReportLogAs(element + " NOT FOUND TO ENTER " + inputData, "FAILED");
				APPLICATION_LOGS.debug(element + " NOT FOUND TO ENTER " + inputData);
				extentTest.log(LogStatus.FAIL, "<b>" + element + " NOT FOUND TO ENTER " + inputData + "</b>");
				Assert.fail(element + " NOT FOUND TO ENTER " + inputData);
			} catch (ElementNotVisibleException e2) {
				ReportLogAs(element + " NOT VISIBLE TO ENTER " + inputData, "FAILED");
				APPLICATION_LOGS.debug(element + " NOT VISIBLE TO ENTER " + inputData);
				extentTest.log(LogStatus.FAIL, "<b>" + element + "</b> NOT VISIBLE TO ENTER <b>" + inputData + "</b>");
				Assert.fail(element + " NOT VISIBLE TO ENTER " + inputData);
			} catch (Exception e3) {
				e3.printStackTrace();
				extentTest.log(LogStatus.FAIL,
						"Cannot select <b>" + inputData + "</b> from <b>" + element.getAttribute("id") + "</b>");
				Assert.fail("Cannot select " + inputData + " from " + element.getAttribute("id"));

			}
		} else {
			element.sendKeys(Keys.TAB);
		}
	}
	
	public void selectFromDDL_Ebank(WebElement element, String inputData) {

		if (inputData != "") {
			try {
				driver.manage().timeouts().implicitlyWait(60, TimeUnit.SECONDS);
				Select ddl = new Select(element);
				ddl.selectByVisibleText(inputData);
				element.sendKeys(Keys.RETURN);	
				element.sendKeys(Keys.TAB);
				APPLICATION_LOGS.debug("Successfully selected " + inputData);
				extentTest.log(LogStatus.PASS, "Successfully selected <b>" + inputData + "</b>");
			} catch (NoSuchElementException e1) {
				APPLICATION_LOGS.debug(element + " NOT FOUND TO ENTER " + inputData + "</b>");
				extentTest.log(LogStatus.FAIL, "<b>" + element + " NOT FOUND TO ENTER " + inputData + "</b>");
				Assert.fail(element + " NOT FOUND TO ENTER " + inputData);
			} catch (ElementNotVisibleException e2) {
				APPLICATION_LOGS.debug(element + " NOT VISIBLE TO ENTER " + inputData);
				extentTest.log(LogStatus.FAIL, "<b>" + element + "</b> NOT VISIBLE TO ENTER <b>" + inputData + "</b>");
				Assert.fail(element + " NOT VISIBLE TO ENTER " + inputData);
			} catch (Exception e3) {
				e3.printStackTrace();
				extentTest.log(LogStatus.FAIL,
						"Cannot select <b>" + inputData + "</b> from <b>" + element.getAttribute("id") + "</b>");
				Assert.fail("Cannot select " + inputData + " from " + element.getAttribute("id"));

			}
		} else {
			element.sendKeys(Keys.RETURN);
		}
	}

	public String getFromDDL(WebElement element) {

		Select ddl = new Select(element);
		WebElement option = ddl.getFirstSelectedOption();
		return option.getText();
	}

	public void setData(WebElement element, String inputData) {// updated on
		// 19thJune2017
		try {
			if (!inputData.isEmpty()) {
				element.click();
				element.clear();
				element.sendKeys(inputData);
				element.sendKeys(Keys.TAB);
				//Thread.sleep(100);
				ReportLogAs("Successfully typed for " + element.getAttribute("id") + " : " + inputData, "PASS");
				APPLICATION_LOGS.debug("Successfully typed for " + element.getAttribute("id") + " : " + inputData);
				extentTest.log(LogStatus.PASS,
						"Successfully typed for <b>" + element.getAttribute("id") + " : " + inputData + "</b>");
			} else if (inputData.isEmpty()) {
				element.sendKeys(Keys.TAB);
				APPLICATION_LOGS.debug("No data for the field" + element.getAttribute("id"));
				extentTest.log(LogStatus.PASS, "No data for the field <b>" + element.getAttribute("id") + "</b>");
			}
		} catch (NoSuchElementException e1) {
			ReportLogAs(element + " NOT FOUND TO ENTER " + inputData, "FAILED");
			APPLICATION_LOGS.debug(element + " NOT FOUND TO ENTER " + inputData);
			extentTest.log(LogStatus.FAIL, "<b>" + element + "</b> NOT FOUND TO ENTER <b>" + inputData + "</b>");
			Assert.fail(element + " NOT FOUND TO ENTER " + inputData);
		} catch (ElementNotVisibleException e2) {
			ReportLogAs(element + " NOT VISIBLE TO ENTER " + inputData, "FAILED");
			APPLICATION_LOGS.debug(element + " NOT VISIBLE TO ENTER " + inputData);
			extentTest.log(LogStatus.FAIL, "<b>" + element + " NOT VISIBLE TO ENTER " + inputData + "</b>");
			Assert.fail(element + " NOT VISIBLE TO ENTER " + inputData);
		} catch (Exception e3) {
			e3.printStackTrace();
			extentTest.log(LogStatus.FAIL, "<b>" + e3.getLocalizedMessage() + "</b>");
		}
	}

	public void setDataForGrid(WebElement element, String inputdata) {

		try {
			if (!inputdata.isEmpty()) {
				doubleClick(element);
				String element_name = element.getText();
				int element_length = element_name.length();
				APPLICATION_LOGS.debug("getting element length" + element_length);
				extentTest.log(LogStatus.INFO, "getting element length : <b>" + element_length + "</b>");
				driver.switchTo().activeElement().sendKeys(Keys.END);
				Thread.sleep(500);
				driver.switchTo().activeElement().sendKeys(Keys.HOME, Keys.chord(Keys.SHIFT, Keys.END), Keys.BACK_SPACE,
						inputdata);
				Thread.sleep(1000);
				driver.switchTo().activeElement().sendKeys(Keys.ENTER);

				APPLICATION_LOGS.debug("Successfully typed " + inputdata);
				extentTest.log(LogStatus.PASS, "Successfully typed : <b>" + inputdata + "</b>");
			}
			else if (inputdata.isEmpty())
			{
				//doubleClick(element);
				driver.switchTo().activeElement().sendKeys(Keys.ENTER);
			}
		} catch (NoSuchElementException e1) {
			APPLICATION_LOGS.debug(element + " NOT FOUND TO ENTER " + inputdata);
			extentTest.log(LogStatus.FAIL, element + " NOT FOUND TO ENTER " + inputdata);
			Assert.fail(element + " NOT FOUND TO ENTER " + inputdata);
		} catch (ElementNotVisibleException e2) {
			APPLICATION_LOGS.debug(element + " NOT VISIBLE TO ENTER " + inputdata);
			extentTest.log(LogStatus.FAIL, element + " NOT VISIBLE TO ENTER " + inputdata);
			Assert.fail(element + " NOT VISIBLE TO ENTER " + inputdata);
		} catch (Exception e3) {
			APPLICATION_LOGS.error(e3.getLocalizedMessage());
			extentTest.log(LogStatus.ERROR, e3.getLocalizedMessage());
		}

	}

	
	public void waitForAlertAndUpdate(WebDriver driver) {
		try {
			WebDriverWait wait = new WebDriverWait(driver, 30);
			wait.until((java.util.function.Function) ExpectedConditions.alertIsPresent());
			APPLICATION_LOGS.debug("Alert Occured");
			APPLICATION_LOGS.debug("Alert Message " + acceptAlert(driver));
			extentTest.log(LogStatus.PASS, "Alert Occured...Alert Message is: <b>" + acceptAlert(driver) + "</b>");
		} catch (Exception e) {
			APPLICATION_LOGS.debug("Alert not occured");
		}
	}

	public void closeAllWindowsExcept(String windowName) throws InterruptedException {
		try {
			Set<String> handles = driver.getWindowHandles();
			for (String windowHandle : handles) {
				if (!(windowHandle.equals(windowName))) {
					driver.switchTo().window(windowHandle);
					driver.close();
					Thread.sleep(1000);
				}
			}
		} catch (NoSuchWindowException e) {
			APPLICATION_LOGS.debug("Switching to window failed");
			//extentTest.log(LogStatus.FAIL, " <b>Switching to window failed</b>");
		} finally {
			driver.switchTo().window(windowName);
			ReportLogAs("Closing All windows Except Homewindow", "PASSED");
			APPLICATION_LOGS.debug("Closing All windows Except Homewindow");
			extentTest.log(LogStatus.PASS, " <b>Closing All windows Except Homewindow</b>");
		}
	}

	/**********************************************************************************************/
	public void switchToChildWindow(String title) throws InterruptedException {
		boolean switched = false;
		try {
			Set<String> handles = driver.getWindowHandles();
			for (String windowHandle : handles) {
				APPLICATION_LOGS.debug("**************************************");
				APPLICATION_LOGS.debug("Total no of windows are " + handles.size());
				extentTest.log(LogStatus.INFO, "Total no of windows are <b>" + handles.size() + "</b>");
				driver.switchTo().window(windowHandle);
				APPLICATION_LOGS.debug("Switched to " + driver.getTitle());
				extentTest.log(LogStatus.PASS, "Switched to <b>" + driver.getTitle() + "</b>");
				if ((driver.getTitle().contains(title.trim()))) {
					APPLICATION_LOGS.debug("Finally Switched to " + driver.getTitle());
					extentTest.log(LogStatus.PASS, "Finally Switched to <b>" + driver.getTitle() + "</b>");
					Thread.sleep(500);
					switched = true;
					break;
				}
			}
			if (!(switched)) {
				Assert.fail("UNABLE TO SWITCH TO " + title + " AS IT IS NOT AVAILABLE");
				extentTest.log(LogStatus.FAIL, "UNABLE TO SWITCH TO <b>" + title + "</b> AS IT IS NOT AVAILABLE");
			}
		} catch (NoSuchWindowException e) {
			APPLICATION_LOGS.debug("Switching to window failed");
			extentTest.log(LogStatus.FAIL, "<b>Switching to window failed</b>");
		}
	}

	/***********************************************************************************************/
	public void waitForNoOfWindows(int noOfWindows) throws InterruptedException {

		for (int i = 0; i < 30; i++) {

			Set<String> handles = driver.getWindowHandles();
			// APPLICATION_LOGS.debug("Total No. of windows: " +
			// handles.size());
			if (handles.size() == noOfWindows) {
				APPLICATION_LOGS.debug("No. Of Windows are " + handles.size() + " Required windows are " + noOfWindows);
				extentTest.log(LogStatus.INFO,
						"No. Of Windows are <b>" + handles.size() + "</b> Required windows are " + noOfWindows);
				break;
			} else {
				APPLICATION_LOGS.debug("Waiting for " + noOfWindows + " in " + i + "th iteration");
				extentTest.log(LogStatus.INFO,
						"Waiting for <b>" + noOfWindows + "</b> in <b>" + i + "th</b> iteration");
				Thread.sleep(500);
			}

		}

	}

	public void doubleClick(WebElement element) {
		try {
			Actions action = new Actions(driver).doubleClick(element);
			action.build().perform();
			System.out.println("Double clicked the element");
		} catch (StaleElementReferenceException e) {
			System.out.println("Element is not attached to the page document " + e.getStackTrace());
		} catch (NoSuchElementException e) {
			System.out.println("Element " + element + " was not found in DOM " + e.getStackTrace());
		} catch (Exception e) {
			System.out.println("Element " + element + " was not clickable " + e.getStackTrace());
		}
	}

	public static void killIeDriverServer() throws IOException, InterruptedException {
		Runtime.getRuntime().exec("taskkill /F /IM iexplore.exe");
		Runtime.getRuntime().exec("taskkill /F /IM IEDriverServer.exe");
		Thread.sleep(1000);
		driver = null;
		Thread.sleep(2000);
	}

	public static void killChromeDriverServer() throws IOException, InterruptedException {
		Runtime.getRuntime().exec("taskkill /F /IM chrome.exe");
		Runtime.getRuntime().exec("taskkill /F /IM chromedriver.exe");
		Runtime.getRuntime().exec("taskkill /F /IM chromedriver_2.41.exe");
		Thread.sleep(1000);
		driver = null;
		Thread.sleep(2000);
	}
	public void waitForFrame(String frame) throws InterruptedException {
		Thread.sleep(1000);
		WebDriverWait wait = new WebDriverWait(driver, 10);
		wait.until((java.util.function.Function)ExpectedConditions.frameToBeAvailableAndSwitchToIt(frame));
		APPLICATION_LOGS.debug("Switched to " + frame + " Successfully");
		extentTest.log(LogStatus.PASS, "Switched to <b>" + frame + "</b> Successfully");
	}

	public void waitExplicitForFrame(String frame, int timeOut) {
		WebDriverWait wait = new WebDriverWait(driver, timeOut);
		wait.until((java.util.function.Function)ExpectedConditions.frameToBeAvailableAndSwitchToIt(frame));
		APPLICATION_LOGS.debug("Switched to " + frame + " Successfully");
		extentTest.log(LogStatus.PASS, "Switched to <b>" + frame + "</b> Successfully");
	}

	/***********************************************************************************************/
	public void clickOn(WebElement element) {
		try {

			if (element.getAttribute("type").contains("checkbox")) {
				element.click();
				APPLICATION_LOGS.debug("Clicked on " + element + " Successfully");
				extentTest.log(LogStatus.PASS, "Clicked on <b>" + element + "</b> Successfully");
				// driver.switchTo().activeElement().sendKeys(Keys.ENTER);
				// APPLICATION_LOGS.debug("Moving to next element by clickin
				// ENTER");
			} else {
				element.click();
				APPLICATION_LOGS.debug("Clicked on " + element + " Successfully\n");
				extentTest.log(LogStatus.PASS, "Clicked on <b>" + element + "</b> Successfully");
			}

		} catch (NoSuchElementException e1) {
			APPLICATION_LOGS.debug(element + " not found");
			extentTest.log(LogStatus.FAIL, "<b>" + element + "</b> not found");
			Assert.fail(element + " NOT FOUND TO ENTER TO BE CLICKED");
		} catch (ElementNotVisibleException e2) {
			APPLICATION_LOGS.debug(element + " not visible");
			extentTest.log(LogStatus.FAIL, "<b>" + element + "</b> not visible");
			Assert.fail(element + " NOT VISIBLE TO ENTER TO BE CLICKED");
		} catch (Exception e3) {
			e3.printStackTrace();
			APPLICATION_LOGS.error("Unknown exception", e3);
			extentTest.log(LogStatus.FAIL, "Unknown exception", e3);
			Assert.fail("UNKNOWN EXCEPTION,CHECK CONSOLE");
		}

	}

	public void clickonCheckbox(WebElement element, String inputData) {
		if (inputData.equalsIgnoreCase("yes")) {
			if (element.isSelected() == false) {
				clickOn(element);
				driver.switchTo().activeElement().sendKeys(Keys.ENTER);
			} else {
				driver.switchTo().activeElement().sendKeys(Keys.ENTER);
			}
		}
		if (inputData.equalsIgnoreCase("no")) {
			if (element.isSelected() == true) {
				clickOn(element);
				driver.switchTo().activeElement().sendKeys(Keys.ENTER);
			} else {
				driver.switchTo().activeElement().sendKeys(Keys.ENTER);
			}
		}
	}

	public boolean getCheckboxValue(WebElement element) {
		return element.isSelected();
	}

	/***********************************************************************************************/
	public void executeJavaFunction(WebDriver driver, String funcName) throws InterruptedException

	{
		Thread.sleep(1000);
		((JavascriptExecutor) driver).executeScript(funcName);
		Thread.sleep(3000);
	}

	/**************************************************************************************************/
	public boolean isAlertPresent() throws Exception {
		boolean alertPresent = false;
		try {
			Thread.sleep(1000);
			Alert alert = driver.switchTo().alert();
			//alert.accept();
			Thread.sleep(500);
			alertPresent = true;

		} catch (org.openqa.selenium.NoAlertPresentException e) {
			APPLICATION_LOGS.debug("NO SUCH ALERT WINDOW");
		}
		return alertPresent;
	}

	/********************************************************************************************/

	public void clickOnRecordUpdate() throws AWTException {
		try {
			Thread.sleep(3000);
			Robot robot = new Robot();
			robot.keyPress(KeyEvent.VK_ENTER);
			robot.keyRelease(KeyEvent.VK_ENTER);
		} catch (Exception e) {
			try {
				APPLICATION_LOGS.debug("NOT ABLE TO CICK ON RECORD UPDATE/AUTHORIZED, RE-CLICKING ONCE AGAIN");
				Thread.sleep(3000);
				Robot robot = new Robot();
				robot.keyPress(KeyEvent.VK_ENTER);
				robot.keyRelease(KeyEvent.VK_ENTER);
				APPLICATION_LOGS.debug("CLICKED SUCCESSFULLY ON RECORD UPDATE");

			} catch (Exception e2) {
				APPLICATION_LOGS.debug("NOT ABLE TO CICK ON RECORD UPDATE/AUTHORIZED AFTER 2 ATTEMPTS");
				Assert.fail();
			}

		}

	}

	/********************************************************************************************/
	public static String generateRefNumber() {

		Random random = new Random();
		int min = 1000000;
		int max = 99999999;
		int result = random.nextInt(max - min) + min;
		return Integer.toString(result);
	}
	
	/***********************************************************************************************/
	public static String generateSessionKey(int length){
		String alphabet = 
		        new String("0123456789");
		int n = alphabet.length();

		String result = new String(); 
		Random r = new Random(); 

		for (int i=0; i<length; i++) 
		    result = result + alphabet.charAt(r.nextInt(n));

		return result;
		}

	/***********************************************************************************************/
	public static String getCurrentDate() {
		DateFormat dateFormat = new SimpleDateFormat("yyyy_MM_dd_HH_mm_ss");
		Date date = new Date();
		return dateFormat.format(date);
	}

	/***********************************************************************************************/
	public static void getScreenShot(String sheetName,String fileName) {
		File srcFile = ((TakesScreenshot)(BaseClass.driver)).getScreenshotAs(OutputType.FILE);
	    try {
	    	//createNewDirectory(System.getProperty("user.dir")+SRC_COM_INTELLECT_CORE_TESTOUTPUT_SCREENSHOTS+scenarioName);
	    	//createNewDirectory(System.getProperty("user.dir")+SRC_COM_INTELLECT_CORE_TESTOUTPUT_SCREENSHOTS+scenarioName+sheetName);
	    	FileUtils.copyFile(srcFile, new File(FilepathsUtility.getScreenshotPath()+fileName+".png"));
	    	extentTest.log(LogStatus.INFO, extentTest.addScreenCapture(FilepathsUtility.getScreenshotPath()+fileName+".png"));	    	
	    } catch (IOException e) {                                                 
			e.printStackTrace();
		}
	}

	/***********************************************************************************************/
	public void getScreenShot(String fileName) {
		File srcFile = ((TakesScreenshot) (driver)).getScreenshotAs(OutputType.FILE);
		try {
			FileUtils.copyFile(srcFile, new File(FilepathsUtility.getScreenshotPath() + fileName + ".png"));
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	/**************************************************************************************************/
	public void getScreenshot(String screenShotName ,String sheetName, int dataRowNo, Xls_Reader xls) {
		// Take screenshot with date
		String Url = FilepathsUtility.getScreenshotPath()+sheetName+"_"+screenShotName+".png";
		
		APPLICATION_LOGS.debug(Url);
		
		getScreenShot(sheetName,sheetName+"_"+screenShotName + getCurrentDate());
		//getScreenShot(screenShotName + getCurrentDate());
		xls.addHyperLink(sheetName, "passedScreenShot", sheetName, dataRowNo, Url, screenShotName + "screenShot");
		xls.setCellData(sheetName, screenShotName + "passedScreenShot", dataRowNo, "", Url);
	}

	/**************************************************************************************************/
	public void getScreenshotForMiddle(String screenShotName, String sheetName, int dataRowNo, Xls_Reader xls) {
		// Take screenshot with date
		String Url = FilepathsUtility.getScreenshotPath()+sheetName+"_"+screenShotName+".png";
		APPLICATION_LOGS.debug(Url);
		getScreenShot(sheetName,sheetName+"_"+screenShotName + getCurrentDate());
		xls.addHyperLink(sheetName, screenShotName, sheetName, dataRowNo, Url, screenShotName + "screenShot");
		xls.setCellData(sheetName, screenShotName + "screenShot", dataRowNo, "", Url);
	}

	/********************************************************************************************/

	public String acceptAlert(WebDriver driver) throws InterruptedException {
		String msg = "";

		try {
			Alert alert = driver.switchTo().alert();
			msg = alert.getText();
			alert.accept();
			APPLICATION_LOGS.debug("Alert Msg: " + msg);
		} catch (NoAlertPresentException e) {
			Thread.sleep(500);
			APPLICATION_LOGS.debug("There is no alert.....!");
			// Assert.fail("No alert present");
		}
		return msg;
	}

	/***********************************************************************************************/
	public void setData1(WebElement element, String inputData) {
		try {
			if (!inputData.isEmpty()) {
				element.clear();
				element.sendKeys(inputData);
				element.sendKeys(Keys.RETURN);
				APPLICATION_LOGS.debug("Successfully typed " + inputData);
				extentTest.log(LogStatus.PASS, "Successfully typed <b>" + inputData + "</b>");
			} else if (inputData.isEmpty()) {
				element.sendKeys(Keys.RETURN);
				APPLICATION_LOGS.debug("Switching to next field by clicking ENTER");
				extentTest.log(LogStatus.PASS, "Switching to next field by clicking ENTER");
			}

		} catch (NoSuchElementException e1) {
			APPLICATION_LOGS.debug(element + " NOT FOUND TO ENTER " + inputData);
			extentTest.log(LogStatus.FAIL, "<b>" + element + "</b> NOT FOUND TO ENTER <b>" + inputData + "</b>");
			Assert.fail(element + " NOT FOUND TO ENTER " + inputData);
		} catch (ElementNotVisibleException e2) {
			APPLICATION_LOGS.debug(element + " NOT VISIBLE TO ENTER " + inputData);
			extentTest.log(LogStatus.FAIL, "<b>" + element + "</b> NOT VISIBLE TO ENTER <b>" + inputData + "</b>");
			Assert.fail(element + " NOT VISIBLE TO ENTER " + inputData);
		} catch (Exception e3) {
			e3.printStackTrace();
			extentTest.log(LogStatus.FAIL, "<b>" + e3.getLocalizedMessage() + "</b>");
		}
	}

	/***********************************************************************************************/

	public String findingGeneratedCifNumber(String value) {
		// Return a substring between the two strings.
		int posA = value.indexOf("Los ID(");
		if (posA == -1) {
			return "";
		}
		int posB = value.lastIndexOf(") Details Saved Successfully");
		if (posB == -1) {
			return "";
		}
		int adjustedPosA = posA + "Los ID(".length();
		if (adjustedPosA >= posB) {
			return "";
		}
		return value.substring(adjustedPosA, posB);
	}

	/***********************************************************************************************/

	public boolean waitelementToBeClickable(WebElement webElement) {

		try {
			WebDriverWait wait = new WebDriverWait(driver, 20);
			wait.until((java.util.function.Function)ExpectedConditions.elementToBeClickable(webElement));
			APPLICATION_LOGS.debug("it is clickable............");
			return true;
		} catch (Exception e) {
			e.printStackTrace();
			return false;
		}
	}

	public boolean waitelementToBeClickable(WebElement webElement, int waitTime) {

		try {
			WebDriverWait wait = new WebDriverWait(driver, waitTime);
			wait.until((java.util.function.Function)ExpectedConditions.elementToBeClickable(webElement));
			return true;
		} catch (Exception e) {
			return false;
		}
	}

	/***********************************************************************************************/
	public boolean waitvisibilityOfElement(WebElement webElement, int waitTime) {
		if (webElement != null) {
			try {
				WebDriverWait wait = new WebDriverWait(driver, waitTime);
				wait.until((java.util.function.Function)ExpectedConditions.visibilityOf(webElement));
				return true;
			} catch (Exception e) {
				return false;
			}
		} else
			APPLICATION_LOGS.error("PageElement " + webElement + " not exist");
		return false;
	}

	/***********************************************************************************************/

	public boolean WaitelementToBeSelected(WebElement webElement) {
		if (webElement != null) {
			try {
				WebDriverWait wait = new WebDriverWait(driver, 20);
				wait.until((java.util.function.Function)ExpectedConditions.elementToBeSelected(webElement));
				return true;
			} catch (Exception e) {
				return false;
			}
		} else
			APPLICATION_LOGS.error("PageElement " + webElement + " not exist");
		return false;
	}

	/***********************************************************************************************/

	public boolean WaittextToBePresentInElementValue(WebElement webElement, String item) {
		if (item != null) {
			try {
				WebDriverWait wait = new WebDriverWait(driver, 20);
				wait.until((java.util.function.Function)ExpectedConditions.textToBePresentInElementValue(webElement, item));
				return true;
			} catch (Exception e) {
				return false;
			}
		} else
			APPLICATION_LOGS.error("PageElement " + webElement + " not exist");
		return false;
	}

	/**************************************************************************************************/
	public boolean waitForAlert(WebDriver driver) {
		boolean isPresent = false;
		try {
			WebDriverWait wait = new WebDriverWait(driver, 10);
			wait.until((java.util.function.Function)ExpectedConditions.alertIsPresent());
			APPLICATION_LOGS.debug("Alert Occured");
			isPresent = true;
		} catch (Exception e) {
			APPLICATION_LOGS.debug("Alert not occured");
		}
		return isPresent;
	}

	protected void waitelementUntilTitleIs(String title, int waitTime) {
		WebDriverWait wait = new WebDriverWait(driver, waitTime);
		wait.until((java.util.function.Function)ExpectedConditions.titleIs(title));
	}

	public void ddlIsDisplayed(WebElement element) {
		Select select = new Select(element);
		WebElement option = select.getFirstSelectedOption();
		APPLICATION_LOGS.debug("ddl--------------------------" + option.getText());
		extentTest.log(LogStatus.PASS, "ddl--------------------------<b>" + option.getText() + "</b>");
	}

	public void validateDDLFieldIsDisplayed(WebElement element, String assertMessage) {
		boolean result = element.isDisplayed();
		APPLICATION_LOGS.debug(result);
		Select select = new Select(element);
		WebElement option = select.getFirstSelectedOption();
		APPLICATION_LOGS.debug("ddl--------------------------" + option.getText());
		extentTest.log(LogStatus.PASS, "ddl--------------------------<b>" + option.getText() + "</b>");
		APPLICATION_LOGS.debug("Full Name Is Displayed");
		Assert.assertTrue(element.isDisplayed(), assertMessage);
	}

	public void validateTextFieldIsDisplayed(WebElement element, String assertMesage) {
		boolean result = element.isDisplayed();
		APPLICATION_LOGS.debug(result);
		String text = element.getAttribute("value");
		APPLICATION_LOGS.debug(text);
		APPLICATION_LOGS.debug("Full Name Is Displayed");
		Assert.assertTrue(element.isDisplayed(), assertMesage);
	}

	
	public void F5window(String inputData) throws InterruptedException {

		Thread.sleep(2000);

		driver.switchTo().activeElement().sendKeys(Keys.F5);
		driver.switchTo().frame("helpframe");
		driver.findElement(By.id("searchtypebut" + "")).click();
		WebElement search_box = driver.findElement(By.name("txtname"));
		setData(search_box, inputData);

		driver.switchTo().activeElement().sendKeys(Keys.ENTER);
		WebElement F5windowpath = driver
				.findElement(By.xpath("//form[@name='menuform']//div//table//tbody//tr//td[2]"));
		Thread.sleep(3000);
		// driver.switchTo().

		Actions action = new Actions(driver).doubleClick(F5windowpath);

		action.build().perform();
		Thread.sleep(3000);
		driver.switchTo().activeElement().sendKeys(Keys.ENTER);
		driver.switchTo().defaultContent();

	}

	public void F5window() throws InterruptedException {

		Thread.sleep(2000);

		driver.switchTo().activeElement().sendKeys(Keys.F5);
		driver.switchTo().frame("helpframe");
		driver.findElement(By.id("searchtypebut" + "")).click();
		// WebElement search_box = driver.findElement(By.name("txtname"));
		// setData(search_box,inputData);

		driver.switchTo().activeElement().sendKeys(Keys.ENTER);
		WebElement F5windowpath = driver
				.findElement(By.xpath("//form[@name='menuform']//div//table//tbody//tr//td[2]"));
		Thread.sleep(3000);
		// driver.switchTo().

		// driver.findElement(By.xpath("//div[@id='hlpTab']//thead//tbody//tr//td[2]").doubleClick());
		Actions action = new Actions(driver).doubleClick(F5windowpath);

		action.build().perform();
		Thread.sleep(3000);
		driver.switchTo().activeElement().sendKeys(Keys.ENTER);
		driver.switchTo().defaultContent();

	}

	public boolean isElementPresent(String xPath) {
		boolean status = false;
		try {
			if (driver.findElement(By.xpath(xPath)).isDisplayed())
				status = true;
			APPLICATION_LOGS.debug("element is displaying");
			extentTest.log(LogStatus.PASS, "<b>element is displaying</b>");
		} catch (Exception e) {
			status = false;
			APPLICATION_LOGS.debug("element is NOT displaying");
			extentTest.log(LogStatus.PASS, "<b>element is NOT displaying</b>");
		}
		return status;
	}

	public boolean WaitelementToBeClickable(WebElement webElement) {

		try {
			WebDriverWait wait = new WebDriverWait(driver, 30);
			wait.until((java.util.function.Function)ExpectedConditions.elementToBeClickable(webElement));
			APPLICATION_LOGS.debug("it is clickable............");
			return true;
		} catch (Exception e) {
			return false;
		}
	}

	

	/*********************************************************************************************************/

	public void checkAlert() {
		try {
			WebDriverWait wait = new WebDriverWait(driver, 2);
			wait.until((java.util.function.Function)ExpectedConditions.alertIsPresent());
			Alert alert = driver.switchTo().alert();
			alert.accept();
		} catch (Exception e) {
			// exception handling
		}
	}

	/******************************************************************************************************************/

	public void setDataForTab(WebElement element, String inputdata) {
		if (inputdata != "") {
			try {
				element.clear();
				element.sendKeys(inputdata);
				Thread.sleep(1000);
				element.sendKeys(Keys.RETURN);
				APPLICATION_LOGS.debug("Entered " + inputdata);
				extentTest.log(LogStatus.PASS, "Entered <b>" + inputdata + "</b>");

			} catch (NoSuchElementException e1) {
				APPLICATION_LOGS.debug(element + " not found to enter data");
				extentTest.log(LogStatus.FAIL, "<b>" + element + "</b> not found to enter data");
				Assert.fail(element + " NOT FOUND TO ENTER " + inputdata);
			} catch (ElementNotVisibleException e2) {
				APPLICATION_LOGS.debug(element + " not visible to enter data");
				extentTest.log(LogStatus.FAIL, "<b>" + element + "</b> not visible to enter data");
				Assert.fail(element + " NOT VISIBLE TO ENTER " + inputdata);
			} catch (Exception e3) {
				APPLICATION_LOGS.debug("UNKNOWN EXCEPTION,CHECK CONSOLE");
				extentTest.log(LogStatus.FAIL, "<b>UNKNOWN EXCEPTION,CHECK CONSOLE</b>");
				e3.printStackTrace();
				Assert.fail("UNKNOWN EXCEPTION,CHECK CONSOLE");

			}
		} else {
			element.sendKeys(Keys.RETURN);
		}
	}

	/*****************************************************************************************************/

	/*****************************************************************************************************/

	public String getErrorMessage() {
		String errorMessage = "";
		try {
			errorMessage = driver.findElement(By.id("errmsg")).getText();
		} catch (NoSuchElementException e1) {
			return errorMessage;
		}
		return errorMessage;
	}
	
	public String ebank_GetErrorMessage() {
		String errorMessage = "";
		try {
			errorMessage = driver.findElement(By.xpath("//span[@class='errorMessage ng-star-inserted']")).getText();
		} catch (NoSuchElementException e1) {
			return errorMessage;
		}
		return errorMessage;
	}
	
	

	/*********************************************************************************************************/
	public void fieldValidation(String screenShotName, String sheetName, int dataRowNo, Xls_Reader xls) {
		if (getErrorMessage().length() != 0) {
			// Take screenshot with date

			String Url = FilepathsUtility.getScreenshotPath()+sheetName+".png";

			APPLICATION_LOGS.debug(Url);
			ReportLogAs("Error message : " + getErrorMessage(), "FAILED");
			APPLICATION_LOGS.debug("Error message : " + getErrorMessage());
			extentTest.log(LogStatus.ERROR, "Error message :  <b>" + getErrorMessage() + "</b>");
			getScreenShot(sheetName,sheetName+"_"+screenShotName + getCurrentDate());
			Reporter.log("Error message : " + getErrorMessage());
			Reporter.log("Error Screenshot :" + Url);
			xls.setCellData(sheetName, "ErrorMessage", dataRowNo, getErrorMessage());
			xls.addHyperLink(sheetName, "failedScreenShot", sheetName, dataRowNo, Url,
					screenShotName + getCurrentDate());
			xls.setCellData(sheetName, screenShotName + "failedScreenShot", dataRowNo, getErrorMessage(), Url);
			// APPLICATION_LOGS.debug("Error message : "+getErrorMessage());

			Assert.fail(getErrorMessage());
		}

	}
	
	/*********************************************************************************************************/
	public void ebank_FieldValidation(String screenShotName, String sheetName, int dataRowNo, Xls_Reader xls) {
		if (ebank_GetErrorMessage().length() != 0) {
			// Take screenshot with date

			String Url = FilepathsUtility.getScreenshotPath()+sheetName+".png";

			APPLICATION_LOGS.debug(Url);
			ReportLogAs("Error message : " + ebank_GetErrorMessage(), "FAILED");
			APPLICATION_LOGS.debug("Error message : " + ebank_GetErrorMessage());
			extentTest.log(LogStatus.ERROR, "Error message :  <b>" + ebank_GetErrorMessage() + "</b>");
			getScreenShot(sheetName,sheetName+"_"+screenShotName + getCurrentDate());
			Reporter.log("Error message : " + ebank_GetErrorMessage());
			Reporter.log("Error Screenshot :" + Url);
			xls.setCellData(sheetName, "ErrorMessage", dataRowNo, ebank_GetErrorMessage());
			xls.addHyperLink(sheetName, "failedScreenShot", sheetName, dataRowNo, Url,
					screenShotName + getCurrentDate());
			xls.setCellData(sheetName, screenShotName + "failedScreenShot", dataRowNo, ebank_GetErrorMessage(), Url);
			// APPLICATION_LOGS.debug("Error message : "+getErrorMessage());

			Assert.fail(ebank_GetErrorMessage());
		}

	}
	
	

	/*********************************************************************************************************/
	public void alertAccept_Robot() throws Exception {
		Thread.sleep(5000);
		String Msg = acceptAlert(driver);
		if (Msg.equals(""))
			clickOnRecordUpdate();
	}

	// To return the number
	public String acceptAlertAutoIt(WebDriver driver) throws InterruptedException {
		String msg = "";

		try {
			Alert alert = driver.switchTo().alert();
			msg = alert.getText();
			alert.accept();
			APPLICATION_LOGS.debug("Alert Msg:" + msg + "\n");
			extentTest.log(LogStatus.PASS, " Generated Messages is:<b>"+ msg + "</b>");
		} catch (NoAlertPresentException e) {
			String jvmBitVersion = System.getProperty("sun.arch.data.model");
			String jacobDllVersionToUse;
			if (jvmBitVersion.contains("32")) {
				jacobDllVersionToUse = "jacob-1.18-x86.dll";
			} else {
				jacobDllVersionToUse = "jacob-1.18-x64.dll";
			}
			File file = new File("Tools\\autoit", jacobDllVersionToUse);
			System.setProperty(LibraryLoader.JACOB_DLL_PATH, file.getAbsolutePath());
			AutoItX x=new AutoItX();
			int count = 0;
			String var = null;

			while (count < 4) {
				x.winActivate("Message from webpage");
				x.winWaitActive("Message from webpage", "", 5);
				Thread.sleep(2000);
				var = x.controlGetText("Message from webpage", "", "2");
				msg = x.controlGetText("Message from webpage", "", "65535");
				APPLICATION_LOGS.debug("Msg in alert box...." + msg);
				extentTest.log(LogStatus.PASS, " Generated Messages is:<b>"+ msg + "</b>");
				APPLICATION_LOGS.debug("Button to be clicked upon : " + var);
				if (var.equals("OK")) {
					x.controlClick("Message from webpage", "", "2");
					count = 4;
				} else {
					count = count + 1;
				}
			}
		}
		return msg;
	}

	public String acceptAlertAutoIt_EOD(WebDriver driver) throws InterruptedException {
		String msg = "";

		try {
			String jvmBitVersion = System.getProperty("sun.arch.data.model");
			String jacobDllVersionToUse;
			if (jvmBitVersion.contains("32")) {
				jacobDllVersionToUse = "jacob-1.18-x86.dll";
			} else {
				jacobDllVersionToUse = "jacob-1.18-x64.dll";
			}
			File file = new File("Tools\\autoit", jacobDllVersionToUse);
			System.setProperty(LibraryLoader.JACOB_DLL_PATH, file.getAbsolutePath());
			AutoItX x = new AutoItX();
			int count = 0;
			String var = null;

			while (count < 4) {
				x.winActivate("Message from webpage");
				x.winWaitActive("Message from webpage", "", 5);
				Thread.sleep(2000);
				var = x.controlGetText("Message from webpage", "", "2");
				msg = x.controlGetText("Message from webpage", "", "65535");
				APPLICATION_LOGS.debug("Msg in alert box...." + msg);
				APPLICATION_LOGS.debug("Button to be clicked upon : " + var);
				if (var.equals("Cancel")) {
					x.controlClick("Message from webpage", "", "2");
					count = 4;
				} else {
					count = count + 1;
				}
			}
		} catch (NoAlertPresentException e) {
			String jvmBitVersion = System.getProperty("sun.arch.data.model");
			String jacobDllVersionToUse;
			if (jvmBitVersion.contains("32")) {
				jacobDllVersionToUse = "jacob-1.18-x86.dll";
			} else {
				jacobDllVersionToUse = "jacob-1.18-x64.dll";
			}
			File file = new File("Tools\\autoit", jacobDllVersionToUse);
			System.setProperty(LibraryLoader.JACOB_DLL_PATH, file.getAbsolutePath());
			AutoItX x = new AutoItX();
			int count = 0;
			String var = null;

			while (count < 4) {
				x.winActivate("Message from webpage");
				x.winWaitActive("Message from webpage", "", 5);
				Thread.sleep(2000);
				var = x.controlGetText("Message from webpage", "", "2");
				msg = x.controlGetText("Message from webpage", "", "65535");
				APPLICATION_LOGS.debug("Msg in alert box...." + msg);
				APPLICATION_LOGS.debug("Button to be clicked upon : " + var);
				if (var.equals("Cancel")) {
					x.controlClick("Message from webpage", "", "2");
					count = 4;
				} else {
					count = count + 1;
				}
			}
		}
		return msg;
	}

	public static String addDays(String effDate, int days) throws ParseException {
		Calendar c = Calendar.getInstance();
		SimpleDateFormat sdf = new SimpleDateFormat("dd-MM-yyyy");
		Date date = sdf.parse(effDate);
		c.setTime(date);
		c.add(Calendar.DATE, days);
		return sdf.format(c.getTime());
	}

	// Reporting function
	public static void ReportLogAs(String msg, String Status) {
		LogAs status = LogAs.INFO;
		if (Status.equalsIgnoreCase("INFO"))
			status = LogAs.INFO;
		if (Status.equalsIgnoreCase("PASS") || Status.equalsIgnoreCase("PASSED") || Status.equalsIgnoreCase("true"))
			status = LogAs.PASSED;
		if (Status.equalsIgnoreCase("FAIL") || Status.equalsIgnoreCase("FAILED") || Status.equalsIgnoreCase("false"))
			status = LogAs.FAILED;
		ATUReports.add(msg, "Testing", "--", "--", status, null);
	}
	
	
	
	
	//Uploading Files by using Robot
		public void uploadFile(String fileLocation) {
	        try {
	        	//Setting clipboard with file location
	            setClipboardData(fileLocation);
	            //native key strokes for CTRL, V and ENTER keys
	           Thread.sleep(3000);
	            Robot r = new Robot();
	            r.keyPress(KeyEvent.VK_CONTROL);
				r.keyPress(KeyEvent.VK_V);
				r.keyRelease(KeyEvent.VK_CONTROL);
				r.keyRelease(KeyEvent.VK_V);
	            APPLICATION_LOGS.info("Before Thread");
	            Thread.sleep(2000);
	            r.keyPress(KeyEvent.VK_ENTER);
	            r.keyRelease(KeyEvent.VK_ENTER);
	            APPLICATION_LOGS.info("After Thread");
	        } catch (Exception exp) {
	        	exp.printStackTrace();
	        }
		}
		public static void setClipboardData(String string) {
			APPLICATION_LOGS.info("FILE PATH ---> "+string);
			//StringSelection is a class that can be used for copy and paste operations.
			   StringSelection stringSelection = new StringSelection(string);
			   Toolkit.getDefaultToolkit().getSystemClipboard().setContents(stringSelection, null);
			}

		public static String ebankaddDays(String effDate, int days,int year) throws ParseException {
			Calendar c = Calendar.getInstance();
			SimpleDateFormat sdf = new SimpleDateFormat("dd/mm/yyyy");
			Date date = sdf.parse(effDate);
			c.setTime(date);
			c.add(Calendar.DATE, days);
			c.add(Calendar.YEAR, year);
			return sdf.format(c.getTime());
		}
		
		public static String ebankBeforeDays(String effDate, int days) throws ParseException {
			Calendar c = Calendar.getInstance();
			SimpleDateFormat sdf = new SimpleDateFormat("dd/mm/yyyy");
			Date date = sdf.parse(effDate);
			c.setTime(date);
			c.add(Calendar.DATE,days);
			//c.add(Calendar.YEAR, 1);
			return sdf.format(c.getTime());
		}
		 public static String getCurrentInterfaceSystemDate() {
			 
					DateFormat originalFormat = new SimpleDateFormat("dd/MM/yyyy");
					Date date = new Date();
					return originalFormat.format(date);
		}
		 public static String getCurrentInterfaceSystemDate1() {
			 
				DateFormat originalFormat = new SimpleDateFormat("dd-MM-yyyy");
				Date date = new Date();
				return originalFormat.format(date);
	}
		 public static String getCurrentInterfaceSystemDate2() {
			 
				DateFormat originalFormat = new SimpleDateFormat("YYYYMMDD");
				Date date = new Date();
				return originalFormat.format(date);
	}
		 
		 public String getsysdate(String idateformat) {
				Date t = new Date();
				DateFormat dateformat = new SimpleDateFormat(idateformat);
				return dateformat.format(t);
			}
		 public static String getCurrentInterfaceSystemDate3() {
			 
				DateFormat originalFormat = new SimpleDateFormat("DDMMYYYY");
				Date date = new Date();
				return originalFormat.format(date);
	}
		 public String getTimeStamp() throws ParseException {
				Date date = new Date();
				SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd'T'HH:mm:ss");
				return sdf.format(date);
			}
	public void reportlog(String message,String flag)
		 {
			
		if (flag.equalsIgnoreCase("pass")) {
			APPLICATION_LOGS.info(message);
			ReportLogAs(message, "PASS");
			extentTest.log(LogStatus.PASS, message);			
		} 
		else if (flag.equalsIgnoreCase("fail")) {
			APPLICATION_LOGS.info(message);
			ReportLogAs(message, "FAIL");
			extentTest.log(LogStatus.FAIL, message);
		} 
		else if (flag.equalsIgnoreCase("info")) {
			APPLICATION_LOGS.info(message);
			ReportLogAs(message,"INFO");
			extentTest.log(LogStatus.INFO, message);
		} 
		else if (flag.equalsIgnoreCase("error")) {			
			ReportLogAs(message, "INFO");			
			extentTest.log(LogStatus.ERROR, message);			
			APPLICATION_LOGS.debug("Error message : " +message);
		} 
		else if (flag.equalsIgnoreCase("skip")) {
			APPLICATION_LOGS.info(message);
			ReportLogAs(message, "INFO");				
			extentTest.log(LogStatus.SKIP, message);		
		} 
		else {
			APPLICATION_LOGS.info(message);
			ReportLogAs(message, "INFO");			
			extentTest.log(LogStatus.UNKNOWN, message);
			APPLICATION_LOGS.info("FILE PATH ---> " + message);
		}
		
			
		 }
}
