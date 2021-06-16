package com.testing.ts.scripts;

import java.util.Hashtable;

import org.openqa.selenium.WebDriver;
import org.testng.annotations.Test;

import com.relevantcodes.extentreports.LogStatus;
import com.testing.ts.hooks.TestNGHooks;
import com.testing.ts.pages.LoginPage;

public class LoginTestRunner extends TestNGHooks{
		//object creation
				public LoginPage login;
				public static final String Login="The Internet";
			//	public static final String ContactUs="Contact us - My Store";
				
				/*********************************************************************/
				
				private final String USERID="userId";
				private final String PASSWORD="password";

				@Test(dataProvider = "getDetails")
				public void testscript(Hashtable<String, String> inputData)
						throws Exception {
					
					WebDriver driver=isTestcaseExecutable(inputData, CONFIG.getProperty("BROWSER"));		
					login=new LoginPage(driver, testCaseNo, this.getClass().getSimpleName(), dataRowNo, DEMOgetVal);
					
					extentTest.log(LogStatus.INFO, "********* Execution Starting for <b>"+Login+"</b> Screen ********");
					getScreenshot(this.getClass().getSimpleName(),testCaseNo, dataRowNo, DEMOgetVal);
					Thread.sleep(3000);
					login.fillUserId(inputData.get(USERID));
					login.fillPassword(inputData.get(PASSWORD));
					login.clicklogin();
					try {
					login.logout().isDisplayed();
					{pass=true;}
					}catch(Exception ex)
					{pass=false;}
					extentTest.log(LogStatus.INFO, "********* Execution Completed for <b>"+Login+"</b> Program Successfully ********");
					getScreenshot(this.getClass().getSimpleName(),testCaseNo, dataRowNo, DEMOgetVal);

					APPLICATION_LOGS.debug("Test case passes successfully\n_______________________________________________________________");	
					}
				}