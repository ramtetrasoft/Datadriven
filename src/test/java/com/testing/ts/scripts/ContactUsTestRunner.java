package com.testing.ts.scripts;

import java.util.Hashtable;

import org.testng.annotations.Test;

import com.relevantcodes.extentreports.LogStatus;
import com.testing.ts.hooks.TestNGHooks;
import com.testing.ts.pages.ContactUsPage;

public class ContactUsTestRunner extends TestNGHooks{
	
	//object creation
		public ContactUsPage poc;
		//public static final String login="Login - My Store";
		public static final String ContactUs="Contact us - My Store";
		
	/*********************************************************************/
		
		private final String SUBJECTHEADING="subjectHeading";
		private final String EMAIL="emailAddress";
		private final String ORDERREF="Ordereference";
		private final String MESSAGE="message";

	/*********************************************************************/
	@Test(dataProvider = "getDetails")
	public void testscript(Hashtable<String, String> inputData)
			throws Exception {
		
		isTestcaseExecutable(inputData, CONFIG.getProperty("BROWSER"));		
		poc=new ContactUsPage(driver, testCaseNo, this.getClass().getSimpleName(), dataRowNo, DEMOgetVal);
		
		extentTest.log(LogStatus.INFO, "********* Execution Starting for <b>"+ContactUs+"</b> Screen ********");
		NavigateTo(ContactUs);
		getScreenshot(this.getClass().getSimpleName(),testCaseNo, dataRowNo, DEMOgetVal);
		poc.select_subjectHeading(inputData.get(SUBJECTHEADING));
		poc.set_emailAddress(inputData.get(EMAIL));
		poc.set_Ordereference(inputData.get(ORDERREF));
		poc.set_message(inputData.get(MESSAGE));
		poc.clickon_Send();
		if(poc.successfullySent().isDisplayed())
		{pass=true;}
		else
		{pass=false;}
		//driver.close();
		extentTest.log(LogStatus.INFO, "********* Execution Completed for <b>"+ContactUs+"</b> Program Successfully ********");
		getScreenshot(this.getClass().getSimpleName(),testCaseNo, dataRowNo, DEMOgetVal);
		
		APPLICATION_LOGS.debug("Test case passes successfully\n_______________________________________________________________");	
		}
	}
