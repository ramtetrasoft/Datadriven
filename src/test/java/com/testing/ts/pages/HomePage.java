package com.testing.ts.pages;

import org.openqa.selenium.By;
import org.openqa.selenium.ElementNotVisibleException;
import org.openqa.selenium.NoSuchElementException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;

import com.relevantcodes.extentreports.LogStatus;
import com.testing.ts.hooks.BrowserOperations;

public class HomePage extends BrowserOperations{
	public HomePage(WebDriver driver)
	{
		super(driver);
		// TODO Auto-generated constructor stub
	}

	private WebElement element = null;
	public WebElement contactUs()
	{
		element = driver.findElement(By.xpath("//a[text()='Contact us']"));
		return element;
	}
	public WebElement signin()
	{
		element = driver.findElement(By.xpath("//a[contains(text(),'Sign in')]"));
		return element;
	}

	public void clickOn_ContactUs() throws Exception
	{	
		try{
			Thread.sleep(2000);
			waitvisibilityOfElement(contactUs(),5);
			
			contactUs().click();
			
			Thread.sleep(1000);
			
			extentTest.log(LogStatus.PASS, "Clicked On <b>"+"Contact Us"+"</b> Button successfully ********");
			
			 APPLICATION_LOGS.debug("Clicked On Contact Us Button successfully");

			}
		catch (NoSuchElementException e1)
		{
			extentTest.log(LogStatus.FAIL, "Clicked On <b>"+"Contact Us"+"</b> Button successfully ********");
			
			APPLICATION_LOGS.debug("Contact Us element not found to click On it....");
		}catch (ElementNotVisibleException e2) {
			APPLICATION_LOGS.debug("Contact Us element not found to click On it....");
		}catch(Exception e){
			e.printStackTrace();
			APPLICATION_LOGS.debug("Exception occured");
		}
		
	}
	public void clickOn_signin() throws Exception
	{	
		try{
			Thread.sleep(2000);
			waitvisibilityOfElement(signin(),5);
			
			signin().click();
			
			Thread.sleep(1000);
			
			extentTest.log(LogStatus.PASS, "Clicked On <b>"+"Sign in"+"</b> Button successfully ********");
			
			 APPLICATION_LOGS.debug("Clicked On Sign in Button successfully");

			}
		catch (NoSuchElementException e1)
		{
			extentTest.log(LogStatus.FAIL, "Clicked On <b>"+"Sign in"+"</b> Button successfully ********");
			
			APPLICATION_LOGS.debug("Sign in element not found to click On it....");
		}catch (ElementNotVisibleException e2) {
			APPLICATION_LOGS.debug("Sign in element not found to click On it....");
		}catch(Exception e){
			e.printStackTrace();
			APPLICATION_LOGS.debug("Exception occured");
		}
		
	}

}
