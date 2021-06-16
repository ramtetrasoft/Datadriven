package com.testing.ts.pages;

import org.apache.log4j.Logger;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;

import com.testing.ts.util.Xls_Reader;

public class LoginPage extends HomePage
{
	private WebElement element=null;
	public String screenshotName=null;
	public String sheetName=null;
	public int dataRowNo;
	Xls_Reader xls;
	public LoginPage (WebDriver driver,String screenshotName,String sheetName,int dataRowNo,Xls_Reader xls) 
	{
		super(driver);
		
		this.screenshotName=screenshotName;
		this.sheetName=sheetName;
		this.dataRowNo=dataRowNo;
		this.xls=xls;
	}
	static Logger log = Logger.getLogger(LoginPage.class);
		
	public WebElement userName() {
		element = driver.findElement(By.name("username"));
		return element;
	}
	public WebElement password() {
		element = driver.findElement(By.name("password"));
		return element;
	}
		
	public WebElement login() {
		element = driver.findElement(By.xpath("//button[@type='submit']"));
		return element;
	}
	public WebElement logout() {
		element = driver.findElement(By.xpath("//i[text()=' Logout']"));
		return element;
	}
	public WebElement usrnameInvalid() {
		element = driver.findElement(By.xpath("//div[contains(text(),'username is invalid!')]"));
		return element;
	}
	public WebElement passwordInvalid() {
		element = driver.findElement(By.xpath("//div[contains(text(),'password is invalid!')]"));
		return element;
	}
		//Operational methods
		public String fillUserId(String x)
		{
			userName().sendKeys(x);
			log.info("fillUserId function is executed");
			return("Done");
		}
		
		public String fillPassword(String x)
		{
			password().sendKeys(x);
			log.info("fillPassword function is executed");
			return("Done");
		}
		
		public String clicklogin()
		{
			login().click();
			log.info("clicklogin function is executed");
			return("Done");
		}

}
