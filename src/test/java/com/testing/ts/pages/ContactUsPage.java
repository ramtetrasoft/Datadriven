package com.testing.ts.pages;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;

import com.testing.ts.util.Xls_Reader;

public class ContactUsPage extends HomePage{
	
	private WebElement element=null;
	public String screenshotName=null;
	public String sheetName=null;
	public int dataRowNo;
	Xls_Reader xls;
	
	public ContactUsPage (WebDriver driver,String screenshotName,String sheetName,int dataRowNo,Xls_Reader xls) 
	{
		super(driver);
		
		this.screenshotName=screenshotName;
		this.sheetName=sheetName;
		this.dataRowNo=dataRowNo;
		this.xls=xls;
	}
	
	public WebElement subjectHeading() {
		element = driver.findElement(By.id("id_contact"));
		return element;
	}
	
	public WebElement emailAddress() {
		element = driver.findElement(By.id("email"));
		return element;
	}
	
	public WebElement Ordereference() {
		element = driver.findElement(By.id("id_order"));
		return element;
	}
	
	public WebElement message() {
		element = driver.findElement(By.id("message"));
		return element;
	}
	
	public WebElement send() {
		element = driver.findElement(By.id("submitMessage"));
		return element;
	}
	////p[contains(text(),'successfully sent')]
	public WebElement successfullySent() {
		element = driver.findElement(By.xpath("//p[contains(text(),'successfully sent')]"));
		return element;
	}
	
	public void select_subjectHeading(String inputdata){
		selectFromDDL(subjectHeading(), inputdata);
	}
	public void set_emailAddress(String inputdata){
		setData(emailAddress(), inputdata);
	}
	
	public void set_Ordereference(String inputdata){
		setData(Ordereference(), inputdata);
	}
	
	public void set_message(String inputdata){
		setData(message(), inputdata);
	}
	
	public void clickon_Send(){
		clickOn(send());
	}
	
		 
}
