package com.testing.ts.hooks;

import java.util.concurrent.TimeUnit;

import org.openqa.selenium.WebDriver;

import com.testing.ts.pages.HomePage;
import com.testing.ts.util.BaseClass;

public class BrowserOperations extends BaseClass {
public static WebDriver driver;
	public BrowserOperations(WebDriver driver) {
		BrowserOperations.driver = driver;
	}

	public HomePage openUrl(String url)
	 {
	 driver.manage().timeouts().implicitlyWait(20, TimeUnit.SECONDS);
	 driver.get(url);
	 driver.manage().window().maximize();
	 APPLICATION_LOGS.debug("URL opened successfully: "+driver.getTitle());
	 return new HomePage(driver);
		
	 }
	public void closeSite()
	{
		driver.quit();
	}
}
