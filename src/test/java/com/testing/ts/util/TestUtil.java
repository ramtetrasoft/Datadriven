package com.testing.ts.util;

import java.io.File;
import java.io.IOException;
import java.util.Hashtable;

import org.apache.commons.io.FileUtils;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;

public class TestUtil {

	BaseClass pom=new BaseClass();
	/**********************************************************************************************************************************/	
		//Checking the Run mode for TestCase and skipping the test case, if Run mode set to "N"
		// true- test has to be executed
		// false- test has to be skipped
		public static boolean isExecutable(String testName, Xls_Reader xls){
			
			for(int rowNum=2;rowNum<=xls.getRowCount("TestCases");rowNum++){
				
				if(xls.getCellData("TestCases", "TSID", rowNum).equals(testName)){
				
					if(xls.getCellData("TestCases", "Runmode", rowNum).equals("Y"))
					{
					//	System.out.println("Run mode for "+testName+" is :Y");
						return true;
					}
						
					else 
						return false;
				}
				// print the test cases with RUnmode Y
			}
			
			return false;
		}
	/**********************************************************************************************************************************/	
		public void takeScreenShot(String fileName) {
			File srcFile = ((TakesScreenshot)(pom.driver)).getScreenshotAs(OutputType.FILE);
		    try {
		    	FileUtils.copyFile(srcFile, new File(System.getProperty("user.dir")+"\\Screenshots\\"+fileName+".jpg"));
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}
		
	/**********************************************************************************************************************************/	
		public static String return_null(String val){
			if(val!=""){
			   return val;
			}
			return "no_data";
		}
		
		// To Fetch data from given xls, Sheet, test name
		public static Object[][] getData(String testName,Xls_Reader xls){
			// find the row num from which test starts
			// number of cols in the test
			// number of rows
			// put the data in hastable and put hastable in array
			
			int testStartRowNum=0;
			// find the row num from which test starts
			for(int rNum=1;rNum<=xls.getRowCount(testName);rNum++){
				if(xls.getCellData(testName, 0, rNum).equals(testName)){
					testStartRowNum=rNum;
					break;
				}
			}
		//	System.out.println("Test "+ testName +" starts from "+ testStartRowNum);
			
			int colStartRowNum=testStartRowNum+2;
			int totalCols=0;
			while(!xls.getCellData(testName, totalCols, colStartRowNum).equals("")){
				totalCols++;
			}
			System.out.println("Total Cols in test "+ testName + " are "+ totalCols);
			
			//rows
			int dataStartRowNum=testStartRowNum+3;
			int totalRows=0;
			while(!xls.getCellData(testName, 0, dataStartRowNum+totalRows).equals("")){
				totalRows++;
			}
			System.out.println("Total Rows in test "+ testName + " are "+ totalRows);

			System.out.println("++++++++++++++++++++++++++++++++++++++++++++++++++++++++++");
			// extract data
			Object[][] data = new Object[totalRows][1];
			int index=0;
			Hashtable<String,String> table=null;
			for(int rNum=dataStartRowNum;rNum<(dataStartRowNum+totalRows);rNum++){
				table = new Hashtable<String,String>();
				for(int cNum=0;cNum<totalCols;cNum++){
					table.put(xls.getCellData(testName, cNum, colStartRowNum), xls.getCellData(testName, cNum, rNum));
					
					System.out.print(xls.getCellData(testName, cNum, colStartRowNum) + " : " + return_null(xls.getCellData(testName, cNum, rNum)) +" -- ");
					//System.out.print(xls.getCellData(testName, cNum, rNum) +" -- ");
				}
				data[index][0]= table;
				index++;
				System.out.println();
			}
			
			
			return data;
		}
		
		
		
		/**********************************************************************************************************************************/		
		
		

	}
