package com.testing.ts.util;

public class FilepathsUtility {

	/*public static final String currentDateTime=new SimpleDateFormat("yyyy-MMM-dd hh-mm-ss").format(new Date());
	public static final String REPORTS = "\\Reports";
	public static final String RUN = "\\RUN_" + currentDateTime;
	private static final String SRC_COM_INTELLECT_CORE_TESTOUTPUT_SCREENSHOTS = "\\screenshots\\";*/
	
	public static boolean checkToRunUtility(Read_XLS filePath, String sheetName, String ToRun, String testSuite){
				
		boolean Flag = false;		
		if(filePath.retrieveToRunFlag(sheetName,ToRun,testSuite).equalsIgnoreCase("y")){
			Flag = true;
		}
		return Flag;		
	}
	
	public static String[] checkToRunUtilityOfData(Read_XLS xls, String sheetName, String ColName){		
		return xls.retrieveToRunFlagTestData(sheetName,ColName);		 	
	}
 
	public static Object[][] GetTestDataUtility(Read_XLS xls, String sheetName){
	 	//System.out.println("Sheeet name "+sheetName);
		return xls.retrieveTestData(sheetName);	
	}
 
	public static boolean WriteResultUtility(Read_XLS xls, String sheetName, String ColName, int rowNum, String Result){			
		return xls.writeResult(sheetName, ColName, rowNum, Result);		 	
	}
 
	public static boolean WriteResultUtility(Read_XLS xls, String sheetName, String ColName, String rowName, String Result){			
		return xls.writeResult(sheetName, ColName, rowName, Result);		 	
	}
	
	public static String getLog4jPath()
	{
		return System.getProperty("user.dir")+"\\src\\log4j.properties";
	}
	
	public static String getProductLogPath()
	{
		return System.getProperty("user.dir")+BaseClass.REPORTS+BaseClass.RUN+"\\logs\\ProductLogs.log";
	}
	
	public static String getExtentReportPath()
	{
		return System.getProperty("user.dir")+BaseClass.REPORTS+BaseClass.RUN+"\\ExtentReports\\" +BaseClass.currentDateTime+ "reports.html";
	}
	
	/*public static String getScreenshotPath()
	{
		return System.getProperty("user.dir")+REPORTS+RUN
				+ SRC_COM_INTELLECT_CORE_TESTOUTPUT_SCREENSHOTS;
	}*/
	
	public static String getScreenshotPath()
	{
		return System.getProperty("user.dir").replace("\\", "\\\\")+BaseClass.REPORTS+BaseClass.RUN+"\\screenshots\\\\";
	}
	
	public static String getCurrentRunPath()
	{
		return System.getProperty("user.dir")+getFileSeparator()+BaseClass.REPORTS+BaseClass.RUN;
	}
	
	public static String getFileSeparator() {
		return System.getProperty("file.separator");
	}
	
}

