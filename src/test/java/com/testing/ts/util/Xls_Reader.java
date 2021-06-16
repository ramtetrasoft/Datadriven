package com.testing.ts.util;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.Hashtable;

import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;

public class Xls_Reader {

	public  String path;
	public  FileInputStream fis = null;
	public  FileOutputStream fileOut =null;
	private HSSFWorkbook workbook = null;
	private HSSFSheet sheet = null;
	private HSSFRow row   =null;
	private HSSFCell cell = null;
	
	public Xls_Reader(String path) {
		
		this.path=path;
		try {
			fis = new FileInputStream(path);
			workbook = new HSSFWorkbook(fis);
			sheet = workbook.getSheetAt(0);
			fis.close();
		} catch (Exception e) {
			e.printStackTrace();
		} 
		
	}
/**********************************************************************************************************************************/	
	
	// returns the row count in a sheet
	public int getRowCount(String sheetName){
		int index = workbook.getSheetIndex(sheetName);
		if(index==-1)
			return 0;
		else{
		sheet = workbook.getSheetAt(index);
		int number=sheet.getLastRowNum()+1;
		return number;
		}
		
	}
/**********************************************************************************************************************************/	

	public int columnIndex(String sheetName,String colName,String identifier)
	{
        int dataRowNo=0;	

		int index = workbook.getSheetIndex(sheetName);
		sheet = workbook.getSheetAt(index);
		int col_Num=-1;
		for(int rNum=1;rNum<=getRowCount(sheetName);rNum++){
			if(getCellData(sheetName, 0, rNum).equals(identifier)){
				dataRowNo=rNum;
				/*System.out.println(dataRowNo);*/
				break;
			}
		}
		
		row=sheet.getRow(dataRowNo+1);
		
		for(int i=0;i<row.getLastCellNum();i++)
		{
			//System.out.println(row.getCell(i).getStringCellValue().trim());
			/*System.out.println(row.getCell(i).getStringCellValue());*/
			if(row.getCell(i).getStringCellValue().trim().equals(colName.trim()))
				col_Num=i;
		}
		
		return col_Num;
	}
	
	/**********************************************************************************************************************************/	

	public int columnIndex_NoMethod(String sheetName,String colName)
	{
		int index = workbook.getSheetIndex(sheetName);
		sheet = workbook.getSheetAt(index);
		int col_Num=-1;
		row=sheet.getRow(2);
		for(int i=0;i<row.getLastCellNum();i++)
		{
			/*System.out.println(row.getCell(i).getStringCellValue().trim());
			System.out.println(colName.trim());*/
			if(row.getCell(i).getStringCellValue().trim().equals(colName.trim()))
			{
				col_Num=i;
			    break;
			}
		}
		
		return col_Num;
	}
	
	/**********************************************************************************************************************************/
	public int columnIndex_NoMethod_Response(String sheetName,String colName)
	{
		int index = workbook.getSheetIndex(sheetName);
		sheet = workbook.getSheetAt(index);
		int col_Num=-1;
		row=sheet.getRow(0);
		for(int i=0;i<row.getLastCellNum();i++)
		{
			//System.out.println(row.getCell(i).getStringCellValue().trim());
			if(row.getCell(i).getStringCellValue().trim().equals(colName.trim()))
				col_Num=i;
		}
		
		return col_Num;
	}
	
	/**********************************************************************************************************************************/
	public int ColumnIndex(Hashtable<String, String> inputData,String columnName)
	{
		int col_Num;
		
		if(inputData.get("methodName").equalsIgnoreCase(""))
		{
			col_Num=columnIndex_NoMethod(inputData.get("sheetName"), columnName);
		}
		
		else
		{
			col_Num=columnIndex(inputData.get("sheetName"),columnName,inputData.get("methodName"));
		}
		
		return col_Num;
	}
/***********************************************************************************************************************************************/

	// returns the data from a cell
	@SuppressWarnings("deprecation")
	public String getCellData(String sheetName,String colName,int rowNum)
	{
		try{
			if(rowNum <=0)
				return "";
		
		int index = workbook.getSheetIndex(sheetName);
		int col_Num=-1;
		if(index==-1)
			return ""; 
		
		sheet = workbook.getSheetAt(index);
		row=sheet.getRow(0);
		for(int i=0;i<row.getLastCellNum();i++)
		{
			//System.out.println(row.getCell(i).getStringCellValue().trim());
			if(row.getCell(i).getStringCellValue().trim().equals(colName.trim()))
				col_Num=i;
		}
		if(col_Num==-1)
			return "";
		
		sheet = workbook.getSheetAt(index);
		row = sheet.getRow(rowNum-1);
		if(row==null)
			return "";
		cell = row.getCell(col_Num);
		
		if(cell==null)
			return "";
		//System.out.println(cell.getCellType());
		if(cell.getCellType()==Cell.CELL_TYPE_STRING)
			  return cell.getStringCellValue();
		else if(cell.getCellType()==Cell.CELL_TYPE_NUMERIC || cell.getCellType()==Cell.CELL_TYPE_FORMULA ){
			  
			  String cellText  = String.valueOf(cell.getNumericCellValue());
			  if (HSSFDateUtil.isCellDateFormatted(cell)) {
		           // format in form of M/D/YY
				  double d = cell.getNumericCellValue();

				  Calendar cal =Calendar.getInstance();
				  cal.setTime(HSSFDateUtil.getJavaDate(d));
		            cellText =
		             (String.valueOf(cal.get(Calendar.YEAR))).substring(2);
		           cellText = cal.get(Calendar.DAY_OF_MONTH) + "/" +
		                      cal.get(Calendar.MONTH)+1 + "/" + 
		                      cellText;
		           
		           //System.out.println(cellText);

		         }		  
			  return cellText;
		  }else if(cell.getCellType()==Cell.CELL_TYPE_BLANK)
		      return ""; 
		  else 
			  return String.valueOf(cell.getBooleanCellValue());
		
		}
		catch(Exception e){
			
			e.printStackTrace();
			return "row "+rowNum+" or column "+colName +" does not exist in xls";
		}
	}
	
/*************************************************************************************************************************************/
	// returns the data from a cell
	public String getCellDataResponse(String sheetName,String colName)
	{
		String response="";
		try{
			
		
		sheet = workbook.getSheet(sheetName);
		
		
		for(int i=1;i<=sheet.getLastRowNum();i++)
		{
					
					    cell=row.getCell(0);
						if(cell.getStringCellValue().equalsIgnoreCase(colName))
						{
							response=row.getCell(2).getStringCellValue();
							break;
						}
			}
		
		}
		catch(Exception e){
			
			System.out.println(e);
			
		}
		return response;
	}
		
/**********************************************************************************************************************************/	

	// returns the data from a cell
	@SuppressWarnings("deprecation")
	public String getCellData(String sheetName,int colNum,int rowNum){	
		try{
			if(rowNum <=0)
				return "";
		
		int index = workbook.getSheetIndex(sheetName);

		if(index==-1)
			return "";
		//System.out.println("Index of sheet name is "+index);
	
		sheet = workbook.getSheetAt(index);
		row = sheet.getRow(rowNum-1);
		if(row==null)
			return "";
		cell = row.getCell(colNum);
		if(cell==null)
			return "";
		
	  if(cell.getCellType()==Cell.CELL_TYPE_STRING)
		  return cell.getStringCellValue();
	  else if(cell.getCellType()==Cell.CELL_TYPE_NUMERIC || cell.getCellType()==Cell.CELL_TYPE_FORMULA ){
		  int temp=(int)cell.getNumericCellValue();
		  String cellText  = String.valueOf(temp);
		  if (HSSFDateUtil.isCellDateFormatted(cell)) {
	           // format in form of M/D/YY
			  double d = cell.getNumericCellValue();

			  
			  
			  Calendar cal =Calendar.getInstance();
			  cal.setTime(HSSFDateUtil.getJavaDate(d));
			  Date dobj=cal.getTime();
			 SimpleDateFormat sdf=new SimpleDateFormat("DD-MMM-YYYY");
			 cellText=(sdf.format(dobj)).toString();
	           /*cellText =
	             (String.valueOf(cal.get(Calendar.YEAR))).substring(2);
	           cellText = cal.get(Calendar.MONTH)+1 + "/" +
	                      cal.get(Calendar.DAY_OF_MONTH) + "/" +
	                      cellText;*/
	         }		  
		  return cellText;
		  
		  
	  }else if(cell.getCellType()==Cell.CELL_TYPE_BLANK)
	      return "";
	  else 
		  return String.valueOf(cell.getBooleanCellValue());
		}
		catch(Exception e){
			
			e.printStackTrace();
			return "row "+rowNum+" or column "+colNum +" does not exist  in xls";
		}
	}
/**********************************************************************************************************************************/	

	// returns true if data is set successfully else false
	@SuppressWarnings("deprecation")
	public boolean setCellData(String sheetName,String colName,int rowNum, String data){
		
		if(data==null){
			System.out.println("Resolving null ptr");
			data=" ";
		}
		try{
		fis = new FileInputStream(path); 
		workbook = new HSSFWorkbook(fis);
//		my_style.setFillPattern(HSSFCellStyle.BORDER_DASHED );
		//New formats
		HSSFCellStyle my_style=workbook.createCellStyle();
		my_style.setAlignment(HorizontalAlignment.CENTER);
		my_style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		
		
		if(rowNum<=0)
			return false;
		int index = workbook.getSheetIndex(sheetName);
		int colNum=-1;
		if(index==-1)
			return false;
		sheet = workbook.getSheetAt(index);
		row=sheet.getRow(2);
		for(int i=0;i<row.getLastCellNum();i++){

			//System.out.println("Printtttttt:"+row.getCell(i).getStringCellValue().trim());
			if(row.getCell(i).getStringCellValue().trim().equals(colName))
				colNum=i;
		}
		if(colNum==-1)
			return false;

		sheet.autoSizeColumn(colNum); 
		row = sheet.getRow(rowNum-1);
		if (row == null)
			row = sheet.createRow(rowNum-1);

		cell = row.getCell(colNum);	
		if (cell == null)
	        cell = row.createCell(colNum);

	    // cell style
	    //CellStyle cs = workbook.createCellStyle();
	    //cs.setWrapText(true);
	    //cell.setCellStyle(cs);

		if(data.equalsIgnoreCase("Failed") || data.equalsIgnoreCase("fail")){
			my_style.setFillForegroundColor(IndexedColors.RED.getIndex());
	        
	        my_style.setFillBackgroundColor((new HSSFColor.BLACK().getIndex()));
	    	cell.setCellStyle(my_style);
	    }
	    else if(data.equalsIgnoreCase("passed") || data.equalsIgnoreCase("pass")){
	    	my_style.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
	    	
	    	short a=my_style.getFillForegroundColor();
	    	if(a!=HSSFColor.LIGHT_GREEN.index)
	    	{
	    		my_style.setFillForegroundColor((HSSFColor.GREEN.index));
	    	}
	    	my_style.setFillBackgroundColor((HSSFColor.WHITE.index));
	    	cell.setCellStyle(my_style);
	    }
	    else if(data.equalsIgnoreCase("skipped") || data.equalsIgnoreCase("skip")){
	    	my_style.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
	    	
		    my_style.setFillBackgroundColor((new HSSFColor.BLACK().getIndex()));
	    	cell.setCellStyle(my_style);
	    }

	    cell.setCellValue(data);

	    fileOut = new FileOutputStream(path);

		workbook.write(fileOut);

	    fileOut.close();

		}
		
		catch(Exception e){
			e.printStackTrace();
			return false;
		}
		
		return true;
	}
	
	/**********************************************************************************************************************************/
	
	public int getCurrentRowNumber()
	{
		int rowNum=sheet.getPhysicalNumberOfRows();
		return rowNum;
	}
	
	public int getCurrentRowNumber(String sheetName)
	{
		int rowNum=workbook.getSheet(sheetName).getPhysicalNumberOfRows();
		return rowNum;
	}
	
	// returns true if data is set successfully else false
		@SuppressWarnings("deprecation")
		public boolean setCellDataResponse(String sheetName,String colName,String data,int rowNum){
			if(data==null){
				System.out.println("Resolving null ptr");
				data=" ";
			}
			try{
			fis = new FileInputStream(path); 
			workbook = new HSSFWorkbook(fis);
			HSSFCellStyle my_style=workbook.createCellStyle();
//			my_style.setFillPattern(HSSFCellStyle.BORDER_DASHED );
			my_style.setAlignment(HorizontalAlignment.CENTER);
			my_style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
			
			
			/*if(rowNum<=0)
				return false;*/
			int index = workbook.getSheetIndex(sheetName);
			int colNum=-1;
			if(index==-1)
				return false;
			sheet = workbook.getSheetAt(index);
			row=sheet.getRow(0);
			for(int i=0;i<row.getPhysicalNumberOfCells();i++){

				//System.out.println("Printtttttt:"+row.getCell(i).getStringCellValue().trim());
				if(row.getCell(i).getStringCellValue().trim().equals(colName))
					colNum=i;
			}
			if(colNum==-1)
				return false;

			sheet.autoSizeColumn(colNum); 
			row = sheet.getRow(rowNum);
			if (row == null)
				row = sheet.createRow(rowNum);

			cell = row.getCell(colNum);	
			if (cell == null)
		        cell = row.createCell(colNum);

		    // cell style
		    //CellStyle cs = workbook.createCellStyle();
		    //cs.setWrapText(true);
		    //cell.setCellStyle(cs);

			if(data.equalsIgnoreCase("Failed") || data.equalsIgnoreCase("fail")){
		        my_style.setFillForegroundColor((new HSSFColor.RED().getIndex()));
		        my_style.setFillBackgroundColor((new HSSFColor.BLACK().getIndex()));
		    	cell.setCellStyle(my_style);
		    }
		    else if(data.equalsIgnoreCase("passed") || data.equalsIgnoreCase("pass")){
		    	my_style.setFillForegroundColor((HSSFColor.LIGHT_GREEN.index));
		    	short a=my_style.getFillForegroundColor();
		    	if(a!=HSSFColor.LIGHT_GREEN.index)
		    	{
		    		my_style.setFillForegroundColor((HSSFColor.GREEN.index));
		    	}
		    	my_style.setFillBackgroundColor((HSSFColor.WHITE.index));
		    	cell.setCellStyle(my_style);
		    }
		    else if(data.equalsIgnoreCase("skipped") || data.equalsIgnoreCase("skip")){
		    	my_style.setFillForegroundColor((new HSSFColor.LIGHT_YELLOW().getIndex()));
			    my_style.setFillBackgroundColor((new HSSFColor.BLACK().getIndex()));
		    	cell.setCellStyle(my_style);
		    }

		    cell.setCellValue(data);

		    fileOut = new FileOutputStream(path);

			workbook.write(fileOut);
		    fileOut.close();

			}
			
			catch(Exception e){
				e.printStackTrace();
				return false;
			}
			
			return true;
		}
		/**********************************************************************************************************************************/
		// returns true if data is set successfully else false
		@SuppressWarnings("deprecation")
		public boolean setCellDataNew(String sheetName,String colName,String data,int rowNum){
			if(data==null){
				System.out.println("Resolving null ptr");
				data=" ";
			}
			try{
			fis = new FileInputStream(path); 
			workbook = new HSSFWorkbook(fis);
			HSSFCellStyle my_style=workbook.createCellStyle();
//			my_style.setFillPattern(HSSFCellStyle.BORDER_DASHED );
			my_style.setAlignment(HorizontalAlignment.CENTER);
			my_style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
			
			
			
			/*if(rowNum<=0)
				return false;*/
			int index = workbook.getSheetIndex(sheetName);
			int colNum=-1;
			if(index==-1)
				return false;
			sheet = workbook.getSheetAt(index);
			row=sheet.getRow(2);
			for(int i=0;i<row.getPhysicalNumberOfCells();i++){

				//System.out.println("Printtttttt:"+row.getCell(i).getStringCellValue().trim());
				if(row.getCell(i).getStringCellValue().trim().equals(colName))
					{
					  colNum=i;
					  break;
					}
			}
			if(colNum==-1)
				return false;

			sheet.autoSizeColumn(colNum); 
			row = sheet.getRow(rowNum-1);
			if (row == null)
				row = sheet.createRow(rowNum-1);

			cell = row.getCell(colNum);	
			if (cell == null)
		        cell = row.createCell(colNum);

		    // cell style
		    //CellStyle cs = workbook.createCellStyle();
		    //cs.setWrapText(true);
		    //cell.setCellStyle(cs);

			if(data.equalsIgnoreCase("Failed") || data.equalsIgnoreCase("fail")){
		        my_style.setFillForegroundColor((new HSSFColor.RED().getIndex()));
		        my_style.setFillBackgroundColor((new HSSFColor.BLACK().getIndex()));
		    	cell.setCellStyle(my_style);
		    }
		    else if(data.equalsIgnoreCase("passed") || data.equalsIgnoreCase("pass")){
		    	my_style.setFillForegroundColor((HSSFColor.LIGHT_GREEN.index));
		    	short a=my_style.getFillForegroundColor();
		    	if(a!=HSSFColor.LIGHT_GREEN.index)
		    	{
		    		my_style.setFillForegroundColor((HSSFColor.GREEN.index));
		    	}
		    	my_style.setFillBackgroundColor((HSSFColor.WHITE.index));
		    	cell.setCellStyle(my_style);
		    }
		    else if(data.equalsIgnoreCase("skipped") || data.equalsIgnoreCase("skip")){
		    	my_style.setFillForegroundColor((new HSSFColor.LIGHT_YELLOW().getIndex()));
			    my_style.setFillBackgroundColor((new HSSFColor.BLACK().getIndex()));
		    	cell.setCellStyle(my_style);
		    }

		    cell.setCellValue(data);
		    fileOut = new FileOutputStream(path);

			workbook.write(fileOut);
		    fileOut.close();

			}
			
			catch(Exception e){
				e.printStackTrace();
				return false;
			}
			
			return true;
		}
/**********************************************************************************************************************************/
	// returns true if data is set successfully else false
		 @SuppressWarnings("deprecation")
		public boolean setCellData(String sheetName,String identifier,String colName,int rowNum, String data){

			if(data==null){
				System.out.println("Resolving null ptr");
				data=" ";
			}
			try{
			fis = new FileInputStream(path); 
			workbook = new HSSFWorkbook(fis);
//			my_style.setFillPattern(HSSFCellStyle.ALIGN_CENTER );
//			my_style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND );
			HSSFCellStyle my_style=workbook.createCellStyle();
			my_style.setAlignment(HorizontalAlignment.CENTER);
			my_style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
			

			if(rowNum<=0)
				return false;
			int index = workbook.getSheetIndex(sheetName);
			int colNum=-1;
			if(index==-1)
				return false;
			sheet = workbook.getSheetAt(index);

			if(getCellRowNum(sheetName, identifier)==-1){
				//Unable to find record with identifier inside sheet,return false
				//System.out.println("Unable to find a record inside "+sheetName+" for Identifier"+identifier);
				return false;
			}
			row=sheet.getRow(getCellRowNum(sheetName, identifier)+2); //+2 is done so as to get column name keys
			for(int i=0;i<row.getLastCellNum();i++){

				/*System.out.println("Printtttttt:"+row.getCell(i).getStringCellValue().trim());*/
				if(row.getCell(i).getStringCellValue().trim().equals(colName)){
					colNum=i;
					break;
				}
			}
			if(colNum==-1)
				return false;

			sheet.autoSizeColumn(colNum); 
			row = sheet.getRow(rowNum-1);
			if (row == null)
				row = sheet.createRow(rowNum-1);

			cell = row.getCell(colNum);	
			if (cell == null)
		        cell = row.createCell(colNum);

		    // cell style
		    //CellStyle cs = workbook.createCellStyle();
		    //cs.setWrapText(true);
		    //cell.setCellStyle(cs);


			if(data.equalsIgnoreCase("Failed") || data.equalsIgnoreCase("fail")){
		        my_style.setFillForegroundColor((new HSSFColor.LIGHT_ORANGE().getIndex()));
		        my_style.setFillBackgroundColor((new HSSFColor.WHITE().getIndex()));
		    	cell.setCellStyle(my_style);
		    }
		    else if(data.equalsIgnoreCase("passed") || data.equalsIgnoreCase("pass")){
		    	my_style.setFillForegroundColor((HSSFColor.LIGHT_GREEN.index));
		    	my_style.setFillBackgroundColor((new HSSFColor.WHITE().getIndex()));
		    	cell.setCellStyle(my_style);
		    }
		    else if(data.equalsIgnoreCase("skipped") || data.equalsIgnoreCase("skip")){
		    	my_style.setFillForegroundColor((new HSSFColor.LIGHT_YELLOW().getIndex()));
		    	my_style.setFillBackgroundColor((new HSSFColor.WHITE().getIndex()));
		    	cell.setCellStyle(my_style);
		    }

		    cell.setCellValue(data);

		    fileOut = new FileOutputStream(path);

			workbook.write(fileOut);

		    fileOut.close();


			}
			catch(Exception e){
				e.printStackTrace();
				return false;
			}
			return true;
		}
/*******************************************************************************************************************************************************************/
	// returns true if data is set successfully else false
	public boolean setCellData(String sheetName,String colName,int rowNum, String data,String url){
		//System.out.println("setCellData setCellData******************");
		try{
		fis = new FileInputStream(path); 
		workbook = new HSSFWorkbook(fis);

		if(rowNum<=0)
			return false;
		
		int index = workbook.getSheetIndex(sheetName);
		int colNum=-1;
		if(index==-1)
			return false;
		
		
		sheet = workbook.getSheetAt(index);
		//System.out.println("A");
		row=sheet.getRow(2);
		for(int i=0;i<row.getLastCellNum();i++){
			//System.out.println(row.getCell(i).getStringCellValue().trim());
			if(row.getCell(i).getStringCellValue().trim().equalsIgnoreCase(colName))
				colNum=i;
		}
		
		if(colNum==-1)
			return false;
		sheet.autoSizeColumn(colNum); //
		row = sheet.getRow(rowNum-1);
		if (row == null)
			row = sheet.createRow(rowNum-1);
		
		cell = row.getCell(colNum);	
		if (cell == null)
	        cell = row.createCell(colNum);
			
	    cell.setCellValue(data);
	    CreationHelper createHelper = workbook.getCreationHelper();

	    //cell style for hyper links
	    //by default hyper links are blue and underlined
	    CellStyle hlink_style = workbook.createCellStyle();
	    HSSFFont hlink_font = workbook.createFont();
	    hlink_font.setUnderline(HSSFFont.U_SINGLE);
	    hlink_font.setColor(IndexedColors.BLUE.getIndex());
	    hlink_style.setFont(hlink_font);
	    //hlink_style.setWrapText(true);

//	    Hyperlink link = createHelper.createHyperlink(HSSFHyperlink.LINK_FILE);
	    Hyperlink link = createHelper.createHyperlink(HyperlinkType.FILE);
	    link.setAddress(url);
	    cell.setHyperlink(link);
	    cell.setCellStyle(hlink_style);
	      
	    fileOut = new FileOutputStream(path);
		workbook.write(fileOut);

	    fileOut.close();	

		}
		catch(Exception e){
			e.printStackTrace();
			return false;
		}
		return true;
	}
	
	
/**********************************************************************************************************************************/	

	// returns true if sheet is created successfully else false
	public boolean addSheet(String  sheetname){		
		
		FileOutputStream fileOut;
		try {
			 workbook.createSheet(sheetname);	
			 fileOut = new FileOutputStream(path);
			 workbook.write(fileOut);
		     fileOut.close();		    
		} catch (Exception e) {			
			e.printStackTrace();
			return false;
		}
		return true;
	}
	
	// returns true if sheet is removed successfully else false if sheet does not exist
	public boolean removeSheet(String sheetName){		
		int index = workbook.getSheetIndex(sheetName);
		if(index==-1)
			return false;
		
		FileOutputStream fileOut;
		try {
			workbook.removeSheetAt(index);
			fileOut = new FileOutputStream(path);
			workbook.write(fileOut);
		    fileOut.close();		    
		} catch (Exception e) {			
			e.printStackTrace();
			return false;
		}
		return true;
	}
/**********************************************************************************************************************************/	

	public boolean addColumn(String sheetName,String colName){
		//System.out.println("**************addColumn*********************");
		
		try{				
			fis = new FileInputStream(path); 
			workbook = new HSSFWorkbook(fis);
			int index = workbook.getSheetIndex(sheetName);
			if(index==-1)
				return false;
			
			HSSFFont font= workbook.createFont();
		    font.setFontHeightInPoints((short)10);
		    font.setFontName("Arial");
		    font.setColor(IndexedColors.RED.getIndex());
		    font.setBold(true);
		    font.setItalic(false);
		HSSFCellStyle style = workbook.createCellStyle();
//		style.setFillForegroundColor(HSSFColor.GREY_40_PERCENT.index);
//		style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		style.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
		style.setAlignment(HorizontalAlignment.CENTER);
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		style.setFont(font);
		
		sheet=workbook.getSheetAt(index);
		
		row = sheet.getRow(0);
		if (row == null)
			row = sheet.createRow(0);
		
		//cell = row.getCell();	
		//if (cell == null)
		//System.out.println(row.getLastCellNum());
		if(row.getLastCellNum() == -1)
			cell = row.createCell(0);
		else
			cell = row.createCell(row.getLastCellNum());
	        
	        cell.setCellValue(colName);
	        cell.setCellStyle(style);
	        
	        fileOut = new FileOutputStream(path);
			workbook.write(fileOut);
		    fileOut.close();		    

		}catch(Exception e){
			e.printStackTrace();
			return false;
		}
		
		return true;
		
		
	}
/**********************************************************************************************************************************/	

	// removes a column and all the contents
	@SuppressWarnings("unused")
	public boolean removeColumn(String sheetName, int colNum) {
		try{
		if(!isSheetExist(sheetName))
			return false;
		fis = new FileInputStream(path); 
		workbook = new HSSFWorkbook(fis);
		sheet=workbook.getSheet(sheetName);
//		style.setFillForegroundColor(HSSFColor.GREY_40_PERCENT.index);
//		@SuppressWarnings("unused")
//		CreationHelper createHelper = workbook.getCreationHelper();
//		style.setFillPattern(HSSFCellStyle.NO_FILL);
		HSSFCellStyle my_style=workbook.createCellStyle();
		my_style.setFillForegroundColor(IndexedColors.GREY_40_PERCENT.getIndex());
		CreationHelper createHelper = workbook.getCreationHelper();
		my_style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		
	    
	
		for(int i =0;i<getRowCount(sheetName);i++){
			row=sheet.getRow(i);	
			if(row!=null){
				cell=row.getCell(colNum);
				if(cell!=null){
					cell.setCellStyle(my_style);
					row.removeCell(cell);
				}
			}
		}
		fileOut = new FileOutputStream(path);
		workbook.write(fileOut);
	    fileOut.close();
		}
		catch(Exception e){
			e.printStackTrace();
			return false;
		}
		return true;
		
	}
/**********************************************************************************************************************************/	

  // find whether sheets exists	
	public boolean isSheetExist(String sheetName){
		int index = workbook.getSheetIndex(sheetName);
		if(index==-1){
			index=workbook.getSheetIndex(sheetName.toUpperCase());
				if(index==-1)
					return false;
				else
					return true;
		}
		else
			return true;
	}
	
	// returns number of columns in a sheet	
	public int getColumnCount(String sheetName){
		// check if sheet exists
		if(!isSheetExist(sheetName))
		 return -1;
		
		sheet = workbook.getSheet(sheetName);
		row = sheet.getRow(0);
		
		if(row==null)
			return -1;
		
		return row.getLastCellNum();
						
	}
/**********************************************************************************************************************************/	

	//String sheetName, String testCaseName,String keyword ,String URL,String message
	public boolean addHyperLink(String sheetName,String screenShotColName,String testCaseName,int dataRowNo,String url,String message){
		
		url=url.replace('\\', '/');
		if(!isSheetExist(sheetName))
		{
			/*System.out.println(sheetName+" doesnt exists please check your sheet name");*/
			 return false;
		}
			
		/*System.out.println("SHEET EXISTS");*/
	    sheet = workbook.getSheet(sheetName);
	    
	    for(int i=1;i<=getRowCount(sheetName);i++)
	    	{
	    	 if(getCellData(sheetName, 0, i).equalsIgnoreCase(testCaseName))
	    	 {
	    		setCellData(sheetName, screenShotColName, dataRowNo, message,url);
	    		break;
	    	 }
	    }
	    /*System.out.println(message);*/
		return true; 
	}
	
/**********************************************************************************************************************************/	
	public int getCellRowNum(String sheetName,String colName,String cellValue){
		
		for(int i=2;i<=getRowCount(sheetName);i++){
	    	if(getCellData(sheetName,colName , i).equalsIgnoreCase(cellValue)){
	    		return i;
	    	}
	    }
		return -1;
		
	}
		

/**********************************************************************************************************************************/	

	public int getCellRowNum(String sheetName,String cellValue){
		int headingRow = -1;
		for(int rn=0;rn<getRowCount(sheetName);rn++){
		    Row temprow=sheet.getRow(rn);
		    if (temprow == null) {
		      // This row is all blank, skip
		      continue;
		    }
		    
		    for(int cn=0;cn<temprow.getLastCellNum();cn++){
		        Cell cell=temprow.getCell(cn);
		        if (cell == null) {
		           // No cell here
		        } 
		        else {
		           try{ 
		               if(cell.getStringCellValue().equals(cellValue)){
		                  headingRow=rn;
		                  break;
		               }
		           	}catch(Exception e){
		              continue;
		           	}
		        }
		    }
		    //break;
		}
	 return headingRow;
	}
}