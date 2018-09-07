package com.basic.excel.operations;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Rows_Count
{
	//1.Count Number of Rows
	//Whenever you want to return number of rows present in excel sheet, 
	//it will return start form zero including FirstName and LastName and DeptName. 
	//Find the Following Table data, Here the Number of rows are 3 and Number of Columns are 3
	//   FirstName     LastName           DeptName
	//    Praveen       Shaik              Testing
	//    Suleman       Shaik              Marketing
	//    Sai          Tokachichu          Developing

	public static void main(String[] args) throws IOException 
	{
		FileInputStream fis=new FileInputStream(System.getProperty("user.dir")+"\\src\\excel\\Sample_Data.xlsx");
		XSSFWorkbook wb=new XSSFWorkbook(fis);
		XSSFSheet    ws=wb.getSheet("StudentData");
		
		      int rows=ws.getLastRowNum();
		      System.out.println("Number of Rows Are "+ rows);
		      
		      fis.close();
		      wb.close();
	}
}
