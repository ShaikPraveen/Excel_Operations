package com.practice.excel_operations;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Count_Rows
{
	public static void main(String[] args) throws IOException {
		
		FileInputStream fis=new FileInputStream(System.getProperty("user.dir")+"\\src\\excel\\Sample_Data.xlsx");
		XSSFWorkbook wb=new XSSFWorkbook(fis);
		XSSFSheet ws=wb.getSheet("LoginData");
		
		        int rc= ws.getLastRowNum();
		        System.out.println("Number of rows are "+ rc);
		        
		        wb.close();
		        fis.close();
	}

}
