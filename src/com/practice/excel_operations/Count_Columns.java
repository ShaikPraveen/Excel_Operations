package com.practice.excel_operations;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Count_Columns 
{
	public static void main(String[] args) throws IOException 
	{
		FileInputStream fis=new FileInputStream(System.getProperty("user.dir")+"\\src\\excel\\Sample_Data.xlsx");
		XSSFWorkbook wb=new XSSFWorkbook(fis);
		XSSFSheet ws=wb.getSheet("LoginData");
		 int rc=ws.getLastRowNum();
		 
		 XSSFRow row;
		 int cc;
		 
		   for (int i = 0; i<=rc; i++)
		   {
			   row=ws.getRow(i);
			   cc=row.getLastCellNum();
			   System.out.println("Number of colums are "+cc);
			
		}
		   wb.close();
		   fis.close();
		
	}

}
