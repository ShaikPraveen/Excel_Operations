package com.practice.excel_operations;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Writing_Data 
{
	public static void main(String[] args) throws IOException 
	{
		FileInputStream fis=new FileInputStream(System.getProperty("user.dir")+"\\src\\excel\\Sample_Data.xlsx");
		XSSFWorkbook wb=new XSSFWorkbook(fis);
		XSSFSheet ws=wb.getSheet("LoginData");
		
		XSSFRow row=ws.getRow(2);
		XSSFCell cell1=row.createCell(3);
		         cell1.setCellValue("Pass");
		         
		         
	     XSSFRow row1=ws.getRow(3);
		 XSSFCell cell2=row1.createCell(3);
		 cell2.setCellValue("Fail");
		 
		 
		 FileOutputStream fo=new FileOutputStream(System.getProperty("user.dir")+"\\src\\excel\\Sample_Data.xlsx");
		 wb.write(fo);
		 wb.close();
		 fis.close();
		 fo.close();	
	}

}
