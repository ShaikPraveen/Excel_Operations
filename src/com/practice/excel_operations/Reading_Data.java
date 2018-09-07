package com.practice.excel_operations;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Reading_Data
{
	public static void main(String[] args) throws IOException {
		
		FileInputStream fis=new FileInputStream(System.getProperty("user.dir")+"\\src\\excel\\Sample_Data.xlsx");
		XSSFWorkbook wb=new XSSFWorkbook(fis);
		XSSFSheet ws=wb.getSheet("LoginData");
		
		int rc=ws.getLastRowNum();
		XSSFRow row;
		XSSFCell cell1, cell2, cell3;
		String fn,ln,dn;
		
		 for (int i = 0; i<=rc; i++) 
		 {
			 row=ws.getRow(i);
			 cell1=row.getCell(0);
			 cell2=row.getCell(1);
			 cell3=row.getCell(2);
			 
			fn= cell1.getStringCellValue();
			ln= cell2.getStringCellValue();
			dn=cell3.getStringCellValue();
			 
			 System.out.println(fn+"  "+ln+"  "+dn);
		}
		
	}

}
