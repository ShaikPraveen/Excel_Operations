package com.basic.excel.operations;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Reading_Student_Data
{
	public static void main(String[] args) throws IOException
	{
		FileInputStream fis=new FileInputStream(System.getProperty("user.dir")+"\\src\\excel\\Sample_Data.xlsx");
		XSSFWorkbook wb=new XSSFWorkbook(fis);
		XSSFSheet ws=wb.getSheet("StudentData");
		
		int rc=ws.getLastRowNum();
		XSSFRow row;
		XSSFCell cell1, cell2,cell3,cell4,cell5,cell6;
		String Sf,Sl,Rn,Marks,Subject,Fn;
		
		for (int i = 0; i <=rc; i++)
		{
			row=ws.getRow(i);
			cell1=row.getCell(0);
			cell2=row.getCell(1);
			cell3=row.getCell(2);
			cell4=row.getCell(3);
			cell5=row.getCell(4);
			//cell6=row.getCell(5);
			
			Sf  =cell1.getStringCellValue();
			Sl  =cell2.getStringCellValue();
			Rn  =cell3.getStringCellValue();
			Marks=cell4.getStringCellValue();
			Subject=cell5.getStringCellValue();
			//Fn=cell6.getStringCellValue();
				
		System.out.println(Sf+" " +Sl+"  "+ Rn+"  "+Marks+" " +Subject);
			
		}
		
		fis.close();
		wb.close();
	}

}
