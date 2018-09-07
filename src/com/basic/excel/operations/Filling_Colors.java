package com.basic.excel.operations;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Filling_Colors 
{
	public static void main(String[] args) throws IOException 
	{
		FileInputStream fis=new FileInputStream(System.getProperty("user.dir")+"\\src\\excel\\Sample_Data.xlsx");
		XSSFWorkbook wb=new XSSFWorkbook(fis);
		XSSFSheet ws=wb.getSheet("LoginData");
		
		XSSFRow row1, row2;
		XSSFCell cell1, cell2;
		
		row1=ws.getRow(2);
		row2=ws.getRow(3);
		cell1=row1.getCell(4);
		cell2=row2.getCell(4);
		
	     CellStyle style=wb.createCellStyle();
	     style.setFillForegroundColor(IndexedColors.GREEN.getIndex());
	     //style.setFillPattern(CellStyle.SOLID_FOREGROUND);
	     cell1.setCellStyle(style);
	     
	     
	     CellStyle style1=wb.createCellStyle();
	     style1.setFillForegroundColor(IndexedColors.RED.getIndex());
	     //style1.setFillPattern(CellStyle.SOLID_FOREGROUND);
	     cell2.setCellStyle(style1);
	
		
		
		FileOutputStream fo=new FileOutputStream(System.getProperty("user.dir")+"\\src\\excel\\Sample_Data.xlsx");
		wb.write(fo);
		wb.close();
		fis.close();
		fo.close();
	}

}
