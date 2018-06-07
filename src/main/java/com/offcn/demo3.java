package com.offcn;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

public class demo3 {

	@Test
	public void testWriter() throws FileNotFoundException, IOException{
		XSSFWorkbook workbook = new XSSFWorkbook();
		
		XSSFSheet sheet = workbook.createSheet("工作表1");
		
		XSSFRow row = sheet.createRow(0);
		
		XSSFCell cell = row.createCell(2);
		
		cell.setCellValue("好好学习");
		workbook.write(new FileOutputStream("d:\\chart\\new.xlsx"));
		
	workbook.close();
	System.out.println("ok");
		
	}
	
	@Test
	public void testRead() throws FileNotFoundException, IOException{
		XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream("d:\\chart\\new.xlsx"));
	
		XSSFSheet sheet = workbook.getSheet("工作表1");
		
		XSSFRow row = sheet.getRow(0);
		
		XSSFCell cell = row.getCell(2);
		
		System.out.println(cell.getStringCellValue());
		
		workbook.close();
	}
}
