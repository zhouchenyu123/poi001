package com.offcn.huigu;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Demo002 {

	public static void main(String[] args) throws FileNotFoundException, IOException {
		XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream("d:\\chart\\huigu\\demo1.xlsx"));

		XSSFSheet sheet = workbook.getSheet("¹¤×÷±í1");
		
		XSSFRow row = sheet.getRow(0);
		
		XSSFCell cell = row.getCell(3);
		System.out.println(cell.getStringCellValue());
	}

}
