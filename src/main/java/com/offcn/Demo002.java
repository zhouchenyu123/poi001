package com.offcn;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class Demo002 {

	public static void main(String[] args) throws FileNotFoundException, IOException {
		HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream("d:\\chart\\hello.xls"));

		HSSFSheet sheet1 = workbook.getSheet("¹¤×÷±í1");
		
		HSSFRow row = sheet1.getRow(0);
		
		HSSFCell cell = row.getCell(2);
		
		System.out.println(cell.getStringCellValue());
		workbook.close();
		
	}

}
