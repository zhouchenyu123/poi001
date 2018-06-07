package com.offcn.huigu;

import java.io.File;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Demo003 {

	public static void main(String[] args) throws EncryptedDocumentException, InvalidFormatException, IOException {
		Workbook workbook=WorkbookFactory.create(new File("d:\\chart\\huigu\\demo1.xls"));

		Sheet sheet = workbook.getSheet("¹¤×÷±í1");
		
		Row row = sheet.getRow(0);
		
		Cell cell = row.getCell(3);
		
		System.out.println(cell.getStringCellValue());
		
		workbook.close();
	}

}
