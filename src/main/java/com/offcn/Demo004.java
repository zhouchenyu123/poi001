package com.offcn;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Demo004 {
//********************************************************************************************
	public static void main(String[] args) throws EncryptedDocumentException, InvalidFormatException, FileNotFoundException, IOException {
		Workbook workbook=WorkbookFactory.create(new FileInputStream("d:\\chart\\hello.xls"));

		Sheet sheet = workbook.getSheet("sheet1");
		
		Row row = sheet.getRow(0);
		
		Cell cell = row.getCell(2);
		
		System.out.println(cell.getStringCellValue());
		
		workbook.close();
	}

}
