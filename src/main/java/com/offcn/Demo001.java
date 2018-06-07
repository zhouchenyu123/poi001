package com.offcn;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class Demo001 {

	public static void main(String[] args) throws FileNotFoundException, IOException {
		//1\�½�������
		HSSFWorkbook workbook = new HSSFWorkbook();
		//2\����������
		HSSFSheet sheet = workbook.createSheet("������1");
		//3\������
		HSSFRow row = sheet.createRow(0);
		//4\������Ԫ�����
		HSSFCell cell = row.createCell(2);
		//5\��Ԫ��д������
		cell.setCellValue("���JAVA");

		//6\���湤����
		workbook.write(new FileOutputStream("d:\\chart\\hello.xls"));
		
		System.out.println("ok");
	}

}
