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
		//1\新建工作簿
		HSSFWorkbook workbook = new HSSFWorkbook();
		//2\创建工作表
		HSSFSheet sheet = workbook.createSheet("工作表1");
		//3\创建行
		HSSFRow row = sheet.createRow(0);
		//4\创建单元格对象
		HSSFCell cell = row.createCell(2);
		//5\向单元格写入内容
		cell.setCellValue("你好JAVA");

		//6\保存工作簿
		workbook.write(new FileOutputStream("d:\\chart\\hello.xls"));
		
		System.out.println("ok");
	}

}
