package com.offcn.huigu;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class Demo001 {
public static void main(String[] args) throws FileNotFoundException, IOException {
	//1\创建工作簿
	HSSFWorkbook workbook = new HSSFWorkbook();
	//2\创建工作表
	HSSFSheet sheet = workbook.createSheet("工作表1");
	
	//3\创建行
	HSSFRow row = sheet.createRow(0);
	
	//4\创建单元格
	HSSFCell cell = row.createCell(3);
	
	//写入内容
	cell.setCellValue("好好学习");
	
	//保存工作簿对象到磁盘
	workbook.write(new FileOutputStream("d:\\chart\\huigu\\demo1.xls"));
    System.out.println("ok");
} 
}
