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
	//1\����������
	HSSFWorkbook workbook = new HSSFWorkbook();
	//2\����������
	HSSFSheet sheet = workbook.createSheet("������1");
	
	//3\������
	HSSFRow row = sheet.createRow(0);
	
	//4\������Ԫ��
	HSSFCell cell = row.createCell(3);
	
	//д������
	cell.setCellValue("�ú�ѧϰ");
	
	//���湤�������󵽴���
	workbook.write(new FileOutputStream("d:\\chart\\huigu\\demo1.xls"));
    System.out.println("ok");
} 
}
