package com.offcn.huigu;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Demo005 {

	public static void main(String[] args) throws FileNotFoundException, IOException {
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFCellStyle style = workbook.createCellStyle();
		//���ö���
		style.setAlignment(HorizontalAlignment.CENTER);
		style.setVerticalAlignment(VerticalAlignment.CENTER);
		//���������ɫ
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		style.setFillForegroundColor(IndexedColors.RED.index);
		XSSFSheet sheet = workbook.createSheet("������1");
		
		//�����ϲ���Ԫ���ַ����
		CellRangeAddress rangaddress = new CellRangeAddress(5, 10, 4, 10);

		//���߹�����Ҫ�ϲ��ķ�Χ
		sheet.addMergedRegion(rangaddress);
		
		XSSFRow row = sheet.createRow(5);
		
		XSSFCell cell = row.createCell(4);
		cell.setCellStyle(style);
		cell.setCellValue("�ϲ���Ԫ��");
		
		workbook.write(new FileOutputStream("d:\\chart\\huigu\\hebing.xlsx"));
	}

}
