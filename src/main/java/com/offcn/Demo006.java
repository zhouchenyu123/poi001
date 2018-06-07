package com.offcn;

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

public class Demo006 {

	public static void main(String[] args) throws FileNotFoundException, IOException {
		XSSFWorkbook workbook = new XSSFWorkbook();
		
		XSSFSheet sheet = workbook.createSheet("������1");
		
		//�����ϲ���Ԫ�����
		CellRangeAddress rangaddress = new CellRangeAddress(4, 11, 3, 7);
		//������ʽ����
		XSSFCellStyle cellstyle = workbook.createCellStyle();
		//ˮƽ����
		cellstyle.setAlignment(HorizontalAlignment.CENTER);
		//��ֱ����
		cellstyle.setVerticalAlignment(VerticalAlignment.CENTER);
		
		//�������ɫ
		cellstyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		cellstyle.setFillForegroundColor(IndexedColors.BLUE.getIndex());
		//�Ѻϲ���ַ����󶨹�����
		sheet.addMergedRegion(rangaddress);
		
		XSSFRow row = sheet.createRow(4);
		
		XSSFCell cell = row.createCell(3);
		cell.setCellStyle(cellstyle);
		
		cell.setCellValue("���Ǻϲ���Ԫ��");
		
		workbook.write(new FileOutputStream("d:\\chart\\hebing.xlsx"));
		
		workbook.close();
	}

}
