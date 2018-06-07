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
		//设置对齐
		style.setAlignment(HorizontalAlignment.CENTER);
		style.setVerticalAlignment(VerticalAlignment.CENTER);
		//设置填充颜色
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		style.setFillForegroundColor(IndexedColors.RED.index);
		XSSFSheet sheet = workbook.createSheet("工作表1");
		
		//创建合并单元格地址对象
		CellRangeAddress rangaddress = new CellRangeAddress(5, 10, 4, 10);

		//告诉工作表要合并的范围
		sheet.addMergedRegion(rangaddress);
		
		XSSFRow row = sheet.createRow(5);
		
		XSSFCell cell = row.createCell(4);
		cell.setCellStyle(style);
		cell.setCellValue("合并单元格");
		
		workbook.write(new FileOutputStream("d:\\chart\\huigu\\hebing.xlsx"));
	}

}
