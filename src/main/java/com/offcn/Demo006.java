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
		
		XSSFSheet sheet = workbook.createSheet("工作表1");
		
		//创建合并单元格对象
		CellRangeAddress rangaddress = new CellRangeAddress(4, 11, 3, 7);
		//创建样式对象
		XSSFCellStyle cellstyle = workbook.createCellStyle();
		//水平居中
		cellstyle.setAlignment(HorizontalAlignment.CENTER);
		//垂直居中
		cellstyle.setVerticalAlignment(VerticalAlignment.CENTER);
		
		//设置填充色
		cellstyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		cellstyle.setFillForegroundColor(IndexedColors.BLUE.getIndex());
		//把合并地址对象绑定工作表
		sheet.addMergedRegion(rangaddress);
		
		XSSFRow row = sheet.createRow(4);
		
		XSSFCell cell = row.createCell(3);
		cell.setCellStyle(cellstyle);
		
		cell.setCellValue("我是合并单元格");
		
		workbook.write(new FileOutputStream("d:\\chart\\hebing.xlsx"));
		
		workbook.close();
	}

}
