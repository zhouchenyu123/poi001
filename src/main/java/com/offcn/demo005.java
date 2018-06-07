package com.offcn;

import java.io.File;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class demo005 {

	public static void main(String[] args) throws EncryptedDocumentException, InvalidFormatException, IOException {
		Workbook workbook=WorkbookFactory.create(new File("d:\\chart\\gz.xlsx"));

		//?工作表数量名字未知!!!
		//获取工作表的数量
		int sheetnum = workbook.getNumberOfSheets();
		
		System.out.println("工作表数量:"+sheetnum);
		//循环读取全部工作表
		for(int i=0;i<sheetnum;i++){
			Sheet sheet = workbook.getSheetAt(i);
			System.out.println("工作表名称:"+sheet.getSheetName());
			//这个工作表有多少行?
			int rownum = sheet.getPhysicalNumberOfRows();
			for(int j=0;j<rownum;j++){
				Row row = sheet.getRow(j);
				//这行包含了多少单元格
				int cellnum = row.getPhysicalNumberOfCells();
				for(int q=0;q<cellnum;q++){
					Cell cell = row.getCell(q);
					//判断单元格的类型
					if(cell.getCellTypeEnum()==CellType.STRING){
						System.out.print(cell.getStringCellValue()+"\t");
					}else if(cell.getCellTypeEnum()==CellType.NUMERIC){
						System.out.print(cell.getNumericCellValue()+"\t");
					}else if(cell.getCellTypeEnum()==CellType.BOOLEAN){
						System.out.print(cell.getBooleanCellValue()+"\t");
					}else if(cell.getCellTypeEnum()==CellType.BLANK){
						System.out.print(""+"\t");
					}else{
						System.out.print(cell.getDateCellValue()+"\t");
					}
				}
				System.out.println("");
			}
		}
	}

}
