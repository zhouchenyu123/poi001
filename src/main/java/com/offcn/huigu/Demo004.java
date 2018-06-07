package com.offcn.huigu;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Demo004 {

	public static void main(String[] args) throws FileNotFoundException, IOException {
		XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream("d:\\chart\\gz.xlsx"));

		//获取工作簿里面包含的工作表数量
		int numsheet = workbook.getNumberOfSheets();
		
		for(int i=0;i<numsheet;i++){
			XSSFSheet sheet = workbook.getSheetAt(i);
			//判断工作表里面包含的有效行数
			int numrow = sheet.getPhysicalNumberOfRows();
			
			for(int j=0;j<numrow;j++){
				XSSFRow row = sheet.getRow(j);
				//判断该行包含的有效单元格数量
				int numcell = row.getPhysicalNumberOfCells();
				for(int q=0;q<numcell;q++){
					XSSFCell cell = row.getCell(q);
					//需要判断单元格的类型
					if(cell.getCellTypeEnum()==CellType.STRING){
						System.out.print(cell.getStringCellValue()+"\t");
					}else if(cell.getCellTypeEnum()==CellType.NUMERIC){
						System.out.print(cell.getNumericCellValue()+"\t");
					}else if(cell.getCellTypeEnum()==CellType.BOOLEAN){
						System.out.print(cell.getBooleanCellValue()+"\t");
					}else if(cell.getCellTypeEnum()==CellType.BLANK){
						System.out.print("null"+"\t");
					}else{
						System.out.print(cell.getDateCellValue()+"\t");
					}
				}
				System.out.println("");
			}
		}
		
		workbook.close();
	}

}
