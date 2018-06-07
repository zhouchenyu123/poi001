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

		//��ȡ��������������Ĺ���������
		int numsheet = workbook.getNumberOfSheets();
		
		for(int i=0;i<numsheet;i++){
			XSSFSheet sheet = workbook.getSheetAt(i);
			//�жϹ����������������Ч����
			int numrow = sheet.getPhysicalNumberOfRows();
			
			for(int j=0;j<numrow;j++){
				XSSFRow row = sheet.getRow(j);
				//�жϸ��а�������Ч��Ԫ������
				int numcell = row.getPhysicalNumberOfCells();
				for(int q=0;q<numcell;q++){
					XSSFCell cell = row.getCell(q);
					//��Ҫ�жϵ�Ԫ�������
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
