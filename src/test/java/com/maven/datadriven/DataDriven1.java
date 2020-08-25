package com.maven.datadriven;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DataDriven1 {

	public static void main(String[] args) throws IOException {
		
		File f=new File("C:\\Users\\AJAY\\Desktop\\Read.xlsx");
		FileInputStream fin=new FileInputStream(f);
		Workbook wb=new XSSFWorkbook(fin);
		Sheet s = wb.getSheet("Sheet1");
		Row r = s.getRow(1);
		int rowcount =s.getPhysicalNumberOfRows();
		System.out.println("no. of rows" + rowcount);
		Cell c = r.getCell(1);
		int cellcount = r.getPhysicalNumberOfCells();
		System.out.println("no. of cells" + cellcount);
		CellType ct = c.getCellType();
		String data=null;
		if(ct.equals(CellType.STRING)) {
		data  =  c.getStringCellValue();
		}
		else if(ct.equals(CellType.NUMERIC))
		{
			double d = c.getNumericCellValue();
			long l=(long) d;
			data = String.valueOf(l);
	}
		System.out.println(data);
}
}