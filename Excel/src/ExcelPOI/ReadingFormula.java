package ExcelPOI;

import java.io.FileInputStream;
import java.util.Iterator;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadingFormula {

	public static void main(String[] args) throws IOException {
		
		String path=".\\datafiles\\cost.xlsx";
		FileInputStream file=new FileInputStream(path);
		
		XSSFWorkbook workbook=new XSSFWorkbook(file);
		XSSFSheet sheet=workbook.getSheetAt(0);
		
		/*
		int rows=sheet.getLastRowNum();
		int cols=sheet.getRow(1).getLastCellNum();
		
		System.out.println(rows);
		System.out.println(cols);
		
		for(int r=0;r<=rows;++r)
		{
			XSSFRow row=sheet.getRow(r);
			for(int c=0;c<cols;++c)
			{
				XSSFCell cell=row.getCell(c);
				switch(cell.getCellType())
				{
				case STRING:System.out.print(cell.getStringCellValue());break;
				case NUMERIC:System.out.print(cell.getNumericCellValue());break;
				case BOOLEAN:System.out.print(cell.getBooleanCellValue());break;
				case FORMULA:System.out.print(cell.getNumericCellValue());break;
				}
				System.out.print(" | ");
			}
			System.out.println();
		}
		*/
		
		Iterator it= sheet.iterator();
		while(it.hasNext())
		{
			XSSFRow row=(XSSFRow)it.next();
			Iterator celliterator=row.iterator();
			while(celliterator.hasNext())
			{
				XSSFCell cell=(XSSFCell)celliterator.next();
				switch(cell.getCellType())
				{
				case STRING:System.out.print(cell.getStringCellValue());break;
				case NUMERIC:System.out.print(cell.getNumericCellValue());break;
				case BOOLEAN:System.out.print(cell.getBooleanCellValue());break;
				case FORMULA:System.out.print(cell.getNumericCellValue());break;
				}
				System.out.print(" | ");
			}
			System.out.println();
		}
		
		workbook.close();
		file.close();
		System.out.println("Reading of Excel Sheet with Formulae Successfully Done!");
	}
}
