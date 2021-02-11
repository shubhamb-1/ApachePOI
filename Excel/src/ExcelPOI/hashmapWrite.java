package ExcelPOI;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class hashmapWrite{
	
	public static void main(String[] args) throws IOException{
		
		/*
		XSSFWorkbook workbook=new XSSFWorkbook();
		XSSFSheet sheet=workbook.createSheet("Sheet1");
		
		Map<String,String> data=new HashMap<String,String>();
		
		data.put("101","John");
		data.put("201","Joe");
		data.put("301","Danny");
		data.put("401","Doe");
		data.put("501","Alex");
		data.put("601","Alfred");
		
		int rowcount=0;
		for(Map.Entry entry:data.entrySet())
		{
			XSSFRow row=sheet.createRow(rowcount++);
			row.createCell(0).setCellValue((String)entry.getKey());
			row.createCell(1).setCellValue((String)entry.getValue());
		}
		
		String path=".\\datafiles\\hashmapWrite.xlsx";
		FileOutputStream file=new FileOutputStream(path);
		workbook.write(file);
		file.close();
		workbook.close();
		
		System.out.println("HashMap to Excel Created Successfully!");
		
		*/
		String path=".\\datafiles\\hashmapWrite.xlsx";
		FileInputStream file=new FileInputStream(path);
		
		XSSFWorkbook workbook=new XSSFWorkbook(file);
		XSSFSheet sheet=workbook.getSheetAt(0);
		
		int rows=sheet.getLastRowNum();
		int cols=sheet.getRow(0).getLastCellNum();
		
		HashMap<String,String> data=new HashMap<String,String>();
		
		for(int r=0;r<=rows;++r)
		{
			String key="aaa";
			String value="aaa";
			for(int c=0;c<cols;++c)
			{
				XSSFRow row=sheet.getRow(r);
				if(c==0)
				{
					key=row.getCell(c).getStringCellValue();
				}
				if(c==1)
				{
					value=row.getCell(c).getStringCellValue();
				}
			}
			data.put(key,value);
		}
		
		for(Map.Entry entry :data.entrySet())
		{
			System.out.print((String)entry.getKey()+" ");
			System.out.print((String)entry.getValue());
			System.out.println();
		}
	}
	
}