package ExcelPOI;

import java.io.*;

import org.apache.poi.xssf.usermodel.*;

public class WritingExcel {

	public static void main(String[] args) throws IOException{
		
		
		XSSFWorkbook workbook=new XSSFWorkbook();
		XSSFSheet sheet=workbook.createSheet("Emp Info");
		
		Object empdata[][]= {
				{"EMP_ID","NAME","JOB"},
				{101,"SHUBHAM BHARDWAJ","SOFTWARE ENGINEER"},
				{451,"Saurav","Software Engineer"},
				{101,"Santanu","Software Developer"}};
		/*
		int rows=empdata.length;
		int cols=empdata[0].length;
		
		System.out.println(rows);
		System.out.println(cols);
		
		for(int r=0;r<rows;++r)
		{
			XSSFRow row=sheet.createRow(r);
			
			for(int c=0;c<cols;++c)
			{
				XSSFCell cell=row.createCell(c);
				
				Object value=empdata[r][c];
				
				if(value instanceof String)
				{
					cell.setCellValue((String)value);
				}
				
				if(value instanceof Integer)
				{
					cell.setCellValue((Integer)value);
				}
				
				if(value instanceof Boolean)
				{
					cell.setCellValue((Boolean)value);
				}
			}
		}
		*/
		int rownum=0;
		for(Object emp[]:empdata)
		{
			XSSFRow row=sheet.createRow(rownum);
			++rownum;
			int colnum=0;
			for(Object value:emp)
			{
				XSSFCell cell=row.createCell(colnum++);
				if(value instanceof String)
					cell.setCellValue((String)value);
				if(value instanceof Integer)
					cell.setCellValue((Integer)value);
				if(value instanceof Boolean)
					cell.setCellValue((Boolean)value);
			}
		}
		
		
		String filepath=".\\datafiles\\employee.xlsx";
		FileOutputStream outstream=new FileOutputStream(filepath);
		workbook.write(outstream);
		
		outstream.close();
		
		System.out.println("Writing of Excel File is Completed");
		
	}
	
	
}
