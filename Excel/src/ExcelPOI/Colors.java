package ExcelPOI;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Colors {

	public static void main(String[] args) throws IOException {
		
		//String path=".\\datafiles\\colors.xlsx";
		//FileOutputStream file=new FileOutputStream(path);
		XSSFWorkbook workbook=new XSSFWorkbook();
		XSSFSheet sheet=workbook.createSheet("sheet1");
		XSSFRow row=sheet.createRow(0);
		
		XSSFCellStyle style=workbook.createCellStyle();
		style.setFillBackgroundColor(IndexedColors.BLUE.getIndex());
		style.setFillPattern(FillPatternType.DIAMONDS);
		
		XSSFCell cell=row.createCell(0);
		cell.setCellValue("Hello");
		cell.setCellStyle(style);
		
		style=workbook.createCellStyle();
		style.setFillForegroundColor(IndexedColors.BRIGHT_GREEN.getIndex());
		style.setFillPattern(FillPatternType.BIG_SPOTS);
		
		cell=row.createCell(1);
		cell.setCellValue("World");
		cell.setCellStyle(style);
		
		String path=".\\datafiles\\colors.xlsx";
		FileOutputStream file=new FileOutputStream(path);
		workbook.write(file);
		workbook.close();
		file.close();
		
		System.out.println("Sheet with color Created Successfully");
	}
	
}
