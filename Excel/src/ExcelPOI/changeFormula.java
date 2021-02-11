package ExcelPOI;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class changeFormula {

	public static void main(String[] args) throws IOException {
		
		String location=".\\datafiles\\mathformula.xlsx";
		FileInputStream file=new FileInputStream(location);
		XSSFWorkbook workbook=new XSSFWorkbook(file);
		XSSFSheet sheet=workbook.getSheetAt(0);
		sheet.getRow(0).getCell(3).setCellFormula("A1+B1+C1");
		
		FileOutputStream path=new FileOutputStream(location);
		workbook.write(path);
		workbook.close();
		path.close();
		file.close();
		
		System.out.println("Cell has been Updated Successfully");
	}
	
}
