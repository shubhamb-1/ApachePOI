import java.io.FileInputStream;
import java.io.IOException;


import java.util.Iterator;
import org.apache.poi.xssf.streaming.SXSSFRow.CellIterator;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadingExcel{
	
	public static void main(String[] args) throws IOException {
	
			String filepath=".\\datafiles\\countries.xlsx";
			FileInputStream fileinput=new FileInputStream(filepath);
			XSSFWorkbook workbook=new XSSFWorkbook(fileinput);
			XSSFSheet sheet=workbook.getSheetAt(0);
			
			int rows=sheet.getLastRowNum();
			int cols=sheet.getRow(0).getLastCellNum();
			
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
					case NUMERIC: System.out.print(cell.getNumericCellValue()+" | ");break;
					case STRING: System.out.print(cell.getStringCellValue()+" | ");break;
					case BOOLEAN: System.out.print(cell.getBooleanCellValue()+" | ");break;
					}
				}
				System.out.println();
			}
			
			/*Iterator iterator=sheet.iterator();
			
			while(iterator.hasNext())
			{
				XSSFRow row=(XSSFRow) iterator.next();
				
				Iterator cellIterator=row.cellIterator();
				
				while(cellIterator.hasNext())
				{
					XSSFCell cell=(XSSFCell)cellIterator.next();
					
					switch(cell.getCellType())
					{
					case NUMERIC: System.out.print(cell.getNumericCellValue()+" | ");break;
					case STRING: System.out.print(cell.getStringCellValue()+" | ");break;
					case BOOLEAN: System.out.print(cell.getBooleanCellValue()+" | ");break;
					}
				}
					System.out.println();
			}
			*/
	}
	
}