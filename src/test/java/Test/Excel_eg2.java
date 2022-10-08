package Test;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;

public class Excel_eg2 {
	
	@Test
	public void sales() throws IOException {
		
	
	
	FileInputStream fis = new FileInputStream("SampleData.xlsx");
	
	XSSFWorkbook book = new XSSFWorkbook(fis);
	
	XSSFSheet sheet = book.getSheet("SalesOrders");
	
	int row = sheet.getLastRowNum();
	
	int column = sheet.getRow(0).getLastCellNum();
	
	String value = sheet.getRow(0).getCell(1).getStringCellValue();
	System.out.println(value);
	
	for(int i=0;i<=row;i++)
	{
		XSSFRow currentrow = sheet.getRow(i);
		
		for(int j=0;j<=column;j++) {
			
			String cell = currentrow.getCell(i).toString();
			
			System.out.print(cell);
		}
		
		System.out.println();
	}
	
	}

}
