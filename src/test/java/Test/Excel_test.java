package Test;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;



public class Excel_test {
	
	 XSSFWorkbook wbook;
	 XSSFSheet wsheet;
	
	 @Test
	public void salesorder() throws IOException {
			
    FileInputStream file = new FileInputStream("SampleData.xlsx");
    
     wbook = new XSSFWorkbook(file);
    wsheet = wbook.getSheet("SalesOrders"); //providing sheet name
  //  wsheet = wbook.getSheetAt(1); //provinding sheet index
   int rowcount= wsheet.getLastRowNum();
   
   int colnum=wsheet.getRow(0).getLastCellNum();
   
   	for(int i =0;i<=rowcount;i++) {
   		
   		XSSFRow currentrow=wsheet.getRow(i); 
   		
   		for(int j=0;j<=colnum;j++)
   		{
   			String value=currentrow.getCell(j).toString();
   			System.out.print(" "+value);
   		}
   		System.out.println( );
   	}
    
   
	}


}
