package exel_File;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.WebDriver;
import org.testng.annotations.Test;

public class ReadDataLooping {
	
	public WebDriver driver;
	@Test
	
	public void readData() throws IOException
	{
		FileInputStream fis=new FileInputStream("./testdata/QSSE10.xlsx");
		XSSFWorkbook workbook=new XSSFWorkbook(fis);
		XSSFSheet sheet=workbook.getSheet("CT college"); // Providing Sheet Name;
		//Sheet sheet=workbook.getSheetAt(0); 			// Providing Sheet index;
		
		int rowcount=sheet.getLastRowNum();				// Return Row count 
		int cellcount=sheet.getRow(2).getLastCellNum();	// Return collen/cell count
		
		for(int i=0; i<rowcount; i++)
		{
			Row currentcount=sheet.getRow(i);			// Focus on current Row 
			
			for(int j=0; j<cellcount; j++)
			{
				String value=currentcount.getCell(j).toString();
				System.out.print(value);
			}
			System.out.println();
		}
		
		
		
	}
}
