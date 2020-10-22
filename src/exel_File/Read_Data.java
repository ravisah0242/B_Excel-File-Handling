package exel_File;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.WebDriver;
import org.testng.annotations.Test;

public class Read_Data 
{
	public WebDriver driver;
	
	@Test
	public void readdata() throws IOException
	{
		FileInputStream fis=new FileInputStream("./testdata/QSSE10.xlsx");
		XSSFWorkbook wb=new XSSFWorkbook(fis);
		
		/*Sheet s=wb.getSheet("Qspider");
		Row r=s.getRow(7);
		Cell c=r.getCell(5);
		System.out.println(c.getStringCellValue());*/
		/*Sheet s=wb.getSheet("Qspider");
		Row r=s.getRow(9);
		Cell c=r.getCell(6);
		System.out.println(c.getStringCellValue());*/
		//OR
		//System.out.println(wb.getSheet("Qspider").getRow(7).getCell(5).getStringCellValue());
		//System.out.println(wb.getSheet("Qspider").getRow(9).getCell(6).getStringCellValue());
		//System.out.println(wb.getSheet("Qspider").getRow(5).getCell(5).getStringCellValue());
		
		//System.out.println(wb.getSheet("Qspider").getRow(7).getCell(5).toString());//-->to String covert String;
		System.out.println(wb.getSheet("Qspider").getRow(5).getCell(5).toString());
		
	}
}
