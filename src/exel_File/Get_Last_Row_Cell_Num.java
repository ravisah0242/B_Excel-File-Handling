package exel_File;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.WebDriver;
import org.testng.annotations.Test;

public class Get_Last_Row_Cell_Num 
{	
	public WebDriver driver;
	@Test
	public void getLastRowCulmnNum() throws IOException
	{
		FileInputStream fis=new FileInputStream("./testdata/QSSE10.xlsx");
		XSSFWorkbook wb=new XSSFWorkbook(fis);
		Sheet s=wb.getSheet("Qspider");
		Row r=s.getRow(7);
		System.out.println(s.getLastRowNum());
		Cell c=r.getCell(5);
		System.out.println(r.getLastCellNum());
	}
}
