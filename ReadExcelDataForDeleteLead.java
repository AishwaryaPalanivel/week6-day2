package week6.day2;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;

public class ReadExcelDataForDeleteLead {
@Test
	
	public static String[][] readDataFromExcelForDeleteLead() throws IOException
	{
		
		
		XSSFWorkbook wb = new XSSFWorkbook("./data/tc002 - Deletelead.xlsx");
		 XSSFSheet sheet = wb.getSheetAt(0);
		 int rowcount = sheet.getLastRowNum();
		 short columncount = sheet.getRow(0).getLastCellNum();
		System.out.println(" The row and column counts are : " +rowcount +" and " + columncount);
		
		String[][] data123 = new String[rowcount][columncount];
		for (int i = 1; i <= rowcount; i++) 
		{
			XSSFRow eachrow = sheet.getRow(i);
			for (int j = 0; j < columncount; j++)
			{
				XSSFCell eachcell = eachrow.getCell(j);
				String stringCellValue = eachcell.getStringCellValue();
				System.out.println("The excel contents are " + stringCellValue);
				data123[i-1][j] =  stringCellValue;
			}
		}
		
		return data123;
	}
}
		
		
		
		
		
		
		
}
