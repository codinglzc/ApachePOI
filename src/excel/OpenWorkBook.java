package excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class OpenWorkBook
{

	public static void main(String[] args) throws IOException
	{
		File file = new File("openworkbook.xlsx");
		FileInputStream fIP = new FileInputStream(file);
		// Get the workbook instance for XLSX file
		XSSFWorkbook workbook = new XSSFWorkbook(fIP);
		if (file.isFile() && file.exists())
		{
			System.out.println("openworkbook.xlsx file open successfully.");
		} else
		{
			System.out.println("Error to open openworkbook.xlsx file.");
		}
	}

}
