package excel;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CreateWorkBook
{

	public static void main(String[] args) throws IOException
	{
		// Create Blank workbook
		XSSFWorkbook workbook = new XSSFWorkbook();
		// Create file system using specific name
		FileOutputStream out = new FileOutputStream(new File("createworkbook.xlsx"));
		// write operation workbook using file out object
		workbook.write(out);
		out.close();
		System.out.println("createworkbook.xlsx written successfully");
	}

}
