package excel;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Date;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class TypesofCells
{

	public static void main(String[] args) throws IOException
	{
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet spreadsheet = workbook.createSheet("cell types");

		XSSFRow row = spreadsheet.createRow((short) 2);
		row.createCell(0).setCellValue("Type of Cell");
		row.createCell(1).setCellValue("cell value");

		row = spreadsheet.createRow((short) 3);
		row.createCell(0).setCellValue("set cell type BLANK");
		row.createCell(1);

		row = spreadsheet.createRow((short) 4);
		row.createCell(0).setCellValue("set cell type BOOLEAN");
		row.createCell(1).setCellValue(true);

		row = spreadsheet.createRow((short) 5);
		row.createCell(0).setCellValue("set cell type ERROR");
		row.createCell(1).setCellValue(XSSFCell.CELL_TYPE_ERROR);

		row = spreadsheet.createRow((short) 6);
		row.createCell(0).setCellValue("set cell type date");
		row.createCell(1).setCellValue(new Date(0));

		row = spreadsheet.createRow((short) 7);
		row.createCell(0).setCellValue("set cell type numeric");
		row.createCell(1).setCellValue(20);

		row = spreadsheet.createRow((short) 8);
		row.createCell(0).setCellValue("set cell type string");
		row.createCell(1).setCellValue("A String");

		FileOutputStream out = new FileOutputStream(new File("typesofcells.xlsx"));
		workbook.write(out);
		out.close();
		System.out.println("typesofcells.xlsx written successfully");
	}

}
