package task8;

import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Readexcel {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub

		// Open the workbook
		XSSFWorkbook book = new XSSFWorkbook("C:\\Users\\DELL\\eclipse-workspace\\Excelfileoperation\\src\\main\\java\\task8\\Sheet1.xlsx");
		
		//Get into the sheet
		XSSFSheet sheet = book.getSheet("Sheet1");
		
		//Get the no.of.rows
		int rowCount = sheet.getLastRowNum();
		
		//Get the no.of coloumns
		int columnCount = sheet.getRow(0).getLastCellNum();
		
		for(int i = 1 ; i <= rowCount; i++) {
			XSSFRow row = sheet.getRow(i);
			
			// Get into the coloumns
			for(int j = 0 ; j<columnCount;j++) {
				XSSFCell cell = row.getCell(j);
				
				//Read the value
				System.out.println(cell.getStringCellValue());
				
			}
			System.out.println();
			
		}
		book.close();
		
	}

}
