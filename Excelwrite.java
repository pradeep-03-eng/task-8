package task8;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excelwrite {


			private static int columnCount;

			public static void main(String[] args) throws IOException {
				// TODO Auto-generated method stub

				//Get into the workbook
				XSSFWorkbook book= new XSSFWorkbook();
				
				//Create the sheet
				XSSFSheet sheet= book.createSheet("Sheet1");
				
				//Store the details->Name(String) Age(int) Email(String)
				Object[][] data = {
						
						{"Name","Age","Email"},
						{"John Doe","30","john@test.com"},
						{"Jane Doe","28","jane@test.com"},
						{"Bob Smith","35","bob@gmail.com"},
						{"Swapnil","37","swapnil@gmail.com"}
						
				};
				//Put the data into the sheet
				
				int rowCount = 0;
				
				//for each get into the each row
				
			for(Object[] row1 :data) {
				XSSFRow row = sheet.createRow(rowCount++);
				
				//for each to get the columns
				for(Object col : row1) {
					
					XSSFCell cell = row.createCell(columnCount++);
					
					//Checking the type of data and making the entry
					if(col instanceof String) {
					
						cell.setCellValue((String)col);
						
					}else if (col instanceof Integer) {
						
						cell.setCellValue((Integer)col);
						
						
					}
				}
			}
			
			try {
				FileOutputStream output = new FileOutputStream("C:\\Users\\DELL\\eclipse-workspace\\Excelfileoperation\\src\\main\\java\\task8\\Sheet1.xlsx");
				book.write(output);
			} catch (Exception e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
			book.close();
			}

		

	}


