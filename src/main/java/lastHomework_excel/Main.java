package lastHomework_excel;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Main {

	public static void main(String[] args) throws FileNotFoundException, 
	IOException {
		
		FirstSheet(); SecondSheet();
	}

	public static void FirstSheet() throws FileNotFoundException, IOException {

		// Creating Excel File 
		// Creating the excel sheet
		// creating sheet ROW
		// creating sheet COLUMN
		// writing inside the column
		// saving the  file
		// closing the file
		
		
		XSSFWorkbook wb = new XSSFWorkbook();
		XSSFSheet excelSheet = wb.createSheet("FirstTestSheet");
		XSSFRow row = excelSheet.createRow(0);
		XSSFCell cell = row.createCell(0);
		cell.setCellValue("testing by Leen");
		wb.write(new FileOutputStream("FirstSheet.xlsx"));
		wb.close();

	}
// The second sheet that contains testing data 
	
	public static void SecondSheet() throws FileNotFoundException, IOException {

		// create excel file
		XSSFWorkbook wb = new XSSFWorkbook();

		// create excel sheet
		XSSFSheet sheet = wb.createSheet("FirstSheet");
		
		Object data[][] = { {"Test Case Name", "UserName", "Password", "Result"},
				{"Apache_POI_TC", "testuser_1","Test@123",""}
				
		};
		
		for(int r=0; r<data.length; r++) {
			
			// create excel row
			XSSFRow row = sheet.createRow(r);
			
			for(int c=0; c<data[0].length; c++) {
				
				// create excel column
				XSSFCell cell = row.createCell(c);

				// write in column
				cell.setCellValue(data[r][c].toString());
			}
		}

		// save file
		wb.write(new FileOutputStream("SecondSheet.xlsx"));

		// close file
		wb.close();

	}

}
