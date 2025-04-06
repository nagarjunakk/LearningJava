package excel_operations;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelDataRead_forLoop {

	
	
	
	public static void main(String[] args) throws FileNotFoundException {
		try {
		String filepath=".\\src\\test\\resources\\excelData\\DataSheet.xlsx";
		FileInputStream inputStream= new FileInputStream(filepath);
		
			XSSFWorkbook workbook= new XSSFWorkbook(inputStream);
			XSSFSheet sheet=workbook.getSheet("QASheet");
				
			int rowCunt= sheet.getLastRowNum();
			int columnCount=sheet.getRow(1).getLastCellNum();
					
//			System.out.println(rowCunt);
//			System.out.println(columnCount);
			// for loop
			
			for (int r = 1; r <= rowCunt; r++) {
				
				
				XSSFRow row=sheet.getRow(r);
				for (int c = 0; c < columnCount; c++) {
				//NUMERIC STRING
					XSSFCell cell=row.getCell(c);
//					System.out.println(cell.getCellType());
//					System.out.println(cell);
				if(cell != null) {					
					switch (cell.getCellType()) {
					case STRING:System.out.print(cell.getStringCellValue());
						break;
					case NUMERIC:System.out.print((int) cell.getNumericCellValue());
						break;
						
					case BOOLEAN:System.out.print(cell.getBooleanCellValue());
						break;

					default:System.out.print("Unsupported format");
						break;
					}
					System.out.print("|| ");
					
				}else System.out.println("null||");
				
				}
				
				System.out.println();
				
			}
			
			
			
			
			
		} catch (IOException e) {
			
			e.printStackTrace();
		}

	}

}
