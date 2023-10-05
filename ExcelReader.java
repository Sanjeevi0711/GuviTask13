package guviTask13;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;

public class ExcelReader {
	


	
	    public static void main(String[] args) {
	        try (FileInputStream fis = new FileInputStream("Autojava.xlsx"); 
	             Workbook workbook = new XSSFWorkbook(fis)) {

	            // Get the first sheet in the Excel workbook
	            Sheet sheet = workbook.getSheetAt(0);

	            // Iterate through rows and columns to read data
	            for (Row row : sheet) {
	                for (Cell cell : row) {
	                    switch (cell.getCellType()) {
	                        case STRING:
	                            System.out.print(cell.getStringCellValue() + "\t");
	                            break;
	                        case NUMERIC:
	                            if (DateUtil.isCellDateFormatted(cell)) {
	                                System.out.print(cell.getDateCellValue() + "\t");
	                            } else {
	                                System.out.print(cell.getNumericCellValue() + "\t");
	                            }
	                            break;
	                        case BOOLEAN:
	                            System.out.print(cell.getBooleanCellValue() + "\t");
	                            break;
	                        case FORMULA:
	                            System.out.print(cell.getCellFormula() + "\t");
	                            break;
	                        default:
	                            System.out.print("\t");
	                    }
	                }
	                System.out.println(); // Move to the next line after each row
	            }
	        } catch (IOException e) {
	            e.printStackTrace();
	        }
	    }
	


}
