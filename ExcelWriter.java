package guviTask13;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;


public class ExcelWriter {
	
	
	
	    public static void main(String[] args) {
	        // Create a new Excel workbook and sheet
	        try (Workbook workbook = new XSSFWorkbook()) {
	            Sheet sheet = workbook.createSheet("Autojava.xlsx");

	            // Create a sample data array
	            Object[][] data = {
	                {"Name", "Age", "Email"},
	                {"John Doe", 30, "John@test.com"},
	                {"Jane Doe", 28, "John@test.com"},
	                {"Bob Johnson", 35, "jacky@example.com"},
	                {"Swapnil", 35, "joy@example.com"},
	            };
	            
	            

	            // Loop through the data and write it to the sheet
	            int rowNum = 0;
	            for (Object[] rowData : data) {
	                Row row = sheet.createRow(rowNum++);
	                int colNum = 0;
	                for (Object field : rowData) {
	                    Cell cell = row.createCell(colNum++);
	                    if (field instanceof String) {
	                        cell.setCellValue((String) field);
	                    } else if (field instanceof Integer) {
	                        cell.setCellValue((Integer) field);
	                    }
	                }
	            }

	            // Save the workbook to a file
	            try (FileOutputStream outputStream = new FileOutputStream("C:\\Users\\sanje\\OneDrive\\Desktop")) {
	                workbook.write(outputStream);
	                System.out.println("Excel file created successfully.");
	            } catch (IOException e) {
	                e.printStackTrace();
	            }
	        } catch (IOException e) {
	            e.printStackTrace();
	        }
	    }
	}



