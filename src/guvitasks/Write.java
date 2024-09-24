package guvitasks;



   
    	import org.apache.poi.ss.usermodel.*;
    	import org.apache.poi.xssf.usermodel.XSSFWorkbook;

    	import java.io.FileOutputStream;
    	import java.io.IOException;

    	public class Write {
    	    public static void main(String[] args) {
    	        // Column headers
    	        String[] columns = {"Name", "Age", "Email"};
    	        
    	        // Data to write to the sheet
    	        String[][] data = {
    	            {"John Doe", "30", "john@test.com"},
    	            {"Jane Doe", "28", "jane@test.com"},
    	            {"Bob Smith", "35", "jacky@example.com"},
    	            {"Swapnil", "37", "swapnil@example.com"}
    	        };

    	        // Create a new workbook and a sheet
    	        Workbook workbook = new XSSFWorkbook();
    	        Sheet sheet = workbook.createSheet("Sheet1");

    	        // Create header row
    	        Row headerRow = sheet.createRow(0);
    	        for (int i = 0; i < columns.length; i++) {
    	            Cell cell = headerRow.createCell(i);
    	            cell.setCellValue(columns[i]);
    	        }

    	        // Write data rows
    	        for (int i = 0; i < data.length; i++) {
    	            Row row = sheet.createRow(i + 1); // Start writing from row 1
    	            for (int j = 0; j < data[i].length; j++) {
    	                Cell cell = row.createCell(j);
    	                cell.setCellValue(data[i][j]);
    	            }
    	        }

    	        // Auto-size all columns to fit the content
    	        for (int i = 0; i < columns.length; i++) {
    	            sheet.autoSizeColumn(i);
    	        }

    	        // Write the output to a file
    	        try (FileOutputStream fileOut = new FileOutputStream("data.xlsx")) {
    	            workbook.write(fileOut);
    	            System.out.println("Excel file 'data.xlsx' created successfully.");
    	        } catch (IOException e) {
    	            e.printStackTrace();
    	        }

    	        // Close the workbook
    	        try {
    	            workbook.close();
    	        } catch (IOException e) {
    	            e.printStackTrace();
    	        }
    	    }
    	}
