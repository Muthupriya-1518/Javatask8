package JavaTask8;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.FileOutputStream;
import java.io.IOException;

public class WriteExcelExample {
    public static void main(String[] args) {
        // Create a new workbook
        Workbook workbook = new XSSFWorkbook();

        // Create a sheet with name "Sheet1"
        Sheet sheet = workbook.createSheet("Sheet1");

        // Data to write
        String[][] data = {
                {"Name", "Age", "Email"},
                {"John Doe", "30", "john@test.com"},
                {"Jane Doe", "28", "john@test.com"},
                {"Bob Smith", "35", "jacky@example.com"},
                {"Swapnil", "37", "swapnil@example.com"}
        };

        // Write data row by row
        for (int i = 0; i < data.length; i++) {
            Row row = sheet.createRow(i);
            for (int j = 0; j < data[i].length; j++) {
                Cell cell = row.createCell(j);
                cell.setCellValue(data[i][j]);
            }
        }

        // Write to Excel file
        try (FileOutputStream fileOut = new FileOutputStream("FileOperations.xlsx")) {
            workbook.write(fileOut);
            System.out.println("Excel file created successfully!");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
