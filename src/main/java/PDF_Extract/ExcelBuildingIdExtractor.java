package PDF_Extract;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;

public class ExcelBuildingIdExtractor {

    public static void main(String[] args) {
        String excelFilePath = "C:\\Users\\Sharath.Gowda\\Downloads\\Bend_BLDGID_Extraction.xlsx";

        try (FileInputStream fis = new FileInputStream(excelFilePath);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0);
            boolean isFirstRow = true;

            for (Row row : sheet) {
                if (isFirstRow) {
                    isFirstRow = false; 
                    continue;
                }

                Cell cell = row.getCell(0);
                if (cell != null && cell.getCellType() == CellType.STRING) {
                    String originalValue = cell.getStringCellValue();
                    String[] parts = originalValue.split(" - ");
                    if (parts.length >= 2) {
                        String updatedValue = parts[1]; // Extract "1005B" from "1005 - 1005B - ..."
                        cell.setCellValue(updatedValue);
                    }
                }
            }

            // Overwrite the original file with updated content
            try (FileOutputStream fos = new FileOutputStream(excelFilePath)) {
                workbook.write(fos);
                System.out.println("Excel file updated successfully.");
            }

        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
