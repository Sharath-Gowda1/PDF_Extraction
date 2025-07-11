package PDF_Extract;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.Iterator;

public class RemoveZeroAmountRows {
    public static void main(String[] args) {
        String inputFilePath = "C:\\\\Users\\\\Sharath.Gowda\\\\Downloads\\\\Remove0RowsBase.xlsx";     // Replace with your file path
        String outputFilePath = "C:\\\\Users\\\\Sharath.Gowda\\\\Downloads\\\\Remove0Rows.xlsx"; // Output file

        try (FileInputStream fis = new FileInputStream(inputFilePath);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0); // Assuming data is in first sheet

            int firstRowNum = sheet.getFirstRowNum();
            int lastRowNum = sheet.getLastRowNum();

            // Iterate from the bottom to avoid row index shifting when removing rows
            for (int i = lastRowNum; i > firstRowNum; i--) {
                Row row = sheet.getRow(i);
                if (row != null && isAllAmountsZero(row)) {
                    removeRow(sheet, i);
                }
            }
            long start = System.currentTimeMillis();
         // code
         long end = System.currentTimeMillis();
         System.out.println("Time: " + (end - start) + "ms");

            try (FileOutputStream fos = new FileOutputStream(outputFilePath)) {
                workbook.write(fos);
                System.out.println("Rows with all-zero amounts removed successfully.");
            }

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    // Check if all relevant amount columns are 0.00
    private static boolean isAllAmountsZero(Row row) {
        // Columns: Amount = 2, Current = 3, 1st = 4, 2nd = 5, 3rd = 6, 4th = 7
        int[] amountColumns = {2, 3, 4, 5, 6, 7};
        for (int colIndex : amountColumns) {
            Cell cell = row.getCell(colIndex, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
            double value = getNumericValue(cell);
            if (value != 0.0) {
                return false;
            }
        }
        return true;
    }

    // Get numeric value from cell (handle string/blank)
    private static double getNumericValue(Cell cell) {
        if (cell == null) return 0.0;
        try {
            if (cell.getCellType() == CellType.NUMERIC) {
                return cell.getNumericCellValue();
            } else if (cell.getCellType() == CellType.STRING) {
                String val = cell.getStringCellValue().replace(",", "").trim();
                return Double.parseDouble(val);
            }
        } catch (Exception e) {
            // Could not parse, assume 0
            return 0.0;
        }
        return 0.0;
    }

    // Utility to remove a row properly in Excel sheet
    private static void removeRow(Sheet sheet, int rowIndex) {
        int lastRowNum = sheet.getLastRowNum();
        if (rowIndex >= 0 && rowIndex < lastRowNum) {
            sheet.shiftRows(rowIndex + 1, lastRowNum, -1);
        } else if (rowIndex == lastRowNum) {
            Row removingRow = sheet.getRow(rowIndex);
            if (removingRow != null) {
                sheet.removeRow(removingRow);
            }
        }
    }
}
