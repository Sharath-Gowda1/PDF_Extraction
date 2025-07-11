package PDF_Extract;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.HashMap;
import java.util.Map;

public class CompareCashReceipts {

    public static void main(String[] args) {
        // Paths for the main and reference Excel files
        String file1Path = "C:\\Users\\Sharath.Gowda\\Downloads\\Credit_test.xlsx";  // Main file (to be updated)
        String file2Path = "C:\\Users\\Sharath.Gowda\\Downloads\\CreditReportToBeCompared.xlsx";     // Reference file

        try {
            // Open the files
            FileInputStream file1 = new FileInputStream(file1Path);
            FileInputStream file2 = new FileInputStream(file2Path);

            Workbook wb1 = new XSSFWorkbook(file1);
            Workbook wb2 = new XSSFWorkbook(file2);

            // Sheets from both workbooks
            Sheet sheet1 = wb1.getSheetAt(0); // Main sheet
            Sheet sheet2 = wb2.getSheetAt(0); // Reference sheet

            // Step 1: Read reference BLDG ID & Cash Receipts into a Map
            Map<String, String> referenceMap = new HashMap<>();

            // Find the columns for BLDG ID and Cash Receipts in the reference file
            int refBldgIdCol = -1, refCashCol = -1;
            Row refHeader = sheet2.getRow(0);
            for (Cell cell : refHeader) {
                if (cell.getStringCellValue().equalsIgnoreCase("BLDG ID")) {
                    refBldgIdCol = cell.getColumnIndex();
                } else if (cell.getStringCellValue().equalsIgnoreCase("Cash Receipts")) {
                    refCashCol = cell.getColumnIndex();
                }
            }

            // Populate the referenceMap with BLDG ID and corresponding Cash Receipts
            for (int i = 1; i <= sheet2.getLastRowNum(); i++) {
                Row row = sheet2.getRow(i);
                if (row == null) continue;
                Cell idCell = row.getCell(refBldgIdCol);
                Cell cashCell = row.getCell(refCashCol);
                if (idCell != null && cashCell != null) {
                    referenceMap.put(getCellValueAsString(idCell).trim(), getCellValueAsString(cashCell));
                }
            }

            // Step 2: Match and update results in sheet1
            int bldgIdCol = -1, cashCol = -1;
            Row headerRow = sheet1.getRow(0);
            for (Cell cell : headerRow) {
                if (cell.getStringCellValue().equalsIgnoreCase("BLDG ID")) {
                    bldgIdCol = cell.getColumnIndex();
                } else if (cell.getStringCellValue().equalsIgnoreCase("Cash Receipts")) {
                    cashCol = cell.getColumnIndex();
                }
            }

            // Add "Result" column
            int resultCol = headerRow.getLastCellNum(); // Next empty column
            Cell resultHeader = headerRow.createCell(resultCol);
            resultHeader.setCellValue("Result");

            // Step 3: Compare each row in the main sheet with the reference map
            for (int i = 1; i <= sheet1.getLastRowNum(); i++) {
                Row row = sheet1.getRow(i);
                if (row == null) continue;

                Cell idCell = row.getCell(bldgIdCol);
                Cell cashCell = row.getCell(cashCol);

                String bldgId = idCell != null ? getCellValueAsString(idCell).trim() : "";
                String cashValue1 = getCellValueAsString(cashCell);

                // Search for the BLDG ID in the reference file and compare Cash Receipts
                String expectedCash = searchInReferenceFile(sheet2, refBldgIdCol, refCashCol, bldgId);
                String result = (expectedCash != null && expectedCash.equals(cashValue1)) ? "PASS" : "FAIL";

                row.createCell(resultCol).setCellValue(result); // Write result into new column
            }

            file1.close();
            file2.close();

            // Save the updated file1 with results
            try (FileOutputStream out = new FileOutputStream(file1Path)) {
                wb1.write(out);
            }

            wb1.close();
            wb2.close();

            System.out.println("✅ Comparison complete. Result column added to: " + file1Path);

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    // Helper method to safely extract cell values as String
    private static String getCellValueAsString(Cell cell) {
        if (cell == null) return "";
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue().trim();
            case NUMERIC:
                return String.valueOf(Math.round(cell.getNumericCellValue()));
            case BOOLEAN:
                return Boolean.toString(cell.getBooleanCellValue());
            case FORMULA:
                return cell.getCellFormula();
            default:
                return "";
        }
    }

    // Search for BLDG ID in the reference sheet and return the corresponding Cash Receipts value
    private static String searchInReferenceFile(Sheet referenceSheet, int bldgIdCol, int cashCol, String bldgId) {
        for (int i = 1; i <= referenceSheet.getLastRowNum(); i++) {
            Row row = referenceSheet.getRow(i);
            if (row == null) continue;

            Cell idCell = row.getCell(bldgIdCol);
            if (idCell != null && getCellValueAsString(idCell).trim().equalsIgnoreCase(bldgId)) {
                // If BLDG ID matches, get the corresponding Cash Receipts value
                Cell cashCell = row.getCell(cashCol);
                if (cashCell != null) {
                    return getCellValueAsString(cashCell);
                }
            }
        }
        return null; // Return null if no matching BLDG ID is found
    }
}
