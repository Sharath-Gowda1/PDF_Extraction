package PDF_Extract;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class PDFExtractor {

    public static void main(String[] args) {
        String filePath = "C:\\\\Users\\\\Sharath.Gowda\\\\Downloads\\\\Boyer Aged.pdf"; // <-- Update this
        String outputExcelPath = "C:\\\\Users\\\\Sharath.Gowda\\\\Downloads\\\\Boyer Aged.xlsx";

        try (PDDocument document = PDDocument.load(new File(filePath))) {
            PDFTextStripper stripper = new PDFTextStripper();
            String text = stripper.getText(document);

            List<String> buildingIds = extractBuildingIDs(text);
            List<List<String>> totalDataRows = extractTotalDataRows(text);

            writeToExcel(buildingIds, totalDataRows, outputExcelPath);
            System.out.println("Excel file created at: " + outputExcelPath);

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static List<String> extractBuildingIDs(String text) {
        List<String> ids = new ArrayList<>();
        Pattern pattern = Pattern.compile("\\b\\d{4}-\\d{6}\\b");
        Matcher matcher = pattern.matcher(text);
        while (matcher.find()) {
            ids.add(matcher.group());
        }
        return ids;
    }

    private static List<List<String>> extractTotalDataRows(String text) {
        List<List<String>> dataRows = new ArrayList<>();
        Pattern pattern = Pattern.compile("Total:\\s+([\\d,]+\\.\\d{2}\\s+)+");
        Matcher matcher = pattern.matcher(text);

        while (matcher.find()) {
            String line = matcher.group().replace("Total:", "").trim();
            String[] values = line.split("\\s+");
            dataRows.add(Arrays.asList(values));
        }

        return dataRows;
    }

    private static void writeToExcel(List<String> buildingIds, List<List<String>> totalRows, String outputPath) throws IOException {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Extracted Data");

        // Custom headers
        String[] headers = {"BLDG ID", "Monthly Base rent", "Monthly cost recovery", "Monthly other income", "2nd Month",
                "3rd Month", "4th Month"};

        // Write header row
        Row headerRow = sheet.createRow(0);
        for (int i = 0; i < headers.length; i++) {
            headerRow.createCell(i).setCellValue(headers[i]);
        }

        int maxRows = Math.max(buildingIds.size(), totalRows.size());

        for (int i = 0; i < maxRows; i++) {
            Row row = sheet.createRow(i + 1);

            // BLDG ID
            if (i < buildingIds.size()) {
                row.createCell(0).setCellValue(buildingIds.get(i));
            }

            // Total row values
            if (i < totalRows.size()) {
                List<String> values = totalRows.get(i);
                for (int j = 0; j < values.size() && j < headers.length - 1; j++) {
                    try {
                        double value = Double.parseDouble(values.get(j).replace(",", ""));
                        row.createCell(j + 1).setCellValue(value);
                    } catch (NumberFormatException e) {
                        row.createCell(j + 1).setCellValue(values.get(j));
                    }
                }
            }
        }

        // Autosize columns
        for (int i = 0; i < headers.length; i++) {
            sheet.autoSizeColumn(i);
        }

        try (FileOutputStream fileOut = new FileOutputStream(outputPath)) {
            workbook.write(fileOut);
        }
        workbook.close();
    }
}
