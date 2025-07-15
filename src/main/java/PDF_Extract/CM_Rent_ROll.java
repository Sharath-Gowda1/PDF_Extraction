package PDF_Extract;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.*;
import java.util.regex.*;

public class CM_Rent_ROll {

    public static void main(String[] args) {
        String filePath = "C:\\\\Users\\\\Sharath.Gowda\\\\Downloads\\\\CM RentRoll_ALL_11.30.24 (18) 1.pdf";
        String outputExcelPath = "C:\\\\Users\\\\Sharath.Gowda\\\\Downloads\\\\RentRoll_Extracted.xlsx";

        try (PDDocument document = PDDocument.load(new File(filePath))) {
            PDFTextStripper stripper = new PDFTextStripper();
            String text = stripper.getText(document);

            List<Map<String, String>> records = extractBuildingData(text);
            writeToExcel(records, outputExcelPath);

            System.out.println("✅ Extracted " + records.size() + " records.");
            System.out.println("✅ Excel saved at: " + outputExcelPath);

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static List<Map<String, String>> extractBuildingData(String text) {
        List<Map<String, String>> records = new ArrayList<>();

        // Match 1 to 4 digit building IDs
        Pattern buildingPattern = Pattern.compile("(?i)(Building Id:\\s*(\\d{1,4}))");
        Matcher buildingMatcher = buildingPattern.matcher(text);

        List<Integer> buildingIndices = new ArrayList<>();
        List<String> buildingIds = new ArrayList<>();

        while (buildingMatcher.find()) {
            String rawId = buildingMatcher.group(2).trim();
            String paddedId = String.format("%04d", Integer.parseInt(rawId));
            buildingIds.add(paddedId);
            buildingIndices.add(buildingMatcher.start());
        }

        // Add end boundary for final section
        buildingIndices.add(text.length());

        for (int i = 0; i < buildingIds.size(); i++) {
            String section = text.substring(buildingIndices.get(i), buildingIndices.get(i + 1));
            String buildingId = buildingIds.get(i);

            // Look for 'Totals:' block with up to next 3 lines
            Pattern totalsBlockPattern = Pattern.compile("Totals:(?:.*\\n?){1,3}");
            Matcher totalsBlockMatcher = totalsBlockPattern.matcher(section);

            if (totalsBlockMatcher.find()) {
                String totalsBlock = totalsBlockMatcher.group();

                // Extract all valid decimal/number entries
                Pattern numberPattern = Pattern.compile("(-?\\d{1,3}(?:,\\d{3})*(?:\\.\\d{2}))");
                Matcher numMatcher = numberPattern.matcher(totalsBlock);

                List<String> numbers = new ArrayList<>();
                while (numMatcher.find()) {
                    numbers.add(numMatcher.group(1));
                }

                // Grab last 3 values only if we have them
                if (numbers.size() >= 3) {
                    int n = numbers.size();
                    Map<String, String> record = new LinkedHashMap<>();
                    record.put("BLDG ID", buildingId);
                    record.put("Monthly Base Rent", numbers.get(n - 3));
                    record.put("Monthly Cost Recovery", numbers.get(n - 2));
                    record.put("Monthly Other Income", numbers.get(n - 1));
                    records.add(record);
                }
            }
        }

        return records;
    }

    private static void writeToExcel(List<Map<String, String>> records, String outputPath) throws IOException {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Rent Roll Totals");

        String[] headers = {"BLDG ID", "Monthly Base Rent", "Monthly Cost Recovery", "Monthly Other Income"};
        Row headerRow = sheet.createRow(0);

        for (int i = 0; i < headers.length; i++) {
            headerRow.createCell(i).setCellValue(headers[i]);
        }

        int rowNum = 1;
        double totalBase = 0.0, totalRecovery = 0.0, totalOther = 0.0;

        for (Map<String, String> record : records) {
            Row row = sheet.createRow(rowNum++);
            for (int i = 0; i < headers.length; i++) {
                String val = record.get(headers[i]);

                if (headers[i].equals("BLDG ID")) {
                    row.createCell(i).setCellValue(val);
                } else {
                    try {
                        double num = Double.parseDouble(val.replace(",", ""));
                        row.createCell(i).setCellValue(num);

                        // Accumulate totals
                        if (headers[i].equals("Monthly Base Rent")) totalBase += num;
                        if (headers[i].equals("Monthly Cost Recovery")) totalRecovery += num;
                        if (headers[i].equals("Monthly Other Income")) totalOther += num;

                    } catch (NumberFormatException e) {
                        row.createCell(i).setCellValue(val);
                    }
                }
            }
        }

        // Write Grand Totals Row
        Row totalRow = sheet.createRow(rowNum);
        totalRow.createCell(0).setCellValue("GRAND TOTAL");
        totalRow.createCell(1).setCellValue(totalBase);
        totalRow.createCell(2).setCellValue(totalRecovery);
        totalRow.createCell(3).setCellValue(totalOther);

        for (int i = 0; i < headers.length; i++) {
            sheet.autoSizeColumn(i);
        }

        try (FileOutputStream out = new FileOutputStream(outputPath)) {
            workbook.write(out);
        }
        workbook.close();
    }
}
