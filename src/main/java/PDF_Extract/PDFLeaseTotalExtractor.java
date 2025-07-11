package PDF_Extract;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.*;
import java.util.regex.*;

public class PDFLeaseTotalExtractor {

    public static void main(String[] args) {
        String pdfFilePath = "C:\\Users\\Sharath.Gowda\\Downloads\\Benderson Aged.pdf";
        String excelOutputPath = "C:\\Users\\Sharath.Gowda\\Downloads\\Benderson_PDFDATA.xlsx";

        Pattern leaseLinePattern = Pattern.compile("(?m)^(\\d{4}-\\d{6})\\s+(.+)$");

        Pattern totalLinePattern = Pattern.compile("(?m)^(.+?)\\s+Total:\\s*([\\d,.-]+)\\s+([\\d,.-]+)\\s+([\\d,.-]+)\\s+([\\d,.-]+)\\s+([\\d,.-]+)\\s+([\\d,.-]+)");

        class LeaseRecord {
            String leaseId;
            String company;
            String[] totals = {"", "", "", "", "", ""};
        }

        List<LeaseRecord> records = new ArrayList<>();

        try (PDDocument document = PDDocument.load(new File(pdfFilePath))) {
            PDFTextStripper stripper = new PDFTextStripper();
            stripper.setStartPage(1);
            stripper.setEndPage(document.getNumberOfPages());

            String text = stripper.getText(document);
            String[] lines = text.split("\\r?\\n");

            LeaseRecord currentRecord = null;

            for (String line : lines) {
                Matcher leaseMatcher = leaseLinePattern.matcher(line);
                Matcher totalMatcher = totalLinePattern.matcher(line);

                if (leaseMatcher.find()) {
                    currentRecord = new LeaseRecord();
                    currentRecord.leaseId = leaseMatcher.group(1).trim();
                    currentRecord.company = leaseMatcher.group(2).trim();
                    records.add(currentRecord);
                } else if (totalMatcher.find() && currentRecord != null) {
                    String[] totals = new String[6];
                    for (int i = 0; i < 6; i++) {
                        totals[i] = totalMatcher.group(i + 2).trim();
                    }
                    currentRecord.totals = totals;
                }
            }

            try (Workbook workbook = new XSSFWorkbook()) {
                Sheet sheet = workbook.createSheet("Lease Totals");
                String[] headers = {"Lease ID", "Company", "Amount", "Current", "1st month", "2nd month", "3rd month", "4th month"};

                Row headerRow = sheet.createRow(0);
                for (int i = 0; i < headers.length; i++) {
                    headerRow.createCell(i).setCellValue(headers[i]);
                }

                int rowIndex = 1;
                for (LeaseRecord record : records) {
                    Row row = sheet.createRow(rowIndex++);
                    row.createCell(0).setCellValue(record.leaseId);
                    row.createCell(1).setCellValue(record.company);
                    for (int i = 0; i < record.totals.length; i++) {
                        row.createCell(i + 2).setCellValue(record.totals[i]);
                    }
                }

                try (FileOutputStream out = new FileOutputStream(excelOutputPath)) {
                    workbook.write(out);
                    System.out.println("Sucessfully extracted the data to the Excel: " + excelOutputPath);
                }
            }

        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
