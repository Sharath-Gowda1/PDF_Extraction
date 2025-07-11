package PDF_Extract;

/*import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class Benderson_Units_Mismatches {
    public static void main(String[] args) {
        try {
            File pdfFile = new File("C:\\\\Users\\\\Sharath.Gowda\\\\Downloads\\\\Benderson Rent Roll.pdf"); // Replace with your PDF file path
            PDDocument document = PDDocument.load(pdfFile);
            PDFTextStripper pdfStripper = new PDFTextStripper();
            String text = pdfStripper.getText(document);
            document.close();

            // Regex for Totals section
            Pattern pattern = Pattern.compile("Totals:\\s*.*?\\n.*?(\\d+) Units.*?\\n.*?(\\d+) Units.*?\\n.*?(\\d+) Units", Pattern.DOTALL);
            Matcher matcher = pattern.matcher(text);

            // Regex for Building ID
            Pattern buildingIdPattern = Pattern.compile("^(\\d{4}[A-Z]?)\\s*-", Pattern.MULTILINE);
            Matcher buildingIdMatcher = buildingIdPattern.matcher(text);

            String buildingId = "Unknown";
            if (buildingIdMatcher.find()) {
                buildingId = buildingIdMatcher.group(1);
            }

            if (matcher.find()) {
                String occupied = matcher.group(1);
                String vacant = matcher.group(2);
                String total = matcher.group(3);

                // Create Excel workbook and sheet
                Workbook workbook = new XSSFWorkbook();
                Sheet sheet = workbook.createSheet("Rent Roll Summary");

                // Create header row
                Row header = sheet.createRow(0);
                header.createCell(0).setCellValue("Building ID");
                header.createCell(1).setCellValue("Occupied Units");
                header.createCell(2).setCellValue("Vacant Units");
                header.createCell(3).setCellValue("Total Units");

                // Write data row
                Row dataRow = sheet.createRow(1);
                dataRow.createCell(0).setCellValue(buildingId);
                dataRow.createCell(1).setCellValue(Integer.parseInt(occupied));
                dataRow.createCell(2).setCellValue(Integer.parseInt(vacant));
                dataRow.createCell(3).setCellValue(Integer.parseInt(total));

                // Write to Excel file
                FileOutputStream fileOut = new FileOutputStream("C:\\\\Users\\\\Sharath.Gowda\\\\Downloads\\\\Bendersondata1.xlsx");
                workbook.write(fileOut);
                fileOut.close();
                workbook.close();

                System.out.println("Data written to RentRollSummary.xlsx");

            } else {
                System.out.println("Totals section not found.");
            }

        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}*/


import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class Benderson_Units_Mismatches 
{
    public static void main(String[] args) 
    {
        String pdfPath = "C:\\Users\\Sharath.Gowda\\Downloads\\Benderson Rent Roll.pdf";
        String excelPath = "C:\\Users\\Sharath.Gowda\\Downloads\\Bendersondata1.xlsx";

        try (PDDocument document = PDDocument.load(new File(pdfPath));
             Workbook workbook = new XSSFWorkbook()) {

            Sheet sheet = workbook.createSheet("Rent Roll Summary");

            Row header = sheet.createRow(0);
            header.createCell(0).setCellValue("Building ID");
            header.createCell(1).setCellValue("Occupied Units");
            header.createCell(2).setCellValue("Vacant Units");
            header.createCell(3).setCellValue("Total Units");
            
            PDFTextStripper pdfStripper = new PDFTextStripper();

            Pattern totalsPattern = Pattern.compile(
                "Totals:\\s*.*?\\n.*?(\\d+) Units.*?\\n.*?(\\d+) Units.*?\\n.*?(\\d+) Units",
                Pattern.DOTALL);
            Pattern buildingIdPattern = Pattern.compile("^(\\d{4}[A-Z]?)\\s*-", Pattern.MULTILINE);

            int excelRowIndex = 1;

            int totalPages = document.getNumberOfPages();
            for (int i = 1; i <= totalPages; i++) {
                pdfStripper.setStartPage(i);
                pdfStripper.setEndPage(i);

                String pageText = pdfStripper.getText(document);

                Matcher buildingIdMatcher = buildingIdPattern.matcher(pageText);
                String buildingId = "Unknown";
                if (buildingIdMatcher.find()) {
                    buildingId = buildingIdMatcher.group(1);
                }

                Matcher totalsMatcher = totalsPattern.matcher(pageText);
                if (totalsMatcher.find()) {
                    String occupied = totalsMatcher.group(1);
                    String vacant = totalsMatcher.group(2);
                    String total = totalsMatcher.group(3);

                    Row dataRow = sheet.createRow(excelRowIndex++);
                    dataRow.createCell(0).setCellValue(buildingId);
                    dataRow.createCell(1).setCellValue(Integer.parseInt(occupied));
                    dataRow.createCell(2).setCellValue(Integer.parseInt(vacant));
                    dataRow.createCell(3).setCellValue(Integer.parseInt(total));
                }
            }

            for (int col = 0; col <= 3; col++) {
                sheet.autoSizeColumn(col);
            }

            try (FileOutputStream fileOut = new FileOutputStream(excelPath)) {
                workbook.write(fileOut);
            }

            System.out.println("Data extracted from all pages and written to Excel successfully.");

        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}


