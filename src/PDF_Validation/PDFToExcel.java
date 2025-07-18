package PDF_Validation;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.regex.*;
import java.util.*;

public class PDFToExcel {

    // Example: Regex pattern to extract data like "Name: John, Age: 30, Email: john@example.com"
    private static final Pattern DATA_PATTERN = Pattern.compile(
            "Name:\\s*(\\w+)\\s*,\\s*Age:\\s*(\\d+)\\s*,\\s*Email:\\s*(\\S+)"
    );

    public static void main(String[] args) {
        String pdfPath = "input.pdf";     // Your PDF file path
        String excelPath = "output.xlsx"; // Output Excel file path

        try {
            // Step 1: Load PDF and extract text
            PDDocument document = PDDocument.load(new File(pdfPath));
            PDFTextStripper stripper = new PDFTextStripper();
            String pdfText = stripper.getText(document);
            document.close();

            // Step 2: Use regex to find matches
            Matcher matcher = DATA_PATTERN.matcher(pdfText);
            List<String[]> extractedData = new ArrayList<>();

            while (matcher.find()) {
                String name = matcher.group(1);
                String age = matcher.group(2);
                String email = matcher.group(3);
                extractedData.add(new String[]{name, age, email});
            }

            // Step 3: Write to Excel
            Workbook workbook = new XSSFWorkbook();
            Sheet sheet = workbook.createSheet("Extracted Data");

            // Write header
            Row header = sheet.createRow(0);
            header.createCell(0).setCellValue("Name");
            header.createCell(1).setCellValue("Age");
            header.createCell(2).setCellValue("Email");

            // Write rows
            int rowIndex = 1;
            for (String[] row : extractedData) {
                Row excelRow = sheet.createRow(rowIndex++);
                for (int i = 0; i < row.length; i++) {
                    excelRow.createCell(i).setCellValue(row[i]);
                }
            }

            // Save to file
            FileOutputStream fileOut = new FileOutputStream(excelPath);
            workbook.write(fileOut);
            fileOut.close();
            workbook.close();

            System.out.println("Data extracted and written to Excel successfully.");

        } catch (IOException e) {
            e.printStackTrace(); //testing
        }
    }
}