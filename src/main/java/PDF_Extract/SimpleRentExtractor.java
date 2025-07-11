package PDF_Extract;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.regex.*;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.regex.*;

public class SimpleRentExtractor {
    public static void main(String[] args) throws Exception {
        // 1. Load the PDF file (update path)
        File pdfFile = new File("C:\\Users\\Sharath.Gowda\\Downloads\\CM RentRoll_ALL_11.30.24 (8) (3).pdf");
        PDDocument document = PDDocument.load(pdfFile);

        // 2. Extract all text from the PDF
        PDFTextStripper pdfStripper = new PDFTextStripper();
        String text = pdfStripper.getText(document);
        document.close();

        // 3. Prepare Excel file
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Rent Data");
        Row header = sheet.createRow(0);
        header.createCell(0).setCellValue("Bldg Id");
        header.createCell(1).setCellValue("Suite Id");
        header.createCell(2).setCellValue("Monthly Rent");

        int rowNum = 1;

        // 4. Define regex for extracting values
        // Pattern: BldgId (4 digits), SuiteId (4 digits), anything, then a number (monthly rent)
        Pattern rentPattern = Pattern.compile("(\\d{4})\\s+(\\d{4})\\s+.*?([\\d,]+)\\s*$");

        // 5. Parse line-by-line
        String[] lines = text.split("\\r?\\n");
        for (String line : lines) {
            Matcher matcher = rentPattern.matcher(line);
            if (matcher.find()) {
                String bldgId = matcher.group(1);
                String suiteId = matcher.group(2);
                String rentStr = matcher.group(3).replace(",", "");

                Row row = sheet.createRow(rowNum++);
                row.createCell(0).setCellValue(bldgId);
                row.createCell(1).setCellValue(suiteId);
                row.createCell(2).setCellValue(Double.parseDouble(rentStr));
            }
        }

        // 6. Write to Excel
        FileOutputStream fileOut = new FileOutputStream("C:\\\\Users\\\\Sharath.Gowda\\\\Downloads\\\\MonthlyRentData.xlsx");
        workbook.write(fileOut);
        fileOut.close();
        workbook.close();

        System.out.println("Excel file created: MonthlyRentData.xlsx");
    }
}
