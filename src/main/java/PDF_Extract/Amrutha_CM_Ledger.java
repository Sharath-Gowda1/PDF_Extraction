package PDF_Extract;

import org.apache.pdfbox.pdmodel.PDDocument;

import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class Amrutha_CM_Ledger {

	public static void main(String[] args) {
		String pdfPath = "C:\\Users\\Sharath.Gowda\\Downloads\\MRI_LEDGER_11.24_ALL_PROD (1).pdf";
		String excelPath = "C:\\Users\\Sharath.Gowda\\Downloads\\Boyer Aged.xlsx";

		List<String[]> extractedData = new ArrayList<>();
		Set<String> bldgIds = new LinkedHashSet<>();
		String currentBLDG = "";

		try (PDDocument document = PDDocument.load(new File(pdfPath))) {
			PDFTextStripper stripper = new PDFTextStripper();
			stripper.setSortByPosition(true);

			for (int page = 1; page <= document.getNumberOfPages(); page++) {
				stripper.setStartPage(page);
				stripper.setEndPage(page);

				String text = stripper.getText(document);
				String[] lines = text.split("\\r?\\n");

				for (String line : lines) {
					line = line.trim();

					if (line.toUpperCase().contains("BLDG:")) {
						Matcher matcher = Pattern.compile("(?i)BLDG:\\s*([A-Za-z0-9-]+)").matcher(line);
						if (matcher.find()) {
							String bldgId = matcher.group(1);
							bldgIds.add(bldgId);
							currentBLDG = bldgId;
						}
					}

					if (line.startsWith("Total:")) {
						String dataPart = line.replace("Total:", "").trim();
						String[] values = dataPart.split("\\s+");

						String[] row = new String[values.length + 1];
						row[0] = currentBLDG.isEmpty() ? "UNKNOWN" : currentBLDG;
						System.arraycopy(values, 0, row, 1, values.length);

						extractedData.add(row);
					}
				}
			}

			writeToExcel(extractedData, excelPath);

			System.out.println("Data extraction complete. Excel saved at: " + excelPath);

		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	private static void writeToExcel(List<String[]> data, String excelPath) {
		Workbook workbook = new XSSFWorkbook();
		Sheet sheet = workbook.createSheet("BLDG Totals");

		// Excel header
		String[] header = { "BLDG ID", "Mo. Rep Charges", "Beg Balance", "Charges", "Cash Receipts", "N/C Credits",
				"Refunds", "End Balance", "Sec Dep Bal" };

		Row headerRow = sheet.createRow(0);
		for (int i = 0; i < header.length; i++) {
			headerRow.createCell(i).setCellValue(header[i]);
		}

		// Data rows
		int rowNum = 1;
		for (String[] rowData : data) {
			Row row = sheet.createRow(rowNum++);
			for (int i = 0; i < rowData.length; i++) {
				row.createCell(i).setCellValue(rowData[i]);
			}
		}
		
		// Save Excel
		try (FileOutputStream fileOut = new FileOutputStream(excelPath)) {
			workbook.write(fileOut);
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
}