package com.amusedgroup;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.Normalizer;
import java.util.regex.Pattern;

public class ExcelPlayerNameNormalizer {

    public static void normalizePlayerNames(String inputFilePath, String outputFilePath) throws IOException {
        try (FileInputStream fileInputStream = new FileInputStream(inputFilePath);
             Workbook workbook = new XSSFWorkbook(fileInputStream);
             Workbook newWorkbook = new XSSFWorkbook()) {

            // Iterate through each sheet in the input workbook
            for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++) {
                Sheet sheet = workbook.getSheetAt(sheetIndex);
                Sheet newSheet = newWorkbook.createSheet(sheet.getSheetName()); // Create a new sheet in the new workbook

                // Iterate through each row in the sheet
                for (int rowIndex = 0; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                    Row row = sheet.getRow(rowIndex);
                    Row newRow = newSheet.createRow(rowIndex); // Create a new row in the new sheet

                    if (row != null) {
                        // Iterate through each cell in the row
                        for (int colIndex = 0; colIndex < row.getLastCellNum(); colIndex++) {
                            Cell cell = row.getCell(colIndex);
                            Cell newCell = newRow.createCell(colIndex); // Create a new cell in the new row

                            if (cell != null && cell.getCellType() == CellType.STRING) {
                                String cellValue = cell.getStringCellValue();
                                String normalizedValue = normalizeString(cellValue);

                                // Check if the value has changed
                                if (!cellValue.equals(normalizedValue)) {
                                    System.out.println("Special characters found in Sheet: " + sheet.getSheetName() +
                                            ", Row: " + (rowIndex + 1) + ", Column: " + (colIndex + 1));
                                    System.out.println("Original: " + cellValue);
                                    System.out.println("Replaced: " + normalizedValue);
                                }

                                newCell.setCellValue(normalizedValue); // Set the normalized value in the new cell
                            } else if (cell != null) {
                                // Copy other cell types without modification
                                copyCell(cell, newCell);
                            }
                        }
                    }
                }
            }

            // Write the new workbook to the output file
            try (FileOutputStream fileOutputStream = new FileOutputStream(outputFilePath)) {
                newWorkbook.write(fileOutputStream);
            }
        }
    }

    // Method to normalize a string by removing accents and special characters
    private static String normalizeString(String input) {
        String normalized = Normalizer.normalize(input, Normalizer.Form.NFD);
        Pattern pattern = Pattern.compile("\\p{InCombiningDiacriticalMarks}+");
        return pattern.matcher(normalized).replaceAll("").replaceAll("[^\\p{ASCII}]", "");
    }

    // Method to copy the content of one cell to another
    private static void copyCell(Cell source, Cell target) {
        switch (source.getCellType()) {
            case STRING:
                target.setCellValue(source.getStringCellValue());
                break;
            case NUMERIC:
                target.setCellValue(source.getNumericCellValue());
                break;
            case BOOLEAN:
                target.setCellValue(source.getBooleanCellValue());
                break;
            case FORMULA:
                target.setCellFormula(source.getCellFormula());
                break;
            case BLANK:
                target.setBlank();
                break;
            default:
                break;
        }
    }

    public static void main(String[] args) {
        try {
            String filePath = "D:\\microservices\\Football Roster 24-25.xlsx";
            normalizePlayerNames(filePath, "D:\\microservices\\players_normalized.xlsx");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
