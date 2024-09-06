package com.amusedgroup;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.Normalizer;
import java.util.HashMap;
import java.util.Map;
import java.util.regex.Pattern;

public class ExcelPlayerNameNormalizer {

    public static void normalizePlayerNames(String inputFilePath, String outputFilePath) throws IOException {
        // Open the input Excel file
        FileInputStream fis = new FileInputStream(inputFilePath);
        Workbook workbook = new XSSFWorkbook(fis);

        // Iterate over all sheets in the workbook
        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            Sheet sheet = workbook.getSheetAt(i);
            System.out.println("=================");
            System.out.println("Tab: " + sheet.getSheetName());
            System.out.println();

            boolean foundSpecialCharacters = false;

            for (Row row : sheet) {
                for (Cell cell : row) {
                    if (cell.getCellType() == CellType.STRING) {
                        String originalValue = cell.getStringCellValue();
                        String normalizedValue = normalizePlayerName(originalValue);

                        if (!originalValue.equals(normalizedValue)) {
                            if (!foundSpecialCharacters) {
                                System.out.println("Special characters found in sheet:");
                                foundSpecialCharacters = true;
                            }
                            System.out.printf("  Row %d, Column %d: '%s' -> '%s'%n",
                                    row.getRowNum() + 1, cell.getColumnIndex() + 1, originalValue, normalizedValue);
                            cell.setCellValue(normalizedValue);
                        }
                    }
                }
            }

            if (!foundSpecialCharacters) {
                System.out.println("No special characters found in sheet.");
            }

            System.out.println(); // Add an empty line for readability between sheets
        }

        // Close the input file stream
        fis.close();

        // Write the normalized content to a new Excel file
        FileOutputStream fos = new FileOutputStream(outputFilePath);
        workbook.write(fos);

        System.out.println("Output File location: " + outputFilePath);

        // Close the output file stream and the workbook
        fos.close();
        workbook.close();
    }

    public static String normalizePlayerName(String playerName) {
        // Normalize the string to remove diacritical marks (accents)
        return Normalizer.normalize(playerName, Normalizer.Form.NFD)
                .replaceAll("\\p{InCombiningDiacriticalMarks}+", "")
                .replaceAll("[øØ]", "o")
                .replaceAll("[óÓ]", "o")
                .replaceAll("[æÆ]", "ae")
                .replaceAll("[œŒ]", "oe")
                .replaceAll("[ß]", "ss")  // ß has no direct uppercase equivalent
                .replaceAll("[ðÐ]", "d")
                .replaceAll("[þÞ]", "th")
                .replaceAll("[łŁ]", "l")
                .replaceAll("[đĐ]", "d")
                .replaceAll("[ŋŊ]", "ng")
                .replaceAll("[ħĦ]", "h")
                .replaceAll("[ıI]", "i")  // ı and I map to lowercase 'i'
                .replaceAll("[ĳĲ]", "ij")
                .replaceAll("[ſ]", "s")   // ſ has no direct uppercase equivalent
                .replaceAll("[ėĖ]", "e")
                .replaceAll("[éÉ]", "e")
                .replaceAll("[ãÃ]", "a")
                .replaceAll("[áÁ]", "a");
    }

//    public static void main(String[] args) {
//        try {
//            String filePath = "D:\\microservices\\Football Roster 24-25.xlsx";
//            ExcelPlayerNameNormalizer.normalizePlayerNames(filePath, "D:\\microservices\\clean-excel\\Mapping\\players_normalized.xlsx");
//        } catch (IOException e) {
//            e.printStackTrace();
//        }
//    }
}