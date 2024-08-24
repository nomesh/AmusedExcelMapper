package com.amusedgroup;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.*;

public class ExcelDuplicatesFinder {
    static String orig_inputFilePath = ".\\Mapping\\Football Roster 24-25.xlsx";
    static String outputFilePath = ".\\Mapping\\cleaned"+new Date().getTime()+".xlsx";

    // Define the index of the Player Name column and other columns
    private static final int PLAYER_NAME_COLUMN_INDEX = 1; // Player Name column index (0-based)
    private static final int TEAM_NAME_COLUMN_INDEX = 3;   // Team Name column index (0-based)
    private static final int COMPETITION_COLUMN_INDEX = 4; // Competition column index (0-based)

    public static void findAndProcessDuplicates(String inputFilePath, String outputFilePath) throws IOException {
        try (FileInputStream fileInputStream = new FileInputStream(orig_inputFilePath);
             Workbook workbook = new XSSFWorkbook(fileInputStream);
             Workbook newWorkbook = new XSSFWorkbook()) {

            // Create a cell style with an orange background
            CellStyle orangeStyle = newWorkbook.createCellStyle();
            orangeStyle.setFillForegroundColor(IndexedColors.ORANGE.getIndex());
            orangeStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            // Iterate through each sheet in the input workbook
            for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++) {
                Sheet sheet = workbook.getSheetAt(sheetIndex);
                Sheet newSheet = newWorkbook.createSheet(sheet.getSheetName()); // Create a new sheet in the new workbook

                System.out.println("Processing sheet: " + sheet.getSheetName());

                // Map to store seen player names with their row numbers
                Map<String, Integer> seenPlayerNames = new HashMap<>();

                // Copy headers for relevant columns only
                Row headerRow = sheet.getRow(0);
                if (headerRow != null) {
                    Row newHeaderRow = newSheet.createRow(0);
                    copySelectedColumns(headerRow, newHeaderRow);
                }

                // Iterate through each row in the sheet
                for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                    Row row = sheet.getRow(rowIndex);
                    Row newRow = newSheet.createRow(rowIndex); // Create a new row in the new sheet

                    if (row != null) {
                        // Handle relevant columns only
                        Cell playerNameCell = row.getCell(PLAYER_NAME_COLUMN_INDEX);
                        Cell teamNameCell = row.getCell(TEAM_NAME_COLUMN_INDEX);
                        Cell competitionCell = row.getCell(COMPETITION_COLUMN_INDEX);

                        Cell newPlayerNameCell = newRow.createCell(PLAYER_NAME_COLUMN_INDEX);
                        Cell newTeamNameCell = newRow.createCell(TEAM_NAME_COLUMN_INDEX);
                        Cell newCompetitionCell = newRow.createCell(COMPETITION_COLUMN_INDEX);

                        if (playerNameCell != null && playerNameCell.getCellType() == CellType.STRING) {
                            String playerName = playerNameCell.getStringCellValue();
                            if (seenPlayerNames.containsKey(playerName)) {
                                int duplicateRowIndex = seenPlayerNames.get(playerName);
                                Row duplicateRow = sheet.getRow(duplicateRowIndex);

                                // Check if the team name cell exists
                                if (teamNameCell != null && teamNameCell.getCellType() == CellType.STRING) {
                                    String teamName = teamNameCell.getStringCellValue();

                                    // Prepend the team name to the player name
                                    String newPlayerName = playerName + " - " + teamName;
                                    System.out.println("duplicate found: "+playerName + " - " + teamName);
                                    // Set the modified value in the new sheet and apply orange style
                                    newPlayerNameCell.setCellValue(newPlayerName);
                                    newPlayerNameCell.setCellStyle(orangeStyle);
                                } else {
                                    // If no team name is found, keep the original value but apply orange style
                                    newPlayerNameCell.setCellValue(playerName);
                                    newPlayerNameCell.setCellStyle(orangeStyle);
                                }
                            } else {
                                // Store the player name and row index in the map
                                seenPlayerNames.put(playerName, rowIndex);
                                newPlayerNameCell.setCellValue(playerName);
                            }
                        } else {
                            // If the player name cell is null, set the cell value as blank
                            newPlayerNameCell.setBlank();
                        }

                        // Copy the team name and competition columns as is
                        if (teamNameCell != null) {
                            newTeamNameCell.setCellValue(teamNameCell.toString());
                        }
                        if (competitionCell != null) {
                            newCompetitionCell.setCellValue(competitionCell.toString());
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

    // Method to copy only selected columns from one row to another
    private static void copySelectedColumns(Row sourceRow, Row targetRow) {
        for (int colIndex = 0; colIndex <= COMPETITION_COLUMN_INDEX; colIndex++) {
            Cell sourceCell = sourceRow.getCell(colIndex);
            Cell targetCell = targetRow.createCell(colIndex);
            if (sourceCell != null) {
                switch (sourceCell.getCellType()) {
                    case STRING:
                        targetCell.setCellValue(sourceCell.getStringCellValue());
                        break;
                    case NUMERIC:
                        targetCell.setCellValue(sourceCell.getNumericCellValue());
                        break;
                    case BOOLEAN:
                        targetCell.setCellValue(sourceCell.getBooleanCellValue());
                        break;
                    case FORMULA:
                        targetCell.setCellFormula(sourceCell.getCellFormula());
                        break;
                    case BLANK:
                        targetCell.setBlank();
                        break;
                    default:
                        break;
                }
            }
        }
    }


    public static void main(String[] args){
        SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMdd_HHmmss");
        String dateTime = sdf.format(new Date());

        try (PrintWriter out = new PrintWriter(new FileWriter("log.txt"))) {
            out.println("Application started");
            // Get the current directory and construct file paths
            File currentDirectory = new File(".");
            String inputFilePath = new File(currentDirectory, "Mapping/Football Roster 24-25.xlsx").getAbsolutePath();
            String outputFilePath = new File(currentDirectory, "Mapping/found-duplicates-" + dateTime + ".xlsx").getAbsolutePath();

            findAndProcessDuplicates(inputFilePath, outputFilePath);
            out.println("Processing complete. Output file: " + outputFilePath);
            System.out.println("Processing complete. Output file: " + outputFilePath);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}