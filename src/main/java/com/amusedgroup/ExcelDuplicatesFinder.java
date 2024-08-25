package com.amusedgroup;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.awt.Color;
import java.io.*;
import java.text.SimpleDateFormat;
import java.util.*;

public class ExcelDuplicatesFinder {
    static String orig_inputFilePath = "Mapping\\Football Roster 24-25.xlsx";

    // Define the index of the Player Name column and other columns
    private static final int PLAYER_NAME_COLUMN_INDEX = 1; // Player Name column index (0-based)
    private static final int TEAM_NAME_COLUMN_INDEX = 3;   // Team Name column index (0-based)
    private static final int COMPETITION_COLUMN_INDEX = 4; // Competition column index (0-based)

    public static void findAndProcessDuplicates(String inputFilePath, String outputFilePath) throws IOException {


        try (FileInputStream fileInputStream = new FileInputStream(inputFilePath);
             Workbook workbook = new XSSFWorkbook(fileInputStream);
             Workbook newWorkbook = new XSSFWorkbook()) {

            // Create a cell style with an orange background for duplicates
            CellStyle orangeStyle = newWorkbook.createCellStyle();
            orangeStyle.setFillForegroundColor(IndexedColors.ORANGE.getIndex());
            orangeStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            // Create a cell style for header with light blue background
            CellStyle headerStyle = newWorkbook.createCellStyle();
            headerStyle.setFillForegroundColor(IndexedColors.LIGHT_BLUE.getIndex());
            headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            // Create a general cell style for all cells with the font configuration
            CellStyle generalStyle = newWorkbook.createCellStyle();

            // Create and configure the font
            Font commonFont = newWorkbook.createFont();
            commonFont.setFontName("Aptos Narrow");
            commonFont.setFontHeightInPoints((short) 12);
            generalStyle.setFont(commonFont);

            // Apply the font to the header style as well
            headerStyle.setFont(commonFont);

            // Create a light green color for the tab
            XSSFColor lightGreenColor = new XSSFColor(new Color(144, 238, 144), null);

            // Iterate through each sheet in the input workbook
            for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++) {
                Sheet sheet = workbook.getSheetAt(sheetIndex);
                XSSFSheet newSheet = (XSSFSheet) newWorkbook.createSheet(sheet.getSheetName()); // Create a new sheet in the new workbook

                System.out.println("\n\n\nProcessing sheet: " + sheet.getSheetName());
                System.out.println("==================================================");

                // Highlight the tab name
                newSheet.setTabColor(lightGreenColor);

                // Map to store seen player names with their row numbers
                Map<String, Integer> seenPlayerNames = new HashMap<>();

                // Copy headers for relevant columns only and apply header style
                Row headerRow = sheet.getRow(0);
                if (headerRow != null) {
                    Row newHeaderRow = newSheet.createRow(0);
                    ExcelDecorator.copySelectedColumns(headerRow, newHeaderRow);
                    // Apply header style to the new header row
                    for (int i = 0; i < newHeaderRow.getLastCellNum(); i++) {
                        Cell headerCell = newHeaderRow.getCell(i);
                        if (headerCell != null) {
                            headerCell.setCellStyle(headerStyle);
                        }
                    }
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
                                    System.out.println("Duplicate found: " + playerName + " - " + teamName);
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

                // Remove empty columns after processing
                ExcelDecorator.removeEmptyColumns(newSheet);
            }

            // Write the new workbook to the output file
            try (FileOutputStream fileOutputStream = new FileOutputStream(outputFilePath)) {
                newWorkbook.write(fileOutputStream);
            }
        }
    }

    // Copy only selected columns (Player Name, Team Name, Competition) from header


    public static void main(String[] args){
        SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMdd_HHmmss");
        String dateTime = sdf.format(new Date());

        try (PrintWriter out = new PrintWriter(new FileWriter("log.txt"))) {
            out.println("Application started");
            // Get the current directory and construct file paths
            File currentDirectory = new File(".");
            String inputFilePath = new File(currentDirectory, "Mapping/Football Roster 24-25.xlsx").getAbsolutePath();
            String outputFilePath = new File(currentDirectory, "Mapping/found-duplicates-" + dateTime + ".xlsx").getAbsolutePath();

            //-------------------------------------------------------------------------------------------------

               // ExcelPlayerNameNormalizer.normalizePlayerNames(inputFilePath, outputFilePath);  /* Identify & Replace Special Chars */
                findAndProcessDuplicates(inputFilePath, outputFilePath); /* Identify Duplicates - Highlights */

            //-------------------------------------------------------------------------------------------------
            out.println("Processing complete. Output file: " + outputFilePath);
            System.out.println("Processing complete. Output file: " + outputFilePath);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}