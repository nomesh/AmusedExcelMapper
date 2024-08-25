package com.amusedgroup;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

public class ExcelDecorator {
    private static final int PLAYER_NAME_COLUMN_INDEX = 1; // Player Name column index (0-based)
    private static final int TEAM_NAME_COLUMN_INDEX = 3;   // Team Name column index (0-based)
    private static final int COMPETITION_COLUMN_INDEX = 4; // Competition column index (0-based)

    public static void copySelectedColumns(Row headerRow, Row newHeaderRow) {
        // Copy Player Name column header
        Cell playerNameHeaderCell = headerRow.getCell(PLAYER_NAME_COLUMN_INDEX, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
        if (playerNameHeaderCell != null) {
            newHeaderRow.createCell(PLAYER_NAME_COLUMN_INDEX).setCellValue(playerNameHeaderCell.getStringCellValue());
        }

        // Copy Team Name column header
        Cell teamNameHeaderCell = headerRow.getCell(TEAM_NAME_COLUMN_INDEX, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
        if (teamNameHeaderCell != null) {
            newHeaderRow.createCell(TEAM_NAME_COLUMN_INDEX).setCellValue(teamNameHeaderCell.getStringCellValue());
        }

        // Copy Competition column header
        Cell competitionHeaderCell = headerRow.getCell(COMPETITION_COLUMN_INDEX, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
        if (competitionHeaderCell != null) {
            newHeaderRow.createCell(COMPETITION_COLUMN_INDEX).setCellValue(competitionHeaderCell.getStringCellValue());
        }
    }

    // Remove empty columns from the sheet
    public static void removeEmptyColumns(Sheet sheet) {
        int maxColumnNum = 0;

        // Find the maximum column number used in the sheet
        for (Row row : sheet) {
            if (row.getLastCellNum() > maxColumnNum) {
                maxColumnNum = row.getLastCellNum();
            }
        }

        // Loop through each column index starting from the last column
        for (int colIndex = maxColumnNum - 1; colIndex >= 0; colIndex--) {
            boolean isEmptyColumn = true;

            // Check each row for the current column index
            for (Row row : sheet) {
                Cell cell = row.getCell(colIndex, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                if (cell != null && cell.getCellType() != CellType.BLANK) {
                    isEmptyColumn = false; // The column has data, so it's not empty
                    break;
                }
            }

            // If the column is empty, remove it
            if (isEmptyColumn) {
                for (Row row : sheet) {
                    Cell cellToRemove = row.getCell(colIndex);
                    if (cellToRemove != null) {
                        row.removeCell(cellToRemove);
                    }
                }

                // Shift columns to the left after removing an empty column
                shiftColumnsLeft(sheet, colIndex);
            }
        }
    }

    // Shift columns to the left after removing an empty column
    public static void shiftColumnsLeft(Sheet sheet, int colIndex) {
        for (Row row : sheet) {
            for (int colToShift = colIndex + 1; colToShift <= row.getLastCellNum(); colToShift++) {
                Cell oldCell = row.getCell(colToShift);
                if (oldCell != null) {
                    Cell newCell = row.createCell(colToShift - 1, oldCell.getCellType());
                    cloneCell(oldCell, newCell);
                    row.removeCell(oldCell);
                }
            }
        }
    }

    // Helper method to clone a cell (copies value and style)
    public static void cloneCell(Cell oldCell, Cell newCell) {
        switch (oldCell.getCellType()) {
            case STRING:
                newCell.setCellValue(oldCell.getStringCellValue());
                break;
            case NUMERIC:
                newCell.setCellValue(oldCell.getNumericCellValue());
                break;
            case BOOLEAN:
                newCell.setCellValue(oldCell.getBooleanCellValue());
                break;
            case FORMULA:
                newCell.setCellFormula(oldCell.getCellFormula());
                break;
            case BLANK:
                newCell.setBlank();
                break;
            case ERROR:
                newCell.setCellErrorValue(oldCell.getErrorCellValue());
                break;
            default:
                break;
        }
        newCell.setCellStyle(oldCell.getCellStyle());
    }
}
