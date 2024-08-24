package com.amusedgroup;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcel implements ExcelUtil {

    public static String filePath = "D:\\microservices\\Football Roster 24-25.xlsx";
    public static String tempCleanedFile = "D:\\microservices\\duplicates - ";

    @Override
    public File loadExcel(String filePath) throws IOException {
        FileInputStream fileInputStream = null;
        Workbook workbook = null;
        File tempFile = null;

        try {
            // Load the Excel file
            fileInputStream = new FileInputStream(new File(filePath));
            workbook = new XSSFWorkbook(fileInputStream);

            // Create a temporary file to return
            tempFile = File.createTempFile(tempCleanedFile, ".xlsx");

            // Write the workbook to the temporary file
            try (FileOutputStream fileOutputStream = new FileOutputStream(tempFile)) {
                workbook.write(fileOutputStream);
            }

        } finally {
            // Clean up resources
            if (workbook != null) {
                workbook.close();
            }
            if (fileInputStream != null) {
                fileInputStream.close();
            }
        }
        return tempFile;
    }

    @Override
    public void findDuplicates(File excelFile) {
        try {
            ReadExcel readExcel = new ReadExcel();
            File tempExcelFile = readExcel.loadExcel(filePath);
           // ExcelDuplicatesFinder.findDuplicatesInColumns(tempExcelFile.getPath());
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    @Override
    public void replacePLayerNameAnomalies(File excelFile) {
        try {
            ExcelPlayerNameNormalizer.normalizePlayerNames(filePath, "D:\\microservices\\players_normalized.xlsx");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
