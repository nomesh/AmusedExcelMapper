package com.amusedgroup;

import java.io.File;
import java.io.IOException;

public class Main {
    public static void main(String[] args) throws IOException {

        ReadExcel readExcel = new ReadExcel();
        String original_filePath = "D:\\microservices\\Football Roster 24-25.xlsx";
        File cleanedExcelFile = readExcel.loadExcel(original_filePath);

        //find duplicates
        readExcel.findDuplicates(new File(original_filePath));
        //readExcel.replacePLayerNameAnomalies(new File(original_filePath));
        System.out.println("File Opened Successfully ="+ cleanedExcelFile);



    }
}