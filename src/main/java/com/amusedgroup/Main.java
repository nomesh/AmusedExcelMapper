package com.amusedgroup;

import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.io.PrintWriter;
import java.text.SimpleDateFormat;
import java.util.Date;

public class Main {
    public static void main(String[] args){
        SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMdd_HHmmss");
        String dateTime = sdf.format(new Date());

        try (PrintWriter out = new PrintWriter(new FileWriter("log.txt"))) {
            System.out.println("Application started");
            // Get the current directory and construct file paths
            File currentDirectory = new File(".");
            String inputFilePath = new File(currentDirectory, "Mapping/Roster_Mapping.xlsx").getAbsolutePath();
            String outputFilePath = new File(currentDirectory, "Mapping/player-normalized-" + dateTime + ".xlsx").getAbsolutePath();
            String outputFilePath2 = new File(currentDirectory, "Mapping/final-with-identified-duplicates-" + dateTime + ".xlsx").getAbsolutePath();
            //-------------------------------------------------------------------------------------------------
            System.out.println("===================================================================");
            System.out.println("Finding Special Characters of the Player Names and Replacing......");
            System.out.println("===================================================================");
            ExcelPlayerNameNormalizer.normalizePlayerNames(inputFilePath, outputFilePath);  /* Identify & Replace Special Chars */

//            Thread.sleep(2000);

            System.out.println("===================================================================");
            System.out.println("Finding Duplicate Players and Highlighting with Team Name Suffix......");
            System.out.println("picking input file in: "+ outputFilePath);
            System.out.println("===================================================================");

            ExcelDuplicatesFinder.findAndProcessDuplicates(outputFilePath, outputFilePath2); /* Identify Duplicates - Highlights */

            System.out.println("Duplicates finding is Completed......");

            System.out.println("Processing complete. Output file: " + outputFilePath);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}