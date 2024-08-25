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
            out.println("Application started");
            // Get the current directory and construct file paths
            File currentDirectory = new File(".");
            String inputFilePath = new File(currentDirectory, "Mapping/Football Roster 24-25.xlsx").getAbsolutePath();
            String outputFilePath = new File(currentDirectory, "Mapping/player-normalized-" + dateTime + ".xlsx").getAbsolutePath();
            String outputFilePath2 = new File(currentDirectory, "Mapping/found-duplicates-" + dateTime + ".xlsx").getAbsolutePath();
            //-------------------------------------------------------------------------------------------------
            out.println("===================================================================");
            out.println("Finding Special Characters of the Player Names and Replacing......");
            out.println("===================================================================");
            ExcelPlayerNameNormalizer.normalizePlayerNames(inputFilePath, outputFilePath);  /* Identify & Replace Special Chars */

            out.println("===================================================================");
            out.println("Finding Duplicate Players and Highlighting with Team Name Suffix......");
            out.println("picking input file in: "+ outputFilePath);
            out.println("===================================================================");
            ExcelDuplicatesFinder.findAndProcessDuplicates(outputFilePath, outputFilePath2); /* Identify Duplicates - Highlights */
            out.println("Duplicates finding is Completed......");

            out.println("Processing complete. Output file: " + outputFilePath);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}