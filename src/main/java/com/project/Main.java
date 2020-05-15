package com.project;

import java.io.IOException;
import java.util.logging.Level;
import java.util.logging.Logger;

public class Main {

    private static final Logger LOGGER = Logger.getLogger(Main.class.getName());

    public static void main(String[] args) throws IOException {
        String inputFilePath = args[0];
        String outputFilePath = args[1];

        Long startTime = System.currentTimeMillis();
        LOGGER.info("Starting converting csv to excel");

        // CsvToExcelInMultipleSheets csvToExcelInMultipleSheets = new CsvToExcelInMultipleSheets();
        // csvToExcelInMultipleSheets.readDataFromCsvAndWriteIntoMultipleSheets(inputFilePath, outputFilePath);

        CsvToExcelInMultipleWorkbooks csvToExcelInMultipleWorkbooks = new CsvToExcelInMultipleWorkbooks();
        csvToExcelInMultipleWorkbooks.readDataFromCsvAndWriteIntoMultipleWorkbooks(inputFilePath, outputFilePath);
        LOGGER.info("Successfully converted csv file to excel");

        Long endTime = System.currentTimeMillis();
        LOGGER.log(Level.INFO, "Total execution time: {0} milliseconds", endTime - startTime);
    }
}
