package com.project;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.BufferedReader;
import java.io.FileReader;
import java.io.IOException;
import java.util.logging.Level;
import java.util.logging.Logger;

public class Main {

    private static final Logger LOGGER = Logger.getLogger(Main.class.getName());

    public static void main(String[] args) {
        String inputFilePath = args[0];
        String outputFilePath = args[1];

        Long startTime = System.currentTimeMillis();
        LOGGER.info("Starting converting csv to excel");
        LOGGER.info("Each sheet will have 50000 records");
        LOGGER.info("Records will be written to excel file in a bunch of 1500");

        CsvToExcel csvToExcel = new CsvToExcel();

        try (HSSFWorkbook workbook = new HSSFWorkbook();
             BufferedReader br = new BufferedReader(new FileReader(inputFilePath))) {

            csvToExcel.readDataFromCsv(workbook, br, outputFilePath);
            LOGGER.info("Successfully converted csv file to excel");

            Long endTime = System.currentTimeMillis();
            LOGGER.log(Level.INFO, "Total execution time: {0} milliseconds", endTime - startTime);
        } catch (IOException e) {
            LOGGER.info("Exception while converting csv to excel");
        }
    }
}
