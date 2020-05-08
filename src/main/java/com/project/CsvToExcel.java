package com.project;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import java.io.*;
import java.util.logging.Logger;

public class CsvToExcel {

    private static final Logger log = Logger.getLogger(CsvToExcel.class.getName());

    private static FileOutputStream fileOutputStream = null;

    private static void readDataFromCsv(HSSFWorkbook workbook, BufferedReader br, String outputFilePath) throws IOException {
        String line;
        int sheetNumber = 1;
        int rowNum = 0;
        int count = 0;

        HSSFSheet sheet = workbook.createSheet(Integer.toString(sheetNumber));
        log.info("Sheet " + sheetNumber + " created for 50000 records");
        Row currentRow;
        String[] nextLine;
        while ((line = br.readLine()) != null) {
            currentRow = sheet.createRow(rowNum++);
            int cellNum = 0;
            nextLine = line.split(",");

            for (int i = 0; i < nextLine.length; i++) {
                Cell cell = currentRow.createCell(cellNum++);
                cell.setCellValue(nextLine[i]);
            }
            if (count == 1500) {
                count = 0;
                writePartialData(workbook, outputFilePath);
            }
            if (rowNum == 3000) {
                sheet = createNewSheet(workbook, ++sheetNumber);
                rowNum = 0;
            }
            count++;
        }
        writeLastRows(count, workbook, outputFilePath);
    }

    private static HSSFSheet createNewSheet(HSSFWorkbook workbook, int sheetNumber) {
        HSSFSheet sheet = workbook.createSheet(Integer.toString(sheetNumber));
        log.info("Sheet " + sheetNumber + " created for 50000 records");
        return sheet;
    }


    private static void writePartialData(HSSFWorkbook workbook, String outputFilePath) throws IOException {
        fileOutputStream = new FileOutputStream(outputFilePath);
        workbook.write(fileOutputStream);
        fileOutputStream.close();
    }

    private static void writeLastRows(int count, HSSFWorkbook workbook, String outputFilePath) throws IOException {
        if (count != 1500) {
            fileOutputStream = new FileOutputStream(new File(outputFilePath));
            workbook.write(fileOutputStream);
            fileOutputStream.close();
        }
    }

    public static void main(String[] args) {
        String inputFilePath = args[0];
        String outputFilePath = args[1];
        Long startTime = System.currentTimeMillis();
        log.info("Starting converting csv to excel");
        log.info("Each shee will have 50000 records");
        log.info("Records will be written to excel file in a bunch of 1500");

        try (HSSFWorkbook workbook = new HSSFWorkbook();
             BufferedReader br = new BufferedReader(new FileReader(inputFilePath))) {

            readDataFromCsv(workbook, br, outputFilePath);
            log.info("Successfully converted csv file to excel");

            Long endTime = System.currentTimeMillis();
            log.info("Total execution time: " + (endTime - startTime) + " milliseconds");
        } catch (IOException e) {
            log.info("Exception while converting csv to excel");
        }
    }
}
