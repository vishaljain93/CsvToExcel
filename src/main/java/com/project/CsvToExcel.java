package com.project;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.logging.Level;
import java.util.logging.Logger;

public class CsvToExcel {

    private static final Logger LOGGER = Logger.getLogger(CsvToExcel.class.getName());

    private FileOutputStream fileOutputStream = null;

    public void readDataFromCsv(HSSFWorkbook workbook, BufferedReader br, String outputFilePath) throws IOException {
        String line;
        int sheetNumber = 1;
        int rowNum = 0;
        int count = 0;

        HSSFSheet sheet = workbook.createSheet(Integer.toString(sheetNumber));
        LOGGER.log(Level.INFO, "Sheet {0} created for 50000 records", sheetNumber);
        Row currentRow;
        String[] nextLine;
        while ((line = br.readLine()) != null) {
            currentRow = sheet.createRow(rowNum++);
            int cellNum = 0;
            nextLine = line.split(",");

            for (String s : nextLine) {
                Cell cell = currentRow.createCell(cellNum++);
                cell.setCellValue(s);
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

    private HSSFSheet createNewSheet(HSSFWorkbook workbook, int sheetNumber) {
        HSSFSheet sheet = workbook.createSheet(Integer.toString(sheetNumber));
        LOGGER.log(Level.INFO, "Sheet {0} created for next 50000 records", sheetNumber);
        return sheet;
    }


    private void writePartialData(HSSFWorkbook workbook, String outputFilePath) throws IOException {
        fileOutputStream = new FileOutputStream(new File(outputFilePath));
        workbook.write(fileOutputStream);
        fileOutputStream.close();
    }

    private void writeLastRows(int count, HSSFWorkbook workbook, String outputFilePath) throws IOException {
        if (count != 1500) {
            fileOutputStream = new FileOutputStream(new File(outputFilePath));
            workbook.write(fileOutputStream);
            fileOutputStream.close();
        }
    }
}
