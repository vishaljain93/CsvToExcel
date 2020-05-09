package com.project;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.logging.Level;
import java.util.logging.Logger;

public class CsvToExcel {

    private static final Logger LOGGER = Logger.getLogger(CsvToExcel.class.getName());

    private FileOutputStream fileOutputStream = null;

    public void readDataFromCsv(XSSFWorkbook workbook, BufferedReader br, String outputFilePath) throws IOException {
        String line;
        int sheetNumber = 1;
        int rowNum = 0;
        int count = 0;

        XSSFSheet sheet = workbook.createSheet(Integer.toString(sheetNumber));
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
            if (count == 3000) {
                count = 0;
                writePartialData(workbook, outputFilePath);
            }
            if (rowNum == 50000) {
                sheet = createNewSheet(workbook, ++sheetNumber);
                rowNum = 0;
            }
            count++;
        }
        writeLastRows(count, workbook, outputFilePath);
    }

    private XSSFSheet createNewSheet(XSSFWorkbook workbook, int sheetNumber) {
        XSSFSheet sheet = workbook.createSheet(Integer.toString(sheetNumber));
        LOGGER.log(Level.INFO, "Sheet {0} created for next 50000 records", sheetNumber);
        return sheet;
    }


    private void writePartialData(XSSFWorkbook workbook, String outputFilePath) throws IOException {
        fileOutputStream = new FileOutputStream(new File(outputFilePath));
        workbook.write(fileOutputStream);
        fileOutputStream.close();
    }

    private void writeLastRows(int count, XSSFWorkbook workbook, String outputFilePath) throws IOException {
        if (count != 3000) {
            fileOutputStream = new FileOutputStream(new File(outputFilePath));
            workbook.write(fileOutputStream);
            fileOutputStream.close();
        }
    }
}
