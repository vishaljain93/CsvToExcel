package com.project;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.logging.Level;
import java.util.logging.Logger;

public class CommonMethods {

    private static final Logger LOGGER = Logger.getLogger(CommonMethods.class.getName());

    public static XSSFSheet createNewSheet(XSSFWorkbook workbook, int sheetNumber) {
        LOGGER.log(Level.INFO, "Sheet {0} created for next 50000 records", sheetNumber);
        return workbook.createSheet(Integer.toString(sheetNumber));
    }

    public static XSSFWorkbook createNewWorkbook(int workbookNumber) {
        LOGGER.log(Level.INFO, "Workbook {0} created for 50000 records", workbookNumber);
        return new XSSFWorkbook();
    }

    public static void setCellData(String[] nextLine, Row currentRow) {
        int cellNum = 0;
        for (String s : nextLine) {
            Cell cell = currentRow.createCell(cellNum++);
            cell.setCellValue(s);
        }
    }

    public static BufferedReader readFile(String inputFilePath) throws IOException {
        return new BufferedReader(new FileReader(inputFilePath));
    }
}
