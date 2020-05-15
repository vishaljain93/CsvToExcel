package com.project;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.logging.Level;
import java.util.logging.Logger;

import static com.project.CommonMethods.*;

public class CsvToExcelInMultipleSheets {

    private static final Logger LOGGER = Logger.getLogger(CsvToExcelInMultipleSheets.class.getName());

    private FileOutputStream fileOutputStream = null;

    private BufferedReader bufferedReader;

    public void readDataFromCsvAndWriteIntoMultipleSheets(String inputFilePath, String outputFilePath) throws IOException {
        String line;
        int sheetNumber = 1;
        int rowNum = 0;
        int count = 0;
        int workbookNumber = 1;

        bufferedReader = readFile(inputFilePath);

        XSSFWorkbook workbook = createNewWorkbook(workbookNumber);
        XSSFSheet sheet = createNewSheet(workbook, sheetNumber);
        Row currentRow;
        String[] nextLine;

        while ((line = bufferedReader.readLine()) != null) {
            currentRow = sheet.createRow(rowNum++);
            nextLine = line.split(",");
            CommonMethods.setCellData(nextLine, currentRow);

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

    private void writePartialData(XSSFWorkbook workbook, String outputFilePath) throws IOException {
        LOGGER.log(Level.INFO, "Writing 3K records");
        fileOutputStream = new FileOutputStream(new File(outputFilePath));
        workbook.write(fileOutputStream);
        fileOutputStream.close();
    }

    private void writeLastRows(int count, XSSFWorkbook workbook, String outputFilePath) throws IOException {
        LOGGER.log(Level.INFO, "Writing remaining records");
        if (count != 3000) {
            fileOutputStream = new FileOutputStream(new File(outputFilePath));
            workbook.write(fileOutputStream);
            fileOutputStream.close();
        }
        workbook.close();
        bufferedReader.close();
    }
}
