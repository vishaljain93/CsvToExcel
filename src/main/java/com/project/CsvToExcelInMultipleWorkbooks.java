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

public class CsvToExcelInMultipleWorkbooks {

    private static final Logger LOGGER = Logger.getLogger(CsvToExcelInMultipleWorkbooks.class.getName());

    private FileOutputStream fileOutputStream = null;

    private BufferedReader bufferedReader;

    public void readDataFromCsvAndWriteIntoMultipleWorkbooks(String inputFilePath, String outputFilePath) throws IOException {
        String line;
        int rowNum = 1;
        int count = 1;
        int workbookNumber = 1;

        bufferedReader = readFile(inputFilePath);

        XSSFWorkbook workbook = createNewWorkbook(workbookNumber);
        XSSFSheet sheet = workbook.createSheet("Sheet");

        Row currentRow;
        String[] nextLine;
        String header = bufferedReader.readLine();

        setHeader(header, workbook, outputFilePath, workbookNumber, sheet);

        while ((line = bufferedReader.readLine()) != null) {

            currentRow = sheet.createRow(rowNum);
            nextLine = line.split(",");
            setCellData(nextLine, currentRow);

            if (count == 3000) {
                count = 0;
                writePartialDataInWorkbook(workbook, outputFilePath, workbookNumber);
            }

            if (rowNum == 50000) {
                workbook.close();
                workbookNumber++;
                workbook = createNewWorkbook(workbookNumber);
                sheet = workbook.createSheet("Sheet");
                setHeader(header, workbook, outputFilePath, workbookNumber, sheet);
                rowNum = 0;
            }
            count++;
            rowNum++;
        }
        writeLastRowsInWorkbook(count, workbook, outputFilePath, workbookNumber);
    }

    private void setHeader(String header, XSSFWorkbook workbook, String outputFilePath, int workbookNumber, XSSFSheet sheet) throws IOException {
        LOGGER.log(Level.INFO, "Setting header");
        Row currentRow = sheet.createRow(0);
        String[] nextLine = header.split(",");
        setCellData(nextLine, currentRow);
        writePartialDataInWorkbook(workbook, outputFilePath, workbookNumber);
    }

    private void writePartialDataInWorkbook(XSSFWorkbook workbook, String outputFilePath, int workbookNumber) throws IOException {
        LOGGER.log(Level.INFO, "Writing 3K records");
        String[] outputFileName = outputFilePath.split("\\.");
        outputFilePath = outputFileName[0].concat(Integer.toString(workbookNumber)).concat(".").concat(outputFileName[1]);
        fileOutputStream = new FileOutputStream(new File(outputFilePath));
        workbook.write(fileOutputStream);
        fileOutputStream.close();
    }

    private void writeLastRowsInWorkbook(int count, XSSFWorkbook workbook, String outputFilePath, int workbookNumber) throws IOException {
        LOGGER.log(Level.INFO, "Writing remaining records");
        if (count != 3000) {
            String[] outputFileName = outputFilePath.split("\\.");
            outputFilePath = outputFileName[0].concat(Integer.toString(workbookNumber)).concat(".").concat(outputFileName[1]);
            fileOutputStream = new FileOutputStream(new File(outputFilePath));
            workbook.write(fileOutputStream);
            fileOutputStream.close();
        }
        workbook.close();
        bufferedReader.close();
    }
}
