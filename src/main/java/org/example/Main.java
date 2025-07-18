package org.example;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.DecimalFormat;

public class Main {
    public static void main(String[] args) {
        // System.out.println("Hello, World!");

        // this tests if apache poi is working correctly
        try {
            Workbook workbook = new XSSFWorkbook();
            //framework for an empty excel sheet
            System.out.println("its working correctly");
            workbook.close();
        } catch (Exception e) {
            System.out.println("error: " + e.getMessage());
            //tells if apache poi isn't working
        }

        String filePath = "David-Infinity-Purchase-History - Copy.xlsx";

        // tests if the file exists
        File file = new File(filePath);
        if (file.exists()) {
            System.out.println("file found at: " + file.getAbsolutePath());
        } else {
            System.out.println("file not found. Looking for: " + file.getAbsolutePath());
            return; // stops here if file doesn't exist
        }

        //OPENING AND READING THE EXCEL FILE
        try {
            FileInputStream fis = new FileInputStream(filePath);
            XSSFWorkbook workbook = new XSSFWorkbook(fis);

            //number of sheets
            System.out.println("\n=== EXCEL FILE INFO ===");
            System.out.println("Number of sheets: " + workbook.getNumberOfSheets());

            // gets the first sheet
            Sheet sheet = workbook.getSheetAt(0);
            System.out.println("Sheet name: " + sheet.getSheetName());
            System.out.println("Number of rows: " + (sheet.getLastRowNum() + 1));

            //READS COLUMN HEADERS
            Row headerRow = sheet.getRow(0);
            System.out.println("\n=== COLUMN HEADERS ===");

            for (int i = 0; i < headerRow.getLastCellNum(); i++) {
                Cell cell = headerRow.getCell(i);
                String header = cell.getStringCellValue();
                System.out.println((i + 1) + ". " + header);
            }

            System.out.println("\n=== ALL DATA ===");

            //  formatter for better number display
            DecimalFormat df = new DecimalFormat("#.##");

            // fixed it to tprocess all the rows
            int totalRows = sheet.getLastRowNum() + 1;
            System.out.println("Processing " + totalRows + " rows total...\n");

            // print headers first
            System.out.print("Row#\t");
            for (int i = 0; i < headerRow.getLastCellNum(); i++) {
                String header = headerRow.getCell(i) != null ? headerRow.getCell(i).getStringCellValue() : "Empty";
                System.out.print(header + "\t");
            }
            System.out.println();

            // print separator
            System.out.println("------------------------------------------------------------");

            //  each data row (starting from row 1, skipping header, idk if all files are like this)
            for (int rowNum = 1; rowNum <= sheet.getLastRowNum(); rowNum++) {
                Row dataRow = sheet.getRow(rowNum);
                if (dataRow != null) {
                    System.out.print(rowNum + "\t");

                    //  each column in this row
                    for (int colNum = 0; colNum < headerRow.getLastCellNum(); colNum++) {
                        Cell cell = dataRow.getCell(colNum);
                        String value = getCellValueAsString(cell, df);
                        System.out.print(value + "\t");
                    }
                    System.out.println();
                }
            }

            //closes workbook
            workbook.close();
            fis.close();

        } catch (IOException e) {
            System.out.println("Error reading Excel file: " + e.getMessage());
        }
    }

    // helper method extracts cell values
    private static String getCellValueAsString(Cell cell, DecimalFormat df) {
        if (cell == null) {
            return "Empty";
        }

        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    // puts the date cells properly
                    return cell.getDateCellValue().toString();
                } else {
                    // numeric cells with proper formatting
                    double numericValue = cell.getNumericCellValue();
                    // checks if it's a whole number
                    if (numericValue == Math.floor(numericValue)) {
                        return String.valueOf((long) numericValue);
                    } else {
                        return df.format(numericValue);
                    }
                }
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                // eval the formula
                try {
                    FormulaEvaluator evaluator = cell.getSheet().getWorkbook().getCreationHelper().createFormulaEvaluator();
                    CellValue cellValue = evaluator.evaluate(cell);
                    switch (cellValue.getCellType()) {
                        case STRING:
                            return cellValue.getStringValue();
                        case NUMERIC:
                            double numVal = cellValue.getNumberValue();
                            if (numVal == Math.floor(numVal)) {
                                return String.valueOf((long) numVal);
                            } else {
                                return df.format(numVal);
                            }
                        case BOOLEAN:
                            return String.valueOf(cellValue.getBooleanValue());
                        default:
                            return "Formula Result: " + cellValue.formatAsString();
                    }
                } catch (Exception e) {
                    return "Formula Error: " + cell.getCellFormula();
                }
            case BLANK:
                return "Empty";
            default:
                return "Unknown Type";
        }
    }
}