package org.example;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

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
            FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();

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



            System.out.println("\n=== SAMPLE DATA FROM EACH COLUMN ===");

            // LOOKING AT THE FIRST 3 ROWS OF DATA
            for (int i = 0; i < headerRow.getLastCellNum(); i++) {
                String header = headerRow.getCell(i) != null ? headerRow.getCell(i).getStringCellValue() : "Empty";
                System.out.println("\nColumn " + (i + 1) + ": " + header);
                System.out.print("  Sample values: ");

                // reads 3 sample values from each column
                for (int rowNum = 1; rowNum <= 3; rowNum++) {
                    Row dataRow = sheet.getRow(rowNum);
                    if (dataRow != null) {
                        Cell cell = dataRow.getCell(i);
                        if (cell != null) {
                            // we need to handle different cell types
                            String value = "";
                            if (cell.getCellType() == CellType.STRING) {
                                value = cell.getStringCellValue();
                            } else if (cell.getCellType() == CellType.NUMERIC) {
                                value = String.valueOf(cell.getNumericCellValue());
                            } else if(cell.getCellType()==CellType.FORMULA){
                                CellValue evaluatedValue = evaluator.evaluate(cell);
                                if(evaluatedValue.getCellType()==CellType.NUMERIC){
                                    value = String.valueOf(evaluatedValue.getNumberValue());
                            } else if(evaluatedValue.getCellType()==CellType.STRING){
                                    value = evaluatedValue.getStringValue();
                        } else {
                            value = "Empty";
                        }
                        if (rowNum < 3) System.out.print(", ");
                    }
                }
                System.out.println();
            }
            //BTW for the sale dates, the number is how excel stores the data im p sure i think thats why it's off

            //closes workbook
            workbook.close();
            fis.close();

        } catch (IOException e) {
            System.out.println("Error reading Excel file: " + e.getMessage());
        }

    }
}
