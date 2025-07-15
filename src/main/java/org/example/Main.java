package org.example;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;

public class Main {
    public static void main(String[] args) {
        System.out.println("Hello, World!");

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

    }
}