package org.example;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Main {
    public static void main(String[] args) {
        System.out.println("Hello, World!");

        // this tests if apach poi is working correctly
        try {
            Workbook workbook = new XSSFWorkbook();
            System.out.println("its working correctly");
            workbook.close();
        } catch (Exception e) {
            System.out.println("error: " + e.getMessage());
        }
    }
}