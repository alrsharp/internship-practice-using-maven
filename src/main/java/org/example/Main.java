package org.example;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

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
            //tells if apache poi isn't working correctly
        }
    }
}