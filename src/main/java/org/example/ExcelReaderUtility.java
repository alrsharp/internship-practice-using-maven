package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.ZoneId;
import java.util.*;

/**
 * Utility class for reading and analyzing Excel files.
 * Handles all Excel-related operations and data extraction.
 */
public class ExcelReaderUtility {

    /**
     * Reads an Excel file and converts it to a list of PurchaseRecord objects.
     *
     * @param filePath Path to the Excel file
     * @return List of PurchaseRecord objects
     * @throws IOException if file cannot be read
     */
    public static List<PurchaseRecord> readExcelFile(String filePath) throws IOException {
        List<PurchaseRecord> records = new ArrayList<>();

        try (FileInputStream fis = new FileInputStream(filePath);
             XSSFWorkbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0);

            // Get header row to understand column structure
            Row headerRow = sheet.getRow(0);
            if (headerRow == null) {
                throw new IOException("Excel file appears to be empty or has no header row");
            }

            // Map column names to indices for flexible reading
            Map<String, Integer> columnMap = createColumnMap(headerRow);

            // Read data rows
            for (int rowNum = 1; rowNum <= sheet.getLastRowNum(); rowNum++) {
                Row dataRow = sheet.getRow(rowNum);
                if (dataRow != null && !isRowEmpty(dataRow)) {
                    try {
                        PurchaseRecord record = createPurchaseRecord(dataRow, columnMap);
                        records.add(record);
                    } catch (Exception e) {
                        System.err.println("Error processing row " + (rowNum + 1) + ": " + e.getMessage());
                        // Continue processing other rows
                    }
                }
            }
        }

        return records;
    }

    /**
     * Creates a map of column names to their indices for flexible column reading.
     */
    private static Map<String, Integer> createColumnMap(Row headerRow) {
        Map<String, Integer> columnMap = new HashMap<>();

        for (int i = 0; i < headerRow.getLastCellNum(); i++) {
            Cell cell = headerRow.getCell(i);
            if (cell != null) {
                String columnName = cell.getStringCellValue().trim().toLowerCase();
                columnMap.put(columnName, i);
            }
        }

        return columnMap;
    }

    /**
     * Creates a PurchaseRecord from a data row using the column map.
     */
    private static PurchaseRecord createPurchaseRecord(Row dataRow, Map<String, Integer> columnMap) {
        PurchaseRecord record = new PurchaseRecord();

        // Extract data based on common column names (flexible approach)
        record.setItemName(getStringValue(dataRow, columnMap, "item", "itemname", "product", "name"));
        record.setPrice(getBigDecimalValue(dataRow, columnMap, "price", "cost", "unitprice"));
        record.setQuantity(getIntValue(dataRow, columnMap, "quantity", "qty", "amount"));
        record.setPurchaseDate(getDateValue(dataRow, columnMap, "date", "purchasedate", "orderdate"));
        record.setCategory(getStringValue(dataRow, columnMap, "category", "type", "group"));
        record.setVendor(getStringValue(dataRow, columnMap, "vendor", "supplier", "store"));
        record.setTotalCost(getBigDecimalValue(dataRow, columnMap, "total", "totalcost", "totalprice"));

        return record;
    }

    /**
     * Gets string value from row using multiple possible column names.
     */
    private static String getStringValue(Row row, Map<String, Integer> columnMap, String... possibleNames) {
        for (String name : possibleNames) {
            Integer columnIndex = columnMap.get(name);
            if (columnIndex != null) {
                Cell cell = row.getCell(columnIndex);
                if (cell != null) {
                    return getCellValueAsString(cell);
                }
            }
        }
        return "";
    }

    /**
     * Gets BigDecimal value from row using multiple possible column names.
     */
    private static BigDecimal getBigDecimalValue(Row row, Map<String, Integer> columnMap, String... possibleNames) {
        for (String name : possibleNames) {
            Integer columnIndex = columnMap.get(name);
            if (columnIndex != null) {
                Cell cell = row.getCell(columnIndex);
                if (cell != null && cell.getCellType() == CellType.NUMERIC) {
                    return BigDecimal.valueOf(cell.getNumericCellValue());
                }
            }
        }
        return BigDecimal.ZERO;
    }

    /**
     * Gets integer value from row using multiple possible column names.
     */
    private static int getIntValue(Row row, Map<String, Integer> columnMap, String... possibleNames) {
        for (String name : possibleNames) {
            Integer columnIndex = columnMap.get(name);
            if (columnIndex != null) {
                Cell cell = row.getCell(columnIndex);
                if (cell != null && cell.getCellType() == CellType.NUMERIC) {
                    return (int) cell.getNumericCellValue();
                }
            }
        }
        return 0;
    }

    /**
     * Gets LocalDate value from row using multiple possible column names.
     */
    private static LocalDate getDateValue(Row row, Map<String, Integer> columnMap, String... possibleNames) {
        for (String name : possibleNames) {
            Integer columnIndex = columnMap.get(name);
            if (columnIndex != null) {
                Cell cell = row.getCell(columnIndex);
                if (cell != null && DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue().toInstant()
                            .atZone(ZoneId.systemDefault())
                            .toLocalDate();
                }
            }
        }
        return null;
    }

    /**
     * Analyzes Excel file structure and provides column information.
     */
    public static ExcelAnalysis analyzeExcelFile(String filePath) throws IOException {
        ExcelAnalysis analysis = new ExcelAnalysis();

        try (FileInputStream fis = new FileInputStream(filePath);
             XSSFWorkbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0);
            Row headerRow = sheet.getRow(0);

            if (headerRow == null) {
                throw new IOException("No header row found");
            }

            // Analyze each column
            for (int colNum = 0; colNum < headerRow.getLastCellNum(); colNum++) {
                Cell headerCell = headerRow.getCell(colNum);
                if (headerCell != null) {
                    String columnName = headerCell.getStringCellValue();
                    ColumnInfo columnInfo = analyzeColumn(sheet, colNum, columnName);
                    analysis.addColumn(columnInfo);
                }
            }
        }

        return analysis;
    }

    /**
     * Analyzes a specific column to determine its data type and calculate statistics.
     */
    private static ColumnInfo analyzeColumn(Sheet sheet, int columnIndex, String columnName) {
        ColumnInfo info = new ColumnInfo(columnName);
        List<Double> numericValues = new ArrayList<>();
        List<String> sampleValues = new ArrayList<>();
        int totalCells = 0;
        int emptyCells = 0;

        // Analyze data in this column
        for (int rowNum = 1; rowNum <= sheet.getLastRowNum() && sampleValues.size() < 5; rowNum++) {
            Row row = sheet.getRow(rowNum);
            if (row != null) {
                Cell cell = row.getCell(columnIndex);
                totalCells++;

                if (cell == null || cell.getCellType() == CellType.BLANK) {
                    emptyCells++;
                } else {
                    String cellValue = getCellValueAsString(cell);
                    if (!cellValue.isEmpty()) {
                        sampleValues.add(cellValue);

                        // Try to parse as number for statistics
                        if (cell.getCellType() == CellType.NUMERIC) {
                            numericValues.add(cell.getNumericCellValue());
                        }
                    }
                }
            }
        }

        info.setSampleValues(sampleValues);
        info.setTotalCells(totalCells);
        info.setEmptyCells(emptyCells);

        // Determine if column is numeric and calculate statistics
        if (!numericValues.isEmpty() && numericValues.size() > totalCells * 0.8) { // 80% numeric threshold
            info.setNumeric(true);
            double sum = numericValues.stream().mapToDouble(Double::doubleValue).sum();
            double average = sum / numericValues.size();
            info.setSum(sum);
            info.setAverage(average);
        }

        return info;
    }

    /**
     * Extracts cell value as string, handling different cell types.
     */
    private static String getCellValueAsString(Cell cell) {
        if (cell == null) return "";

        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue().trim();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue().toString();
                } else {
                    double numValue = cell.getNumericCellValue();
                    if (numValue == Math.floor(numValue)) {
                        return String.valueOf((long) numValue);
                    } else {
                        return String.valueOf(numValue);
                    }
                }
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                try {
                    FormulaEvaluator evaluator = cell.getSheet().getWorkbook().getCreationHelper().createFormulaEvaluator();
                    CellValue cellValue = evaluator.evaluate(cell);
                    return cellValue.formatAsString();
                } catch (Exception e) {
                    return "Formula Error";
                }
            default:
                return "";
        }
    }

    /**
     * Checks if a row is completely empty.
     */
    private static boolean isRowEmpty(Row row) {
        for (int cellNum = row.getFirstCellNum(); cellNum < row.getLastCellNum(); cellNum++) {
            Cell cell = row.getCell(cellNum);
            if (cell != null && cell.getCellType() != CellType.BLANK) {
                String cellValue = getCellValueAsString(cell);
                if (!cellValue.isEmpty()) {
                    return false;
                }
            }
        }
        return true;
    }

    /**
     * Inner class to hold column analysis information.
     */
    public static class ColumnInfo {
        private String name;
        private boolean isNumeric;
        private double sum;
        private double average;
        private List<String> sampleValues;
        private int totalCells;
        private int emptyCells;

        public ColumnInfo(String name) {
            this.name = name;
            this.sampleValues = new ArrayList<>();
        }

        // Getters and setters
        public String getName() { return name; }
        public boolean isNumeric() { return isNumeric; }
        public void setNumeric(boolean numeric) { isNumeric = numeric; }
        public double getSum() { return sum; }
        public void setSum(double sum) { this.sum = sum; }
        public double getAverage() { return average; }
        public void setAverage(double average) { this.average = average; }
        public List<String> getSampleValues() { return sampleValues; }
        public void setSampleValues(List<String> sampleValues) { this.sampleValues = sampleValues; }
        public int getTotalCells() { return totalCells; }
        public void setTotalCells(int totalCells) { this.totalCells = totalCells; }
        public int getEmptyCells() { return emptyCells; }
        public void setEmptyCells(int emptyCells) { this.emptyCells = emptyCells; }
    }

    /**
     * Class to hold complete Excel analysis results.
     */
    public static class ExcelAnalysis {
        private List<ColumnInfo> columns;

        public ExcelAnalysis() {
            this.columns = new ArrayList<>();
        }

        public void addColumn(ColumnInfo columnInfo) {
            columns.add(columnInfo);
        }

        public List<ColumnInfo> getColumns() {
            return columns;
        }
    }
}