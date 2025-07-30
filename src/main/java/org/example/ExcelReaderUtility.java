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
 * Fixed utility class for reading and analyzing Excel files.
 * Handles all Excel-related operations and data extraction.
 * NOW SUPPORTS FORMULA EVALUATION!
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

            // Create formula evaluator for handling formulas
            FormulaEvaluator formulaEvaluator = workbook.getCreationHelper().createFormulaEvaluator();

            Sheet sheet = workbook.getSheetAt(0);

            // Get header row to understand column structure
            Row headerRow = sheet.getRow(0);
            if (headerRow == null) {
                throw new IOException("Excel file appears to be empty or has no header row");
            }

            // Map column names to indices for flexible reading
            Map<String, Integer> columnMap = createColumnMap(headerRow);

            // Debug: Print the column map
            System.out.println("Column mapping:");
            for (Map.Entry<String, Integer> entry : columnMap.entrySet()) {
                System.out.println("  '" + entry.getKey() + "' -> column " + entry.getValue());
            }

            // Find the first non-empty data row and last data row
            int firstDataRow = findFirstDataRow(sheet);
            int lastDataRow = findLastDataRow(sheet);

            System.out.println("Data rows: " + firstDataRow + " to " + lastDataRow + " (total sheet rows: " + sheet.getLastRowNum() + ")");

            // Read data rows (excluding empty and summary rows)
            for (int rowNum = firstDataRow; rowNum <= lastDataRow; rowNum++) {
                Row dataRow = sheet.getRow(rowNum);
                if (dataRow != null && !isRowEmpty(dataRow)) {
                    try {
                        PurchaseRecord record = createPurchaseRecord(dataRow, columnMap, formulaEvaluator);
                        // Only add records that have essential data
                        if (record.getItemName() != null && !record.getItemName().trim().isEmpty()) {
                            records.add(record);
                        }
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
     * Finds the first row that contains actual data (not empty).
     */
    private static int findFirstDataRow(Sheet sheet) {
        // Start from row 1 (after header) and find first non-empty row
        for (int rowNum = 1; rowNum <= sheet.getLastRowNum(); rowNum++) {
            Row row = sheet.getRow(rowNum);
            if (row != null && !isRowEmpty(row)) {
                // Check if this row has meaningful data (not just random values)
                if (hasValidData(row)) {
                    return rowNum;
                }
            }
        }
        return 1; // Default to row 1 if no data found
    }

    /**
     * Checks if a row has valid purchase data.
     */
    private static boolean hasValidData(Row row) {
        // A valid data row should have at least a date and some text/numeric data
        Cell firstCell = row.getCell(0); // Assuming first column is date
        if (firstCell != null && (DateUtil.isCellDateFormatted(firstCell) || firstCell.getCellType() == CellType.NUMERIC)) {
            // Check if there's some text data in the row (product name, customer, etc.)
            for (int i = 1; i < Math.min(8, row.getLastCellNum()); i++) {
                Cell cell = row.getCell(i);
                if (cell != null && cell.getCellType() == CellType.STRING) {
                    String value = cell.getStringCellValue().trim();
                    if (!value.isEmpty() && value.length() > 2) {
                        return true;
                    }
                }
            }
        }
        return false;
    }

    /**
     * Finds the last row that contains actual data (not summary rows).
     */
    private static int findLastDataRow(Sheet sheet) {
        int totalRows = sheet.getLastRowNum();

        System.out.println("Total rows in sheet: " + (totalRows + 1) + " (0-indexed: " + totalRows + ")");

        // Look backwards from the end to find the last row with meaningful data
        for (int rowNum = totalRows; rowNum >= 1; rowNum--) {
            Row row = sheet.getRow(rowNum);
            if (row != null && !isRowEmpty(row) && hasValidData(row)) {
                System.out.println("Last data row found at: " + (rowNum + 1));
                return rowNum;
            }
        }

        return totalRows; // Default to all rows if unsure
    }

    /**
     * Creates a map of column names to their indices for flexible column reading.
     */
    private static Map<String, Integer> createColumnMap(Row headerRow) {
        Map<String, Integer> columnMap = new HashMap<>();

        for (int i = 0; i < headerRow.getLastCellNum(); i++) {
            Cell cell = headerRow.getCell(i);
            if (cell != null && cell.getCellType() == CellType.STRING) {
                String columnName = cell.getStringCellValue().trim().toLowerCase();
                columnMap.put(columnName, i);
                System.out.println("Found column: '" + columnName + "' at index " + i);
            }
        }

        return columnMap;
    }

    /**
     * Creates a PurchaseRecord from a data row using the column map.
     * Updated to match the actual Excel column names from the file.
     * NOW INCLUDES FORMULA EVALUATOR!
     */
    private static PurchaseRecord createPurchaseRecord(Row dataRow, Map<String, Integer> columnMap, FormulaEvaluator formulaEvaluator) {
        PurchaseRecord record = new PurchaseRecord();

        // Extract data based on actual Excel column names (case-insensitive)
        record.setItemName(getStringValue(dataRow, columnMap, formulaEvaluator, "procuct name", "product name", "item", "itemname", "name"));
        record.setPrice(getBigDecimalValue(dataRow, columnMap, formulaEvaluator, "unit price", "price", "cost", "unitprice"));
        record.setQuantity(getIntValue(dataRow, columnMap, formulaEvaluator, "qty sold", "quantity sold", "quantity", "qty", "amount"));
        record.setPurchaseDate(getDateValue(dataRow, columnMap, formulaEvaluator, "sale date", "date", "purchasedate", "orderdate"));
        record.setCategory(getStringValue(dataRow, columnMap, formulaEvaluator, "category", "type", "group", "sku")); // Include SKU as category for now
        record.setVendor(getStringValue(dataRow, columnMap, formulaEvaluator, "customer name", "vendor", "supplier", "store", "customer"));
        record.setTotalCost(getBigDecimalValue(dataRow, columnMap, formulaEvaluator, "total amount", "total", "totalcost", "totalprice"));

        return record;
    }

    /**
     * Gets string value from row using multiple possible column names.
     * NOW SUPPORTS FORMULAS!
     */
    private static String getStringValue(Row row, Map<String, Integer> columnMap, FormulaEvaluator formulaEvaluator, String... possibleNames) {
        for (String name : possibleNames) {
            Integer columnIndex = columnMap.get(name);
            if (columnIndex != null) {
                Cell cell = row.getCell(columnIndex);
                if (cell != null) {
                    String value = getCellValueAsString(cell, formulaEvaluator);
                    if (!value.isEmpty()) {
                        return value;
                    }
                }
            }
        }
        return "";
    }

    /**
     * Gets BigDecimal value from row using multiple possible column names.
     * FIXED TO HANDLE FORMULAS!
     */
    private static BigDecimal getBigDecimalValue(Row row, Map<String, Integer> columnMap, FormulaEvaluator formulaEvaluator, String... possibleNames) {
        for (String name : possibleNames) {
            Integer columnIndex = columnMap.get(name);
            if (columnIndex != null) {
                Cell cell = row.getCell(columnIndex);
                if (cell != null) {
                    try {
                        double numericValue = getNumericCellValue(cell, formulaEvaluator);
                        if (!Double.isNaN(numericValue)) {
                            return BigDecimal.valueOf(numericValue);
                        }
                    } catch (Exception e) {
                        System.err.println("Error getting numeric value from cell: " + e.getMessage());
                    }
                }
            }
        }
        return BigDecimal.ZERO;
    }

    /**
     * Gets integer value from row using multiple possible column names.
     * FIXED TO HANDLE FORMULAS!
     */
    private static int getIntValue(Row row, Map<String, Integer> columnMap, FormulaEvaluator formulaEvaluator, String... possibleNames) {
        for (String name : possibleNames) {
            Integer columnIndex = columnMap.get(name);
            if (columnIndex != null) {
                Cell cell = row.getCell(columnIndex);
                if (cell != null) {
                    try {
                        double numericValue = getNumericCellValue(cell, formulaEvaluator);
                        if (!Double.isNaN(numericValue)) {
                            return (int) numericValue;
                        }
                    } catch (Exception e) {
                        System.err.println("Error getting integer value from cell: " + e.getMessage());
                    }
                }
            }
        }
        return 0;
    }

    /**
     * Gets LocalDate value from row using multiple possible column names.
     */
    private static LocalDate getDateValue(Row row, Map<String, Integer> columnMap, FormulaEvaluator formulaEvaluator, String... possibleNames) {
        for (String name : possibleNames) {
            Integer columnIndex = columnMap.get(name);
            if (columnIndex != null) {
                Cell cell = row.getCell(columnIndex);
                if (cell != null) {
                    if (DateUtil.isCellDateFormatted(cell)) {
                        return cell.getDateCellValue().toInstant()
                                .atZone(ZoneId.systemDefault())
                                .toLocalDate();
                    } else if (cell.getCellType() == CellType.NUMERIC) {
                        // Handle Excel date as numeric value
                        Date date = DateUtil.getJavaDate(cell.getNumericCellValue());
                        return date.toInstant().atZone(ZoneId.systemDefault()).toLocalDate();
                    }
                }
            }
        }
        return null;
    }

    /**
     * NEW METHOD: Gets numeric value from cell, handling both direct values and formulas.
     */
    private static double getNumericCellValue(Cell cell, FormulaEvaluator formulaEvaluator) {
        if (cell == null) {
            return Double.NaN;
        }

        switch (cell.getCellType()) {
            case NUMERIC:
                return cell.getNumericCellValue();
            case FORMULA:
                try {
                    // Evaluate the formula and get the result
                    CellValue cellValue = formulaEvaluator.evaluate(cell);
                    if (cellValue.getCellType() == CellType.NUMERIC) {
                        return cellValue.getNumberValue();
                    }
                } catch (Exception e) {
                    System.err.println("Error evaluating formula in cell: " + e.getMessage());
                    // Try to get cached value if evaluation fails
                    try {
                        return cell.getNumericCellValue();
                    } catch (Exception e2) {
                        System.err.println("Error getting cached value: " + e2.getMessage());
                    }
                }
                break;
            case STRING:
                // Try to parse string as number
                try {
                    return Double.parseDouble(cell.getStringCellValue().trim());
                } catch (NumberFormatException e) {
                    // Not a number
                }
                break;
        }
        return Double.NaN;
    }

    /**
     * Analyzes Excel file structure and provides column information.
     * UPDATED TO SUPPORT FORMULAS!
     */
    public static ExcelAnalysis analyzeExcelFile(String filePath) throws IOException {
        ExcelAnalysis analysis = new ExcelAnalysis();

        try (FileInputStream fis = new FileInputStream(filePath);
             XSSFWorkbook workbook = new XSSFWorkbook(fis)) {

            // Create formula evaluator
            FormulaEvaluator formulaEvaluator = workbook.getCreationHelper().createFormulaEvaluator();

            Sheet sheet = workbook.getSheetAt(0);
            Row headerRow = sheet.getRow(0);

            if (headerRow == null) {
                throw new IOException("No header row found");
            }

            // Find actual data range
            int firstDataRow = findFirstDataRow(sheet);
            int lastDataRow = findLastDataRow(sheet);

            // Analyze each column
            for (int colNum = 0; colNum < headerRow.getLastCellNum(); colNum++) {
                Cell headerCell = headerRow.getCell(colNum);
                if (headerCell != null) {
                    String columnName = headerCell.getStringCellValue();
                    ColumnInfo columnInfo = analyzeColumn(sheet, colNum, columnName, firstDataRow, lastDataRow, formulaEvaluator);
                    analysis.addColumn(columnInfo);
                }
            }
        }

        return analysis;
    }

    /**
     * Analyzes a specific column to determine its data type and calculate statistics.
     * UPDATED TO SUPPORT FORMULAS!
     */
    private static ColumnInfo analyzeColumn(Sheet sheet, int columnIndex, String columnName, int firstDataRow, int lastDataRow, FormulaEvaluator formulaEvaluator) {
        ColumnInfo info = new ColumnInfo(columnName);
        List<Double> numericValues = new ArrayList<>();
        List<String> sampleValues = new ArrayList<>();
        int totalCells = 0;
        int emptyCells = 0;

        System.out.println("\nAnalyzing column: " + columnName + " (index " + columnIndex + ")");
        System.out.println("  Processing rows " + firstDataRow + " to " + lastDataRow);

        // Analyze only actual data rows
        for (int rowNum = firstDataRow; rowNum <= lastDataRow; rowNum++) {
            Row row = sheet.getRow(rowNum);
            if (row != null) {
                Cell cell = row.getCell(columnIndex);
                totalCells++;

                if (cell == null || cell.getCellType() == CellType.BLANK) {
                    emptyCells++;
                } else {
                    String cellValue = getCellValueAsString(cell, formulaEvaluator);
                    if (!cellValue.isEmpty()) {
                        // Collect sample values
                        if (sampleValues.size() < 5) {
                            sampleValues.add(cellValue);
                        }

                        // Try to get numeric value for statistics
                        double numValue = getNumericCellValue(cell, formulaEvaluator);
                        if (!Double.isNaN(numValue)) {
                            // Skip date values (Excel dates are large numbers like 45000+)
                            if (!columnName.toLowerCase().contains("date") || numValue <= 40000) {
                                numericValues.add(numValue);
                            }
                        }
                    }
                }
            }
        }

        info.setSampleValues(sampleValues);
        info.setTotalCells(totalCells);
        info.setEmptyCells(emptyCells);

        // Calculate statistics from actual data
        if (!numericValues.isEmpty()) {
            info.setNumeric(true);
            double sum = numericValues.stream().mapToDouble(Double::doubleValue).sum();
            double average = sum / numericValues.size();
            info.setSum(sum);
            info.setAverage(average);

            System.out.println("  Column '" + columnName + "': " + numericValues.size() + " numeric values, sum = " + sum + ", avg = " + average);
        } else {
            info.setNumeric(false);
            System.out.println("  Column '" + columnName + "': No numeric values found");
        }

        return info;
    }

    /**
     * Extracts cell value as string, handling different cell types.
     * UPDATED TO SUPPORT FORMULAS!
     */
    private static String getCellValueAsString(Cell cell, FormulaEvaluator formulaEvaluator) {
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
                    CellValue cellValue = formulaEvaluator.evaluate(cell);
                    switch (cellValue.getCellType()) {
                        case NUMERIC:
                            double numValue = cellValue.getNumberValue();
                            if (numValue == Math.floor(numValue)) {
                                return String.valueOf((long) numValue);
                            } else {
                                return String.valueOf(numValue);
                            }
                        case STRING:
                            return cellValue.getStringValue();
                        case BOOLEAN:
                            return String.valueOf(cellValue.getBooleanValue());
                        default:
                            return "";
                    }
                } catch (Exception e) {
                    System.err.println("Error evaluating formula: " + e.getMessage());
                    return "Formula Error";
                }
            default:
                return "";
        }
    }

    /**
     * OVERLOADED METHOD: Extracts cell value as string without formula evaluator (for backward compatibility).
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
        if (row == null) return true;

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