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

            // Debug: Print the column map
            System.out.println("Column mapping:");
            for (Map.Entry<String, Integer> entry : columnMap.entrySet()) {
                System.out.println("  '" + entry.getKey() + "' -> column " + entry.getValue());
            }

            // Detect where the actual data ends (before summary rows)
            int lastDataRow = findLastDataRow(sheet);
            System.out.println("Data rows: 1 to " + lastDataRow + " (total sheet rows: " + sheet.getLastRowNum() + ")");

            // Read data rows (excluding summary rows)
            for (int rowNum = 1; rowNum <= lastDataRow; rowNum++) {
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
     * Finds the last row that contains actual data (not summary rows).
     * Uses a simple but effective approach: exclude last 2 rows if they look like summaries.
     */
    private static int findLastDataRow(Sheet sheet) {
        int totalRows = sheet.getLastRowNum();

        System.out.println("Total rows in sheet: " + (totalRows + 1) + " (0-indexed: " + totalRows + ")");

        // Simple approach: check if the last 2 rows look like summary rows
        // If so, exclude them. If not, include all rows.
        boolean lastRowsAreSummary = checkLastTwoRowsForSummary(sheet);

        if (lastRowsAreSummary) {
            int lastDataRow = totalRows - 2;
            System.out.println("Last 2 rows detected as summary rows. Data ends at row " + (lastDataRow + 1));
            return lastDataRow;
        } else {
            System.out.println("No summary rows detected. Processing all rows.");
            return totalRows;
        }
    }

    /**
     * Specifically checks if the last two rows contain summary data.
     */
    private static boolean checkLastTwoRowsForSummary(Sheet sheet) {
        int totalRows = sheet.getLastRowNum();

        System.out.println("Checking rows " + (totalRows - 1) + " and " + totalRows + " for summary patterns");

        // Check both of the last two rows (0-indexed: totalRows-1 and totalRows)
        for (int rowNum = totalRows - 1; rowNum <= totalRows; rowNum++) {
            Row row = sheet.getRow(rowNum);
            if (row != null) {
                System.out.println("Checking row " + (rowNum + 1) + " for summary pattern:");

                int emptyCells = 0;
                int totalCells = 0;
                int numericCells = 0;
                boolean hasLargeNumbers = false;

                // Check up to 10 cells or the actual number of cells in the row
                int cellsToCheck = Math.min(10, row.getLastCellNum());

                for (int cellNum = 0; cellNum < cellsToCheck; cellNum++) {
                    Cell cell = row.getCell(cellNum);
                    totalCells++;

                    if (cell == null || cell.getCellType() == CellType.BLANK) {
                        emptyCells++;
                    } else if (cell.getCellType() == CellType.NUMERIC) {
                        numericCells++;
                        double value = cell.getNumericCellValue();
                        System.out.println("  Cell " + cellNum + ": " + value);

                        // Look for numbers that could be totals (avoid dates which are 40000+)
                        if (value >= 1000 && value <= 100000) {
                            hasLargeNumbers = true;
                        }
                    } else if (cell.getCellType() == CellType.STRING) {
                        System.out.println("  Cell " + cellNum + ": '" + cell.getStringCellValue().trim() + "'");
                    }
                }

                System.out.println("  Empty: " + emptyCells + "/" + totalCells + ", Numeric: " + numericCells + ", Has large numbers: " + hasLargeNumbers);

                // If this row is mostly empty (60%+) and has large numbers, it's likely a summary
                double emptyPercentage = totalCells > 0 ? (double) emptyCells / totalCells : 0;
                if (emptyPercentage >= 0.6 && hasLargeNumbers) {
                    System.out.println("  -> Row " + (rowNum + 1) + " looks like a summary row");
                    return true;
                }
            }
        }

        System.out.println("No summary rows detected in last 2 rows");
        return false;
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
     * Updated to match the actual Excel column names.
     */
    private static PurchaseRecord createPurchaseRecord(Row dataRow, Map<String, Integer> columnMap) {
        PurchaseRecord record = new PurchaseRecord();

        // Extract data based on actual Excel column names
        record.setItemName(getStringValue(dataRow, columnMap, "procuct name", "product name", "item", "itemname", "name"));
        record.setPrice(getBigDecimalValue(dataRow, columnMap, "unit price", "price", "cost", "unitprice"));
        record.setQuantity(getIntValue(dataRow, columnMap, "qty sold", "quantity sold", "quantity", "qty", "amount"));
        record.setPurchaseDate(getDateValue(dataRow, columnMap, "sale date", "date", "purchasedate", "orderdate"));
        record.setCategory(getStringValue(dataRow, columnMap, "category", "type", "group"));
        record.setVendor(getStringValue(dataRow, columnMap, "customer name", "vendor", "supplier", "store", "customer"));
        record.setTotalCost(getBigDecimalValue(dataRow, columnMap, "total amount", "total", "totalcost", "totalprice"));

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
                    String value = getCellValueAsString(cell);
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

            // Find last data row (exclude summary rows)
            int lastDataRow = findLastDataRow(sheet);

            // Analyze each column
            for (int colNum = 0; colNum < headerRow.getLastCellNum(); colNum++) {
                Cell headerCell = headerRow.getCell(colNum);
                if (headerCell != null) {
                    String columnName = headerCell.getStringCellValue();
                    ColumnInfo columnInfo = analyzeColumn(sheet, colNum, columnName, lastDataRow);
                    analysis.addColumn(columnInfo);
                }
            }
        }

        return analysis;
    }

    /**
     * Analyzes a specific column to determine its data type and calculate statistics.
     * Only analyzes actual data rows, excluding summary rows.
     */
    private static ColumnInfo analyzeColumn(Sheet sheet, int columnIndex, String columnName, int lastDataRow) {
        ColumnInfo info = new ColumnInfo(columnName);
        List<Double> numericValues = new ArrayList<>();
        List<String> sampleValues = new ArrayList<>();
        int totalCells = 0;
        int emptyCells = 0;

        System.out.println("\nAnalyzing column: " + columnName + " (index " + columnIndex + ")");
        System.out.println("  Processing rows 1 to " + lastDataRow);

        // Analyze only data rows (excluding summary rows)
        for (int rowNum = 1; rowNum <= lastDataRow; rowNum++) {
            Row row = sheet.getRow(rowNum);
            if (row != null) {
                Cell cell = row.getCell(columnIndex);
                totalCells++;

                if (cell == null || cell.getCellType() == CellType.BLANK) {
                    emptyCells++;
                } else {
                    String cellValue = getCellValueAsString(cell);
                    if (!cellValue.isEmpty()) {
                        // Collect sample values
                        if (sampleValues.size() < 5) {
                            sampleValues.add(cellValue);
                        }

                        // Try to parse as number for statistics
                        if (cell.getCellType() == CellType.NUMERIC) {
                            double numValue = cell.getNumericCellValue();

                            // Skip date values (Excel dates are large numbers like 45000+)
                            if (columnName.toLowerCase().contains("date") && numValue > 40000) {
                                continue;
                            }

                            numericValues.add(numValue);
                        } else {
                            // Try to parse string as number for columns that should be numeric
                            if (columnName.toLowerCase().contains("qty") ||
                                    columnName.toLowerCase().contains("total") ||
                                    columnName.toLowerCase().contains("price")) {
                                try {
                                    String cleanValue = cellValue.replace(",", "").replace("$", "").trim();
                                    double numValue = Double.parseDouble(cleanValue);
                                    numericValues.add(numValue);
                                } catch (NumberFormatException e) {
                                    // Not a number, skip
                                }
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