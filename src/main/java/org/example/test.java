//import org.apache.poi.ss.usermodel.*;
//import java.io.FileInputStream;
//
//public class test {
//    public static void main(String[] args) throws Exception {
//        // Replace "your-file.xlsx" with your actual file path
//        FileInputStream file = new FileInputStream("David-Infinity-Purchase-History - Copy.xlsx");
//        Workbook workbook = WorkbookFactory.create(file);
//        Sheet sheet = workbook.getSheetAt(0); // Gets first sheet
//
//        // Define the column indices we want (0-based): A, C, D, E, F, G, H, I
//        int[] columnsToShow = {0, 2, 3, 4, 5, 6, 7, 8}; // A=0, C=2, D=3, etc.
//
//        int lastRowNum = sheet.getLastRowNum();
//        for (int i = 0; i <= lastRowNum; i++) {
//            Row row = sheet.getRow(i);
//            if (row != null) {
//                // Print only the specified columns
//                for (int colIndex : columnsToShow) {
//                    Cell cell = row.getCell(colIndex);
//                    System.out.print(getCellValue(cell) + "\t");
//                }
//                System.out.println(); // New line after each row
//            } else {
//                // Handle null rows by printing empty values for our columns
//                for (int j = 0; j < columnsToShow.length; j++) {
//                    System.out.print("\t");
//                }
//                System.out.println();
//            }
//        }
//
//        workbook.close();
//        file.close();
//    }
//
//    private static String getCellValue(Cell cell) {
//        if (cell == null) return "";
//
//        switch (cell.getCellType()) {
//            case STRING: return cell.getStringCellValue();
//            case NUMERIC:
//                if (DateUtil.isCellDateFormatted(cell)) {
//                    return cell.getDateCellValue().toString();
//                } else {
//                    return String.valueOf(cell.getNumericCellValue());
//                }
//            case BOOLEAN: return String.valueOf(cell.getBooleanCellValue());
//            default: return "";
//        }
//    }
//}