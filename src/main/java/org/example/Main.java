package org.example;

/**
 * Main application entry point.
 * This class now only handles application startup.
 * All Excel processing logic has been moved to ExcelReaderUtility.
 * All GUI logic has been moved to ExcelViewer.
 */
public class Main {

    /**
     * Main method - launches the GUI application.
     *
     * @param args Command line arguments (not used)
     */
    public static void main(String[] args) {
        // Launch the GUI application
        ExcelViewer.main(args);
    }
}