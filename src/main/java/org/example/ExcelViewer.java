package org.example;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.table.DefaultTableModel;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.IOException;
import java.math.BigDecimal;
import java.time.format.DateTimeFormatter;
import java.util.List;

/**
 * Fixed GUI class for Excel file viewer and analyzer.
 * Handles all user interface interactions and display logic.
 */
public class ExcelViewer extends JFrame {

    // GUI Components
    private JButton loadFileButton;
    private JTable dataTable;
    private DefaultTableModel tableModel;
    private JTextArea analysisArea;
    private JScrollPane tableScrollPane;
    private JScrollPane analysisScrollPane;
    private JLabel statusLabel;
    private JSplitPane splitPane;

    // Data
    private List<PurchaseRecord> currentRecords;
    private String currentFilePath;

    /**
     * Constructor - sets up the GUI interface.
     */
    public ExcelViewer() {
        initializeGUI();
        setupEventHandlers();
    }

    /**
     * Initializes all GUI components and layout.
     */
    private void initializeGUI() {
        // Set up main window
        setTitle("Sales Data Analyzer");
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setSize(1200, 800);
        setLocationRelativeTo(null); // Center on screen

        // Create main panel with BorderLayout
        JPanel mainPanel = new JPanel(new BorderLayout());

        // Create top panel with controls
        JPanel topPanel = createTopPanel();
        mainPanel.add(topPanel, BorderLayout.NORTH);

        // Create center panel with table and analysis
        splitPane = createCenterPanel();
        mainPanel.add(splitPane, BorderLayout.CENTER);

        // Create bottom panel with status
        JPanel bottomPanel = createBottomPanel();
        mainPanel.add(bottomPanel, BorderLayout.SOUTH);

        // Add main panel to frame
        add(mainPanel);

        // Set look and feel
        try {
            UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
            SwingUtilities.updateComponentTreeUI(this);
        } catch (Exception e) {
            // Use default look and feel if system L&F fails
        }
    }

    /**
     * Creates the top panel with file loading controls.
     */
    private JPanel createTopPanel() {
        JPanel topPanel = new JPanel(new FlowLayout(FlowLayout.LEFT));
        topPanel.setBorder(BorderFactory.createEmptyBorder(10, 10, 5, 10));

        loadFileButton = new JButton("Load Excel File");
        loadFileButton.setPreferredSize(new Dimension(150, 30));

        JLabel instructionLabel = new JLabel("Select an Excel file (.xlsx) to analyze:");
        instructionLabel.setFont(new Font(Font.SANS_SERIF, Font.PLAIN, 12));

        topPanel.add(instructionLabel);
        topPanel.add(Box.createHorizontalStrut(10));
        topPanel.add(loadFileButton);

        return topPanel;
    }

    /**
     * Creates the center panel with split view of table and analysis.
     */
    private JSplitPane createCenterPanel() {
        // Create table panel
        JPanel tablePanel = createTablePanel();

        // Create analysis panel
        JPanel analysisPanel = createAnalysisPanel();

        // Create split pane
        JSplitPane splitPane = new JSplitPane(JSplitPane.HORIZONTAL_SPLIT);
        splitPane.setLeftComponent(tablePanel);
        splitPane.setRightComponent(analysisPanel);
        splitPane.setDividerLocation(700); // Initial divider position
        splitPane.setResizeWeight(0.6); // Table gets 60% of space

        return splitPane;
    }

    /**
     * Creates the data table panel.
     */
    private JPanel createTablePanel() {
        JPanel tablePanel = new JPanel(new BorderLayout());
        tablePanel.setBorder(BorderFactory.createTitledBorder("Sales Data"));

        // Initialize table model with cleaned up column names
        String[] columnNames = {"Product Name", "Unit Price", "Qty Sold", "Sale Date", "Customer Name", "Total Amount"};
        tableModel = new DefaultTableModel(columnNames, 0) {
            @Override
            public boolean isCellEditable(int row, int column) {
                return false; // Make table read-only
            }
        };

        dataTable = new JTable(tableModel);
        dataTable.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
        dataTable.setAutoResizeMode(JTable.AUTO_RESIZE_ALL_COLUMNS);
        dataTable.getTableHeader().setReorderingAllowed(false);

        // Style the table
        dataTable.setRowHeight(25);
        dataTable.setGridColor(Color.LIGHT_GRAY);
        dataTable.setShowGrid(true);

        tableScrollPane = new JScrollPane(dataTable);
        tableScrollPane.setPreferredSize(new Dimension(700, 600));

        tablePanel.add(tableScrollPane, BorderLayout.CENTER);

        return tablePanel;
    }

    /**
     * Creates the analysis panel for statistics and column information.
     */
    private JPanel createAnalysisPanel() {
        JPanel analysisPanel = new JPanel(new BorderLayout());
        analysisPanel.setBorder(BorderFactory.createTitledBorder("Data Analysis & Statistics"));

        analysisArea = new JTextArea();
        analysisArea.setEditable(false);
        analysisArea.setFont(new Font(Font.MONOSPACED, Font.PLAIN, 12));
        analysisArea.setBackground(new Color(248, 248, 248));
        analysisArea.setText("Load an Excel file to see analysis results...");

        analysisScrollPane = new JScrollPane(analysisArea);
        analysisScrollPane.setPreferredSize(new Dimension(400, 600));

        analysisPanel.add(analysisScrollPane, BorderLayout.CENTER);

        return analysisPanel;
    }

    /**
     * Creates the bottom status panel.
     */
    private JPanel createBottomPanel() {
        JPanel bottomPanel = new JPanel(new FlowLayout(FlowLayout.LEFT));
        bottomPanel.setBorder(BorderFactory.createEmptyBorder(5, 10, 10, 10));

        statusLabel = new JLabel("Ready - No file loaded");
        statusLabel.setFont(new Font(Font.SANS_SERIF, Font.PLAIN, 11));
        statusLabel.setForeground(Color.GRAY);

        bottomPanel.add(statusLabel);

        return bottomPanel;
    }

    /**
     * Sets up event handlers for GUI interactions.
     */
    private void setupEventHandlers() {
        loadFileButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                loadExcelFile();
            }
        });
    }

    /**
     * Handles Excel file loading and processing.
     */
    private void loadExcelFile() {
        JFileChooser fileChooser = new JFileChooser();
        fileChooser.setDialogTitle("Select Excel File");
        fileChooser.setFileSelectionMode(JFileChooser.FILES_ONLY);

        // Set file filter for Excel files
        FileNameExtensionFilter filter = new FileNameExtensionFilter(
                "Excel Files (*.xlsx)", "xlsx");
        fileChooser.setFileFilter(filter);

        // Set default directory (optional)
        fileChooser.setCurrentDirectory(new File(System.getProperty("user.dir")));

        int result = fileChooser.showOpenDialog(this);

        if (result == JFileChooser.APPROVE_OPTION) {
            File selectedFile = fileChooser.getSelectedFile();
            currentFilePath = selectedFile.getAbsolutePath();

            // Update status
            statusLabel.setText("Loading file: " + selectedFile.getName() + "...");
            loadFileButton.setEnabled(false);

            // Process file in background thread to prevent UI freezing
            SwingWorker<Void, Void> worker = new SwingWorker<Void, Void>() {
                private String errorMessage = null;

                @Override
                protected Void doInBackground() throws Exception {
                    try {
                        // Read the Excel file
                        currentRecords = ExcelReaderUtility.readExcelFile(currentFilePath);

                        // Perform analysis
                        ExcelReaderUtility.ExcelAnalysis analysis =
                                ExcelReaderUtility.analyzeExcelFile(currentFilePath);

                        // Update UI on EDT
                        SwingUtilities.invokeLater(() -> {
                            populateTable(currentRecords);
                            displayAnalysis(analysis);
                            updateStatus(selectedFile.getName(), currentRecords.size());
                        });

                    } catch (IOException ex) {
                        errorMessage = "Error reading file: " + ex.getMessage();
                    } catch (Exception ex) {
                        errorMessage = "Unexpected error: " + ex.getMessage();
                    }

                    return null;
                }

                @Override
                protected void done() {
                    loadFileButton.setEnabled(true);

                    if (errorMessage != null) {
                        showErrorDialog("File Loading Error", errorMessage);
                        statusLabel.setText("Error loading file");
                    }
                }
            };

            worker.execute();
        }
    }

    /**
     * Populates the JTable with data from PurchaseRecord list.
     * Updated to properly format and display only relevant columns.
     */
    private void populateTable(List<PurchaseRecord> records) {
        // Clear existing data
        tableModel.setRowCount(0);

        // DateTimeFormatter for better date formatting
        DateTimeFormatter dateFormatter = DateTimeFormatter.ofPattern("MM/dd/yyyy");

        // Add new data - only showing relevant columns without category
        for (PurchaseRecord record : records) {
            Object[] rowData = {
                    record.getItemName() != null ? record.getItemName() : "",                    // Product Name
                    record.getPrice() != null ? String.format("$%.2f", record.getPrice()) : "$0.00",  // Unit Price formatted
                    record.getQuantity(),                                                        // Qty Sold
                    record.getPurchaseDate() != null ? record.getPurchaseDate().format(dateFormatter) : "", // Sale Date formatted
                    record.getVendor() != null ? record.getVendor() : "",                       // Customer Name
                    record.getTotalCost() != null ? String.format("$%.2f", record.getTotalCost()) : "$0.00" // Total Amount formatted
            };
            tableModel.addRow(rowData);
        }

        // Auto-resize columns
        dataTable.setAutoResizeMode(JTable.AUTO_RESIZE_ALL_COLUMNS);
    }

    /**
     * Displays analysis results in the analysis text area.
     */
    private void displayAnalysis(ExcelReaderUtility.ExcelAnalysis analysis) {
        StringBuilder sb = new StringBuilder();
        sb.append("SALES DATA ANALYSIS\n");
        sb.append("===================\n\n");

        List<ExcelReaderUtility.ColumnInfo> columns = analysis.getColumns();

        sb.append("DATA OVERVIEW:\n");
        sb.append("--------------\n");
        sb.append(String.format("Records loaded: %d\n", currentRecords != null ? currentRecords.size() : 0));
        sb.append(String.format("Columns found: %d\n\n", columns.size()));

        sb.append("COLUMN INFORMATION:\n");
        sb.append("-------------------\n");

        for (int i = 0; i < columns.size(); i++) {
            ExcelReaderUtility.ColumnInfo column = columns.get(i);
            sb.append(String.format("%d. %s\n", (i + 1), column.getName()));
            sb.append(String.format("   Type: %s\n",
                    column.isNumeric() ? "Numeric" : "Text"));
            sb.append(String.format("   Non-empty cells: %d/%d\n",
                    (column.getTotalCells() - column.getEmptyCells()), column.getTotalCells()));

            // Show sample values
            if (!column.getSampleValues().isEmpty()) {
                sb.append("   Sample values: ");
                for (int j = 0; j < Math.min(3, column.getSampleValues().size()); j++) {
                    sb.append("\"").append(column.getSampleValues().get(j)).append("\"");
                    if (j < Math.min(2, column.getSampleValues().size() - 1)) {
                        sb.append(", ");
                    }
                }
                sb.append("\n");
            }

            // Show statistics for numeric columns
            if (column.isNumeric()) {
                sb.append(String.format("   Sum: $%.2f\n", column.getSum()));
                sb.append(String.format("   Average: $%.2f\n", column.getAverage()));
            }

            sb.append("\n");
        }

        // Sales summary statistics
        if (currentRecords != null && !currentRecords.isEmpty()) {
            sb.append("SALES SUMMARY:\n");
            sb.append("--------------\n");

            BigDecimal totalRevenue = currentRecords.stream()
                    .filter(r -> r.getTotalCost() != null)
                    .map(PurchaseRecord::getTotalCost)
                    .reduce(BigDecimal.ZERO, BigDecimal::add);

            int totalQuantity = currentRecords.stream()
                    .mapToInt(PurchaseRecord::getQuantity)
                    .sum();

            long uniqueProducts = currentRecords.stream()
                    .map(PurchaseRecord::getItemName)
                    .filter(name -> name != null && !name.trim().isEmpty())
                    .distinct()
                    .count();

            long uniqueCustomers = currentRecords.stream()
                    .map(PurchaseRecord::getVendor)
                    .filter(vendor -> vendor != null && !vendor.trim().isEmpty())
                    .distinct()
                    .count();

            sb.append(String.format("Total Revenue: $%.2f\n", totalRevenue));
            sb.append(String.format("Total Units Sold: %d\n", totalQuantity));
            sb.append(String.format("Unique Products: %d\n", uniqueProducts));
            sb.append(String.format("Unique Customers: %d\n", uniqueCustomers));

            if (totalQuantity > 0) {
                BigDecimal avgRevenuePerUnit = totalRevenue.divide(BigDecimal.valueOf(totalQuantity), 2, BigDecimal.ROUND_HALF_UP);
                sb.append(String.format("Average Revenue per Unit: $%.2f\n", avgRevenuePerUnit));
            }
        }

        analysisArea.setText(sb.toString());
        analysisArea.setCaretPosition(0); // Scroll to top
    }

    /**
     * Updates the status label with file information.
     */
    private void updateStatus(String fileName, int recordCount) {
        statusLabel.setText(String.format("Loaded: %s (%d records)", fileName, recordCount));
    }

    /**
     * Shows error dialog with proper formatting and user-friendly message.
     */
    private void showErrorDialog(String title, String message) {
        JOptionPane.showMessageDialog(
                this,
                message,
                title,
                JOptionPane.ERROR_MESSAGE
        );
    }

    /**
     * Shows information dialog.
     */
    private void showInfoDialog(String title, String message) {
        JOptionPane.showMessageDialog(
                this,
                message,
                title,
                JOptionPane.INFORMATION_MESSAGE
        );
    }

    /**
     * Main method to launch the application.
     */
    public static void main(String[] args) {
        // Set up the GUI to run on the Event Dispatch Thread
        SwingUtilities.invokeLater(new Runnable() {
            @Override
            public void run() {
                try {
                    // Set system look and feel
                    UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
                } catch (Exception e) {
                    // Fall back to default if system L&F is not available
                }

                // Create and show the GUI
                new ExcelViewer().setVisible(true);
            }
        });
    }
}