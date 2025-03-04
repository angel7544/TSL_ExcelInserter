
import java.awt.*;
import java.awt.event.*;
import java.io.*;
import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.table.DefaultTableModel;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class Main extends JFrame {
    private JButton importButton;
    private JButton insertRowsButton;
    private JButton insertColumnsButton;
    private JButton saveButton;
    private JButton previewButton;
    private JLabel statusLabel;
    private JTextField filePathField;
    private JTextField outputFileField;

    private File selectedFile;
    private Workbook workbook;
    private JDialog previewDialog;

    public Main() {
        setTitle("TSL - Excel Editor");
        setSize(600, 250);
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setLocationRelativeTo(null);

        JPanel panel = new JPanel();
        panel.setLayout(new GridLayout(5, 1, 10, 10));
        panel.setBorder(BorderFactory.createEmptyBorder(10, 10, 10, 10));

        JPanel filePanel = new JPanel(new BorderLayout(5, 0));
        filePathField = new JTextField();
        filePathField.setEditable(false);
        importButton = new JButton("Import Excel");
        filePanel.add(filePathField, BorderLayout.CENTER);
        filePanel.add(importButton, BorderLayout.EAST);

        JPanel outputPanel = new JPanel(new BorderLayout(5, 0));
        outputFileField = new JTextField();
        JLabel outputLabel = new JLabel("Output file: ");
        outputPanel.add(outputLabel, BorderLayout.WEST);
        outputPanel.add(outputFileField, BorderLayout.CENTER);

        insertRowsButton = new JButton("Insert Rows Between");
        insertRowsButton.setEnabled(false);
        
        insertColumnsButton = new JButton("Insert Columns Between");
        insertColumnsButton.setEnabled(false);

        saveButton = new JButton("Save File");
        saveButton.setEnabled(false);
        
        previewButton = new JButton("Preview");
        previewButton.setEnabled(false);

        statusLabel = new JLabel("Select an Excel file to start");
        statusLabel.setHorizontalAlignment(JLabel.CENTER);

        panel.add(filePanel);
        panel.add(outputPanel);

        JPanel buttonPanel = new JPanel(new FlowLayout(FlowLayout.CENTER));
        buttonPanel.add(insertRowsButton);
        buttonPanel.add(insertColumnsButton);
        panel.add(buttonPanel);
        
        JPanel actionPanel = new JPanel(new FlowLayout(FlowLayout.CENTER));
        actionPanel.add(previewButton);
        actionPanel.add(saveButton);
        panel.add(actionPanel);

        panel.add(statusLabel);

        add(panel);

        // Adding MenuBar
        JMenuBar menuBar = new JMenuBar();
        JMenu helpMenu = new JMenu("Help");
        JMenuItem aboutItem = new JMenuItem("About");

        aboutItem.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                showAboutDialog();
            }
        });

        helpMenu.add(aboutItem);
        menuBar.add(helpMenu);
        setJMenuBar(menuBar);

        // Add action listeners
        importButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                importExcelFile();
            }
        });

        insertRowsButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                insertRowsBetween();
            }
        });
        
        insertColumnsButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                insertColumnsBetween();
            }
        });

        saveButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                saveExcelFile();
            }
        });
        
        previewButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                previewExcelFile();
            }
        });
    }

    private void importExcelFile() {
        JFileChooser fileChooser = new JFileChooser();
        fileChooser.setDialogTitle("Select Excel File");
        FileNameExtensionFilter filter = new FileNameExtensionFilter(
                "Excel Files", "xls", "xlsx");
        fileChooser.setFileFilter(filter);

        int result = fileChooser.showOpenDialog(this);
        if (result == JFileChooser.APPROVE_OPTION) {
            selectedFile = fileChooser.getSelectedFile();
            filePathField.setText(selectedFile.getAbsolutePath());

            String inputPath = selectedFile.getAbsolutePath();
            String outputPath = inputPath.substring(0, inputPath.lastIndexOf('.')) + "_modified" +
                    inputPath.substring(inputPath.lastIndexOf('.'));
            outputFileField.setText(outputPath);

            try {
                FileInputStream fileInputStream = new FileInputStream(selectedFile);
                if (selectedFile.getName().endsWith(".xlsx")) {
                    workbook = new XSSFWorkbook(fileInputStream);
                } else {
                    workbook = new HSSFWorkbook(fileInputStream);
                }
                fileInputStream.close();

                insertRowsButton.setEnabled(true);
                insertColumnsButton.setEnabled(true);
                previewButton.setEnabled(true);
                statusLabel.setText("File loaded successfully. Choose an action to process.");
            } catch (IOException ex) {
                JOptionPane.showMessageDialog(this,
                        "Error loading file: " + ex.getMessage(),
                        "Error", JOptionPane.ERROR_MESSAGE);
                statusLabel.setText("Error loading file");
            }
        }
    }

    private void insertRowsBetween() {
        if (workbook == null) {
            statusLabel.setText("No workbook loaded");
            return;
        }

        try {
            Sheet sheet = workbook.getSheetAt(0);

            int lastRow = -1;
            for (int i = sheet.getLastRowNum(); i >= 0; i--) {
                Row row = sheet.getRow(i);
                if (row != null && row.getCell(0) != null &&
                        row.getCell(0).getCellType() != CellType.BLANK) {
                    lastRow = i;
                    break;
                }
            }

            if (lastRow < 0) {
                statusLabel.setText("No data found in column A");
                return;
            }

            for (int i = lastRow - 1; i >= 0; i--) {
                sheet.shiftRows(i + 1, sheet.getLastRowNum(), 1);
            }

            saveButton.setEnabled(true);
            statusLabel.setText("Rows inserted successfully. Click 'Preview' to view or 'Save File' to save the changes.");
        } catch (Exception ex) {
            JOptionPane.showMessageDialog(this,
                    "Error processing file: " + ex.getMessage(),
                    "Error", JOptionPane.ERROR_MESSAGE);
            statusLabel.setText("Error processing file");
        }
    }
    
    private void insertColumnsBetween() {
        if (workbook == null) {
            statusLabel.setText("No workbook loaded");
            return;
        }

        try {
            Sheet sheet = workbook.getSheetAt(0);
            int lastColIndex = -1;
            
            // Find the last non-empty column
            for (int r = 0; r <= sheet.getLastRowNum(); r++) {
                Row row = sheet.getRow(r);
                if (row != null) {
                    for (int c = row.getLastCellNum() - 1; c >= 0; c--) {
                        Cell cell = row.getCell(c);
                        if (cell != null && cell.getCellType() != CellType.BLANK) {
                            lastColIndex = Math.max(lastColIndex, c);
                            break;
                        }
                    }
                }
            }
            
            if (lastColIndex < 0) {
                statusLabel.setText("No data found in the sheet");
                return;
            }
            
            // Insert columns between existing columns
            for (int i = lastColIndex; i > 0; i--) {
                for (int r = 0; r <= sheet.getLastRowNum(); r++) {
                    Row row = sheet.getRow(r);
                    if (row != null) {
                        // Shift cells right for each row
                        for (int c = row.getLastCellNum() - 1; c >= i; c--) {
                            Cell oldCell = row.getCell(c);
                            Cell newCell = row.createCell(c + 1);
                            
                            if (oldCell != null) {
                                // Copy cell style and value
                                CellStyle newStyle = workbook.createCellStyle();
                                newStyle.cloneStyleFrom(oldCell.getCellStyle());
                                newCell.setCellStyle(newStyle);
                                
                                switch (oldCell.getCellType()) {
                                    case STRING:
                                        newCell.setCellValue(oldCell.getStringCellValue());
                                        break;
                                    case NUMERIC:
                                        newCell.setCellValue(oldCell.getNumericCellValue());
                                        break;
                                    case BOOLEAN:
                                        newCell.setCellValue(oldCell.getBooleanCellValue());
                                        break;
                                    case FORMULA:
                                        newCell.setCellFormula(oldCell.getCellFormula());
                                        break;
                                    default:
                                        break;
                                }
                            }
                        }
                        // Create the new blank cell
                        row.createCell(i);
                    }
                }
            }
            
            saveButton.setEnabled(true);
            statusLabel.setText("Columns inserted successfully. Click 'Preview' to view or 'Save File' to save the changes.");
        } catch (Exception ex) {
            JOptionPane.showMessageDialog(this,
                    "Error processing file: " + ex.getMessage(),
                    "Error", JOptionPane.ERROR_MESSAGE);
            statusLabel.setText("Error processing file");
        }
    }

    private void saveExcelFile() {
        if (workbook == null) {
            statusLabel.setText("No workbook to save");
            return;
        }

        String outputPath = outputFileField.getText();
        if (outputPath.isEmpty()) {
            JOptionPane.showMessageDialog(this,
                    "Please specify an output file name",
                    "Error", JOptionPane.WARNING_MESSAGE);
            return;
        }

        try {
            FileOutputStream fileOut = new FileOutputStream(outputPath);
            workbook.write(fileOut);
            fileOut.close();

            statusLabel.setText("File saved successfully as: " + outputPath);
            JOptionPane.showMessageDialog(this,
                    "File saved successfully",
                    "Success", JOptionPane.INFORMATION_MESSAGE);
        } catch (IOException ex) {
            JOptionPane.showMessageDialog(this,
                    "Error saving file: " + ex.getMessage(),
                    "Error", JOptionPane.ERROR_MESSAGE);
            statusLabel.setText("Error saving file");
        }
    }
    
    private void previewExcelFile() {
        if (workbook == null) {
            statusLabel.setText("No workbook to preview");
            return;
        }
        
        try {
            Sheet sheet = workbook.getSheetAt(0);
            
            // Determine the number of rows and columns to display
            int rowCount = Math.min(sheet.getLastRowNum() + 1, 100); // Limit to 100 rows for preview
            int colCount = 0;
            
            for (int i = 0; i < rowCount; i++) {
                Row row = sheet.getRow(i);
                if (row != null) {
                    colCount = Math.max(colCount, row.getLastCellNum());
                }
            }
            
            colCount = Math.min(colCount, 15); // Limit to 15 columns for preview
            
            // Create column headers (A, B, C, ...)
            String[] columnHeaders = new String[colCount];
            for (int i = 0; i < colCount; i++) {
                columnHeaders[i] = getColumnName(i);
            }
            
            // Populate data
            Object[][] data = new Object[rowCount][colCount];
            
            for (int r = 0; r < rowCount; r++) {
                Row row = sheet.getRow(r);
                if (row != null) {
                    for (int c = 0; c < colCount; c++) {
                        Cell cell = row.getCell(c);
                        if (cell != null) {
                            switch (cell.getCellType()) {
                                case STRING:
                                    data[r][c] = cell.getStringCellValue();
                                    break;
                                case NUMERIC:
                                    if (DateUtil.isCellDateFormatted(cell)) {
                                        data[r][c] = cell.getDateCellValue();
                                    } else {
                                        data[r][c] = cell.getNumericCellValue();
                                    }
                                    break;
                                case BOOLEAN:
                                    data[r][c] = cell.getBooleanCellValue();
                                    break;
                                case FORMULA:
                                    data[r][c] = "=" + cell.getCellFormula();
                                    break;
                                default:
                                    data[r][c] = "";
                            }
                        } else {
                            data[r][c] = "";
                        }
                    }
                }
            }
            
            // Create and show the preview dialog
            if (previewDialog != null && previewDialog.isVisible()) {
                previewDialog.dispose();
            }
            
            previewDialog = new JDialog(this, "Excel Preview", false);
            previewDialog.setSize(800, 600);
            previewDialog.setLocationRelativeTo(this);
            
            JTable previewTable = new JTable(new DefaultTableModel(data, columnHeaders));
            JScrollPane scrollPane = new JScrollPane(previewTable);
            
            previewDialog.add(scrollPane);
            previewDialog.setVisible(true);
            
        } catch (Exception ex) {
            JOptionPane.showMessageDialog(this,
                    "Error previewing file: " + ex.getMessage(),
                    "Error", JOptionPane.ERROR_MESSAGE);
        }
    }
    
    // Helper method to convert column index to Excel-style column name (A, B, ..., Z, AA, AB, ...)
    private String getColumnName(int columnIndex) {
        StringBuilder columnName = new StringBuilder();
        int dividend = columnIndex + 1;
        int modulo;
        
        while (dividend > 0) {
            modulo = (dividend - 1) % 26;
            columnName.insert(0, (char) (65 + modulo));
            dividend = (int)((dividend - modulo) / 26);
        }
        
        return columnName.toString();
    }

    private void showAboutDialog() {
        JOptionPane.showMessageDialog(this,
                "<html><b>TSL - Excel Editor</b><br>" +
                        "Version: 1.1.0<br>" +
                        "Developer: Angel Singh<br>" +
                        "Contact: angelsingh2199@gmail.com<br>" +
                        "Description: This tool allows inserting blank rows or columns between every row/column of an Excel file.<br>" +
                        "Built using Apache POI and Swing.</html>",
                "About",
                JOptionPane.INFORMATION_MESSAGE);
    }

    public static void main(String[] args) {
        SwingUtilities.invokeLater(new Runnable() {
            @Override
            public void run() {
                Main app = new Main();
                app.setVisible(true);
            }
        });
    }
}
