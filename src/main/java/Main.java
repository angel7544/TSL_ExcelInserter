import java.awt.*;
import java.awt.event.*;
import java.io.*;
import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class Main extends JFrame {
    private JButton importButton;
    private JButton processButton;
    private JButton saveButton;
    private JLabel statusLabel;
    private JTextField filePathField;
    private JTextField outputFileField;

    private File selectedFile;
    private Workbook workbook;

    public Main() {

        setTitle("TSL - Excel Row Inserter");
        setSize(600, 200);
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setLocationRelativeTo(null);

        JPanel panel = new JPanel();
        panel.setLayout(new GridLayout(4, 1, 10, 10));
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

        processButton = new JButton("Insert Rows Between");
        processButton.setEnabled(false);

        saveButton = new JButton("Save File");
        saveButton.setEnabled(false);

        statusLabel = new JLabel("Select an Excel file to start");
        statusLabel.setHorizontalAlignment(JLabel.CENTER);

        panel.add(filePanel);
        panel.add(outputPanel);

        JPanel buttonPanel = new JPanel(new FlowLayout(FlowLayout.CENTER));
        buttonPanel.add(processButton);
        buttonPanel.add(saveButton);
        panel.add(buttonPanel);

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

        processButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                insertRowsBetween();
            }
        });

        saveButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                saveExcelFile();
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

                processButton.setEnabled(true);
                statusLabel.setText("File loaded successfully. Click 'Insert Rows Between' to process.");
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
            statusLabel.setText("Rows inserted successfully. Click 'Save File' to save the changes.");
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

    private void showAboutDialog() {
        JOptionPane.showMessageDialog(this,
                "<html><b>TSL - Excel Row Inserter</b><br>" +
                        "Version: 1.0.2<br>" +
                        "Developer: Angel Singh<br>" +
                        "Contact: angelsingh2199@gmail.com<br>" +
                        "Description: This tool allows inserting blank rows between every row of an Excel file.<br>" +
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
