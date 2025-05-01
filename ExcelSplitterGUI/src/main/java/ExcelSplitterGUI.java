import javax.swing.*;
import java.awt.*;
import java.awt.event.*;
import java.io.*;
import java.util.ArrayList;
import java.util.List;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class ExcelSplitterGUI extends JFrame {
    private JTextField filePathField;
    private JTextField rowsPerFileField;
    private JTextField outputDirField;
    private JTextField fileNameField;
    private JButton browseButton;
    private JButton outputBrowseButton;
    private JButton startButton;
    private JProgressBar progressBar;
    private JComboBox<String> formatComboBox;

    public ExcelSplitterGUI() {
        setTitle("ExcelSplitter");
        setSize(550, 310);
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setLayout(new GridBagLayout());

        initComponents();
        layoutComponents();
    }

    private void initComponents() {
        filePathField = new JTextField(25);
        rowsPerFileField = new JTextField("500", 5);
        outputDirField = new JTextField(25);
        fileNameField = new JTextField("", 15);

        browseButton = new JButton("Browse...");
        browseButton.addActionListener(e -> chooseInputFile());

        outputBrowseButton = new JButton("Browse...");
        outputBrowseButton.addActionListener(e -> chooseOutputDirectory());

        startButton = new JButton("Split");
        startButton.setPreferredSize(new Dimension(524, 40));
        startButton.addActionListener(e -> startSplitting());

        progressBar = new JProgressBar();
        progressBar.setPreferredSize(new Dimension(524, 40));
        progressBar.setStringPainted(true);

        // Выбор формата файла
        formatComboBox = new JComboBox<>(new String[]{
                "Excel sheet XLSX (*.xlsx)",
                "Excel sheet XLS 97-2003 (*.xls)",
                "Excel CSV (*.csv)",
                "Notebook Text (*.txt)",
                "Web XML (*.xml)"
        });
        formatComboBox.setSelectedIndex(0);
    }

    private void layoutComponents() {
        GridBagConstraints gbc = new GridBagConstraints();
        gbc.insets = new Insets(5, 5, 5, 5);
        gbc.anchor = GridBagConstraints.WEST;
        gbc.fill = GridBagConstraints.HORIZONTAL;

        // Выбор файла
        gbc.gridy = 0;
        add(new JLabel("Load Excel file:"), gbc);

        gbc.gridx = 1;
        gbc.weightx = 1;
        add(filePathField, gbc);

        gbc.gridx = 2;
        gbc.weightx = 0;
        add(browseButton, gbc);

        // Количество строк
        gbc.gridy = 1;
        gbc.gridx = 0;
        add(new JLabel("Split to strings:"), gbc);

        gbc.gridx = 1;
        add(rowsPerFileField, gbc);

        // Имя файла
        gbc.gridy = 1;
        gbc.gridx = 2;
        add(new JLabel("E.g. -> 500"), gbc);

        // Поле для имени файла
        gbc.gridy = 2;
        gbc.gridx = 0;
        add(new JLabel("Output file name:"), gbc);

        gbc.gridx = 1;
        add(fileNameField, gbc);

        // Аннотация для имени файла
        gbc.gridy = 2;
        gbc.gridx = 2;
        add(new JLabel("«_» included"), gbc);

        // Папка сохранения
        gbc.gridy = 3;
        gbc.gridx = 0;
        add(new JLabel("Save to:"), gbc);

        gbc.gridx = 1;
        add(outputDirField, gbc);

        gbc.gridx = 2;
        add(outputBrowseButton, gbc);

        // Выбор формата
        gbc.gridy = 4;
        gbc.gridx = 0;
        add(new JLabel("Output format:"), gbc);

        gbc.gridx = 1;
        gbc.gridwidth = 2;
        add(formatComboBox, gbc);

        // Кнопка старта
        gbc.gridy = 5;
        gbc.gridx = 0;
        gbc.gridwidth = 3;
        gbc.fill = GridBagConstraints.CENTER;
        add(startButton, gbc);

        // Прогресс-бар
        gbc.gridy = 6;
        gbc.fill = GridBagConstraints.HORIZONTAL;
        add(progressBar, gbc);
    }

    private void chooseInputFile() {
        JFileChooser fileChooser = new JFileChooser();
        fileChooser.setFileFilter(new javax.swing.filechooser.FileNameExtensionFilter(
                "Excel Files", "xlsx", "xls"));

        if (fileChooser.showOpenDialog(this) == JFileChooser.APPROVE_OPTION) {
            filePathField.setText(fileChooser.getSelectedFile().getAbsolutePath());
        }
    }

    private void chooseOutputDirectory() {
        JFileChooser dirChooser = new JFileChooser();
        dirChooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);

        if (dirChooser.showOpenDialog(this) == JFileChooser.APPROVE_OPTION) {
            outputDirField.setText(dirChooser.getSelectedFile().getAbsolutePath());
        }
    }

    private void startSplitting() {
        String inputPath = filePathField.getText();
        String outputPath = outputDirField.getText();
        int rowsPerFile;

        try {
            rowsPerFile = Integer.parseInt(rowsPerFileField.getText());
        } catch (NumberFormatException e) {
            JOptionPane.showMessageDialog(this, "Choose correct number of strings (exp 500)", "Error", JOptionPane.ERROR_MESSAGE);
            return;
        }

        if (inputPath.isEmpty() || outputPath.isEmpty()) {
            JOptionPane.showMessageDialog(this, "Choose file and folder to save", "Error", JOptionPane.ERROR_MESSAGE);
            return;
        }

        new Thread(() -> {
            try {
                splitExcelFile(inputPath, outputPath, rowsPerFile);
                JOptionPane.showMessageDialog(ExcelSplitterGUI.this,
                        "Split well done", "Done", JOptionPane.INFORMATION_MESSAGE);
            } catch (Exception e) {
                JOptionPane.showMessageDialog(ExcelSplitterGUI.this,
                        "Error: " + e.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
                e.printStackTrace();
            }
        }).start();
    }

    private void splitExcelFile(String inputPath, String outputDir, int rowsPerFile) throws IOException {
        progressBar.setValue(0);
        progressBar.setString("Processing...");

        try (InputStream inp = new FileInputStream(inputPath);
             Workbook workbook = WorkbookFactory.create(inp)) {

            Sheet sheet = workbook.getSheetAt(0);
            int firstDataRow = findFirstDataRow(sheet);
            Row headerRow = sheet.getRow(firstDataRow);

            List<Row> dataRows = new ArrayList<>();
            for (int i = firstDataRow + 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row != null) dataRows.add(row);
            }

            new File(outputDir).mkdirs();
            int totalFiles = (int) Math.ceil((double) dataRows.size() / rowsPerFile);

            String format = (String) formatComboBox.getSelectedItem();
            String fileExtension = getFileExtension(format);
            String baseFileName = fileNameField.getText().trim();

            for (int i = 0, fileCount = 1; i < dataRows.size(); i += rowsPerFile, fileCount++) {
                int end = Math.min(i + rowsPerFile, dataRows.size());
                List<Row> chunk = dataRows.subList(i, end);

                String outputPath = outputDir+File.separator+baseFileName+"_"+fileCount+fileExtension;

                switch (format) {
                    case "Excel sheet XLSX (*.xlsx)":
                        saveAsExcel(new XSSFWorkbook(), headerRow, chunk, outputPath);
                        break;
                    case "Excel sheet XLS 97-2003 (*.xls)":
                        saveAsExcel(new HSSFWorkbook(), headerRow, chunk, outputPath);
                        break;
                    case "Excel CSV (*.csv)":
                        saveAsCsv(headerRow, chunk, outputPath);
                        break;
                    case "Notebook Text (*.txt)":
                        saveAsText(headerRow, chunk, outputPath);
                        break;
                    case "Web XML (*.xml)":
                        saveAsXml(headerRow, chunk, outputPath);
                        break;
                }

                final int progress = (int) ((double) i / dataRows.size() * 100);
                int finalFileCount = fileCount;
                SwingUtilities.invokeLater(() -> {
                    progressBar.setValue(progress);
                    progressBar.setString(String.format("%d/%d files", finalFileCount, totalFiles));
                });
            }
        }

        SwingUtilities.invokeLater(() -> {
            progressBar.setValue(100);
            progressBar.setString("Done");
        });
    }

    private String getFileExtension(String format) {
        try {
            int start = format.indexOf("(*") + 2;
            int end = format.indexOf(")");
            if (start > 0 && end > start) {
                return format.substring(start, end);
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return ".xlsx"; // Значение по умолчанию
    }

    private void saveAsExcel(Workbook workbook, Row headerRow, List<Row> rows, String outputPath) throws IOException {
        try (Workbook newWorkbook = workbook) {
            Sheet newSheet = newWorkbook.createSheet("Data");
            copyRow(headerRow, newSheet.createRow(0), newWorkbook);

            for (int j = 0; j < rows.size(); j++) {
                copyRow(rows.get(j), newSheet.createRow(j + 1), newWorkbook);
            }

            try (FileOutputStream out = new FileOutputStream(outputPath)) {
                newWorkbook.write(out);
            }
        }
    }

    private void saveAsCsv(Row headerRow, List<Row> rows, String outputPath) throws IOException {
        try (BufferedWriter writer = new BufferedWriter(new FileWriter(outputPath))) {
            writeCsvRow(writer, headerRow);
            for (Row row : rows) {
                writeCsvRow(writer, row);
            }
        }
    }

    private void saveAsText(Row headerRow, List<Row> rows, String outputPath) throws IOException {
        try (BufferedWriter writer = new BufferedWriter(new FileWriter(outputPath))) {
            writeTextRow(writer, headerRow);
            for (Row row : rows) {
                writeTextRow(writer, row);
            }
        }
    }

    private void saveAsXml(Row headerRow, List<Row> rows, String outputPath) throws IOException {
        try (BufferedWriter writer = new BufferedWriter(new FileWriter(outputPath))) {
            writer.write("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n");
            writer.write("<data>\n");

            // Заголовки
            writer.write("  <headers>\n");
            for (Cell cell : headerRow) {
                writer.write("    <header>" + escapeXml(getCellValue(cell)) + "</header>\n");
            }
            writer.write("  </headers>\n");

            // Данные
            writer.write("  <rows>\n");
            for (Row row : rows) {
                writer.write("    <row>\n");
                for (Cell cell : row) {
                    writer.write("      <cell>" + escapeXml(getCellValue(cell)) + "</cell>\n");
                }
                writer.write("    </row>\n");
            }
            writer.write("  </rows>\n");
            writer.write("</data>");
        }
    }

    private String getCellValue(Cell cell) {
        switch (cell.getCellType()) {
            case STRING: return cell.getStringCellValue();
            case NUMERIC: return String.valueOf(cell.getNumericCellValue());
            case BOOLEAN: return String.valueOf(cell.getBooleanCellValue());
            default: return "";
        }
    }

    private String escapeXml(String input) {
        return input.replace("&", "&amp;")
                .replace("<", "&lt;")
                .replace(">", "&gt;")
                .replace("\"", "&quot;")
                .replace("'", "&apos;");
    }

    private void writeCsvRow(BufferedWriter writer, Row row) throws IOException {
        StringBuilder sb = new StringBuilder();
        for (Cell cell : row) {
            if (sb.length() > 0) sb.append(",");
            sb.append(escapeCsv(getCellValue(cell)));
        }
        writer.write(sb.toString());
        writer.newLine();
    }

    private void writeTextRow(BufferedWriter writer, Row row) throws IOException {
        StringBuilder sb = new StringBuilder();
        for (Cell cell : row) {
            if (sb.length() > 0) sb.append("\t");
            sb.append(getCellValue(cell));
        }
        writer.write(sb.toString());
        writer.newLine();
    }

    private String escapeCsv(String value) {
        if (value.contains(",") || value.contains("\"") || value.contains("\n")) {
            return "\"" + value.replace("\"", "\"\"") + "\"";
        }
        return value;
    }

    private int findFirstDataRow(Sheet sheet) {
        for (int i = 0; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            if (row != null && isDataRow(row)) return i;
        }
        return 0;
    }

    private boolean isDataRow(Row row) {
        Cell cell = row.getCell(0);
        return cell != null && cell.getCellType() != CellType.BLANK;
    }

    private void copyRow(Row source, Row target, Workbook workbook) {
        for (Cell cell : source) {
            Cell newCell = target.createCell(cell.getColumnIndex());
            switch (cell.getCellType()) {
                case STRING: newCell.setCellValue(cell.getStringCellValue()); break;
                case NUMERIC: newCell.setCellValue(cell.getNumericCellValue()); break;
                case BOOLEAN: newCell.setCellValue(cell.getBooleanCellValue()); break;
                case FORMULA: newCell.setCellFormula(cell.getCellFormula()); break;
                default: newCell.setCellValue("");
            }
        }
    }

    public static void main(String[] args) {
        SwingUtilities.invokeLater(() -> {
            try {
                UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
            } catch (Exception e) {
                e.printStackTrace();
            }

            ExcelSplitterGUI app = new ExcelSplitterGUI();
            app.setLocationRelativeTo(null);
            app.setVisible(true);
        });
    }
}