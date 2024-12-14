package com.fis.services;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.*;

import com.fis.types.ReportParameters;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileWriter;
import java.io.IOException;

import com.opencsv.CSVWriter;

import java.sql.ResultSetMetaData;
import java.util.Map;

public class ExcelGenerator {
    private final String MAX_NUM_ROWS = "600000";
    final int ROWS_PER_FILE = 1000000;
    private Workbook workbook;

    public ExcelGenerator() {
        workbook = new SXSSFWorkbook();
    }

    public Sheet getSheet(String sheetName) {
        if (workbook.getSheet(sheetName) == null) {
            return this.workbook.createSheet(sheetName);
        }
        return this.workbook.getSheet(sheetName);
    }

    public Sheet generateExcel(int rowIndex, List<DynamicObject> dataList, Boolean optionalHeader)
        throws FileNotFoundException, IOException, InterruptedException {
        workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Report");
        if (!optionalHeader && !dataList.isEmpty()) {
            createHeaderRow(sheet, dataList.get(0).getColumns(), rowIndex - 1);
        }
        if (dataList.get(0).getProperties().isEmpty()) {
            return sheet;
        }
        CellStyle cellStyle = sheet.getWorkbook().createCellStyle();
        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setBorderTop(BorderStyle.THIN);
        cellStyle.setBorderRight(BorderStyle.THIN);
        cellStyle.setBorderLeft(BorderStyle.THIN);
        int _rowIndex = rowIndex;
        int stt = 1;

        // System.out.println("Start autosize");
        // for (int i = 0; i < dataList.get(0).getProperties().size(); i++) {
        // System.out.println("autosize column " + i);
        // sheet.autoSizeColumn(i);
        // }
        // System.out.println("End autosize");

        for (DynamicObject dynamicObject : dataList) {
            // System.out.println(stt);
            // System.out.println(dynamicObject.getProperties());
            Row row = sheet.createRow(_rowIndex++);
            row.createCell(0).setCellValue(stt++);
            row.getCell(0).setCellStyle(cellStyle);
            populateRowWithObjectData(row, dynamicObject.getProperties(), cellStyle);
            // if(stt>600000){
            // break;
            // }
            // if (stt > dataList.size() -10000) {
            // Thread.sleep(200);
            // }
        }
        return sheet;
        // try (FileOutputStream fos = new FileOutputStream(fileName)) {
        // workbook.write(fos);
        // } finally {
        // workbook.close();
        // }

    }

    // generate excel case data more than 1 million rows then split into multiple
    // sheets still return sheet to style and write to file
    public Sheet generateExcel(int rowIndex, List<DynamicObject> dataList, Boolean optionalHeader, int maxRowPerSheet)
        throws FileNotFoundException, IOException, InterruptedException {
        // check sheet name exist
        Sheet sheet;

        if (workbook.getSheet("Report") == null) {
            sheet = workbook.createSheet("Report");
        } else {
            sheet = workbook.getSheet("Report");
        }

        if (!optionalHeader && !dataList.isEmpty()) {
            createHeaderRow(sheet, dataList.get(0).getColumns(), rowIndex - 1);
        }
        if (dataList.get(0).getProperties().isEmpty()) {
            return sheet;
        }
//        CellStyle cellStyle = sheet.getWorkbook().createCellStyle();
//        cellStyle.setBorderBottom(BorderStyle.THIN);
//        cellStyle.setBorderTop(BorderStyle.THIN);
//        cellStyle.setBorderRight(BorderStyle.THIN);
//        cellStyle.setBorderLeft(BorderStyle.THIN);
        int _rowIndex = rowIndex;
        int stt = 1;
        int sheetIndex = 0;
        Sheet currentSheet = sheet;
        for (DynamicObject dynamicObject : dataList) {
            // System.out.println(stt + "/" + dataList.size());
            if (_rowIndex % maxRowPerSheet == 0) {
                ++sheetIndex;
//                writeFooter(currentSheet, _rowIndex + 6);
                currentSheet = workbook.createSheet("Report" + sheetIndex);
                System.out.println("Create sheet");
                if (!optionalHeader && !dataList.isEmpty()) {
                    createHeaderRow(currentSheet, dynamicObject.getColumns(), 0);
                }
                if (optionalHeader) {
                    // Code here
                }
                _rowIndex = 1;
            }
            Row row = currentSheet.createRow(_rowIndex);
            row.createCell(0).setCellValue(stt++);
//            row.getCell(0).setCellStyle(cellStyle);
//            populateRowWithObjectData(row, dynamicObject.getProperties(), cellStyle);
            // if (stt > dataList.size() -10000) {
            // Thread.sleep(200);
            // }
        }
        if (sheetIndex == 0) {
//            writeFooter(currentSheet, _rowIndex + 6);
        }
        return workbook.getSheetAt(0);
        // try (FileOutputStream fos = new FileOutputStream(fileName)) {
        // workbook.write(fos);
        // } finally {
        // workbook.close();
        // }

    }

    public List<DynamicObject> mapResultSetToDynamicObject(ResultSet resultSet, Map<String, String> columns) throws SQLException, IOException, InterruptedException {
        List<DynamicObject> resultList = new ArrayList<>();
        ResultSetMetaData metaData = resultSet.getMetaData();
        int columnCount = metaData.getColumnCount();

        // Pre-compute column name mapping to avoid repeated lookups
        Map<Integer, String> columnIndexToKey = new HashMap<>();
        for (int i = 1; i <= columnCount; i++) {
            String columnName = metaData.getColumnName(i);
            for (String key : columns.keySet()) {
                if (key.equalsIgnoreCase(columnName)) {
                    columnIndexToKey.put(i, key);
                    break;
                }
            }
        }

        // Create date formatter once
        DateFormat dateFormatter = new SimpleDateFormat("dd/MM/yyyy");

        // Process in batches
        int batchSize = 1000;
        int rowCount = 0;
        int rowIndex = 8;
        while (resultSet.next()) {
            Map<String, Object> properties = new LinkedHashMap<>(columns.size());

            // Initialize properties with empty values only for relevant columns
            for (String key : columns.keySet()) {
                if (!key.equalsIgnoreCase("STT") && !key.equalsIgnoreCase("TT")) {
                    properties.put(key, "");
                }
            }

            // Direct column mapping using pre-computed index
            for (int i = 1; i <= columnCount; i++) {
                String key = columnIndexToKey.get(i);
                if (key != null) {
                    Object value = resultSet.getObject(i);
                    if (value != null) {
                        if ("DATE".equals(columns.get(key))) {
                            value = dateFormatter.format(resultSet.getDate(i));
                        }
                        properties.put(key, value);
                    }
                }
            }

            resultList.add(new DynamicObject(properties, columns));
            generateExcel(rowIndex, resultList, true, Integer.parseInt(MAX_NUM_ROWS));
            rowIndex += 1;
            resultList.clear();
            // Log only every 1000 records
            if (rowIndex >= 600000) {
                rowIndex = 2;
            }
            if (rowIndex % batchSize == 0) {
                System.out.println("Processed " + rowIndex + " records");
            }
        }

        return resultList;
    }

    public void writeFooter(Sheet sheet, int rowIndex) {
        Row row = sheet.createRow(rowIndex);
        CellStyle cellStyle = sheet.getWorkbook().createCellStyle();
        Font font = sheet.getWorkbook().createFont();
        font.setBold(true);
        cellStyle.setFont(font);
        cellStyle.setWrapText(true);
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        row.createCell(2).setCellValue("LẬP BẢNG");
        row.createCell(8).setCellValue("NGƯỜI KIỂM SOÁT");
        row.createCell(13).setCellValue("ĐẠI DIỆN CHI NHÁNH");
        row.getCell(2).setCellStyle(cellStyle);
        row.getCell(8).setCellStyle(cellStyle);
        row.getCell(13).setCellStyle(cellStyle);
    }

    public void writeResultSetToCSV(ReportParameters reportParameters) throws SQLException, IOException {
        ResultSet rs = reportParameters.getResultSet();
        String baseFileName = reportParameters.getBaseFileName();
        Map<String, String> columns = reportParameters.getColumns();
        String[] groupHeaders = reportParameters.getGroupHeaders();
        String reportName = reportParameters.getReportName();
        String branchCode = reportParameters.getBranchCode();
        String branchName = reportParameters.getBranchName();
        String reportCode = reportParameters.getReportCode();
        Date reportDate = reportParameters.getReportDate();

        int currentFileIndex = 1;
        int rowCount = 0;
        CSVWriter writer = null;
        int ind = 0;

        try {
            while (rs.next()) {
                if (ind % 1000 == 0) {
                    System.out.println("row = " + ind);
                }
                ind += 1;

                if (writer == null || rowCount % ROWS_PER_FILE == 0) {
                    if (writer != null) {
                        writer.close();
                    }

                    String fileName = baseFileName.replace(".csv", String.format("_%d.csv", currentFileIndex));

                    writer = new CSVWriter(new FileWriter(fileName));

                    writeReportHeaders(writer, reportName, branchCode, branchName, reportCode, reportDate);

                    Vector<String> columnsInMap = new Vector<>();
                    columnsInMap.add("STT"); // Thêm cột STT
                    for (String key : columns.keySet()) {
                        String value = columns.get(key);
                        columnsInMap.add(value);
                    }
                    String[] columnInMapList = columnsInMap.toArray(new String[0]);
                    writeHeaders(writer, columnInMapList,groupHeaders);

                    currentFileIndex++;
                }

                // Tạo mảng row có độ dài lớn hơn columns.size() để chứa STT
                String[] row = new String[columns.size() + 1];
                row[0] = String.valueOf(ind); // Gán giá trị STT ở cột đầu tiên
                int i = 1; // Bắt đầu từ 1 vì 0 đã dành cho STT
                for (String columnKey : columns.keySet()) {
                    String value = rs.getString(columnKey);
                    row[i] = value != null ? value : "";
                    i++;
                }
                writer.writeNext(row);
                rowCount++;
            }
        } finally {
            if (writer != null) {
                writer.close();
            }
        }
    }


    private void writeReportHeaders(CSVWriter writer, String reportName, String branchCode, String branchName, String reportCode, Date reportDate) {
        writer.writeNext(new String[]{reportName}); // Sử dụng tên báo cáo từ tham số
        writer.writeNext(new String[]{"Mã chi nhánh: ", branchCode, "", "", "", "", "", "", "", "Mã báo cáo: " + reportCode});
        writer.writeNext(new String[]{"Tên chi nhánh: ", branchName});

        // Format ngày báo cáo
        DateFormat dateFormat = new SimpleDateFormat("dd/MM/yyyy");
        String dateStr = dateFormat.format(reportDate); // Sử dụng ngày từ tham số
        writer.writeNext(new String[]{"", "", "", "", "", "", "", "", "", "Ngày báo cáo: " + dateStr});

        // Dòng trống
        writer.writeNext(new String[]{});
    }

    private void writeHeaders(CSVWriter writer,String[] columns,String[] headers) {

        String[] rows = new String[columns.length];


        // Fill subheaders for each group
        if( headers.length > 0){

            for (int i = 0; i < columns.length; i++) {
                if(!Objects.equals(headers[i], "")){
                    rows[i] = headers[i]  + " ( "+columns[i]+ " ) ";
                }else{
                    rows[i] = columns[i];
                }
            }
        }else{
            for (int i = 0; i < columns.length; i++) {
                rows[i] = columns[i];
            }
        }

        writer.writeNext(rows);
    }
    public void writeExcel(String fileName) throws FileNotFoundException, IOException {
        // if directory not exist, create it
        if (fileName.lastIndexOf("/") > 0) {
            String directory = fileName.substring(0, fileName.lastIndexOf("/"));
            java.io.File directoryFile = new java.io.File(directory);
            if (!directoryFile.exists()) {
                directoryFile.mkdirs();
            }
        }
        try (FileOutputStream fos = new FileOutputStream(fileName)) {
            workbook.write(fos);
        } catch (Exception e) {
            System.out.println(e);
        } finally {
            workbook.close();
        }
    }

    private void createHeaderRow(Sheet sheet, Map<String, String> columns, int rowIndex) {
        Row headerRow = sheet.createRow(rowIndex);
        CellStyle cellStyle = sheet.getWorkbook().createCellStyle();
        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setBorderTop(BorderStyle.THIN);
        cellStyle.setBorderRight(BorderStyle.THIN);
        cellStyle.setBorderLeft(BorderStyle.THIN);
        // bold font, wrap text, middle alignment
        Font font = sheet.getWorkbook().createFont();
        font.setBold(true);
        cellStyle.setFont(font);
        cellStyle.setWrapText(true);
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        headerRow.createCell(0).setCellValue("STT");
        headerRow.getCell(0).setCellStyle(cellStyle);
        // sheet.autoSizeColumn(0);
        int colIndex = 1;
        for (String columnName : columns.values()) {
            Cell cell = headerRow.createCell(colIndex++);
            // sheet.autoSizeColumn(colIndex);
            cell.setCellValue(columnName);
            cell.setCellStyle(cellStyle);
        }
    }

    private void populateRowWithObjectData(Row row, Map<String, Object> properties, CellStyle cellStyle) {
        int colIndex = 1;
        for (Object value : properties.values()) {
            Cell cell = row.createCell(colIndex++);
            cell.setCellStyle(cellStyle);
            if (value instanceof String) {
                cell.setCellValue((String) value);
            } else if (value instanceof Number) {
                cell.setCellValue(((Number) value).doubleValue());
            } else if (value instanceof Boolean) {
                cell.setCellValue((Boolean) value);
            } else {
                cell.setCellValue(value != null ? value.toString() : ""); // Fallback for unknown types
            }
        }
    }

    private void copyLastHeaderWithMerges(Sheet destinationSheet, Map<String, String> columns, int rowIndex) {
        Row headerRow = destinationSheet.createRow(rowIndex);
        CellStyle cellStyle = destinationSheet.getWorkbook().createCellStyle();
        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setBorderTop(BorderStyle.THIN);
        cellStyle.setBorderRight(BorderStyle.THIN);
        cellStyle.setBorderLeft(BorderStyle.THIN);

        // Apply bold, center, wrap text, etc.
        Font font = destinationSheet.getWorkbook().createFont();
        font.setBold(true);
        cellStyle.setFont(font);
        cellStyle.setWrapText(true);
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        // Create "STT" column
        headerRow.createCell(0).setCellValue("STT");
        headerRow.getCell(0).setCellStyle(cellStyle);

        int colIndex = 1;
        for (String columnName : columns.values()) {
            Cell cell = headerRow.createCell(colIndex++);
            cell.setCellValue(columnName);
            cell.setCellStyle(cellStyle);
        }

        // Copy any merged regions from the original header row (assumed as rowIndex - 1)
        for (int i = 0; i < destinationSheet.getNumMergedRegions(); i++) {
            CellRangeAddress mergedRegion = destinationSheet.getMergedRegion(i);
            if (mergedRegion.getFirstRow() == rowIndex) {
                destinationSheet.addMergedRegion(new CellRangeAddress(
                    rowIndex, rowIndex, mergedRegion.getFirstColumn(), mergedRegion.getLastColumn()));
            }
        }
    }
}
