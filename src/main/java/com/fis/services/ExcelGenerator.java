package com.fis.services;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelGenerator {

    private Workbook workbook;

    public ExcelGenerator() {
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
        //     System.out.println("autosize column " + i);
        //     sheet.autoSizeColumn(i);
        // }
        // System.out.println("End autosize");

        for (DynamicObject dynamicObject : dataList) {
            System.out.println(stt);
            System.out.println(dynamicObject.getProperties());
            Row row = sheet.createRow(_rowIndex++);
            row.createCell(0).setCellValue(stt++);
            row.getCell(0).setCellStyle(cellStyle);
            populateRowWithObjectData(row, dynamicObject.getProperties(), cellStyle);
            // if (stt > dataList.size() -10000) {
            //     Thread.sleep(200);
            // }
        }
        return sheet;
        // try (FileOutputStream fos = new FileOutputStream(fileName)) {
        // workbook.write(fos);
        // } finally {
        // workbook.close();
        // }

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
        }catch(Exception e){
            System.out.println(e);
        }
        finally {
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
        sheet.autoSizeColumn(0);
        int colIndex = 1;
        for (String columnName : columns.values()) {
            Cell cell = headerRow.createCell(colIndex++);
            sheet.autoSizeColumn(colIndex);
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
}
