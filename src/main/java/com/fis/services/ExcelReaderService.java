package com.fis.services;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.fis.domain.Branch;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class ExcelReaderService {

    // Method to read an Excel file and extract branch data
    public static List<Branch> readBranchExcel(String excelFilePath) {
        List<Branch> branches = new ArrayList<>();

        try (FileInputStream fis = new FileInputStream(new File(excelFilePath));
                Workbook workbook = new XSSFWorkbook(fis)) {

            // Assuming data is on the first sheet
            Sheet sheet = workbook.getSheetAt(0);

            // Iterate over the rows, skipping the header row (starting at row 1)
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                // Check if the row is not empty and Collumn M is not 0
                if (row != null && !getCellValueAsString(row.getCell(12)).equalsIgnoreCase("0")) {
                    // Read each cell for branchCode, branchName, and folderPath
                    String branchCode = getCellValueAsString(row.getCell(5)); // Column F
                    String branchName = getCellValueAsString(row.getCell(4)); // Column E
                    String folderPath = getCellValueAsString(row.getCell(11)); // Column L
                    String[] reports; // Column N
                    if (!getCellValueAsString(row.getCell(13)).contains(";")) {
                        reports = new String[0];
                    } else {
                        reports = getCellValueAsString(row.getCell(13)).split(";");
                    }

                    // Create Branch object and add to list
                    Branch branch = new Branch(branchCode, branchName, folderPath,reports);
                    branches.add(branch);
                }
            }
        } catch (IOException e) {
            e.printStackTrace();
        }

        return branches;
    }

    // Helper method to get the cell value as a String, handling different cell
    // types
    private static String getCellValueAsString(Cell cell) {
        if (cell == null)
            return "";

        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue().toString();
                } else {
                    return String.valueOf((int) cell.getNumericCellValue());
                }
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                return cell.getCellFormula();
            default:
                return "";
        }
    }
}
