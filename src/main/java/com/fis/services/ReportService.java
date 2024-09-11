package com.fis.services;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.sql.SQLException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import com.fis.domain.Branch;

public class ReportService {

    private static DatabaseService databaseService = new DatabaseService();

    public static void CMS018Report() throws FileNotFoundException, IOException {
        String fileName = "CMS018Report.xlsx";
        System.out.println("Generating CMS018 Report: " + fileName);
        // new dynacmic object
        DynamicObject dynamicObject = new DynamicObject();
        HashMap<String, String> columns = new LinkedHashMap<>();
        columns.put("date", "Ngày dữ liệu");
        columns.put("branch", "CN PHT");
        columns.put("account", "Số tài khoản thẻ");
        columns.put("card_nbr", "Số thẻ");
        columns.put("emboss_name", "Tên chủ thẻ");
        columns.put("cred_limit", "HMTD");
        columns.put("tongduno", "Dư nợ thẻ");
        columns.put("stm_mindue", "Số tiền cần thanh toán tối thiểu");
        columns.put("stm_expirydue", "Ngày hết hạn thanh toán");
        columns.put("cash_advce", "Dư nợ Cash");
        columns.put("sale_advce", "Dư nợ Sale");
        columns.put("cash_adfees", "Tổng phí giao dịch rút tiền trong kỳ");
        columns.put("other_fees", "Tổng phí giao dịch khác trong kỳ");
        columns.put("some_fees", "Phí phạt chậm thanh toán trong kỳ");
        columns.put("total_cash_adv", "Tổng tiền lãi giao dịch rút tiền trong kỳ");
        columns.put("total_other_adv", "Tổng tiền lãi giao dịch khác trong kỳ");
        dynamicObject.setColumns(columns);

        // new dynamic object
        Map<String, Object> properties = new LinkedHashMap<>();
        properties.put("date", "2021-09-01");
        properties.put("branch", "CN PHT");
        properties.put("account", "1234567890");
        properties.put("card_nbr", "1234567890");
        properties.put("emboss_name", "Nguyen Van A");
        properties.put("cred_limit", 100000000);
        properties.put("tongduno", 50000000);
        properties.put("stm_mindue", 5000000);
        properties.put("stm_expirydue", "2021-09-01");
        properties.put("cash_advce", 10000000);
        properties.put("sale_advce", 10000000);
        properties.put("cash_adfees", 100000);
        properties.put("other_fees", 100000);
        properties.put("some_fees", 100000);
        properties.put("total_cash_adv", 10000);
        properties.put("total_other_adv", 10000);
        dynamicObject.setProperties(properties);

        List<DynamicObject> dynamicObjects = new ArrayList<>();
        dynamicObjects.add(dynamicObject);

        ExcelGenerator excelGenerator = new ExcelGenerator();
        Sheet sheet = excelGenerator.generateExcel(6, dynamicObjects, false);
        // title row 0
        Row titleRow = sheet.createRow(0);
        titleRow.createCell(0)
                .setCellValue("BÁO CÁO CHI TIẾT THÔNG TIN DƯ NỢ, GỐC, LÃI THẺ TDQT TRƯỚC KHI CHUYỂN ĐỔI ");
        // font bold, size 16 for title
        Font font = sheet.getWorkbook().createFont();
        font.setBold(true);
        font.setFontHeightInPoints((short) 16);
        CellStyle style = sheet.getWorkbook().createCellStyle();
        style.setFont(font);
        style.setAlignment(HorizontalAlignment.CENTER);
        titleRow.getCell(0).setCellStyle(style);
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 16));

        Font fontBold = sheet.getWorkbook().createFont();
        fontBold.setBold(true);
        CellStyle styleBold = sheet.getWorkbook().createCellStyle();
        styleBold.setFont(fontBold);

        // date now format dd/MM/yyyy
        Date date = new Date();

        // header row 1
        Row row1 = sheet.createRow(1);
        row1.createCell(1).setCellValue("Mã chi nhánh");
        row1.getCell(1).setCellStyle(styleBold);
        // header row 2
        Row row2 = sheet.createRow(2);
        row2.createCell(1).setCellValue("Tên chi nhánh");
        row2.createCell(15).setCellValue("Mã bc: CMS_018");
        row2.getCell(1).setCellStyle(styleBold);
        row2.getCell(15).setCellStyle(styleBold);

        Row row4 = sheet.createRow(4);
        row4.createCell(5).setCellValue("Ngày báo cáo: " + date);
        row4.getCell(5).setCellStyle(styleBold);

        int endRow = 6 + dynamicObjects.size() + 2;
        Row eRow = sheet.createRow(endRow);
        eRow.createCell(3).setCellValue("LẬP BẢNG");
        eRow.createCell(9).setCellValue("NGƯỜI KIỂM SOÁT");
        eRow.createCell(14).setCellValue("ĐẠI DIỆN CHI NHÁNH");
        eRow.getCell(3).setCellStyle(styleBold);
        eRow.getCell(9).setCellStyle(styleBold);
        eRow.getCell(14).setCellStyle(styleBold);

        excelGenerator.writeExcel(fileName);

    }

    public static void ISS011Report(Branch branch) throws FileNotFoundException, IOException, SQLException {
        System.out.println("Generating ISS011 Report");
        Date date = new Date();
        DateFormat dateFNFormat = new SimpleDateFormat("ddMMyyyy");
        String dateFN = dateFNFormat.format(date);
        String fileName = branch.getFolderPath() + "ISS_011_" + branch.getBranchCode() + "_" + branch.getBranchName()
                + "_" + dateFN
                + ".xlsx";
        DynamicObject dynamicObject = new DynamicObject();
        Map<String, String> columns = new LinkedHashMap<>();
        columns.put("acct_style", "Loại thẻ");
        columns.put("branch", "CN quản lý");
        columns.put("custr_ref", "Số CIF");
        columns.put("account", "Số tài khoản thẻ");
        columns.put("acc_name1", "Tên tài khoản");
        columns.put("recla_code", "Trạng thái tài khoản");
        columns.put("card_nbr", "Số thẻ");
        columns.put("cancl_code", "Trạng thái thẻ");
        columns.put("product", "Sản phẩm thẻ");
        columns.put("reference", "AM");
        columns.put("cred_limit", "HMTD");
        columns.put("class_code", "ClassCode");
        columns.put("int_code", "Interest Code");
        columns.put("fee_code", "Feecode");
        columns.put("xrefacct_a", "ID TSBĐ (nếu có)");
        columns.put("dpd", "Số ngày quá hạn");
        columns.put("cycle_nbr", "Ngày sao kê");
        columns.put("address1", "Địa chỉ nhận sao kê");
        columns.put("tk1", "TK1 liên kết đến thẻ");
        columns.put("tk2", "TK2 liên kết đến thẻ");
        columns.put("tk3", "TK3 liên kết đến thẻ");
        columns.put("tk4", "TK4 liên kết đến thẻ");
        columns.put("tk5", "TK5 liên kết đến thẻ");
        columns.put("tk6", "TK6 liên kết đến thẻ");
        columns.put("tk7", "TK7 liên kết đến thẻ");
        columns.put("tk8", "TK8 liên kết đến thẻ");
        columns.put("tk9", "TK9 liên kết đến thẻ");
        columns.put("tk10", "TK10 liên kết đến thẻ");
        columns.put("repay_code", "Tỷ lệ trích nợ tự động");
        columns.put("tktntd", "Tài khoản trích nợ tự động");
        columns.put("tongduno", "Dư nợ thẻ");
        columns.put("query_amt", "Số tiền khiếu nại đang treo");
        columns.put("SUM(mp.orig_purch)", "Tổng số dư giao dịch trả góp");
        columns.put("COUNT(mp.status)", "Tổng số giao dịch trả góp đang hoạt động");
        columns.put("AUTHS_AMT", "Tổng giá trị các giao dịch chưa lên dư nợ");
        columns.put("FEE_SPENDA", "Doanh số chi tiêu lũy kế");
        dynamicObject.setColumns(columns);

        Map<Integer, Object> inputParams = new HashMap<>();
        inputParams.put(1, branch.getBranchCode());
        List<Integer> outParams = new ArrayList<>();
        outParams.add(2);
        List<DynamicObject> dynamicObjects =
        databaseService.callProcedure("REPORT_MIGRATE", "ISS_011", columns,
        inputParams, outParams);
        // List<DynamicObject> dynamicObjects = new ArrayList<>();
        if (dynamicObjects.size() == 0) {
            DynamicObject dynamicObject1 = new DynamicObject();
            dynamicObject1.setColumns(columns);
            dynamicObject1.setProperties(new LinkedHashMap<>());
            dynamicObjects.add(dynamicObject1);
        }
        ExcelGenerator excelGenerator = new ExcelGenerator();
        Sheet sheet = excelGenerator.generateExcel(6, dynamicObjects, false);
        // title row 0
        Row titleRow = sheet.createRow(0);
        titleRow.createCell(0)
                .setCellValue("BÁO CÁO CHI TIẾT THÔNG TIN THẺ TRƯỚC KHI CHUYỂN ĐỔI");
        // font bold, size 16 for title
        Font font = sheet.getWorkbook().createFont();
        font.setBold(true);
        font.setFontHeightInPoints((short) 16);
        CellStyle style = sheet.getWorkbook().createCellStyle();
        style.setFont(font);
        style.setAlignment(HorizontalAlignment.CENTER);
        titleRow.getCell(0).setCellStyle(style);
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 36));

        Font fontBold = sheet.getWorkbook().createFont();
        fontBold.setBold(true);
        CellStyle styleBold = sheet.getWorkbook().createCellStyle();
        styleBold.setFont(fontBold);

        // date now format dd/MM/yyyy
        DateFormat dateFormat = new SimpleDateFormat("dd/MM/yyyy");
        String dateStr = dateFormat.format(date);

        // header row 1
        Row row1 = sheet.createRow(1);
        row1.createCell(1).setCellValue("Mã chi nhánh: " + branch.getBranchCode());
        row1.getCell(1).setCellStyle(styleBold);
        // header row 2
        Row row2 = sheet.createRow(2);
        row2.createCell(1).setCellValue("Tên chi nhánh: " + branch.getBranchName());
        row2.createCell(15).setCellValue("Mã bc: ISS_011");
        row2.getCell(1).setCellStyle(styleBold);
        row2.getCell(15).setCellStyle(styleBold);

        Row row4 = sheet.createRow(4);
        row4.createCell(5).setCellValue("Ngày báo cáo: " + dateStr);
        row4.getCell(5).setCellStyle(styleBold);

        int endRow = 6 + dynamicObjects.size() + 2;
        Row eRow = sheet.createRow(endRow);
        eRow.createCell(3).setCellValue("LẬP BẢNG");
        eRow.createCell(9).setCellValue("NGƯỜI KIỂM SOÁT");
        eRow.createCell(14).setCellValue("ĐẠI DIỆN CHI NHÁNH");
        eRow.getCell(3).setCellStyle(styleBold);
        eRow.getCell(9).setCellStyle(styleBold);
        eRow.getCell(14).setCellStyle(styleBold);

        excelGenerator.writeExcel(fileName);

    }

    public static void ISS012Report(Branch branch) throws FileNotFoundException, IOException, SQLException {
        System.out.println("Generating ISS012 Report");
        Date date = new Date();
        DateFormat dateFNFormat = new SimpleDateFormat("ddMMyyyy");
        String dateFN = dateFNFormat.format(date);
        String fileName = branch.getFolderPath() + "ISS_012_" + branch.getBranchCode() + "_" + branch.getBranchName()
                + "_" + dateFN
                + ".xlsx";
        DynamicObject dynamicObject = new DynamicObject();
        Map<String, String> columns = new LinkedHashMap<>();
        columns.put("acct_style", "Loại thẻ");
        columns.put("branch", "CN quản lý");
        columns.put("custr_ref", "Số CIF");
        columns.put("account", "Số tài khoản thẻ");
        columns.put("acc_name1", "Tên tài khoản");
        columns.put("recla_code", "Trạng thái tài khoản");
        columns.put("card_nbr", "Số thẻ");
        columns.put("cancl_code", "Trạng thái thẻ");
        columns.put("product", "Sản phẩm thẻ");
        columns.put("reference", "AM");
        columns.put("cred_limit", "HMTD");
        columns.put("class_code", "ClassCode");
        columns.put("int_code", "Interest Code");
        columns.put("fee_code", "Feecode");
        columns.put("xrefacct_a", "ID TSBĐ (nếu có)");
        columns.put("dpd", "Số ngày quá hạn");
        columns.put("cycle_nbr", "Ngày sao kê");
        columns.put("address1", "Địa chỉ nhận sao kê");
        columns.put("tk1", "TK1 liên kết đến thẻ");
        columns.put("tk2", "TK2 liên kết đến thẻ");
        columns.put("tk3", "TK3 liên kết đến thẻ");
        columns.put("tk4", "TK4 liên kết đến thẻ");
        columns.put("tk5", "TK5 liên kết đến thẻ");
        columns.put("tk6", "TK6 liên kết đến thẻ");
        columns.put("tk7", "TK7 liên kết đến thẻ");
        columns.put("tk8", "TK8 liên kết đến thẻ");
        columns.put("tk9", "TK9 liên kết đến thẻ");
        columns.put("tk10", "TK10 liên kết đến thẻ");
        columns.put("repay_code", "Tỷ lệ trích nợ tự động");
        columns.put("tktntd", "Tài khoản trích nợ tự động");
        columns.put("tongduno", "Dư nợ thẻ");
        columns.put("query_amt", "Số tiền khiếu nại đang treo");
        columns.put("SUM(mp.orig_purch)", "Tổng số dư giao dịch trả góp");
        columns.put("COUNT(mp.status)", "Tổng số giao dịch trả góp đang hoạt động");
        columns.put("AUTHS_AMT", "Tổng giá trị các giao dịch chưa lên dư nợ");
        columns.put("FEE_SPENDA", "Tổng Doanh số thanh toán");
        dynamicObject.setColumns(columns);

        Map<Integer, Object> inputParams = new HashMap<>();
        inputParams.put(1, branch.getBranchCode());
        List<Integer> outParams = new ArrayList<>();
        outParams.add(2);

        List<DynamicObject> dynamicObjects =
        databaseService.callProcedure("REPORT_MIGRATE", "ISS_012", columns,
        inputParams, outParams);

        // List<DynamicObject> dynamicObjects = new ArrayList<>();
        if (dynamicObjects.size() == 0) {
            DynamicObject dynamicObject1 = new DynamicObject();
            dynamicObject1.setColumns(columns);
            dynamicObject1.setProperties(new LinkedHashMap<>());
            dynamicObjects.add(dynamicObject1);
        }

        ExcelGenerator excelGenerator = new ExcelGenerator();
        Sheet sheet = excelGenerator.generateExcel(6, dynamicObjects, false);
        // title row 0
        Row titleRow = sheet.createRow(0);
        titleRow.createCell(0)
                .setCellValue("BÁO CÁO CHI TIẾT THÔNG TIN THẺ KHÔNG CHUYỂN ĐỔI");
        // font bold, size 16 for title
        Font font = sheet.getWorkbook().createFont();
        font.setBold(true);
        font.setFontHeightInPoints((short) 16);
        CellStyle style = sheet.getWorkbook().createCellStyle();
        style.setFont(font);
        style.setAlignment(HorizontalAlignment.CENTER);
        titleRow.getCell(0).setCellStyle(style);
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 36));

        Font fontBold = sheet.getWorkbook().createFont();
        fontBold.setBold(true);
        CellStyle styleBold = sheet.getWorkbook().createCellStyle();
        styleBold.setFont(fontBold);

        // date now format dd/MM/yyyy

        DateFormat dateFormat = new SimpleDateFormat("dd/MM/yyyy");
        String dateStr = dateFormat.format(date);

        // header row 1
        Row row1 = sheet.createRow(1);
        row1.createCell(1).setCellValue("Mã chi nhánh: " + branch.getBranchCode());
        row1.getCell(1).setCellStyle(styleBold);
        // header row 2
        Row row2 = sheet.createRow(2);
        row2.createCell(1).setCellValue("Tên chi nhánh: " + branch.getBranchName());
        row2.createCell(15).setCellValue("Mã bc: ISS_012");
        row2.getCell(1).setCellStyle(styleBold);
        row2.getCell(15).setCellStyle(styleBold);

        Row row4 = sheet.createRow(4);
        row4.createCell(5).setCellValue("Ngày báo cáo: " + dateStr);
        row4.getCell(5).setCellStyle(styleBold);

        int endRow = 6 + dynamicObjects.size() + 2;
        Row eRow = sheet.createRow(endRow);
        eRow.createCell(3).setCellValue("LẬP BẢNG");
        eRow.createCell(9).setCellValue("NGƯỜI KIỂM SOÁT");
        eRow.createCell(14).setCellValue("ĐẠI DIỆN CHI NHÁNH");
        eRow.getCell(3).setCellStyle(styleBold);
        eRow.getCell(9).setCellStyle(styleBold);
        eRow.getCell(14).setCellStyle(styleBold);

        excelGenerator.writeExcel(fileName);

    }

    public static void ATM002REPORT() throws FileNotFoundException, IOException {
        System.out.println("Generating ATM002 Report");
        Date date = new Date();
        DateFormat dateFNFormat = new SimpleDateFormat("ddMMyyyy");
        String dateFN = dateFNFormat.format(date);
        String fileName = "/FTPData/HSC/ATM_002_" + dateFN + ".xlsx";
        System.out.println("Generating ATM002 Report: " + fileName);
        DynamicObject dynamicObject = new DynamicObject();
        Map<String, String> columns = new LinkedHashMap<>();
        columns.put("TERMID", "Terminal ID");
        columns.put("GROUP_NAME", "Branch quản lý");
        columns.put("GROUP_CRED", "Branch tiếp quỹ");
        columns.put("TMS", "Mã AM quản lý máy");
        columns.put("ATM TYPE", "ATM TYPE");
        columns.put("Hãng ATM", "Hãng ATM");
        columns.put("Model", "Model");
        columns.put("DE43", "ATM location");
        columns.put("ATMBAL.STARTBALCASS1", "Mệnh giá hộp tiền 1");
        columns.put("ATMBAL.STARTBALCASS2", "Mệnh giá hộp tiền 2");
        columns.put("ATMBAL.STARTBALCASS3", "Mệnh giá hộp tiền 3");
        columns.put("ATMBAL.STARTBALCASS4", "Mệnh giá hộp tiền 4");
        columns.put("ATM group", "ATM group");
        columns.put("visa", "%Phí Visa off us nước ngoài");
        columns.put("visa_min", "Phí visa off us nước ngoài Min");
        columns.put("visa_off", "%Phí MC off us nước ngoài");
        columns.put("visa_off_min", "%Phí MC off us nước ngoài Min");
        dynamicObject.setColumns(columns);

        Map<String, Object> properties = new LinkedHashMap<>();
        properties.put("TERMID", "1234567890");
        properties.put("GROUP_NAME", "Branch quản lý");
        properties.put("GROUP_CRED", "Branch tiếp quỹ");
        properties.put("TMS", "Mã AM quản lý máy");
        properties.put("ATM TYPE", "ATM TYPE");
        properties.put("Hãng ATM", "Hãng ATM");
        properties.put("Model", "Model");
        properties.put("DE43", "ATM location");
        properties.put("ATMBAL.STARTBALCASS1", 100000000);
        properties.put("ATMBAL.STARTBALCASS2", 100000000);
        properties.put("ATMBAL.STARTBALCASS3", 100000000);
        properties.put("ATMBAL.STARTBALCASS4", 100000000);
        properties.put("ATM group", "ATM group");
        properties.put("visa", 0.1);
        properties.put("visa_min", 100000);
        properties.put("visa_off", 0.1);
        properties.put("visa_off_min", 100000);
        dynamicObject.setProperties(properties);

        List<DynamicObject> dynamicObjects = new ArrayList<>();
        dynamicObjects.add(dynamicObject);

        ExcelGenerator excelGenerator = new ExcelGenerator();
        Sheet sheet = excelGenerator.generateExcel(3, dynamicObjects, false);
        // title row 0
        Row code = sheet.createRow(0);
        code.createCell(0)
                .setCellValue("Mã báo cáo: ATM_002");
        // title row 1
        Row titleRow = sheet.createRow(1);
        titleRow.createCell(0)
                .setCellValue("Báo cáo chi tiết trước chuyển đổi");
        // font bold, size 16 for title
        Font font = sheet.getWorkbook().createFont();
        font.setBold(true);
        font.setFontHeightInPoints((short) 16);
        CellStyle style = sheet.getWorkbook().createCellStyle();
        style.setFont(font);
        style.setAlignment(HorizontalAlignment.CENTER);
        titleRow.getCell(0).setCellStyle(style);
        sheet.addMergedRegion(new CellRangeAddress(1, 1, 0, 8));

        excelGenerator.writeExcel(fileName);

    }

    public static void ACQ009Report(Branch branch) throws FileNotFoundException, IOException {
        Date date = new Date();
        DateFormat dateFNFormat = new SimpleDateFormat("ddMMyyyy");
        String dateFN = dateFNFormat.format(date);
        String fileName = branch.getFolderPath() + "ACQ_009_" + branch.getBranchCode() + "_" + branch.getBranchName()
                + "_" + dateFN
                + ".xlsx";
        DynamicObject dynamicObject = new DynamicObject();
        Map<String, String> columns = new LinkedHashMap<>();
        columns.put("BRANCH", "CN quản lý");
        columns.put("BRCD", "Số Cif khách hàng");
        columns.put("desc", "Tên khách hàng");
        columns.put("IMP_NBR", "Merchant main ID");
        columns.put("OWN_MERCH", "Merchant liên kết");
        columns.put("MERCHANT", "Merchant ID");
        columns.put("TERM_NBR", "Terminal ID");
        columns.put("MCC", "MCC");
        columns.put("AM", "AM");
        columns.put("COMM_ACCT", "Dep_acct");
        columns.put("MC", "MC");
        columns.put("MD", "MD");
        columns.put("VC", "VC");
        columns.put("VD", "VD");
        columns.put("JC", "JC");
        columns.put("JD", "JD");
        columns.put("PD", "PD");
        columns.put("PC", "PC");
        dynamicObject.setColumns(columns);

        Map<String, Object> properties = new LinkedHashMap<>();
        properties.put("BRANCH", "CN quản lý");
        properties.put("BRCD", "Số Cif khách hàng");
        properties.put("desc", "Tên khách hàng");
        properties.put("IMP_NBR", "Merchant main ID");
        properties.put("OWN_MERCH", "Merchant liên kết");
        properties.put("MERCHANT", "Merchant ID");
        properties.put("TERM_NBR", "Terminal ID");
        properties.put("MCC", "MCC");
        properties.put("AM", "AM");
        properties.put("COMM_ACCT", "Dep_acct");
        properties.put("MC", "MC");
        properties.put("MD", "MD");
        properties.put("VC", "VC");
        properties.put("VD", "VD");
        properties.put("JC", "JC");
        properties.put("JD", "JD");
        properties.put("PD", "PD");
        properties.put("PC", "PC");
        dynamicObject.setProperties(properties);

        List<DynamicObject> dynamicObjects = new ArrayList<>();
        dynamicObjects.add(dynamicObject);

        ExcelGenerator excelGenerator = new ExcelGenerator();
        Sheet sheet = excelGenerator.generateExcel(3, dynamicObjects, true);

        // title row 0
        // title row 0
        Row titleRow = sheet.createRow(0);
        titleRow.createCell(2)
                .setCellValue(
                        "Báo cáo chi tiết Đơn vị chấp nhận thẻ và POS trước chuyển đổi đủ điều kiện chuyển đổi theo chi nhánh");
        Row code = sheet.createRow(2);
        code.createCell(0)
                .setCellValue("Mã báo cáo: ACQ_009");
        // title row 1
        Row dateRow = sheet.createRow(3);
        DateFormat dateFormat = new SimpleDateFormat("dd/MM/yyyy");
        dateRow.createCell(0)
                .setCellValue("Ngày báo cáo: " + dateFormat.format(date));
        Row branchRow = sheet.createRow(1); // Changed index from 0 to 1
        branchRow.createCell(4)
                .setCellValue("Chi nhánh: ");
        // font bold, size 16 for title
        Font font = sheet.getWorkbook().createFont();
        font.setBold(true);
        font.setFontHeightInPoints((short) 16);
        CellStyle style = sheet.getWorkbook().createCellStyle();
        style.setFont(font);
        style.setAlignment(HorizontalAlignment.CENTER);
        titleRow.getCell(2).setCellStyle(style);
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 2, 16));

        Row headerRow = sheet.createRow(6);
        Row headerRow2 = sheet.createRow(7);
        CellStyle cellStyle = sheet.getWorkbook().createCellStyle();
        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setBorderTop(BorderStyle.THIN);
        cellStyle.setBorderRight(BorderStyle.THIN);
        cellStyle.setBorderLeft(BorderStyle.THIN);
        // bold font, wrap text, middle alignment
        Font font2 = sheet.getWorkbook().createFont();
        font2.setBold(true);
        cellStyle.setFont(font2);
        cellStyle.setWrapText(true);
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        headerRow.createCell(0).setCellValue("STT");
        headerRow.getCell(0).setCellStyle(cellStyle);
        // Cell mergeCell = headerRow.getCell(0);
        // mergeCell.setCellValue("Phí MDR áp dụng VND");
        // mergeCell.setCellStyle(cellStyle);
        for (int i = 1; i < columns.size() - 8; i++) {
            headerRow.createCell(i).setCellValue((String) columns.values().toArray()[i]);
            headerRow.getCell(i).setCellStyle(cellStyle);
        }

        for (int i = columns.size() - 8; i < columns.size(); i++) {
            headerRow2.createCell(i).setCellValue((String) columns.values().toArray()[i]);
            if (headerRow.getCell(i) == null) {
                headerRow.createCell(i).setCellStyle(cellStyle);
                headerRow.getCell(i).setCellValue("Phí MDR áp dụng VND");
                ;
            } else {
                headerRow.createCell(i).setCellStyle(cellStyle);

            }
            headerRow2.getCell(i).setCellStyle(cellStyle);
        }
        // merge cell for header
        // headerRow.getCell(9).setCellStyle(cellStyle);
        sheet.addMergedRegion(new CellRangeAddress(6, 6, columns.size() - 8, columns.size() - 1));
        excelGenerator.writeExcel(fileName);
    }

    public static void GL007Report() throws FileNotFoundException, IOException {
        System.out.println("Generating GL007 Report");
        DynamicObject dynamicObject = new DynamicObject();
        Map<String, String> columns = new LinkedHashMap<>();
        columns.put("branch", "Chi nhánh");
        columns.put("tsc", "TSC");
        columns.put("account", "Số tài khoản");
        columns.put("acc_name", "Tên tài khoản");
        columns.put("user", "Người thực hiện");
        columns.put("chanel_id", "Kênh giao dịch");
        columns.put("date", "Ngày giao dịch");
        columns.put("terminal_id", "Mã thiết bị");
        columns.put("nbr", "Số giao dịch");
        columns.put("scc", "Mã SCC");
        columns.put("profile", "Hồ sơ giao dịch");
        columns.put("amount", "Số tiền");
        columns.put("content", "Nội dung");
        dynamicObject.setColumns(columns);

        // not fill properties for now
        dynamicObject.setProperties(new LinkedHashMap<>());

        List<DynamicObject> dynamicObjects = new ArrayList<>();
        dynamicObjects.add(dynamicObject);

        ExcelGenerator excelGenerator = new ExcelGenerator();
        Sheet sheet = excelGenerator.generateExcel(7, dynamicObjects, true);

        // title row 2
        Row titleRow = sheet.createRow(2);
        titleRow.createCell(0)
                .setCellValue("BÁO CÁO CHI TIẾT GIAO DỊCH CHUYỂN KHOẢN");
        // font bold, size 16 for title
        Font font = sheet.getWorkbook().createFont();
        font.setBold(true);
        font.setFontHeightInPoints((short) 16);
        CellStyle style = sheet.getWorkbook().createCellStyle();
        style.setFont(font);
        style.setAlignment(HorizontalAlignment.CENTER);
        titleRow.getCell(0).setCellStyle(style);
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 12));

        Font fontBold = sheet.getWorkbook().createFont();
        fontBold.setBold(true);
        CellStyle styleBold = sheet.getWorkbook().createCellStyle();
        styleBold.setFont(fontBold);

        // date now format dd/MM/yyyy
        Date date = new Date();
        DateFormat dateFormat = new SimpleDateFormat("dd/MM/yyyy");
        String dateStr = dateFormat.format(date);

        // header row 0
        Row row1 = sheet.createRow(0);
        row1.createCell(0).setCellValue("Mã chi nhánh: ");
        // header row 1
        Row row2 = sheet.createRow(1);
        row2.createCell(0).setCellValue("Tên chi nhánh: ");

        Row row4 = sheet.createRow(3);
        row4.createCell(0).setCellValue("Ngày báo cáo: " + dateStr);
        row4.getCell(0).setCellStyle(styleBold);
        row4.createCell(columns.size()).setCellValue("Mã BC: GL_007");

        Row row5 = sheet.createRow(4);
        row5.createCell(columns.size()).setCellValue("Loại tiền: ");

        sheet.addMergedRegion(new CellRangeAddress(2, 2, 0, 12));
        DateFormat dateFNFormat = new SimpleDateFormat("ddMMyyyy");
        String dateFN = dateFNFormat.format(date);
        String fileName = "/FTPData/HSC/GL_007_" + dateFN + ".xlsx";
        excelGenerator.writeExcel(fileName);
    }

    public static void GL005ISSReport() throws FileNotFoundException, IOException {
        System.out.println("Generating GL005_ISS Report");
        Date date = new Date();
        DateFormat dateFNFormat = new SimpleDateFormat("ddMMyyyy");
        String dateFN = dateFNFormat.format(date);
        String fileName = "/FTPData/HSC/GL_005_ISS_" + dateFN + ".xlsx";
        DynamicObject dynamicObject = new DynamicObject();
        Map<String, String> columns = new LinkedHashMap<>();
        columns.put("branch", "TT");
        columns.put("CN_CAD", "Mã CN CAD");
        columns.put("am", "MÃ CN 6 số");
        columns.put("plan", "Loại thẻ (phân loại theo TCT)");
        columns.put("nbr", "Số thẻ");
        columns.put("cif", "CIF");
        columns.put("zcomcode", "Mã KH (ZCOMCDE)");
        columns.put("customer", "Khách hàng");
        columns.put("nhomno", "Phân loại nợ");
        columns.put("dunotronghan", "Dư nợ trong hạn");
        columns.put("dunoquahan", "Dư nợ quá hạn");
        columns.put("lai", "Lãi");
        columns.put("phi", "Phí");
        columns.put("cong", "Cộng");
        columns.put("sodu", "Số dư tài khoản Cho vay phát hành thẻ trước thời điểm chuyển đổi");

        dynamicObject.setColumns(columns);

        // not fill properties for now
        dynamicObject.setProperties(new LinkedHashMap<>());

        List<DynamicObject> dynamicObjects = new ArrayList<>();
        dynamicObjects.add(dynamicObject);

        ExcelGenerator excelGenerator = new ExcelGenerator();
        Sheet sheet = excelGenerator.generateExcel(7, dynamicObjects, true);

        // row 0 mã chi nhánh, row 1 tên chi nhánh
        // row 1 title: BÁO CÁO TỔNG HỢP SỐ LIỆU CHUYỂN ĐỔI
        // row 2 mã báo cáo
        // row 3 ngày dữ liệu
        // row 4 loại tiền tệ

    }
}
