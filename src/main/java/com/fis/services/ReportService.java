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

import oracle.net.aso.q;

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
        columns.put("CLOSE_CODE", "Trạng thái tài khoản");
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
        // columns.put("query_amt", "Số tiền khiếu nại đang treo");
        columns.put("TONG_SO_DU_GD_TRAGOP", "Tổng số dư giao dịch trả góp");
        columns.put("TONG_SO_GD_TRAGOP_DANG_HOAT_DONG", "Tổng số giao dịch trả góp đang hoạt động");
        columns.put("TONG_GIATRI_GD_CHUA_LEN_DU_NO", "Tổng giá trị các giao dịch chưa lên dư nợ");
        columns.put("TONG_DOANH_SO_THANH_TOAN", "Tổng Doanh số thanh toán");
        dynamicObject.setColumns(columns);

        Map<Integer, Object> inputParams = new HashMap<>();
        inputParams.put(1, branch.getBranchCode());
        List<Integer> outParams = new ArrayList<>();
        outParams.add(2);
        List<DynamicObject> dynamicObjects = databaseService.callProcedure("REPORT_MIGRATE", "ISS_011", columns,
                inputParams, outParams);
        // List<DynamicObject> dynamicObjects = new ArrayList<>();
        if (dynamicObjects.size() == 0) {
            DynamicObject dynamicObject1 = new DynamicObject();
            dynamicObject1.setColumns(columns);
            dynamicObject1.setProperties(new LinkedHashMap<>());
            dynamicObjects.add(dynamicObject1);
        }
        // else {
        // dynamicObjects.forEach(dynamicObject1 -> {
        // if (dynamicObject1.getProperties().get("class_code") == null) {
        // dynamicObject1.getProperties().put("class_code", "");
        // }
        // });
        // }
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
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 35));

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
        columns.put("CLOSE_CODE", "Trạng thái tài khoản");
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
        // columns.put("query_amt", "Số tiền khiếu nại đang treo");
        columns.put("TONG_SO_DU_GD_TRAGOP", "Tổng số dư giao dịch trả góp");
        columns.put("TONG_SO_GD_TRAGOP_DANG_HOAT_DONG", "Tổng số giao dịch trả góp đang hoạt động");
        columns.put("TONG_GIATRI_GD_CHUA_LEN_DU_NO", "Tổng giá trị các giao dịch chưa lên dư nợ");
        columns.put("TONG_DOANH_SO_THANH_TOAN", "Tổng Doanh số thanh toán");
        columns.put("note", "Lý do không chuyển đổi");
        dynamicObject.setColumns(columns);

        Map<Integer, Object> inputParams = new HashMap<>();
        inputParams.put(1, branch.getBranchCode());
        List<Integer> outParams = new ArrayList<>();
        outParams.add(2);

        List<DynamicObject> dynamicObjects = databaseService.callProcedure("REPORT_MIGRATE", "ISS_012", columns,
                inputParams, outParams);

        dynamicObjects.forEach(dynamicObject1 -> {
            if (dynamicObject1.getProperties().get("note") == null) {
                dynamicObject1.getProperties().put("note", "");
            }
        });

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

    public static void ATM002REPORT() throws FileNotFoundException, IOException, SQLException {
        System.out.println("Generating ATM002 Report");
        Date date = new Date();
        DateFormat dateFNFormat = new SimpleDateFormat("ddMMyyyy");
        String dateFN = dateFNFormat.format(date);
        String fileName = "/FTPData/HSC/ATM_002_" + dateFN + ".xlsx";
        System.out.println("Generating ATM002 Report: " + fileName);
        DynamicObject dynamicObject = new DynamicObject();
        Map<String, String> columns = new LinkedHashMap<>();
        columns.put("ATM_ID", "Terminal ID");
        columns.put("zbrcd", "Branch quản lý");
        columns.put("ma_chi_nhanh_tq6", "Branch tiếp quỹ");
        columns.put("MA_AM", "Mã AM quản lý máy");
        columns.put("ATM_TYPE", "ATM TYPE");
        columns.put("HANG_ATM", "Hãng ATM");
        columns.put("ATM_MODEL", "Model");
        columns.put("ACCEPTORNAME", "ATM location");
        columns.put("STARTBALCASS1", "Mệnh giá hộp tiền 1");
        columns.put("STARTBALCASS2", "Mệnh giá hộp tiền 2");
        columns.put("STARTBALCASS3", "Mệnh giá hộp tiền 3");
        columns.put("STARTBALCASS4", "Mệnh giá hộp tiền 4");
        columns.put("ATM_GROUP", "ATM group");
        // columns.put("visa", "%Phí Visa off us nước ngoài");
        // columns.put("visa_min", "Phí visa off us nước ngoài Min");
        // columns.put("visa_off", "%Phí MC off us nước ngoài");
        // columns.put("visa_off_min", "%Phí MC off us nước ngoài Min");
        dynamicObject.setColumns(columns);

        Map<Integer, Object> inputParams = new HashMap<>();

        List<Integer> outParams = new ArrayList<>();
        outParams.add(1);

        List<DynamicObject> dynamicObjects = databaseService.callProcedure("REPORT_MIGRATE", "ATM_002", columns,
                inputParams, outParams);

        if (dynamicObjects.size() == 0) {
            DynamicObject dynamicObject1 = new DynamicObject();
            dynamicObject1.setColumns(columns);
            dynamicObject1.setProperties(new LinkedHashMap<>());
            dynamicObjects.add(dynamicObject1);
        }

        ExcelGenerator excelGenerator = new ExcelGenerator();
        Sheet sheet = excelGenerator.generateExcel(4, dynamicObjects, false);
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

        Font fontBold = sheet.getWorkbook().createFont();
        fontBold.setBold(true);
        CellStyle styleBold = sheet.getWorkbook().createCellStyle();
        styleBold.setFont(fontBold);
        // int endRow = 6 + dynamicObjects.size() + 2;
        // Row eRow = sheet.createRow(endRow);
        // eRow.createCell(3).setCellValue("LẬP BẢNG");
        // eRow.createCell(9).setCellValue("NGƯỜI KIỂM SOÁT");
        // eRow.createCell(14).setCellValue("ĐẠI DIỆN CHI NHÁNH");
        // eRow.getCell(3).setCellStyle(styleBold);
        // eRow.getCell(9).setCellStyle(styleBold);
        // eRow.getCell(14).setCellStyle(styleBold);

        excelGenerator.writeExcel(fileName);

    }

    public static void ACQ009Report(Branch branch) throws FileNotFoundException, IOException, SQLException {
        Date date = new Date();
        DateFormat dateFNFormat = new SimpleDateFormat("ddMMyyyy");
        String dateFN = dateFNFormat.format(date);
        String fileName = branch.getFolderPath() + "ACQ_009_" + branch.getBranchCode() + "_" + branch.getBranchName()
                + "_" + dateFN
                + ".xlsx";
        DynamicObject dynamicObject = new DynamicObject();
        Map<String, String> columns = new LinkedHashMap<>();
        columns.put("STT", "STT");
        columns.put("BRANCH", "CN quản lý");
        columns.put("IMP_NBR", "Số Cif khách hàng");
        columns.put("NAM", "Tên khách hàng");
        columns.put("1", "Merchant main ID");
        columns.put("OWN_MERCH", "Merchant liên kết");
        columns.put("MERCHANT", "Merchant ID");
        columns.put("TERM_NBR", "Terminal ID");
        columns.put("DEP_ACCT", "MCC");
        columns.put("CAID", "AM");
        columns.put("DEP_ACCT", "Dep_acct");
        columns.put("MC", "MC");
        columns.put("MD", "MD");
        columns.put("VC", "VC");
        columns.put("VD", "VD");
        columns.put("JC", "JC");
        columns.put("JD", "JD");
        columns.put("PD", "PD");
        columns.put("PC", "PC");
        dynamicObject.setColumns(columns);

        Map<Integer, Object> inputParams = new HashMap<>();
        inputParams.put(1, branch.getBranchCode());
        List<Integer> outParams = new ArrayList<>();
        outParams.add(2);

        List<DynamicObject> dynamicObjects = databaseService.callProcedure("REPORT_MIGRATE", "ACQ_009", columns,
                inputParams, outParams);

        if (dynamicObjects.size() == 0) {
            DynamicObject dynamicObject1 = new DynamicObject();
            dynamicObject1.setColumns(columns);
            dynamicObject1.setProperties(new LinkedHashMap<>());
            dynamicObjects.add(dynamicObject1);
        }

        // map dynamicObjects add properties value JD = "" if JD is null
        for (DynamicObject dynamicObject2 : dynamicObjects) {
            if (dynamicObject2.getProperties().get("JD") == null) {
                dynamicObject2.getProperties().put("JD", "");
            }
            dynamicObject2.getProperties().put("1", dynamicObject2.getProperties().get("OWN_MERCH"));
        }

        ExcelGenerator excelGenerator = new ExcelGenerator();
        Sheet sheet = excelGenerator.generateExcel(9, dynamicObjects, true);

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
        // cellStyle.setWrapText(true);
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        headerRow.createCell(0).setCellValue("STT");
        headerRow.getCell(0).setCellStyle(cellStyle);
        // Cell mergeCell = headerRow.getCell(0);
        // mergeCell.setCellValue("Phí MDR áp dụng VND");
        // mergeCell.setCellStyle(cellStyle);
        for (int i = 1; i < columns.size() - 8; i++) {
            headerRow.createCell(i).setCellValue((String) columns.values().toArray()[i]);
            // resize column width
            sheet.autoSizeColumn(i);
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
            sheet.autoSizeColumn(i);
            headerRow2.getCell(i).setCellStyle(cellStyle);
        }
        // merge cell for header
        // headerRow.getCell(9).setCellStyle(cellStyle);
        sheet.addMergedRegion(new CellRangeAddress(6, 6, columns.size() - 8, columns.size() - 1));

        Font fontBold = sheet.getWorkbook().createFont();
        fontBold.setBold(true);
        CellStyle styleBold = sheet.getWorkbook().createCellStyle();
        styleBold.setFont(fontBold);
        int endRow = 9 + dynamicObjects.size() + 2;
        Row eRow = sheet.createRow(endRow);
        eRow.createCell(3).setCellValue("LẬP BẢNG");
        eRow.createCell(9).setCellValue("NGƯỜI KIỂM SOÁT");
        eRow.createCell(14).setCellValue("ĐẠI DIỆN CHI NHÁNH");
        eRow.getCell(3).setCellStyle(styleBold);
        eRow.getCell(9).setCellStyle(styleBold);
        eRow.getCell(14).setCellStyle(styleBold);

        excelGenerator.writeExcel(fileName);
    }

    public static void GL007Report() throws FileNotFoundException, IOException {
        System.out.println("Generating GL007 Report");
        DynamicObject dynamicObject = new DynamicObject();
        Map<String, String> columns = new LinkedHashMap<>();
        columns.put("branch", "Chi nhánh");
        columns.put("tsc", "Mã phòng/ban (TSC CN/ PGD)");
        columns.put("account", "Số hiệu tài khoản");
        columns.put("acc_name", "USER hạch toán");
        columns.put("user", "Người thực hiện");
        columns.put("chanel_id", "CHANEL ID/Chương trình hạch toán");
        columns.put("date", "Ngày phát sinh");
        columns.put("terminal_id", "Terminal ID");
        columns.put("nbr", "Số thẻ");
        columns.put("scc", "Số chuẩn chi");
        columns.put("profile", "ZPRFREFNO Profile ");
        columns.put("amount", "Số tiền ");
        columns.put("content", "Nội dung giao dịch");
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
        Sheet sheet = excelGenerator.generateExcel(7, dynamicObjects, false);

        Font fontBold = sheet.getWorkbook().createFont();
        fontBold.setBold(true);
        CellStyle styleBold = sheet.getWorkbook().createCellStyle();
        styleBold.setFont(fontBold);

        Font font = sheet.getWorkbook().createFont();
        font.setBold(true);
        font.setFontHeightInPoints((short) 16);
        CellStyle style = sheet.getWorkbook().createCellStyle();
        style.setFont(font);
        style.setAlignment(HorizontalAlignment.CENTER);

        Row row0 = sheet.createRow(0);
        row0.createCell(0).setCellValue("Mã chi nhánh: ");
        Row row1 = sheet.createRow(1);
        row1.createCell(0).setCellValue("Tên chi nhánh: ");

        Row rowTitle = sheet.createRow(2);
        rowTitle.createCell(0).setCellValue("BÁO CÁO TỔNG HỢP SỐ LIỆU CHUYỂN ĐỔI");
        rowTitle.getCell(0).setCellStyle(style);
        Row row2 = sheet.createRow(3);
        row2.createCell(0).setCellValue("Mã báo cáo: GL_005_ISS");
        Row row3 = sheet.createRow(4);
        DateFormat dateFormat = new SimpleDateFormat("dd/MM/yyyy");
        row3.createCell(0).setCellValue("Ngày dữ liệu: " + dateFormat.format(date));
        Row row4 = sheet.createRow(5);
        row4.createCell(0).setCellValue("Loại tiền tệ: VND");

        sheet.addMergedRegion(new CellRangeAddress(2, 2, 0, 14));

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
        headerRow.createCell(0).setCellValue("TT");
        headerRow.getCell(0).setCellStyle(cellStyle);
        headerRow2.createCell(0).setCellStyle(cellStyle);
        for (int i = 1; i < columns.size() - 5; i++) {
            headerRow.createCell(i).setCellValue((String) columns.values().toArray()[i]);
            headerRow.getCell(i).setCellStyle(cellStyle);
            headerRow2.createCell(i).setCellStyle(cellStyle);
        }
        for (int i = columns.size() - 6; i < columns.size() - 1; i++) {
            headerRow2.createCell(i).setCellValue((String) columns.values().toArray()[i]);
            if (headerRow.getCell(i) == null) {
                headerRow.createCell(i).setCellStyle(cellStyle);
                headerRow.getCell(i).setCellValue("Giá trị chuyển đổi");
            } else {
                headerRow.createCell(i).setCellStyle(cellStyle);
                headerRow.getCell(i).setCellValue("Giá trị chuyển đổi");
            }
            headerRow2.getCell(i).setCellStyle(cellStyle);
        }
        Cell headerLastCell = headerRow.createCell(columns.size() - 1);
        headerLastCell.setCellValue((String) columns.values().toArray()[columns.size() - 1]);
        headerLastCell.setCellStyle(cellStyle);
        headerRow2.createCell(columns.size() - 1).setCellStyle(cellStyle);

        sheet.addMergedRegion(new CellRangeAddress(6, 6, 9, 13));
        for (int i = 0; i < 9; i++) {
            // merge cell row 6 -7
            sheet.addMergedRegion(new CellRangeAddress(6, 7, i, i));
        }
        sheet.addMergedRegion(new CellRangeAddress(6, 7, columns.size() - 1, columns.size() - 1));

        int endRow = 6 + dynamicObjects.size() + 2;
        Row eRow = sheet.createRow(endRow);
        eRow.createCell(2).setCellValue("LẬP BẢNG");
        eRow.createCell(8).setCellValue("NGƯỜI KIỂM SOÁT");
        eRow.createCell(13).setCellValue("ĐẠI DIỆN CHI NHÁNH");
        eRow.getCell(2).setCellStyle(styleBold);
        eRow.getCell(8).setCellStyle(styleBold);
        eRow.getCell(13).setCellStyle(styleBold);

        excelGenerator.writeExcel(fileName);

    }

    public static void ISS009Report(Branch branch) throws FileNotFoundException, IOException, SQLException {
        System.out.println("Generating ISS009 Report");
        Date date = new Date();
        DateFormat dateFNFormat = new SimpleDateFormat("ddMMyyyy");
        String dateFN = dateFNFormat.format(date);
        String fileName = branch.getFolderPath() + "ISS_009_" + branch.getBranchCode() + "_" + branch.getBranchName()
                + "_" + dateFN
                + ".xlsx";
        DynamicObject dynamicObject = new DynamicObject();
        Map<String, String> columns = new LinkedHashMap<>();
        columns.put("date", "Ngày dữ liệu");
        columns.put("cif", "Số CIF");
        columns.put("full_name", "Họ tên khách hàng");
        columns.put("GENDER", "Giới tính");
        columns.put("BIRTH_DATE", "Ngày tháng năm sinh");
        columns.put("REG_NUMBER", "Số ID");
        columns.put("REG_DETAILS", "Nơi cấp ID");
        columns.put("REG_DETAILS2", "Ngày cấp ID");
        columns.put("ADD_DATE_01", "Ngày hết hạn ID");
        columns.put("MARITAL_STATUS", "Tình trạng hôn nhân");
        columns.put("CITIZENSHIP", "Quốc tịch");
        columns.put("BIRTH_PLACE", "Nơi sinh");
        columns.put("PHONE_M", "SĐT di động");
        columns.put("PHONE_H", "SĐT cố định");
        columns.put("PHONE", "SĐT cơ quan");
        columns.put("E_MAIL", "Email");
        columns.put("COMPANY_NAM", "Tên cơ quan công tác");
        columns.put("PROFESSION", "Chức vụ");
        columns.put("a", "Mã nhóm KH");
        columns.put("b", "Mã GLP (mã dặm thường)");
        columns.put("c", "Mã VTV");
        columns.put("SOCIAL_NUMBER", "Câu trả lời câu hỏi bảo mật");
        columns.put("reference_name", "Tên người tham chiếu");
        columns.put("TITLE", "Giới tính người tham chiếu");
        columns.put("ADD_INFO", "Mối quan hệ với chủ thẻ của người tham chiếu");
        columns.put("CLIENT_ADDRESS.PHONE_M", "SDDT người tham chiếu");
        columns.put("CLIENT_ADDRESS.E_MAIL", "Email người tham chiếu");
        columns.put("CLIENT_ADDRESS.ADDRESS_LINE_1", "Cơ quan công tác người tham chiếu");

        dynamicObject.setColumns(columns);

        Map<Integer, Object> inputParams = new HashMap<>();
        inputParams.put(1, branch.getBranchCode());
        List<Integer> outParams = new ArrayList<>();
        outParams.add(2);

        List<DynamicObject> dynamicObjects = databaseService.callProcedure("REPORT_MIGRATE", "ISS_009", columns,
                inputParams, outParams);

        if (dynamicObjects.size() == 0) {
            DynamicObject dynamicObject1 = new DynamicObject();
            dynamicObject1.setColumns(columns);
            dynamicObject1.setProperties(new LinkedHashMap<>());
            dynamicObjects.add(dynamicObject1);
        }

        // map dynamicObjects add properties value a = "" if a is null
        for (DynamicObject dynamicObject2 : dynamicObjects) {
            if (dynamicObject2.getProperties().get("cif") != null
                    && dynamicObject2.getProperties().get("REG_DETAILS") == null) {
                dynamicObject2.getProperties().put("REG_DETAILS", "");
            }
        }

        ExcelGenerator excelGenerator = new ExcelGenerator();
        Sheet sheet = excelGenerator.generateExcel(7, dynamicObjects, false);

        // title row 0 title
        Row titleRow = sheet.createRow(0);
        titleRow.createCell(1)
                .setCellValue("BÁO CÁO CHI TIẾT THÔNG TIN CIF KHCN CHUYỂN ĐỔI THÀNH CÔNG LÊN WAY4");
        // row 1 branch
        Row row1 = sheet.createRow(1);
        row1.createCell(1).setCellValue("Mã chi nhánh: " + branch.getBranchCode());

        // row 2 branch name
        Row row2 = sheet.createRow(2);
        row2.createCell(1).setCellValue("Tên chi nhánh: " + branch.getBranchName());
        row2.createCell(15).setCellValue("Mã bc: ISS_009");

        // row 4 date
        DateFormat dateFormat = new SimpleDateFormat("dd/MM/yyyy");
        String dateStr = dateFormat.format(date);
        Row row4 = sheet.createRow(4);
        row4.createCell(5).setCellValue("Ngày báo cáo: " + dateStr);

        Font fontBold = sheet.getWorkbook().createFont();
        fontBold.setBold(true);
        CellStyle styleBold = sheet.getWorkbook().createCellStyle();
        styleBold.setFont(fontBold);

        int endRow = 6 + dynamicObjects.size() + 2;
        Row eRow = sheet.createRow(endRow);
        eRow.createCell(3).setCellValue("LẬP BẢNG");
        eRow.createCell(9).setCellValue("NGƯỜI KIỂM SOÁT");
        eRow.createCell(14).setCellValue("ĐẠI DIỆN CHI NHÁNH");
        eRow.getCell(3).setCellStyle(styleBold);
        eRow.getCell(9).setCellStyle(styleBold);
        eRow.getCell(14).setCellStyle(styleBold);

        // merge cell for header
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 1, 15));
        excelGenerator.writeExcel(fileName);
    }

    public static void ISS010Report(Branch branch) throws FileNotFoundException, IOException, SQLException {
        System.out.println("Generating ISS010 Report");
        Date date = new Date();
        DateFormat dateFNFormat = new SimpleDateFormat("ddMMyyyy");
        String dateFN = dateFNFormat.format(date);
        String fileName = branch.getFolderPath() + "ISS_010_" + branch.getBranchCode() + "_" + branch.getBranchName()
                + "_" + dateFN
                + ".xlsx";
        DynamicObject dynamicObject = new DynamicObject();
        Map<String, String> columns = new LinkedHashMap<>();
        columns.put("date", "Ngày dữ liệu");
        columns.put("branch", "CN PHT");
        columns.put("DATE_OPEN", "Ngày khai báo thông tin");
        columns.put("CLIENT_NUMBER", "CIF");
        columns.put("SHORT_NAME", "Tên doanh nghiệp");
        columns.put("REG_NUMBER", "Mã số ĐKKD của Doanh nghiệp");
        columns.put("fullname", "Tên người đại diện theo pháp luật");
        columns.put("TR_COMPANY_NAM", "Tên công ty được dập nổi trên thẻ");
        columns.put("AUTH_LIMIT_AMOUNT", "HMTD doanh nghiệp");
        columns.put("SL_THE_TIN_DUNG", "Số lượng thẻ TDDN");
        columns.put("Addre", "Địa chỉ DN");
        columns.put("E_MAIL", "Email DN");
        columns.put("V_CS_ALL_ACNT_STATUS.STATUS_TYPE_CODE", "Mã lãi suất");
        columns.put("CONTR_STATUS", "Trạng thái TK thẻ Doanh nghiệp");
        columns.put("fullname2", "Họ tên người liên hệ");
        columns.put("ADDRESS_LINE_2", "Phòng/Ban công tác của người liên hệ");
        columns.put("PHONE_M", "SĐT liên lạc");
        columns.put("NGAY_SAO_KE_THE", "Ngày sao kê thẻ ");
        columns.put("TY_LE_TRICH_NO_TU_DONG", "Tỷ lệ trích nợ tự động");
        columns.put("SO_TIEN_DK_TT_TD", "Số tiền đk thanh toán tự động ");
        columns.put("SO_TK_DK_TRICH_NO_TU_DONG", "Số TK đăng ký trích nợ tự động");

        dynamicObject.setColumns(columns);

        Map<Integer, Object> inputParams = new HashMap<>();
        inputParams.put(1, branch.getBranchCode());
        List<Integer> outParams = new ArrayList<>();
        outParams.add(2);

        List<DynamicObject> dynamicObjects = databaseService.callProcedure("REPORT_MIGRATE", "ISS_010", columns,
                inputParams, outParams);

        if (dynamicObjects.size() == 0) {
            DynamicObject dynamicObject1 = new DynamicObject();
            dynamicObject1.setColumns(columns);
            dynamicObject1.setProperties(new LinkedHashMap<>());
            dynamicObjects.add(dynamicObject1);
        } else {
            for (DynamicObject dynamicObject2 : dynamicObjects) {
                if (dynamicObject2.getProperties().get("fullname") != null) {
                    String fullname = (String) dynamicObject2.getProperties().get("fullname");
                    String addr1 = (String) dynamicObject2.getProperties().get("ADDRESS_LINE_1");
                    dynamicObject2.getProperties().put("fullname2", fullname.trim());
                    dynamicObject2.getProperties().put("ADDRESS_LINE_2", addr1.trim());
                }
            }
        }

        ExcelGenerator excelGenerator = new ExcelGenerator();
        Sheet sheet = excelGenerator.generateExcel(7, dynamicObjects, false);

        // title row 0 title
        Row titleRow = sheet.createRow(0);
        titleRow.createCell(1)
                .setCellValue("BÁO CÁO CHI TIẾT THÔNG TIN CIF KHDN CHUYỂN ĐỔI THÀNH CÔNG LÊN WAY4");
        Font font = sheet.getWorkbook().createFont();
        font.setBold(true);
        font.setFontHeightInPoints((short) 16);
        CellStyle style = sheet.getWorkbook().createCellStyle();
        style.setFont(font);
        style.setAlignment(HorizontalAlignment.CENTER);
        titleRow.getCell(1).setCellStyle(style);
        // row 1 branch
        Row row1 = sheet.createRow(1);
        row1.createCell(1).setCellValue("Mã chi nhánh: " + branch.getBranchCode());

        // row 2 branch name
        Row row2 = sheet.createRow(2);
        row2.createCell(1).setCellValue("Tên chi nhánh: " + branch.getBranchName());
        row2.createCell(15).setCellValue("Mã bc: ISS_010");

        // row 4 date
        DateFormat dateFormat = new SimpleDateFormat("dd/MM/yyyy");
        String dateStr = dateFormat.format(date);
        Row row4 = sheet.createRow(4);
        row4.createCell(5).setCellValue("Ngày báo cáo: " + dateStr);

        Font fontBold = sheet.getWorkbook().createFont();
        fontBold.setBold(true);
        CellStyle styleBold = sheet.getWorkbook().createCellStyle();
        styleBold.setFont(fontBold);
        row1.getCell(1).setCellStyle(styleBold);
        row2.getCell(1).setCellStyle(styleBold);
        row2.getCell(15).setCellStyle(styleBold);
        row4.getCell(5).setCellStyle(styleBold);

        int endRow = 6 + dynamicObjects.size() + 2;
        Row eRow = sheet.createRow(endRow);
        eRow.createCell(3).setCellValue("LẬP BẢNG");
        eRow.createCell(9).setCellValue("NGƯỜI KIỂM SOÁT");
        eRow.createCell(14).setCellValue("ĐẠI DIỆN CHI NHÁNH");
        eRow.getCell(3).setCellStyle(styleBold);
        eRow.getCell(9).setCellStyle(styleBold);
        eRow.getCell(14).setCellStyle(styleBold);

        // merge cell for header
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 1, 15));
        excelGenerator.writeExcel(fileName);
    }

    public static void ATM001REPORT() throws FileNotFoundException, IOException {
        System.out.println("Generating ATM001 Report");
        Date date = new Date();
        DateFormat dateFNFormat = new SimpleDateFormat("ddMMyyyy");
        String dateFN = dateFNFormat.format(date);
        String fileName = "/FTPData/HSC/ATM_001_" + dateFN + ".xlsx";
        DynamicObject dynamicObject = new DynamicObject();
        Map<String, String> columns = new LinkedHashMap<>();
        columns.put("s_excel_atm_cnQuanly.CNQUANLY", "Chi nhánh");
        columns.put("s_tms_tblAtm.atmid", "Số lượng ATM");
        columns.put("s_tms_tblCrm.atmid", "Số lượng CRM");
        columns.put("ACNT_CONTRACT.BRANCH", "Branch quản lý");
        columns.put("ACNT_CONTRACT.PRODUCT", "Số lượng CRM");
        columns.put("ACNT_CONTRACT.PRODUCT2", "Số lượng CRM");

        dynamicObject.setColumns(columns);

        // not fill properties for now
        dynamicObject.setProperties(new LinkedHashMap<>());

        List<DynamicObject> dynamicObjects = new ArrayList<>();
        dynamicObjects.add(dynamicObject);

        ExcelGenerator excelGenerator = new ExcelGenerator();
        Sheet sheet = excelGenerator.generateExcel(7, dynamicObjects, false);

        // title row 0
        Row code = sheet.createRow(0);
        code.createCell(0)
                .setCellValue("Mã báo cáo: ATM_001");
        // title row 1
        Row titleRow = sheet.createRow(1);
        titleRow.createCell(0)
                .setCellValue("Báo cáo đối chiếu tổng hợp số lượng ATM chuyển đổi");
        // font bold, size 16 for title
        Font font = sheet.getWorkbook().createFont();
        font.setBold(true);
        font.setFontHeightInPoints((short) 14);
        CellStyle style = sheet.getWorkbook().createCellStyle();
        style.setFont(font);
        style.setAlignment(HorizontalAlignment.CENTER);
        titleRow.getCell(0).setCellStyle(style);
        sheet.addMergedRegion(new CellRangeAddress(1, 1, 0, 6));

        Row row5 = sheet.createRow(5);
        row5.createCell(1).setCellValue("Trước chuyển đổi");
        row5.createCell(4).setCellValue("Sau chuyển đổi");

        // merge cell for header
        sheet.addMergedRegion(new CellRangeAddress(5, 5, 1, 3));
        sheet.addMergedRegion(new CellRangeAddress(5, 5, 4, 6));

        Font fontBold = sheet.getWorkbook().createFont();
        fontBold.setBold(true);
        CellStyle styleBold = sheet.getWorkbook().createCellStyle();
        styleBold.setFont(fontBold);

        excelGenerator.writeExcel(fileName);
    }

    public static void ISS0010Report() throws FileNotFoundException, IOException {
        System.out.println("Generating ISS0010 Report");
        Date date = new Date();
        DateFormat dateFNFormat = new SimpleDateFormat("ddMMyyyy");
        String dateFN = dateFNFormat.format(date);
        String fileName = "/FTPData/HSC/ISS_001_0_" + dateFN + ".xlsx";
        DynamicObject dynamicObject = new DynamicObject();
        Map<String, String> columns = new LinkedHashMap<>();
        columns.put("CIF.ZBOC", "BDS");
        columns.put("CUSTR.CUSTR_REF1", "Tổng số CIF KHCN Cadencie");
        columns.put("CUSTR.CUSTR_REF2", "CIF không chuyển đổi");
        columns.put("CUSTR.CUSTR_REF3", "Chuyển đổi (Migration)");
        columns.put("CLIENT.CLIENT_NUMBER", "Cập nhật lên hệ thống Way4");
        columns.put("5", "Chênh lệch");
        columns.put("cif1", "Tổng số CIF KHDN Cadencie");
        columns.put("cif2", "CIF không chuyển đổi");
        columns.put("BUSI_NAME/ACN", "Chuyển đổi (Migration)");
        columns.put("CLIENT_NUMBER", "Cập nhật lên hệ thống Way4");
        columns.put("8", "Chênh lệch");

        dynamicObject.setColumns(columns);

        // not fill properties for now
        dynamicObject.setProperties(new LinkedHashMap<>());

        List<DynamicObject> dynamicObjects = new ArrayList<>();
        dynamicObjects.add(dynamicObject);

        ExcelGenerator excelGenerator = new ExcelGenerator();
        Sheet sheet = excelGenerator.generateExcel(7, dynamicObjects, true);

        // title row 0
        Row titleRow = sheet.createRow(0);
        titleRow.createCell(0)
                .setCellValue(
                        "BÁO CÁO TỔNG HỢP SỐ LIỆU  KHÁCH HÀNG THẺ QUỐC TẾ CHUYỂN ĐỔI TOÀN HỆ THỐNG THEO CHI NHÁNH");
        // title row 1
        Row code = sheet.createRow(1);
        code.createCell(0)
                .setCellValue("Mã chi nhánh: ");
        code.createCell(9)
                .setCellValue("Mã báo cáo: ISS_001.0");
        // font bold, size 16 for title
        Font font = sheet.getWorkbook().createFont();
        font.setBold(true);
        font.setFontHeightInPoints((short) 16);
        CellStyle style = sheet.getWorkbook().createCellStyle();
        style.setFont(font);
        style.setAlignment(HorizontalAlignment.CENTER);
        titleRow.getCell(0).setCellStyle(style);
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 10));

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
        headerRow2.createCell(0).setCellStyle(cellStyle);
        headerRow.createCell(1).setCellValue((String) columns.values().toArray()[0]);
        headerRow.getCell(1).setCellStyle(cellStyle);
        headerRow2.createCell(1).setCellValue((String) columns.values().toArray()[0]);
        headerRow2.getCell(1).setCellStyle(cellStyle);
        headerRow.getCell(0).setCellStyle(cellStyle);

        for (int i = 2; i < columns.size(); i++) {
            if (i < 7) {
                headerRow.createCell(i).setCellStyle(cellStyle);
                headerRow.getCell(i).setCellValue("Chuyển đổi CIF - KHCN");
            } else {
                headerRow.createCell(i).setCellStyle(cellStyle);
                headerRow.getCell(i).setCellValue("Chuyển đổi CIF - KHDN");
            }
            headerRow2.createCell(i).setCellValue((String) columns.values().toArray()[i - 1]);
            sheet.autoSizeColumn(i);
            headerRow2.getCell(i).setCellStyle(cellStyle);
        }
        // row sum after data
        Row sumRow = sheet.createRow(7 + dynamicObjects.size());
        sumRow.createCell(0).setCellValue("SUM");
        // sum column 3,4,5,9,10
        for (int i = 2; i < 6; i++) {
            sumRow.createCell(i).setCellValue("***");
        }
        for (int i = 9; i < 11; i++) {
            sumRow.createCell(i).setCellValue("***");
        }

        // merge cell for header
        sheet.addMergedRegion(new CellRangeAddress(6, 6, 2, 6));
        sheet.addMergedRegion(new CellRangeAddress(6, 6, 7, 10));
        sheet.addMergedRegion(new CellRangeAddress(6, 7, 0, 0));
        sheet.addMergedRegion(new CellRangeAddress(6, 7, 1, 1));

        Font fontBold = sheet.getWorkbook().createFont();
        fontBold.setBold(true);
        CellStyle styleBold = sheet.getWorkbook().createCellStyle();
        styleBold.setFont(fontBold);

        int endRow = 9 + dynamicObjects.size() + 2;
        Row eRow = sheet.createRow(endRow);
        eRow.createCell(0).setCellValue("LẬP BẢNG");
        eRow.createCell(4).setCellValue("NGƯỜI KIỂM SOÁT");
        eRow.createCell(9).setCellValue("ĐẠI DIỆN CHI NHÁNH");
        eRow.getCell(0).setCellStyle(styleBold);
        eRow.getCell(4).setCellStyle(styleBold);
        eRow.getCell(9).setCellStyle(styleBold);

        excelGenerator.writeExcel(fileName);
    }

    public static void ISS0011Report() throws FileNotFoundException, IOException {
        System.out.println("Generating ISS0011 Report");
        Date date = new Date();
        DateFormat dateFNFormat = new SimpleDateFormat("ddMMyyyy");
        String dateFN = dateFNFormat.format(date);
        String fileName = "/FTPData/HSC/ISS_001_1_" + dateFN + ".xlsx";
        DynamicObject dynamicObject = new DynamicObject();
        Map<String, String> columns = new LinkedHashMap<>();
        columns.put("CIF.ZBOC", "BDS");
        columns.put("CARD.PROFIL", "Product code");
        columns.put("ACCT.ACCOUNT1", "Số lượng trên Cadencie");
        columns.put("ACCT.ACCOUNT2", "Không chuyển đổi");
        columns.put("ACCT.ACCOUNT3", "Chuyển đổi (Migration)");
        columns.put("ACNT_CONTRACT.CONTRACT_NUMBER", "Cập nhật lên hệ thống Way4");
        columns.put("7", "Chênh lệch");
        columns.put("CARD.CARD_NBR1", "Số lượng thẻ trên Cadencie");
        columns.put("CARD.CARD_NBR2", "Số lượng thẻ không CĐ");
        columns.put("CARD.CARD_NBR3", "Số lượng thẻ CĐ");
        columns.put("CARD_INFO.CARD_NUMBER", "Cập nhật lên W4");
        columns.put("12", "Chênh lệch");

        dynamicObject.setColumns(columns);

        // not fill properties for now
        dynamicObject.setProperties(new LinkedHashMap<>());

        List<DynamicObject> dynamicObjects = new ArrayList<>();
        dynamicObjects.add(dynamicObject);

        ExcelGenerator excelGenerator = new ExcelGenerator();
        Sheet sheet = excelGenerator.generateExcel(7, dynamicObjects, true);

        // title row
        Row titleRow = sheet.createRow(0);
        titleRow.createCell(0)
                .setCellValue(
                        "BÁO CÁO TỔNG HỢP SỐ LIỆU TÀI KHOẢN THẺ QUỐC TẾ CHUYỂN ĐỔI TOÀN HỆ THỐNG THEO CHI NHÁNH VÀ MÃ SẢN PHẨM THẺ");
        // font bold, size 16 for title
        Font font = sheet.getWorkbook().createFont();
        font.setBold(true);
        font.setFontHeightInPoints((short) 16);
        CellStyle style = sheet.getWorkbook().createCellStyle();
        style.setFont(font);
        style.setAlignment(HorizontalAlignment.CENTER);
        titleRow.getCell(0).setCellStyle(style);
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 11));

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
        headerRow2.createCell(0).setCellStyle(cellStyle);
        headerRow.createCell(1).setCellValue((String) columns.values().toArray()[0]);
        headerRow.getCell(1).setCellStyle(cellStyle);
        headerRow2.createCell(1).setCellValue((String) columns.values().toArray()[0]);
        headerRow2.getCell(1).setCellStyle(cellStyle);
        headerRow.createCell(2).setCellValue((String) columns.values().toArray()[1]);
        headerRow.getCell(2).setCellStyle(cellStyle);
        headerRow2.createCell(2).setCellValue((String) columns.values().toArray()[1]);
        headerRow2.getCell(2).setCellStyle(cellStyle);
        headerRow.getCell(0).setCellStyle(cellStyle);

        for (int i = 3; i < columns.size(); i++) {
            headerRow.createCell(i).setCellStyle(cellStyle);
            headerRow.getCell(i).setCellValue("Chuyển đổi Tài khoản thẻ");
            headerRow2.createCell(i).setCellValue((String) columns.values().toArray()[i - 1]);
            sheet.autoSizeColumn(i);
            headerRow2.getCell(i).setCellStyle(cellStyle);
        }

        // row sum after data
        Row sumRow = sheet.createRow(7 + dynamicObjects.size());
        sumRow.createCell(0).setCellValue("SUM");

        sumRow.createCell(3).setCellValue("***");
        sumRow.createCell(6).setCellValue("***");
        sumRow.createCell(7).setCellValue("***");
        sumRow.createCell(8).setCellValue("***");
        sumRow.createCell(9).setCellValue("***");

        // merge cell for header
        sheet.addMergedRegion(new CellRangeAddress(6, 6, 3, 11));
        sheet.addMergedRegion(new CellRangeAddress(6, 7, 0, 0));
        sheet.addMergedRegion(new CellRangeAddress(6, 7, 1, 1));
        sheet.addMergedRegion(new CellRangeAddress(6, 7, 2, 2));

        Font fontBold = sheet.getWorkbook().createFont();
        fontBold.setBold(true);
        CellStyle styleBold = sheet.getWorkbook().createCellStyle();
        styleBold.setFont(fontBold);

        int endRow = 9 + dynamicObjects.size() + 2;
        Row eRow = sheet.createRow(endRow);
        eRow.createCell(0).setCellValue("LẬP BẢNG");
        eRow.createCell(4).setCellValue("NGƯỜI KIỂM SOÁT");

        eRow.getCell(0).setCellStyle(styleBold);
        eRow.getCell(4).setCellStyle(styleBold);

        excelGenerator.writeExcel(fileName);
    }

    public static void ISS002Report() throws FileNotFoundException, IOException {
        System.out.println("Generating ISS002 Report");
        Date date = new Date();
        DateFormat dateFNFormat = new SimpleDateFormat("ddMMyyyy");
        String dateFN = dateFNFormat.format(date);
        String fileName = "/FTPData/HSC/ISS_002_" + dateFN + ".xlsx";
        DynamicObject dynamicObject = new DynamicObject();
        Map<String, String> columns = new LinkedHashMap<>();
        columns.put("CLIENT.CLIENT_NUMBER", "Số CIF");
        columns.put("fullname1", "Cadencie");
        columns.put("fullname2", "Way4");
        columns.put("1", "So sánh (Pass/Fail)");
        columns.put("CUSTR.GENDER", "Cadencie");
        columns.put("CLIENT.GENDER", "Way4");
        columns.put("2", "So sánh (Pass/Fail)");
        columns.put("CUSTR.DAY_BIRTH", "Cadencie");
        columns.put("CLIENT.BIRTH_DATE", "Way4");
        columns.put("3", "So sánh (Pass/Fail)");
        columns.put("CUSTR.ID_NBR", "Cadencie");
        columns.put("CLIENT.REG_NUMER", "Way4");
        columns.put("4", "So sánh (Pass/Fail)");
        columns.put("CIF.OED", "Cadencie");
        columns.put("CLIENT.add_date_02", "Way4");
        columns.put("5", "So sánh (Pass/Fail)");
        columns.put("ADDR1", "Cadencie");
        columns.put("ADDR2", "Way4");
        columns.put("6", "So sánh (Pass/Fail)");
        columns.put("ADDR3", "Cadencie");
        columns.put("ADDR4", "Way4");
        columns.put("7", "So sánh (Pass/Fail)");
        columns.put("ADDR5", "Cadencie");
        columns.put("ADDR6", "Way4");
        columns.put("8", "So sánh (Pass/Fail)");
        columns.put("CUSTR.MOBL_PHONE", "Cadencie");
        columns.put("CLIENT.PHONE", "Way4");
        columns.put("9", "So sánh (Pass/Fail)");
        columns.put("CUSTR.EMAIL_ADDR", "Cadencie");
        columns.put("CLIENT.E_MAIL", "Way4");
        columns.put("10", "So sánh (Pass/Fail)");
        columns.put("CIF.STAT", "Cadencie");
        columns.put("STATUS_TYPE_CODE1", "Way4");
        columns.put("11", "So sánh (Pass/Fail)");
        columns.put("CUSTR.ID_EXPDATE", "Cadencie");
        columns.put("CLIENT.ADD_DATE_01", "Way4");
        columns.put("12", "So sánh (Pass/Fail)");
        columns.put("CIF.DARCOVR", "Cadencie");
        columns.put("STATUS_TYPE_CODE2", "Way4");
        columns.put("13", "So sánh (Pass/Fail)");
        columns.put("CIF.CCODE", "Cadencie");
        columns.put("STATUS_TYPE_CODE3", "Way4");
        columns.put("14", "So sánh (Pass/Fail)");

        dynamicObject.setColumns(columns);

        // not fill properties for now
        dynamicObject.setProperties(new LinkedHashMap<>());

        List<DynamicObject> dynamicObjects = new ArrayList<>();
        dynamicObjects.add(dynamicObject);

        ExcelGenerator excelGenerator = new ExcelGenerator();
        Sheet sheet = excelGenerator.generateExcel(7, dynamicObjects, true);

        // title row
        Row titleRow = sheet.createRow(0);
        titleRow.createCell(0)
                .setCellValue(
                        "BÁO CÁO ĐỐI CHIẾU DỮ LIỆU CIF KHCN TẠI CADENCIE - WAY4");
        // font bold, size 16 for title
        Font font = sheet.getWorkbook().createFont();
        font.setBold(true);
        font.setFontHeightInPoints((short) 16);
        CellStyle style = sheet.getWorkbook().createCellStyle();
        style.setFont(font);
        style.setAlignment(HorizontalAlignment.CENTER);
        titleRow.getCell(0).setCellStyle(style);
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 14));

        Row row1 = sheet.createRow(1);
        row1.createCell(0).setCellValue("Mã chi nhánh: ");
        row1.createCell(9).setCellValue("Mã báo cáo: ISS_002");
        Row row2 = sheet.createRow(2);
        row2.createCell(0).setCellValue("Tên chi nhánh: ");
        DateFormat dateFormat = new SimpleDateFormat("dd/MM/yyyy");
        String dateStr = dateFormat.format(date);
        Row row3 = sheet.createRow(3);
        row3.createCell(9).setCellValue("Ngày báo cáo: " + dateStr);

        Font fontBold = sheet.getWorkbook().createFont();
        fontBold.setBold(true);
        CellStyle styleBold = sheet.getWorkbook().createCellStyle();
        styleBold.setFont(fontBold);

        row1.getCell(0).setCellStyle(styleBold);
        row1.getCell(9).setCellStyle(styleBold);
        row2.getCell(0).setCellStyle(styleBold);
        row3.getCell(9).setCellStyle(styleBold);

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
        headerRow2.createCell(0).setCellStyle(cellStyle);
        headerRow.createCell(1).setCellValue((String) columns.values().toArray()[0]);
        headerRow.getCell(1).setCellStyle(cellStyle);
        headerRow2.createCell(1).setCellValue((String) columns.values().toArray()[0]);
        headerRow2.getCell(1).setCellStyle(cellStyle);
        headerRow.getCell(0).setCellStyle(cellStyle);

        String[] header = { "Tên KH", "Giới tính", "Ngày tháng năm sinh", "Số CMND của KH", "Ngày hết hạn thị lực",
                "Địa chỉ thường chú", "Địa chỉ cơ quan", "Địa chỉ cư chú", "Số điện thoại", "Email", "CIF status",
                "Ngày hết hạn CMND KH (ADD_DATE_01)", "Nhóm nợ CIC (STATUS_TYPE_CODE)", "Hạng khách hàng" };

        for (int i = 2; i < columns.size() + 1; i++) {
            headerRow.createCell(i).setCellStyle(cellStyle);
            headerRow2.createCell(i).setCellValue((String) columns.values().toArray()[i - 1]);
            sheet.autoSizeColumn(i);
            headerRow2.getCell(i).setCellStyle(cellStyle);
        }

        // no sum
        // merge cell for header: every 3 columns begin from 2
        for (int i = 2; i < columns.size(); i += 3) {
            headerRow.getCell(i).setCellValue(header[i / 3]);
            sheet.addMergedRegion(new CellRangeAddress(6, 6, i, i + 2));
        }

        sheet.addMergedRegion(new CellRangeAddress(6, 7, 0, 0));
        sheet.addMergedRegion(new CellRangeAddress(6, 7, 1, 1));

        int endRow = 9 + dynamicObjects.size() + 2;
        Row eRow = sheet.createRow(endRow);
        eRow.createCell(4).setCellValue("NGƯỜI KIỂM SOÁT");
        eRow.createCell(9).setCellValue("ĐẠI DIỆN CHI NHÁNH");

        eRow.getCell(4).setCellStyle(styleBold);
        eRow.getCell(9).setCellStyle(styleBold);

        excelGenerator.writeExcel(fileName);
    }

    public static void ISS005Report() throws SQLException, FileNotFoundException, IOException {
        System.out.println("Generating ISS005 Report");
        Date date = new Date();
        DateFormat dateFNFormat = new SimpleDateFormat("ddMMyyyy");
        String dateFN = dateFNFormat.format(date);
        String fileName = "/FTPData/HSC/ISS_005_" + dateFN + ".xlsx";
        DynamicObject dynamicObject = new DynamicObject();
        Map<String, String> columns = new LinkedHashMap<>();
        columns.put("CARD_NBR", "Số thẻ");
        // columns.put("ACCT.CARD_NBR", "Cadencie");
        // columns.put("card_number", "Way4");
        // columns.put("SS_ST", "So sánh (Pass/ Fail)");
        columns.put("branch", "Cadencie");
        columns.put("W4_BRANCH", "Way4");
        columns.put("SS_CN_PH", "So sánh (Pass/ Fail)");
        columns.put("STKT", "Cadencie");
        columns.put("CONTRACT_NUMBER", "Way4");
        columns.put("SS_STK_THE", "So sánh (Pass/ Fail)");
        columns.put("CUSTR_REF", "Cadencie");
        columns.put("CLIENT__ID", "Way4");
        columns.put("SS_CIF", "So sánh (Pass/ Fail)");
        columns.put("EMBOSS_NME", "Cadencie");
        columns.put("CONTRACT_NAME", "Way4");
        columns.put("SS_TEN_CT", "So sánh (Pass/ Fail)");
        columns.put("ISSUE_DAY", "Cadencie");
        columns.put("DATE_OPEN", "Way4");
        columns.put("SS_NGAY_PH", "So sánh (Pass/ Fail)");
        columns.put("EXPIRY_DTE", "Cadencie");
        columns.put("CARD_EXPIRE", "Way4");
        columns.put("SS_THOIGIAN_HIEU_LUC", "So sánh (Pass/ Fail)");
        columns.put("CANCL_CODE", "Cadencie");
        columns.put("CONTR_STATUS", "Way4");
        columns.put("SS_TRANGTHAI_THE", "So sánh (Pass/ Fail)");
        columns.put("ISS_SERIAL", "Cadencie");
        columns.put("ADD_INFO_01", "Way4");
        columns.put("SS_ISN", "So sánh (Pass/ Fail)");
        columns.put("HOLD_REAS", "Cadencie");
        columns.put("STATUS_TYPE_CODE", "Way4");
        columns.put("SS_GIAHAN_THE", "So sánh (Pass/ Fail)");
        columns.put("PIN_FAILS", "Cadencie");
        columns.put("PIN2", "Way4");
        columns.put("SS_PIN_FAIL", "So sánh (Pass/ Fail)");
        columns.put("limit_code", "Cadencie");
        columns.put("limit_code2", "Way4");
        columns.put("SS_LIMIT_CODE", "So sánh (Pass/ Fail)");
        columns.put("FEE_MONTH", "Cadencie");
        columns.put("ADD_INFO_02", "Way4");
        columns.put("SS_FEE", "So sánh (Pass/ Fail)");

        dynamicObject.setColumns(columns);
        Map<Integer, Object> inputParams = new HashMap<>();

        List<Integer> outParams = new ArrayList<>();
        outParams.add(1);

        List<DynamicObject> dynamicObjects = databaseService.callProcedure("REPORT_MIGRATE", "ISS_005", columns,
                inputParams, outParams);

        if (dynamicObjects.size() == 0) {
            DynamicObject dynamicObject1 = new DynamicObject();
            dynamicObject1.setColumns(columns);
            dynamicObject1.setProperties(new LinkedHashMap<>());
            dynamicObjects.add(dynamicObject1);
        } else {
            // SS_LIMIT_CODE null -> set to 0
            // for (DynamicObject dynamicObject2 : dynamicObjects) {
            // if (dynamicObject2.getProperties().get("SS_LIMIT_CODE") == null) {
            // dynamicObject2.getProperties().put("SS_LIMIT_CODE", "");
            // }
            // }

            // check if any column is null -> set to ""
            for (DynamicObject dynamicObject2 : dynamicObjects) {
                for (String key : columns.keySet()) {
                    if (dynamicObject2.getProperties().get(key) == null) {
                        dynamicObject2.getProperties().put(key, "");
                        System.out.println(key);
                    }
                }
            }
        }
        ExcelGenerator excelGenerator = new ExcelGenerator();

        Sheet sheet = excelGenerator.generateExcel(7, dynamicObjects, true);

        // title row
        Row titleRow = sheet.createRow(0);
        titleRow.createCell(0)
                .setCellValue(
                        "BÁO CÁO ĐỐI CHIẾU DỮ LIỆU THÔNG TIN THẺ TẠI CADENCIE - WAY4");

        // font bold, size 16 for title
        Font font = sheet.getWorkbook().createFont();
        font.setBold(true);
        font.setFontHeightInPoints((short) 16);
        CellStyle style = sheet.getWorkbook().createCellStyle();
        style.setFont(font);
        style.setAlignment(HorizontalAlignment.CENTER);
        titleRow.getCell(0).setCellStyle(style);
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 14));

        Row row1 = sheet.createRow(1);
        row1.createCell(0).setCellValue("Mã chi nhánh: ");
        row1.createCell(9).setCellValue("Mã báo cáo: ISS_005");
        Row row2 = sheet.createRow(2);
        row2.createCell(0).setCellValue("Tên chi nhánh: ");
        DateFormat dateFormat = new SimpleDateFormat("dd/MM/yyyy");
        String dateStr = dateFormat.format(date);
        Row row3 = sheet.createRow(3);
        row3.createCell(4).setCellValue("Ngày báo cáo: " + dateStr);

        Font fontBold = sheet.getWorkbook().createFont();
        fontBold.setBold(true);
        CellStyle styleBold = sheet.getWorkbook().createCellStyle();
        styleBold.setFont(fontBold);

        row1.getCell(0).setCellStyle(styleBold);
        row1.getCell(9).setCellStyle(styleBold);
        row2.getCell(0).setCellStyle(styleBold);
        row3.getCell(4).setCellStyle(styleBold);

        Row headerRow = sheet.createRow(5);
        Row headerRow2 = sheet.createRow(6);
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
        headerRow2.createCell(0).setCellStyle(cellStyle);
        headerRow.createCell(1).setCellValue((String) columns.values().toArray()[0]);
        headerRow.getCell(1).setCellStyle(cellStyle);
        headerRow2.createCell(1).setCellValue((String) columns.values().toArray()[0]);
        headerRow2.getCell(1).setCellStyle(cellStyle);
        headerRow.getCell(0).setCellStyle(cellStyle);

        String[] header = { "CN PHT", "Số tài khoản thẻ", "Số CIF", "Tên chủ thẻ", "Ngày phát hành",
                "Thời hạn hiệu lực", "Trạng thái thẻ", "Issue Serial Number", "Hold Reason Code/ Gia hạn thẻ",
                "Số lần nhập sai PIN", "Hạn mức giao dịch thẻ",
                "Fee month" };

        for (int i = 2; i < columns.size() + 1; i++) {
            headerRow.createCell(i).setCellStyle(cellStyle);
            headerRow2.createCell(i).setCellValue((String) columns.values().toArray()[i - 1]);
            sheet.autoSizeColumn(i);
            headerRow2.getCell(i).setCellStyle(cellStyle);
        }

        for (int i = 2; i < columns.size(); i += 3) {
            headerRow.getCell(i).setCellValue(header[i / 3]);
            sheet.addMergedRegion(new CellRangeAddress(5, 5, i, i + 2));
        }

        sheet.addMergedRegion(new CellRangeAddress(5, 6, 0, 0));
        sheet.addMergedRegion(new CellRangeAddress(5, 6, 1, 1));

        int endRow = 9 + dynamicObjects.size() + 2;
        Row eRow = sheet.createRow(endRow);
        eRow.createCell(3).setCellValue("LẬP BẢNG");
        eRow.createCell(6).setCellValue("NGƯỜI KIỂM SOÁT");
        eRow.createCell(9).setCellValue("ĐẠI DIỆN CHI NHÁNH");

        eRow.getCell(3).setCellStyle(styleBold);
        eRow.getCell(6).setCellStyle(styleBold);
        eRow.getCell(9).setCellStyle(styleBold);

        excelGenerator.writeExcel(fileName);
    }

    public static void ISS003Report() throws SQLException, FileNotFoundException, IOException {
        System.out.println("Generating ISS003 Report");
        Date date = new Date();
        DateFormat dateFNFormat = new SimpleDateFormat("ddMMyyyy");
        String dateFN = dateFNFormat.format(date);
        String fileName = "/FTPData/HSC/ISS_003_" + dateFN + ".xlsx";
        DynamicObject dynamicObject = new DynamicObject();
        Map<String, String> columns = new LinkedHashMap<>();
        columns.put("CUSTR_REF", "Số CIF");
        // columns.put("ACCT.CARD_NBR", "Cadencie");
        // columns.put("card_number", "Way4");
        // columns.put("SS_ST", "So sánh (Pass/ Fail)");
        columns.put("ZBOC", "Cadencie");
        columns.put("BRANCH", "Way4");
        columns.put("SS_CNPHT", "So sánh (Pass/ Fail)");
        columns.put("BUSI_NAME", "Cadencie");
        columns.put("SHORT_NAME", "Way4");
        columns.put("SS_TKH", "So sánh (Pass/ Fail)");
        columns.put("OIN", "Cadencie");
        columns.put("REG_NUMBER", "Way4");
        columns.put("SS_MSDKKD", "So sánh (Pass/ Fail)");
        columns.put("OED", "Cadencie");
        columns.put("ADD_DATE_01", "Way4");
        columns.put("SS_NHH_DKKD", "So sánh (Pass/ Fail)");
        columns.put("taxid", "Cadencie");
        columns.put("ITN", "Way4");
        columns.put("SS_MST", "So sánh (Pass/ Fail)");
        columns.put("NGUOI_DAI_DIEN", "Cadencie");
        columns.put("NGUOI_DAI_DIEN_W4", "Way4");
        columns.put("SS_NGUOI_DAI_DIEN", "So sánh (Pass/ Fail)");
        columns.put("SDT", "Cadencie");
        columns.put("SDTW4", "Way4");
        columns.put("SS_SDT", "So sánh (Pass/ Fail)");
        columns.put("email", "Cadencie");
        columns.put("e_mail", "Way4");
        columns.put("SS_EMAIL", "So sánh (Pass/ Fail)");
        columns.put("stat", "Cadencie");
        columns.put("CIF_STATUS", "Way4");
        columns.put("SS_CIF_STATUS", "So sánh (Pass/ Fail)");
        columns.put("ccode", "Cadencie");
        columns.put("CLIENT_SEC_CS", "Way4");
        columns.put("SS_HANG_KH", "So sánh (Pass/ Fail)");
        columns.put("darcovr", "Cadencie");
        columns.put("CBS_LOAN_GROUP_CS", "Way4");
        columns.put("SS_NHOM_NO", "So sánh (Pass/ Fail)");

        dynamicObject.setColumns(columns);
        Map<Integer, Object> inputParams = new HashMap<>();

        List<Integer> outParams = new ArrayList<>();
        outParams.add(1);

        List<DynamicObject> dynamicObjects = databaseService.callProcedure("REPORT_MIGRATE", "ISS_003", columns,
                inputParams, outParams);

        if (dynamicObjects.size() == 0) {
            DynamicObject dynamicObject1 = new DynamicObject();
            dynamicObject1.setColumns(columns);
            dynamicObject1.setProperties(new LinkedHashMap<>());
            dynamicObjects.add(dynamicObject1);
        } else {

        }
        ExcelGenerator excelGenerator = new ExcelGenerator();

        Sheet sheet = excelGenerator.generateExcel(7, dynamicObjects, true);

        // title row
        Row titleRow = sheet.createRow(0);
        titleRow.createCell(0)
                .setCellValue(
                        "BÁO CÁO ĐỐI CHIẾU DỮ LIỆU CIF KHDN TẠI CADENCIE - WAY4");

        // font bold, size 16 for title
        Font font = sheet.getWorkbook().createFont();
        font.setBold(true);
        font.setFontHeightInPoints((short) 16);
        CellStyle style = sheet.getWorkbook().createCellStyle();
        style.setFont(font);
        style.setAlignment(HorizontalAlignment.CENTER);
        titleRow.getCell(0).setCellStyle(style);
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 14));

        Row row1 = sheet.createRow(1);
        row1.createCell(0).setCellValue("Mã chi nhánh: ");
        row1.createCell(9).setCellValue("Mã báo cáo: ISS_003");
        Row row2 = sheet.createRow(2);
        row2.createCell(0).setCellValue("Tên chi nhánh: ");
        DateFormat dateFormat = new SimpleDateFormat("dd/MM/yyyy");
        String dateStr = dateFormat.format(date);
        Row row3 = sheet.createRow(3);
        row3.createCell(4).setCellValue("Ngày báo cáo: " + dateStr);

        Font fontBold = sheet.getWorkbook().createFont();
        fontBold.setBold(true);
        CellStyle styleBold = sheet.getWorkbook().createCellStyle();
        styleBold.setFont(fontBold);

        row1.getCell(0).setCellStyle(styleBold);
        row1.getCell(9).setCellStyle(styleBold);
        row2.getCell(0).setCellStyle(styleBold);
        row3.getCell(4).setCellStyle(styleBold);

        Row headerRow = sheet.createRow(5);
        Row headerRow2 = sheet.createRow(6);
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
        headerRow2.createCell(0).setCellStyle(cellStyle);
        headerRow.createCell(1).setCellValue((String) columns.values().toArray()[0]);
        headerRow.getCell(1).setCellStyle(cellStyle);
        headerRow2.createCell(1).setCellValue((String) columns.values().toArray()[0]);
        headerRow2.getCell(1).setCellStyle(cellStyle);
        headerRow.getCell(0).setCellStyle(cellStyle);

        String[] header = { "CN PHT", "Tên Doanh nghiệp tiếng Việt", "Mã số ĐKKD",
                "Ngày hết hạn ĐKKD",
                "Mã số thuế", "Người đại diện", "Số điện thoại", "Email",
                "Trạng thái CIF", "Phân đoạn KH",
                "Nhóm nợ Core Prf" };

        for (int i = 2; i < columns.size() + 1; i++) {
            headerRow.createCell(i).setCellStyle(cellStyle);
            headerRow2.createCell(i).setCellValue((String) columns.values().toArray()[i - 1]);
            sheet.autoSizeColumn(i);
            headerRow2.getCell(i).setCellStyle(cellStyle);
        }

        for (int i = 2; i < columns.size(); i += 3) {
            headerRow.getCell(i).setCellValue(header[i / 3]);
            sheet.addMergedRegion(new CellRangeAddress(5, 5, i, i + 2));
        }

        sheet.addMergedRegion(new CellRangeAddress(5, 6, 0, 0));
        sheet.addMergedRegion(new CellRangeAddress(5, 6, 1, 1));

        int endRow = 9 + dynamicObjects.size() + 2;
        Row eRow = sheet.createRow(endRow);
        eRow.createCell(3).setCellValue("LẬP BẢNG");
        eRow.createCell(6).setCellValue("NGƯỜI KIỂM SOÁT");
        eRow.createCell(9).setCellValue("ĐẠI DIỆN CHI NHÁNH");

        eRow.getCell(3).setCellStyle(styleBold);
        eRow.getCell(6).setCellStyle(styleBold);
        eRow.getCell(9).setCellStyle(styleBold);

        excelGenerator.writeExcel(fileName);
    }

    public static void ISS013Report(Branch branch) throws SQLException, FileNotFoundException, IOException {
        System.out.println("Generating ISS013 Report");
        Date date = new Date();
        DateFormat dateFNFormat = new SimpleDateFormat("ddMMyyyy");
        String dateFN = dateFNFormat.format(date);
        String fileName = branch.getFolderPath() + "ISS_013_" + branch.getBranchCode() + "_" + branch.getBranchName()
                + "_" + dateFN
                + ".xlsx";
        DynamicObject dynamicObject = new DynamicObject();
        Map<String, String> columns = new LinkedHashMap<>();
        columns.put("acct_style", "Loại thẻ");
        columns.put("branch", "CN quản lý");
        columns.put("custr_ref", "Số CIF");
        columns.put("account", "Số tài khoản thẻ");
        columns.put("acc_name1", "Tên tài khoản");
        columns.put("CLOSE_CODE", "Trạng thái tài khoản");
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
        // columns.put("query_amt", "Số tiền khiếu nại đang treo");
        columns.put("TONG_SO_DU_GD_TRAGOP", "Tổng số dư giao dịch trả góp");
        columns.put("TONG_SO_GD_TRAGOP_DANG_HOAT_DONG", "Tổng số giao dịch trả góp đang hoạt động");
        columns.put("TONG_GIATRI_GD_CHUA_LEN_DU_NO", "Tổng giá trị các giao dịch chưa lên dư nợ");
        columns.put("TONG_DOANH_SO_THANH_TOAN", "Tổng Doanh số thanh toán");
        dynamicObject.setColumns(columns);

        Map<Integer, Object> inputParams = new HashMap<>();
        inputParams.put(1, branch.getBranchCode());
        List<Integer> outParams = new ArrayList<>();
        outParams.add(2);
        // List<DynamicObject> dynamicObjects =
        // databaseService.callProcedure("REPORT_MIGRATE", "ISS_013", columns,
        // inputParams, outParams);
        List<DynamicObject> dynamicObjects = new ArrayList<>();
        if (dynamicObjects.size() == 0) {
            DynamicObject dynamicObject1 = new DynamicObject();
            dynamicObject1.setColumns(columns);
            dynamicObject1.setProperties(new LinkedHashMap<>());
            dynamicObjects.add(dynamicObject1);
        }
        // else {
        // dynamicObjects.forEach(dynamicObject1 -> {
        // if (dynamicObject1.getProperties().get("class_code") == null) {
        // dynamicObject1.getProperties().put("class_code", "");
        // }
        // });
        // }
        ExcelGenerator excelGenerator = new ExcelGenerator();
        Sheet sheet = excelGenerator.generateExcel(6, dynamicObjects, false);
        // title row 0
        Row titleRow = sheet.createRow(0);
        titleRow.createCell(0)
                .setCellValue("BÁO CÁO CHI TIẾT THÔNG TIN THẺ CHUYỂN ĐỔI THÀNH CÔNG");
        // font bold, size 16 for title
        Font font = sheet.getWorkbook().createFont();
        font.setBold(true);
        font.setFontHeightInPoints((short) 16);
        CellStyle style = sheet.getWorkbook().createCellStyle();
        style.setFont(font);
        style.setAlignment(HorizontalAlignment.CENTER);
        titleRow.getCell(0).setCellStyle(style);
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 35));

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
        row2.createCell(15).setCellValue("Mã bc: ISS_013");
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

    public static void ISS006Report() throws FileNotFoundException, IOException {
        System.out.println("Generating ISS006 Report");
        Date date = new Date();
        DateFormat dateFNFormat = new SimpleDateFormat("ddMMyyyy");
        String dateFN = dateFNFormat.format(date);
        String fileName = "/FTPData/HSC/ISS_006_" + dateFN + ".xlsx";
        DynamicObject dynamicObject = new DynamicObject();
        Map<String, String> columns = new LinkedHashMap<>();
        columns.put("CUSTR_REF", "Loại thẻ");
        columns.put("BRANCH", "Tiêu chí");
        columns.put("ACCOUNT", "Số lượng giao dịch");
        columns.put("ACC_NAME1", "Tổng giá trị giao dịch");
        columns.put("CLOSE_CODE", "Số lượng giao dịch");
        columns.put("CARD_NBR", "Tổng số lượng giao dịch");
        columns.put("CANCL_CODE", "Cadencie");
        columns.put("PRODUCT", "Way4");
        columns.put("REFERENCE", "Cadencie");
        columns.put("CRED_LIMIT", "Way4");
        columns.put("CLASS_CODE", "Số lượng");
        columns.put("INT_CODE", "Giá trị");

        dynamicObject.setColumns(columns);
        Map<Integer, Object> inputParams = new HashMap<>();
        List<Integer> outParams = new ArrayList<>();
        outParams.add(1);
        List<DynamicObject> dynamicObjects = new ArrayList<>();
        if (dynamicObjects.size() == 0) {
            DynamicObject dynamicObject1 = new DynamicObject();
            dynamicObject1.setColumns(columns);
            dynamicObject1.setProperties(new LinkedHashMap<>());
            dynamicObjects.add(dynamicObject1);
        }

        ExcelGenerator excelGenerator = new ExcelGenerator();
        Sheet sheet = excelGenerator.generateExcel(6, dynamicObjects, true);

        // title row
        Row titleRow = sheet.createRow(0);
        titleRow.createCell(0)
                .setCellValue(
                        "<ISS_006>: BÁO CÁO TỔNG HỢP SỐ LIỆU  KHÁCH HÀNG THẺ QUỐC TẾ CHUYỂN ĐỔI TOÀN HỆ THỐNG");
        // font bold, size 16 for title
        Font font = sheet.getWorkbook().createFont();
        font.setBold(true);
        font.setFontHeightInPoints((short) 16);
        CellStyle style = sheet.getWorkbook().createCellStyle();
        style.setFont(font);
        style.setAlignment(HorizontalAlignment.CENTER);
        titleRow.getCell(0).setCellStyle(style);
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 12));

        String[] header = { "Loại thẻ", "Tiêu chí", "Dữ liệu tổng trên Cadencie", "Dữ liệu không chuyển đổi",
                "Số lượng giao dịch chuyển đổi", "Tổng giá trị giao dịch chuyển đổi", "Chênh lệch giao dịch" };

        Row headerRow = sheet.createRow(5);
        Row headerRow2 = sheet.createRow(6);
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
        headerRow2.createCell(0).setCellStyle(cellStyle);
        headerRow.createCell(1).setCellValue((String) columns.values().toArray()[0]);
        headerRow.getCell(1).setCellStyle(cellStyle);
        headerRow2.createCell(1).setCellValue((String) columns.values().toArray()[0]);
        headerRow2.getCell(1).setCellStyle(cellStyle);
        headerRow.getCell(0).setCellStyle(cellStyle);
        headerRow.createCell(2).setCellValue((String) columns.values().toArray()[1]);
        headerRow.getCell(2).setCellStyle(cellStyle);
        headerRow2.createCell(2).setCellValue((String) columns.values().toArray()[1]);
        headerRow2.getCell(2).setCellStyle(cellStyle);

        for (int i = 3; i < columns.size() + 1; i++) {
            headerRow.createCell(i).setCellStyle(cellStyle);
            headerRow2.createCell(i).setCellValue((String) columns.values().toArray()[i - 1]);
            sheet.autoSizeColumn(i);
            headerRow2.getCell(i).setCellStyle(cellStyle);
        }

        for (int i = 3; i < columns.size(); i += 2) {
            headerRow.getCell(i).setCellValue(header[i / 2 + 1]);
            sheet.addMergedRegion(new CellRangeAddress(5, 5, i, i + 1));
        }

        sheet.addMergedRegion(new CellRangeAddress(5, 6, 0, 0));
        sheet.addMergedRegion(new CellRangeAddress(5, 6, 1, 1));
        sheet.addMergedRegion(new CellRangeAddress(5, 6, 2, 2));

        // total row
        int endRow = 6 + dynamicObjects.size();
        Row eRow = sheet.createRow(endRow);
        eRow.createCell(0).setCellValue("Tổng cộng");
        eRow.getCell(0).setCellStyle(cellStyle);
        eRow.createCell(1).setCellStyle(cellStyle);
        eRow.createCell(2).setCellStyle(cellStyle);
        for (int i = 3; i < columns.size() + 1; i++) {
            eRow.createCell(i).setCellStyle(cellStyle);
        }

        excelGenerator.writeExcel(fileName);
    }

    public static void ISS007Report() throws FileNotFoundException, IOException {
        System.out.println("Generating ISS007 Report");
        Date date = new Date();
        DateFormat dateFNFormat = new SimpleDateFormat("ddMMyyyy");
        String dateFN = dateFNFormat.format(date);
        String fileName = "/FTPData/HSC/ISS_007_" + dateFN + ".xlsx";
        DynamicObject dynamicObject = new DynamicObject();
        Map<String, String> columns = new LinkedHashMap<>();
        columns.put("AUTH.CARD_PLAN", "Loại thẻ");
        columns.put("tieuchi", "Tiêu chí");
        columns.put("AUTH.CARD_NBR", "Cadencie");
        columns.put("DOC.TARGET_NUMBER", "Way4");
        columns.put("ss1", "So sánh (Pass/ Fail)");
        columns.put("AUTH.RETRV_REF", "Cadencie");
        columns.put("DOC.RET_REF_NUMBER", "Way4");
        columns.put("ss2", "So sánh (Pass/ Fail)");
        columns.put("AUTH.AUTH_CODE", "Cadencie");
        columns.put("DOC.AUTH_CODE", "Way4");
        columns.put("ss3", "So sánh (Pass/ Fail)");
        columns.put("AUTH.ISS_AMT", "Cadencie");
        columns.put("DOC.TRANS_AMOUNT", "Way4");
        columns.put("ss4", "So sánh (Pass/ Fail)");
        columns.put("AUTH.BILL_AMT", "Cadencie");
        columns.put("DOC.SETTL_AMOUNT", "Way4");
        columns.put("ss5", "So sánh (Pass/ Fail)");
        columns.put("AUTH.RETRV_REF", "Cadencie");
        columns.put("DOC.RET_REF_NUMBER", "Way4");
        columns.put("ss6", "So sánh (Pass/ Fail)");
        columns.put("AUTH.ACQUIR_ID", "Cadencie");
        columns.put("DOC.SOURCE_MEMBER_ID", "Way4");
        columns.put("ss7", "So sánh (Pass/ Fail)");
        columns.put("AUTH.ACPTOR_ID", "Cadencie");
        columns.put("DOC.MERCHANT_ID", "Way4");
        columns.put("ss8", "So sánh (Pass/ Fail)");
        columns.put("AUTH.ACCOUNT", "Cadencie");
        columns.put("DOC.TARGET_NUMBER", "Way4");
        columns.put("ss9", "So sánh (Pass/ Fail)");

        dynamicObject.setColumns(columns);
        Map<Integer, Object> inputParams = new HashMap<>();
        List<Integer> outParams = new ArrayList<>();
        outParams.add(1);
        List<DynamicObject> dynamicObjects = new ArrayList<>();
        if (dynamicObjects.size() == 0) {
            DynamicObject dynamicObject1 = new DynamicObject();
            dynamicObject1.setColumns(columns);
            dynamicObject1.setProperties(new LinkedHashMap<>());
            dynamicObjects.add(dynamicObject1);
        }

        ExcelGenerator excelGenerator = new ExcelGenerator();
        Sheet sheet = excelGenerator.generateExcel(6, dynamicObjects, true);

        // title row
        Row titleRow = sheet.createRow(1);
        titleRow.createCell(0)
                .setCellValue(
                        "BÁO CÁO ĐỐI CHIẾU DỮ LIỆU GIAO DỊCH THẺ TẠI CADENCIE - WAY4");
        // font bold, size 16 for title
        Font font = sheet.getWorkbook().createFont();
        font.setBold(true);
        font.setFontHeightInPoints((short) 16);
        CellStyle style = sheet.getWorkbook().createCellStyle();
        style.setFont(font);
        style.setAlignment(HorizontalAlignment.CENTER);
        titleRow.getCell(0).setCellStyle(style);
        sheet.addMergedRegion(new CellRangeAddress(1, 1, 0, 14));

        // date now format dd/MM/yyyy
        DateFormat dateFormat = new SimpleDateFormat("dd/MM/yyyy");
        String dateStr = dateFormat.format(date);

        // header row 2
        Row row2 = sheet.createRow(2);
        row2.createCell(3).setCellValue("Mã báo cáo: ISS_007");
        // header row 3
        Row row3 = sheet.createRow(3);
        row3.createCell(3).setCellValue("Ngày báo cáo: " + dateStr);

        Font fontBold = sheet.getWorkbook().createFont();
        fontBold.setBold(true);
        CellStyle styleBold = sheet.getWorkbook().createCellStyle();
        styleBold.setFont(fontBold);

        row2.getCell(3).setCellStyle(styleBold);
        row3.getCell(3).setCellStyle(styleBold);

        String[] header = { "Loại thẻ", "Tiêu chí", "Số thẻ masked", "Chuyển đổi",
                "Số authcode", "Số tiền giao dịch gốc", "Số tiền giao dịch quy đổi", "Số Refference number", "TID",
                "MID", "Số tài khoản hạch toán thẻ GNQT" };

        Row headerRow = sheet.createRow(5);
        Row headerRow2 = sheet.createRow(6);
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
        headerRow2.createCell(0).setCellStyle(cellStyle);
        headerRow.createCell(1).setCellValue((String) columns.values().toArray()[0]);
        headerRow.getCell(1).setCellStyle(cellStyle);
        headerRow2.createCell(1).setCellValue((String) columns.values().toArray()[0]);
        headerRow2.getCell(1).setCellStyle(cellStyle);
        headerRow.getCell(0).setCellStyle(cellStyle);
        headerRow.createCell(2).setCellValue((String) columns.values().toArray()[1]);
        headerRow.getCell(2).setCellStyle(cellStyle);
        headerRow2.createCell(2).setCellValue((String) columns.values().toArray()[1]);
        headerRow2.getCell(2).setCellStyle(cellStyle);

        for (int i = 3; i < columns.size() + 1; i++) {
            headerRow.createCell(i).setCellStyle(cellStyle);
            headerRow2.createCell(i).setCellValue((String) columns.values().toArray()[i - 1]);
            sheet.autoSizeColumn(i);
            headerRow2.getCell(i).setCellStyle(cellStyle);
        }

        for (int i = 3; i < columns.size(); i += 3) {
            headerRow.getCell(i).setCellValue(header[i / 3]);
            sheet.addMergedRegion(new CellRangeAddress(5, 5, i, i + 2));
        }

        sheet.addMergedRegion(new CellRangeAddress(5, 6, 0, 0));
        sheet.addMergedRegion(new CellRangeAddress(5, 6, 1, 1));
        sheet.addMergedRegion(new CellRangeAddress(5, 6, 2, 2));

        excelGenerator.writeExcel(fileName);
    }

    public static void ATM003REPORT() throws FileNotFoundException, IOException, SQLException {
        System.out.println("Generating ATM003 Report");
        Date date = new Date();
        DateFormat dateFNFormat = new SimpleDateFormat("ddMMyyyy");
        String dateFN = dateFNFormat.format(date);
        String fileName = "/FTPData/HSC/ATM_003_" + dateFN + ".xlsx";
        DynamicObject dynamicObject = new DynamicObject();
        Map<String, String> columns = new LinkedHashMap<>();
        columns.put("CONTRACT_NUMBER", "Terminal ID");
        columns.put("BRANCH", "Branch quản lý");
        columns.put("BRANCH_TIEP_QUY", "Branch tiếp quỹ");
        columns.put("MA_AM_QL_MAY", "Mã AM quản lý máy");
        columns.put("ATM_TYPE", "ATM TYPE");
        columns.put("BRAND", "Hãng ATM");
        columns.put("MODEL", "Model");
        columns.put("LOCATION", "ATM location");
        columns.put("MENH_GIA_HOP_TIEN_1", "Mệnh giá hộp tiền 1");
        columns.put("MENH_GIA_HOP_TIEN_2", "Mệnh giá hộp tiền 2");
        columns.put("MENH_GIA_HOP_TIEN_3", "Mệnh giá hộp tiền 3");
        columns.put("MENH_GIA_HOP_TIEN_4", "Mệnh giá hộp tiền 4");
        columns.put("ATM_GROUP", "ATM group");
        // columns.put("ATM_STATUS_NOTE", "%Phí Visa off us nước ngoài");
        // columns.put("ATM_STATUS_TIME", "Phí visa off us nước ngoài Min");
        // columns.put("ATM_STATUS_USER_IP", "%Phí MC off us nước ngoài");
        // columns.put("ATM_STATUS_USER_MAC", "%Phí MC off us nước ngoài Min");

        dynamicObject.setColumns(columns);

        Map<Integer, Object> inputParams = new HashMap<>();
        List<Integer> outParams = new ArrayList<>();
        outParams.add(1);
        List<DynamicObject> dynamicObjects = databaseService.callProcedure("REPORT_MIGRATE", "ATM_003", columns,
                inputParams, outParams);

        if (dynamicObjects.size() == 0) {
            DynamicObject dynamicObject1 = new DynamicObject();
            dynamicObject1.setColumns(columns);
            dynamicObject1.setProperties(new LinkedHashMap<>());
            dynamicObjects.add(dynamicObject1);
        }

        ExcelGenerator excelGenerator = new ExcelGenerator();
        Sheet sheet = excelGenerator.generateExcel(5, dynamicObjects, false);

        // ma bao cao
        Row row1 = sheet.createRow(0);
        row1.createCell(0).setCellValue("Mã báo cáo: ATM_003");
        // title row 1
        Row titleRow = sheet.createRow(1);
        titleRow.createCell(0).setCellValue("BÁO CÁO CHI TIẾT THÔNG TIN ATM");
        // font bold, size 16 for title
        Font font = sheet.getWorkbook().createFont();
        font.setBold(true);
        font.setFontHeightInPoints((short) 16);
        CellStyle style = sheet.getWorkbook().createCellStyle();
        style.setFont(font);
        style.setAlignment(HorizontalAlignment.CENTER);
        titleRow.getCell(0).setCellStyle(style);
        sheet.addMergedRegion(new CellRangeAddress(1, 1, 0, columns.size()));

        // no date

        excelGenerator.writeExcel(fileName);

    }

    public static void ISS008Report() throws FileNotFoundException, IOException {
        System.out.println("Generating ISS008 Report");
        Date date = new Date();
        DateFormat dateFNFormat = new SimpleDateFormat("ddMMyyyy");
        String dateFN = dateFNFormat.format(date);
        String fileName = "/FTPData/HSC/ISS_008_" + dateFN + ".xlsx";
        DynamicObject dynamicObject = new DynamicObject();
        Map<String, String> columns = new LinkedHashMap<>();
        columns.put("TDQT/GNQT", "Số tài khoản thẻ");
        columns.put("1", "Cadencie");
        columns.put("2", "Way4");
        columns.put("ss1", "So sánh (Pass/ Fail)");
        columns.put("3", "Cadencie");
        columns.put("4", "Way4");
        columns.put("ss2", "So sánh (Pass/ Fail)");
        columns.put("5", "Cadencie");
        columns.put("6", "Way4");
        columns.put("ss3", "So sánh (Pass/ Fail)");
        columns.put("7", "Cadencie(1)");
        columns.put("8", "\"W4 Open cash(2)");
        columns.put("9", "W4 Open Sale(3)");
        columns.put("10", "W4 Grace cash (4)");
        columns.put("11", "W4 Grace sale (5)");
        columns.put("12", "W4 Close Cash (6)");
        columns.put("13", "W4 Close Sale(7)");
        columns.put("14", "W4 Open principle Instalment(7)");
        columns.put("cl1", "Chênh lệch (8)=(1)-(2)-(3)-(4)- (5)- (6)- (7)- (7')");
        columns.put("15", "Cadencie (9) Lãi");
        columns.put("16", "W4 Grace Interest (10)");
        columns.put("17", "W4 Close Interest (11)");
        columns.put("cl2", "Chênh lệch (12)=(9)-(10)-(11)");
        columns.put("18", "Cadencie (13) Phí");
        columns.put("19", "W4 Open Fee (14)");
        columns.put("20", "W4 Grace Fee(15)");
        columns.put("21", "W4 Close Fee (16)");
        columns.put("22", "W4 Open Fee Instalment (17)");
        columns.put("cl3", "Chênh lệch (18)=(13)-(14)-(15)-(16)- (17)");
        columns.put("23", "Cadencie (19)");
        columns.put("24", "W4 Waiting Principle (20)");
        columns.put("ss4", "So sánh (Pass/ Fail)");
        columns.put("25", "Cadencie");
        columns.put("26", "Way4");
        columns.put("ss5", "So sánh (Pass/ Fail)");
        columns.put("27", "Cadencie");
        columns.put("28", "Way4");
        columns.put("ss6", "So sánh (Pass/ Fail)");
        columns.put("29", "Cadencie");
        columns.put("30", "Way4");
        columns.put("ss7", "So sánh (Pass/ Fail)");
        columns.put("31", "Cadencie");
        columns.put("32", "Way4");
        columns.put("ss8", "So sánh (Pass/ Fail)");
        columns.put("33", "Cadencie");
        columns.put("34", "Way4");
        columns.put("ss9", "So sánh (Pass/ Fail)");

        dynamicObject.setColumns(columns);

        Map<Integer, Object> inputParams = new HashMap<>();
        List<Integer> outParams = new ArrayList<>();
        outParams.add(1);
        List<DynamicObject> dynamicObjects = new ArrayList<>();

        if (dynamicObjects.size() == 0) {
            DynamicObject dynamicObject1 = new DynamicObject();
            dynamicObject1.setColumns(columns);
            dynamicObject1.setProperties(new LinkedHashMap<>());
            dynamicObjects.add(dynamicObject1);
        }

        ExcelGenerator excelGenerator = new ExcelGenerator();

        Sheet sheet = excelGenerator.generateExcel(6, dynamicObjects, true);

        // title row
        Row titleRow = sheet.createRow(0);

        titleRow.createCell(0)
                .setCellValue(
                        "BÁO CÁO ĐỐI CHIẾU DƯ NỢ HÀNG NGÀY TỪ NGÀY 10 ĐẾN NGÀY 20");
        // font bold, size 16 for title
        Font font = sheet.getWorkbook().createFont();
        font.setBold(true);
        font.setFontHeightInPoints((short) 16);
        CellStyle style = sheet.getWorkbook().createCellStyle();
        style.setFont(font);
        style.setAlignment(HorizontalAlignment.CENTER);
        titleRow.getCell(0).setCellStyle(style);

        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 35));

        // date now format dd/MM/yyyy
        DateFormat dateFormat = new SimpleDateFormat("dd/MM/yyyy");

        String dateStr = dateFormat.format(date);

        // header row 1
        Row row1 = sheet.createRow(1);
        row1.createCell(0).setCellValue("Mã báo cáo: ISS_008");
        // header row 2
        Row row2 = sheet.createRow(2);
        row2.createCell(0).setCellValue("Ngày báo cáo: " + dateStr);

        Font fontBold = sheet.getWorkbook().createFont();
        fontBold.setBold(true);
        CellStyle styleBold = sheet.getWorkbook().createCellStyle();
        styleBold.setFont(fontBold);

        row1.getCell(0).setCellStyle(styleBold);
        row2.getCell(0).setCellStyle(styleBold);

        String[] header = { "Số tài khoản thẻ", "CN quản lý", "Số tài khoản thẻ", "Hạn mức tín dụng", "Dư nợ gốc",
                "Dư lãi",
                "Dư phí", "Dư nợ trả góp chưa phân bổ", "Kỳ hạn trả góp còn lại", "Hạn mức khả dụng",
                "Tổng doanh số thanh toán", "Số ngày quá hạn",
                "Số tiền quá hạn" };

        Row headerRow = sheet.createRow(5);
        Row headerRow2 = sheet.createRow(6);
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
        headerRow2.createCell(0).setCellStyle(cellStyle);
        headerRow.createCell(1).setCellValue((String) columns.values().toArray()[0]);
        headerRow.getCell(1).setCellStyle(cellStyle);
        headerRow2.createCell(1).setCellValue((String) columns.values().toArray()[0]);
        headerRow2.getCell(1).setCellStyle(cellStyle);
        headerRow.getCell(0).setCellStyle(cellStyle);

        for (int i = 2; i < columns.size() + 1; i++) {
            headerRow.createCell(i).setCellStyle(cellStyle);
            headerRow2.createCell(i).setCellValue((String) columns.values().toArray()[i - 1]);
            sheet.autoSizeColumn(i);
            headerRow2.getCell(i).setCellStyle(cellStyle);
        }

        for (int i = 2; i < 11; i += 3) {
            headerRow.getCell(i).setCellValue(header[i / 3 + 1]);
            sheet.addMergedRegion(new CellRangeAddress(5, 5, i, i + 2));
        }

        // merge 11 to 19
        sheet.addMergedRegion(new CellRangeAddress(5, 5, 11, 19));
        headerRow.getCell(11).setCellValue("Dư nợ gốc");

        sheet.addMergedRegion(new CellRangeAddress(5, 5, 20, 23));
        headerRow.getCell(20).setCellValue("Dư lãi");

        sheet.addMergedRegion(new CellRangeAddress(5, 5, 24, 29));
        headerRow.getCell(24).setCellValue("Dư phí");

        for (int i = 30; i < columns.size(); i += 3) {
            int j = 6;
            headerRow.getCell(i).setCellValue(header[header.length - j]);
            // last 6 header
            sheet.addMergedRegion(new CellRangeAddress(5, 5, i, i + 2));
            j--;
        }

        sheet.addMergedRegion(new CellRangeAddress(5, 6, 0, 0));
        sheet.addMergedRegion(new CellRangeAddress(5, 6, 1, 1));

        int endRow = 9 + dynamicObjects.size() + 2;
        Row eRow = sheet.createRow(endRow);
        eRow.createCell(3).setCellValue("LẬP BẢNG");
        eRow.createCell(6).setCellValue("NGƯỜI KIỂM SOÁT");
        eRow.createCell(9).setCellValue("ĐẠI DIỆN CHI NHÁNH");

        eRow.getCell(3).setCellStyle(styleBold);
        eRow.getCell(6).setCellStyle(styleBold);
        eRow.getCell(9).setCellStyle(styleBold);

        excelGenerator.writeExcel(fileName);
    }

    public static void ISS0081Report() throws FileNotFoundException, IOException {
        System.out.println("Generating ISS0081 Report");
        Date date = new Date();
        DateFormat dateFNFormat = new SimpleDateFormat("ddMMyyyy");
        String dateFN = dateFNFormat.format(date);
        String fileName = "/FTPData/HSC/ISS_008_1_" + dateFN + ".xlsx";
        DynamicObject dynamicObject = new DynamicObject();
        Map<String, String> columns = new LinkedHashMap<>();
        // columns.put("TDQT/GNQT", "Số tài khoản thẻ");
        columns.put("1", "Cadencie");
        columns.put("2", "Way4");
        columns.put("ss1", "So sánh (Pass/ Fail)");
        columns.put("3", "Cadencie");
        columns.put("4", "Way4");
        columns.put("ss2", "So sánh (Pass/ Fail)");
        columns.put("5", "Cadencie");
        columns.put("6", "Way4");
        columns.put("ss3", "So sánh (Pass/ Fail)");
        columns.put("7", "Cadencie");
        columns.put("8", "Way4");
        columns.put("ss4", "So sánh (Pass/ Fail)");
        columns.put("9", "Cadencie");
        columns.put("10", "Way4");
        columns.put("ss5", "So sánh (Pass/ Fail)");
        columns.put("11", "Cadencie");
        columns.put("12", "Way4");
        columns.put("ss6", "So sánh (Pass/ Fail)");
        columns.put("13", "Cadencie");
        columns.put("14", "Way4");
        columns.put("ss7", "So sánh (Pass/ Fail)");
        columns.put("16", "Cadencie");
        columns.put("17", "Way4");
        columns.put("ss8", "So sánh (Pass/ Fail)");
        columns.put("18", "Cadencie");
        columns.put("19", "Way4");
        columns.put("ss9", "So sánh (Pass/ Fail)");

        dynamicObject.setColumns(columns);

        Map<Integer, Object> inputParams = new HashMap<>();
        List<Integer> outParams = new ArrayList<>();
        outParams.add(1);
        List<DynamicObject> dynamicObjects = new ArrayList<>();

        if (dynamicObjects.size() == 0) {
            DynamicObject dynamicObject1 = new DynamicObject();
            dynamicObject1.setColumns(columns);
            dynamicObject1.setProperties(new LinkedHashMap<>());
            dynamicObjects.add(dynamicObject1);
        }

        ExcelGenerator excelGenerator = new ExcelGenerator();

        Sheet sheet = excelGenerator.generateExcel(7, dynamicObjects, true);

        // title row
        Row titleRow = sheet.createRow(0);
        titleRow.createCell(0)
                .setCellValue(
                        "BÁO CÁO CHI TIẾT GIAO DỊCH TRẢ GÓP");
        // font bold, size 16 for title
        Font font = sheet.getWorkbook().createFont();
        font.setBold(true);
        font.setFontHeightInPoints((short) 16);
        CellStyle style = sheet.getWorkbook().createCellStyle();
        style.setFont(font);

        titleRow.getCell(0).setCellStyle(style);

        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 17));

        // date now format dd/MM/yyyy
        DateFormat dateFormat = new SimpleDateFormat("dd/MM/yyyy");

        String dateStr = dateFormat.format(date);

        // header row 1
        Row row1 = sheet.createRow(1);
        row1.createCell(0).setCellValue("Mã báo cáo: ISS_008_1");
        // header row 2
        Row row2 = sheet.createRow(2);
        row2.createCell(0).setCellValue("Ngày báo cáo: " + dateStr);

        Font fontBold = sheet.getWorkbook().createFont();
        fontBold.setBold(true);
        CellStyle styleBold = sheet.getWorkbook().createCellStyle();
        styleBold.setFont(fontBold);

        row1.getCell(0).setCellStyle(styleBold);
        row2.getCell(0).setCellStyle(styleBold);

        String[] header = { "CN quản lý", "Số tài khoản thẻ", "Số tiền GD gốc", "Số kỳ đăng kí trả góp",
                "Gốc trả góp từng kỳ",
                "Số kỳ đăng kí trả góp", "Gốc trả góp từng kỳ",
                "Dư nợ trả góp chưa lên sao kê", "Ngày kết thúc",
                "Số tiền quá hạn" };

        Row headerRow = sheet.createRow(5);
        Row headerRow2 = sheet.createRow(6);
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
        headerRow2.createCell(0).setCellValue("STT");
        headerRow2.getCell(0).setCellStyle(cellStyle);

        for (int i = 1; i < columns.size() + 1; i++) {
            headerRow.createCell(i).setCellStyle(cellStyle);
            headerRow2.createCell(i).setCellValue((String) columns.values().toArray()[i - 1]);
            sheet.autoSizeColumn(i);
            headerRow2.getCell(i).setCellStyle(cellStyle);
        }

        for (int i = 1; i < columns.size(); i += 3) {
            headerRow.getCell(i).setCellValue(header[i / 3]);
            sheet.addMergedRegion(new CellRangeAddress(5, 5, i, i + 2));
        }

        sheet.addMergedRegion(new CellRangeAddress(5, 6, 0, 0));

        int endRow = 9 + dynamicObjects.size() + 2;
        Row eRow = sheet.createRow(endRow);
        eRow.createCell(3).setCellValue("LẬP BẢNG");
        eRow.createCell(6).setCellValue("NGƯỜI KIỂM SOÁT");
        eRow.createCell(9).setCellValue("ĐẠI DIỆN CHI NHÁNH");

        eRow.getCell(3).setCellStyle(styleBold);
        eRow.getCell(6).setCellStyle(styleBold);
        eRow.getCell(9).setCellStyle(styleBold);

        excelGenerator.writeExcel(fileName);
    }

    public static void ACQ007Report() throws FileNotFoundException, IOException {
        System.out.println("Generating ACQ007 Report");
        Date date = new Date();
        DateFormat dateFNFormat = new SimpleDateFormat("ddMMyyyy");
        String dateFN = dateFNFormat.format(date);
        String fileName = "/FTPData/HSC/ACQ_007_" + dateFN + ".xlsx";
        DynamicObject dynamicObject = new DynamicObject();
        Map<String, String> columns = new LinkedHashMap<>();
        columns.put("0", "STT");
        columns.put("1", "CN quản lý");
        columns.put("2", "Số Cif Khách hàng");
        columns.put("3", "Tên Khách hàng");
        columns.put("4", "Parent Merchant ID");
        columns.put("5", "Merchant ID");
        columns.put("6", "Tên Merchant");
        columns.put("7", "Terminal ID");
        columns.put("8", "MCC");
        columns.put("9", "AM");
        columns.put("10", "RBS number");
        columns.put("11", "MC");
        columns.put("12", "MD");
        columns.put("13", "VC");
        columns.put("14", "VD");
        columns.put("15", "JC");
        columns.put("16", "JD");
        columns.put("17", "PD");
        columns.put("18", "PC");

        dynamicObject.setColumns(columns);

        Map<Integer, Object> inputParams = new HashMap<>();
        List<Integer> outParams = new ArrayList<>();
        outParams.add(1);
        List<DynamicObject> dynamicObjects = new ArrayList<>();

        if (dynamicObjects.size() == 0) {
            DynamicObject dynamicObject1 = new DynamicObject();
            dynamicObject1.setColumns(columns);
            dynamicObject1.setProperties(new LinkedHashMap<>());
            dynamicObjects.add(dynamicObject1);
        }

        ExcelGenerator excelGenerator = new ExcelGenerator();

        Sheet sheet = excelGenerator.generateExcel(5, dynamicObjects, true);

        // title row
        Row titleRow = sheet.createRow(0);
        titleRow.createCell(0)
                .setCellValue(
                        "Báo cáo chi tiết các khách hàng, mid, tid chuyển đổi không thành công (có CAD, không có WAy4)");
        // font bold, size 16 for title
        Font font = sheet.getWorkbook().createFont();
        font.setBold(true);
        font.setFontHeightInPoints((short) 16);
        CellStyle style = sheet.getWorkbook().createCellStyle();
        style.setFont(font);
        style.setAlignment(HorizontalAlignment.CENTER);
        titleRow.getCell(0).setCellStyle(style);
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 18));

        // date now format dd/MM/yyyy
        DateFormat dateFormat = new SimpleDateFormat("dd/MM/yyyy");

        String dateStr = dateFormat.format(date);

        // header row 1
        Row row1 = sheet.createRow(1);
        row1.createCell(0).setCellValue("Mã báo cáo: ACQ_007");
        // header row 2
        Row row2 = sheet.createRow(2);
        row2.createCell(0).setCellValue("Ngày báo cáo: " + dateStr);

        Font fontBold = sheet.getWorkbook().createFont();
        fontBold.setBold(true);
        CellStyle styleBold = sheet.getWorkbook().createCellStyle();
        styleBold.setFont(fontBold);

        row1.getCell(0).setCellStyle(styleBold);
        row2.getCell(0).setCellStyle(styleBold);

        Row headerRow = sheet.createRow(3);
        Row headerRow2 = sheet.createRow(4);
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
        headerRow2.createCell(0).setCellStyle(cellStyle);

        for (int i = 1; i < columns.size() - 8; i++) {
            headerRow.createCell(i).setCellValue((String) columns.values().toArray()[i]);
            // resize column width
            sheet.autoSizeColumn(i);
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
            sheet.autoSizeColumn(i);
            headerRow2.getCell(i).setCellStyle(cellStyle);
        }

        sheet.addMergedRegion(new CellRangeAddress(6, 6, columns.size() - 8, columns.size() - 1));

        excelGenerator.writeExcel(fileName);
    }

    public static void ACQ008Report() throws FileNotFoundException, IOException {
        System.out.println("Generating AQC008 Report");
        Date date = new Date();
        DateFormat dateFNFormat = new SimpleDateFormat("ddMMyyyy");
        String dateFN = dateFNFormat.format(date);
        String fileName = "/FTPData/HSC/AQC_008_" + dateFN + ".xlsx";
        DynamicObject dynamicObject = new DynamicObject();
        Map<String, String> columns = new LinkedHashMap<>();
        columns.put("1", "CN quản lý");
        columns.put("2", "Số Cif Khách hàng");
        columns.put("3", "Tên Khách hàng");
        columns.put("4", "Merchant main ID");
        columns.put("5", "Merchant liên kết (Add_info_03)");
        columns.put("6", "Merchant ID");
        columns.put("7", "Tên Merchant");
        columns.put("8", "Terminal ID");
        columns.put("9", "MCC");
        columns.put("10", "AM");
        columns.put("11", "RBS number");

        dynamicObject.setColumns(columns);

        Map<Integer, Object> inputParams = new HashMap<>();
        List<Integer> outParams = new ArrayList<>();
        outParams.add(1);
        List<DynamicObject> dynamicObjects = new ArrayList<>();

        if (dynamicObjects.size() == 0) {
            DynamicObject dynamicObject1 = new DynamicObject();
            dynamicObject1.setColumns(columns);
            dynamicObject1.setProperties(new LinkedHashMap<>());
            dynamicObjects.add(dynamicObject1);
        }

        ExcelGenerator excelGenerator = new ExcelGenerator();

        Sheet sheet = excelGenerator.generateExcel(5, dynamicObjects, false);

        // title row
        Row titleRow = sheet.createRow(0);
        titleRow.createCell(0)
                .setCellValue(
                        "Báo cáo chi tiết các khách hàng, mid, tid không có tại CAD nhưng lại có tại way4");
        // font bold, size 16 for title
        Font font = sheet.getWorkbook().createFont();
        font.setBold(true);
        font.setFontHeightInPoints((short) 16);
        CellStyle style = sheet.getWorkbook().createCellStyle();
        style.setFont(font);
        style.setAlignment(HorizontalAlignment.CENTER);
        titleRow.getCell(0).setCellStyle(style);
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 11));

        // date now format dd/MM/yyyy
        DateFormat dateFormat = new SimpleDateFormat("dd/MM/yyyy");

        String dateStr = dateFormat.format(date);

        // header row 1
        Row row1 = sheet.createRow(1);
        row1.createCell(0).setCellValue("Mã báo cáo: AQC_008");
        // header row 2

        Row row2 = sheet.createRow(2);
        row2.createCell(0).setCellValue("Ngày báo cáo: " + dateStr);

        Font fontBold = sheet.getWorkbook().createFont();
        fontBold.setBold(true);
        CellStyle styleBold = sheet.getWorkbook().createCellStyle();
        styleBold.setFont(fontBold);

        row1.getCell(0).setCellStyle(styleBold);
        row2.getCell(0).setCellStyle(styleBold);

        excelGenerator.writeExcel(fileName);
    }

    public static void ACQ006Report() throws FileNotFoundException, IOException {
        System.out.println("Generating ACQ006 Report");
        Date date = new Date();
        DateFormat dateFNFormat = new SimpleDateFormat("ddMMyyyy");
        String dateFN = dateFNFormat.format(date);
        String fileName = "/FTPData/HSC/ACQ_006_" + dateFN + ".xlsx";
        DynamicObject dynamicObject = new DynamicObject();
        Map<String, String> columns = new LinkedHashMap<>();

        // columns.put("0", "STT");
        columns.put("1", "Cadencie");
        columns.put("2", "Way4");
        columns.put("3", "Way4");
        columns.put("4", "MC");
        columns.put("5", "MD");
        columns.put("6", "VC");
        columns.put("7", "VD");
        columns.put("8", "JC");
        columns.put("9", "JD");
        columns.put("10", "PD");
        columns.put("11", "PC");
        columns.put("12", "ACI");
        columns.put("13", "ACV");
        columns.put("14", "ADI");
        columns.put("15", "ADV");
        columns.put("16", "API");
        columns.put("17", "APV");
        columns.put("18", "DCI");
        columns.put("19", "DCV");
        columns.put("20", "DDI");
        columns.put("21", "DDV");
        columns.put("22", "DPI");
        columns.put("23", "DPV");
        columns.put("24", "JCI");
        columns.put("25", "JCO");
        columns.put("26", "JCV");
        columns.put("27", "JDI");
        columns.put("28", "JDO");
        columns.put("29", "JDV");
        columns.put("30", "JPI");
        columns.put("31", "JPV");
        columns.put("32", "MCI");
        columns.put("33", "MCO");
        columns.put("34", "MCV");
        columns.put("35", "MDI");
        columns.put("36", "MDO");
        columns.put("37", "MDV");
        columns.put("38", "MPI");
        columns.put("39", "MPV");
        columns.put("40", "NCI");
        columns.put("41", "NDI");
        columns.put("42", "PCV");
        columns.put("43", "PDO");
        columns.put("44", "PDV");
        columns.put("45", "UCI");
        columns.put("46", "UCO");
        columns.put("47", "UCV");
        columns.put("48", "UDI");
        columns.put("49", "UDO");
        columns.put("50", "UDV");
        columns.put("51", "UPI");
        columns.put("52", "UPV");
        columns.put("53", "VCI");
        columns.put("54", "VCO");
        columns.put("55", "VCV");
        columns.put("56", "VDI");
        columns.put("57", "VDO");
        columns.put("58", "VDV");
        columns.put("59", "VPI");
        columns.put("60", "VPO");
        columns.put("61", "VPV");
        columns.put("62", "LDI");
        columns.put("63", "CDI");
        dynamicObject.setColumns(columns);

        Map<Integer, Object> inputParams = new HashMap<>();
        List<Integer> outParams = new ArrayList<>();
        outParams.add(1);
        List<DynamicObject> dynamicObjects = new ArrayList<>();

        if (dynamicObjects.size() == 0) {
            DynamicObject dynamicObject1 = new DynamicObject();
            dynamicObject1.setColumns(columns);
            dynamicObject1.setProperties(new LinkedHashMap<>());
            dynamicObjects.add(dynamicObject1);
        }

        ExcelGenerator excelGenerator = new ExcelGenerator();

        Sheet sheet = excelGenerator.generateExcel(7, dynamicObjects, true);

        // title row
        Row titleRow = sheet.createRow(0);
        titleRow.createCell(0)
                .setCellValue(
                        "Báo cáo chi tiết phí MDR khai báo theo merchant (gắn personal tariff đến từng acnt_contract)");
        // font bold, size 16 for title
        Font font = sheet.getWorkbook().createFont();
        font.setBold(true);
        font.setFontHeightInPoints((short) 16);
        CellStyle style = sheet.getWorkbook().createCellStyle();
        style.setFont(font);
        style.setAlignment(HorizontalAlignment.CENTER);

        titleRow.getCell(0).setCellStyle(style);
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, columns.size()));

        // date now format dd/MM/yyyy
        DateFormat dateFormat = new SimpleDateFormat("dd/MM/yyyy");

        String dateStr = dateFormat.format(date);

        // header row 1
        Row row1 = sheet.createRow(1);
        row1.createCell(0).setCellValue("Mã báo cáo: ACQ_006");
        // header row 2
        Row row2 = sheet.createRow(2);
        row2.createCell(0).setCellValue("Ngày báo cáo: " + dateStr);

        Font fontBold = sheet.getWorkbook().createFont();
        fontBold.setBold(true);
        CellStyle styleBold = sheet.getWorkbook().createCellStyle();
        styleBold.setFont(fontBold);

        row1.getCell(0).setCellStyle(styleBold);
        row2.getCell(0).setCellStyle(styleBold);

        Row headerRow = sheet.createRow(5);
        Row headerRow2 = sheet.createRow(6);
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

        headerRow2.createCell(0).setCellValue("TT");
        headerRow.createCell(0).setCellStyle(cellStyle);
        headerRow2.getCell(0).setCellStyle(cellStyle);

        for (int i = 1; i < columns.size() + 1; i++) {
            headerRow2.createCell(i).setCellValue((String) columns.values().toArray()[i - 1]);
            // resize column width
            sheet.autoSizeColumn(i);
            headerRow2.getCell(i).setCellStyle(cellStyle);
            headerRow.createCell(i).setCellStyle(cellStyle);
        }

        headerRow.getCell(1).setCellValue("Merchant ID");
        headerRow.getCell(3).setCellValue("Contract number");
        headerRow.getCell(4).setCellValue("Phí MDR VND CAD");
        headerRow.getCell(12).setCellValue("Phí MDR VND Way4");

        sheet.addMergedRegion(new CellRangeAddress(5, 5, 1, 2));
        sheet.addMergedRegion(new CellRangeAddress(5, 5, 4, 11));
        sheet.addMergedRegion(new CellRangeAddress(5, 5, 12, columns.size()));

        excelGenerator.writeExcel(fileName);
    }

    public static void ACQ010Report(Branch branch) throws FileNotFoundException, IOException {
        System.out.println("Generating ACQ010 Report");
        Date date = new Date();
        DateFormat dateFNFormat = new SimpleDateFormat("ddMMyyyy");
        String dateFN = dateFNFormat.format(date);
        String fileName = branch.getFolderPath() + "ACQ_010_" + branch.getBranchCode() + "_" + branch.getBranchName()
                + "_" + dateFN + ".xlsx";
        DynamicObject dynamicObject = new DynamicObject();
        Map<String, String> columns = new LinkedHashMap<>();

        columns.put("ACNT_CONTRACT.BRANCH", "CAD");
        columns.put("x1", "Way4");
        columns.put("x2", "CAD");
        columns.put("x3", "Way4");
        columns.put("ACNT_CONTRACT.PRODUCT", "Product");
        columns.put("ACNT_CONTRACT.CONTRACT_NAME", "Main Contract name");
        columns.put("ACNT_CONTRACT.CONTRACT_NUM", "Main Contract number");
        columns.put("x4", "CAD");
        columns.put("x5", "Way4");
        columns.put("x6", "Sub Contract name");
        columns.put("x7", "Sub Contract number");
        columns.put("x8", "CAD");
        columns.put("x9", "Way4");
        columns.put("x10", "CAD");
        columns.put("x11", "Way4");
        columns.put("x12", "CAD");
        columns.put("x13", "Way4");
        columns.put("x14", "CAD");
        columns.put("x15", "Way4");
        columns.put("x16", "CAD");
        columns.put("x17", "Way4");

        columns.put("4", "MC");
        columns.put("5", "MD");
        columns.put("6", "VC");
        columns.put("7", "VD");
        columns.put("8", "JC");
        columns.put("9", "JD");
        columns.put("10", "PD");
        columns.put("11", "PC");
        columns.put("12", "ACI");
        columns.put("13", "ACV");
        columns.put("14", "ADI");
        columns.put("15", "ADV");
        columns.put("16", "API");
        columns.put("17", "APV");
        columns.put("18", "DCI");
        columns.put("19", "DCV");
        columns.put("20", "DDI");
        columns.put("21", "DDV");
        columns.put("22", "DPI");
        columns.put("23", "DPV");
        columns.put("24", "JCI");
        columns.put("25", "JCO");
        columns.put("26", "JCV");
        columns.put("27", "JDI");
        columns.put("28", "JDO");
        columns.put("29", "JDV");
        columns.put("30", "JPI");
        columns.put("31", "JPV");
        columns.put("32", "MCI");
        columns.put("33", "MCO");
        columns.put("34", "MCV");
        columns.put("35", "MDI");
        columns.put("36", "MDO");
        columns.put("37", "MDV");
        columns.put("38", "MPI");
        columns.put("39", "MPV");
        columns.put("40", "NCI");
        columns.put("41", "NDI");
        columns.put("42", "PCV");
        columns.put("43", "PDO");
        columns.put("44", "PDV");
        columns.put("45", "UCI");
        columns.put("46", "UCO");
        columns.put("47", "UCV");
        columns.put("48", "UDI");
        columns.put("49", "UDO");
        columns.put("50", "UDV");
        columns.put("51", "UPI");
        columns.put("52", "UPV");
        columns.put("53", "VCI");
        columns.put("54", "VCO");
        columns.put("55", "VCV");
        columns.put("56", "VDI");
        columns.put("57", "VDO");
        columns.put("58", "VDV");
        columns.put("59", "VPI");
        columns.put("60", "VPO");
        columns.put("61", "VPV");
        columns.put("62", "LDI");
        columns.put("63", "CDI");

        dynamicObject.setColumns(columns);

        Map<Integer, Object> inputParams = new HashMap<>();
        inputParams.put(1, branch.getBranchCode());
        List<Integer> outParams = new ArrayList<>();
        outParams.add(2);
        List<DynamicObject> dynamicObjects = new ArrayList<>();

        if (dynamicObjects.size() == 0) {
            DynamicObject dynamicObject1 = new DynamicObject();
            dynamicObject1.setColumns(columns);
            dynamicObject1.setProperties(new LinkedHashMap<>());
            dynamicObjects.add(dynamicObject1);
        }

        ExcelGenerator excelGenerator = new ExcelGenerator();

        Sheet sheet = excelGenerator.generateExcel(7, dynamicObjects, true);

        // title row
        Row titleRow = sheet.createRow(0);
        titleRow.createCell(0)
                .setCellValue(
                        "Báo cáo chi tiết Đơn vị chấp nhận thẻ và POS sau chuyển đổi theo chi nhánh");
        // font bold, size 16 for title
        Font font = sheet.getWorkbook().createFont();
        font.setBold(true);
        font.setFontHeightInPoints((short) 16);
        CellStyle style = sheet.getWorkbook().createCellStyle();
        style.setFont(font);
        style.setAlignment(HorizontalAlignment.CENTER);

        titleRow.getCell(0).setCellStyle(style);
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, columns.size()));

        // date now format dd/MM/yyyy
        DateFormat dateFormat = new SimpleDateFormat("dd/MM/yyyy");

        String dateStr = dateFormat.format(date);

        // header row 1
        Row row1 = sheet.createRow(1);
        row1.createCell(0).setCellValue("Mã báo cáo: ACQ_010");
        // header row 2
        Row row2 = sheet.createRow(2);
        row2.createCell(0).setCellValue("Ngày báo cáo: " + dateStr);

        Row row3 = sheet.createRow(3);
        row3.createCell(0).setCellValue("Chi nhánh: " + branch.getBranchName());

        Font fontBold = sheet.getWorkbook().createFont();
        fontBold.setBold(true);
        CellStyle styleBold = sheet.getWorkbook().createCellStyle();
        styleBold.setFont(fontBold);

        row1.getCell(0).setCellStyle(styleBold);
        row2.getCell(0).setCellStyle(styleBold);
        row3.getCell(0).setCellStyle(styleBold);

        Row headerRow = sheet.createRow(5);
        Row headerRow2 = sheet.createRow(6);
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
        headerRow2.createCell(0).setCellStyle(cellStyle);

        for (int i = 1; i < columns.size() + 1; i++) {
            headerRow2.createCell(i).setCellStyle(cellStyle);
            // resize column width
            sheet.autoSizeColumn(i);
            headerRow.createCell(i).setCellStyle(cellStyle);
            // check columns value length <=3 then set value headerRow2
            if (columns.values().toArray()[i - 1].toString().length() <= 4
                    && !columns.values().toArray()[i - 1].toString().equalsIgnoreCase("MCC")
                    && !columns.values().toArray()[i - 1].toString().equalsIgnoreCase("AM")) {
                headerRow2.getCell(i).setCellValue((String) columns.values().toArray()[i - 1]);
            } else {
                headerRow.getCell(i).setCellValue((String) columns.values().toArray()[i - 1]);
            }
        }

        headerRow.getCell(1).setCellValue("CN quản lý");
        headerRow.getCell(3).setCellValue("Số Cif Khách hàng");
        headerRow.getCell(8).setCellValue("Merchant ID (Main/ Ho)");
        headerRow.getCell(12).setCellValue("Merchant ID (Sub)");
        headerRow.getCell(14).setCellValue("Terminal ID");
        headerRow.getCell(16).setCellValue("MCC");
        headerRow.getCell(18).setCellValue("AM");
        headerRow.getCell(20).setCellValue("RBS number");
        headerRow.getCell(22).setCellValue("Phí MDR áp dụng VND - CAD");
        headerRow.getCell(30).setCellValue("Phí MDR Way4");

        sheet.addMergedRegion(new CellRangeAddress(5, 5, 1, 2));
        sheet.addMergedRegion(new CellRangeAddress(5, 5, 3, 4));
        sheet.addMergedRegion(new CellRangeAddress(5, 5, 8, 9));
        sheet.addMergedRegion(new CellRangeAddress(5, 5, 12, 13));
        sheet.addMergedRegion(new CellRangeAddress(5, 5, 14, 15));
        sheet.addMergedRegion(new CellRangeAddress(5, 5, 16, 17));
        sheet.addMergedRegion(new CellRangeAddress(5, 5, 18, 19));
        sheet.addMergedRegion(new CellRangeAddress(5, 5, 20, 21));
        sheet.addMergedRegion(new CellRangeAddress(5, 5, 22, 29));
        sheet.addMergedRegion(new CellRangeAddress(5, 5, 30, columns.size()));

        excelGenerator.writeExcel(fileName);
    }

    public static void ACQ011Report() throws FileNotFoundException, IOException {
        System.out.println("Generating ACQ011 Report");
        Date date = new Date();
        DateFormat dateFNFormat = new SimpleDateFormat("ddMMyyyy");
        String dateFN = dateFNFormat.format(date);
        String fileName = "/FTPData/HSC/ACQ_011_" + dateFN + ".xlsx";
        DynamicObject dynamicObject = new DynamicObject();
        Map<String, String> columns = new LinkedHashMap<>();
        columns.put("1", "IST");
        columns.put("2", "IST");
        columns.put("3", "IST");
        columns.put("4", "IST");
        columns.put("5", "IST");
        columns.put("6", "IST");
        columns.put("7", "WAY4");
        columns.put("8", "Check (True/ False)");
        columns.put("6", "IST");
        columns.put("7", "WAY4");
        columns.put("8", "Check (True/ False)");
        columns.put("9", "IST");
        columns.put("10", "WAY4");
        columns.put("11", "Check (True/ False)");
        columns.put("12", "IST");
        columns.put("13", "WAY4");
        columns.put("14", "Check (True/ False)");
        columns.put("15", "IST");
        columns.put("16", "WAY4");
        columns.put("17", "Check (True/ False)");
        columns.put("18", "IST");
        columns.put("19", "WAY4");
        columns.put("20", "Check (True/ False)");
        columns.put("21", "IST");
        columns.put("22", "WAY4");
        columns.put("23", "Check (True/ False)");
        columns.put("24", "IST");
        columns.put("25", "WAY4");
        columns.put("26", "Check (True/ False)");
        columns.put("27", "SRN (Source Registration Number)");
        columns.put("28", "Target channel");
        columns.put("29", "Source channel");

        dynamicObject.setColumns(columns);

        Map<Integer, Object> inputParams = new HashMap<>();
        List<Integer> outParams = new ArrayList<>();
        outParams.add(1);
        List<DynamicObject> dynamicObjects = new ArrayList<>();

        if (dynamicObjects.size() == 0) {
            DynamicObject dynamicObject1 = new DynamicObject();
            dynamicObject1.setColumns(columns);
            dynamicObject1.setProperties(new LinkedHashMap<>());
            dynamicObjects.add(dynamicObject1);
        }

        ExcelGenerator excelGenerator = new ExcelGenerator();

        Sheet sheet = excelGenerator.generateExcel(7, dynamicObjects, true);

        // title row
        Row titleRow = sheet.createRow(0);
        titleRow.createCell(0)
                .setCellValue(
                        "Báo cáo chi tiết Giao dịch preauth thẻ offus trên thiết bị BIDV");
        // font bold, size 16 for title
        Font font = sheet.getWorkbook().createFont();
        font.setBold(true);
        font.setFontHeightInPoints((short) 16);
        CellStyle style = sheet.getWorkbook().createCellStyle();
        style.setFont(font);
        style.setAlignment(HorizontalAlignment.CENTER);

        titleRow.getCell(0).setCellStyle(style);
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, columns.size()));

        // date now format dd/MM/yyyy
        DateFormat dateFormat = new SimpleDateFormat("dd/MM/yyyy");

        String dateStr = dateFormat.format(date);

        // header row 1
        Row row1 = sheet.createRow(1);
        row1.createCell(0).setCellValue("Mã báo cáo: ACQ_011");
        // header row 2
        Row row2 = sheet.createRow(2);
        row2.createCell(0).setCellValue("Ngày báo cáo: " + dateStr);

        Font fontBold = sheet.getWorkbook().createFont();
        fontBold.setBold(true);
        CellStyle styleBold = sheet.getWorkbook().createCellStyle();
        styleBold.setFont(fontBold);

        row1.getCell(0).setCellStyle(styleBold);
        row2.getCell(0).setCellStyle(styleBold);

        Row headerRow = sheet.createRow(5);
        Row headerRow2 = sheet.createRow(6);
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
        headerRow2.createCell(0).setCellValue("Nguồn dữ liệu");
        headerRow2.getCell(0).setCellStyle(cellStyle);

        for (int i = 1; i < columns.size() - 2; i++) {
            headerRow.createCell(i).setCellStyle(cellStyle);
            headerRow2.createCell(i).setCellValue((String) columns.values().toArray()[i - 1]);
            // resize column width
            sheet.autoSizeColumn(i);
            headerRow2.getCell(i).setCellStyle(cellStyle);
        }

        for (int i = columns.size() - 2; i < columns.size() + 1; i++) {
            headerRow.createCell(i).setCellValue((String) columns.values().toArray()[i - 1]);
            sheet.autoSizeColumn(i);
            headerRow.getCell(i).setCellStyle(cellStyle);
        }
        String[] header = { "MSGTYPE", "PCODE", "TXNDEST", "RESPCODE", "SHCERROR", "Transaction DATE",
                "Target Number (PAN)", "RRN (Retrieval Reference Number)", "Auth Code", "Merchant ID", "Terminal ID",
                "Transaction amount" };
        for (int i = 1; i < 6; i++) {
            sheet.autoSizeColumn(i);
            headerRow.getCell(i).setCellValue(header[i - 1]);
        }
        for (int i = 6; i < 26; i += 3) {
            headerRow.getCell(i).setCellValue(header[i / 3 + 3]);
            sheet.autoSizeColumn(i);
            sheet.addMergedRegion(new CellRangeAddress(5, 5, i, i + 2));
        }

        excelGenerator.writeExcel(fileName);
    }
}
