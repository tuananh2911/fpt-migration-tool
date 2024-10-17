package com.fis.services;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.math.BigDecimal;
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
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.fis.domain.Branch;

import oracle.net.aso.c;

public class ReportService {

    private static DatabaseService databaseService = new DatabaseService();

    private final static String MAX_NUM_ROWS = "600000";

    public static void initData() throws SQLException {
        databaseService.initData("REPORT_MIGRATE", "INIT_DATA");
    }

    public static void CMS018Report() throws FileNotFoundException, IOException, InterruptedException {
        String fileName = "CMS018Report.xlsx";

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

    public static void ISS_011(Branch branch)
            throws FileNotFoundException, IOException, InterruptedException, SQLException, InterruptedException {

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
        columns.put("dcnsk", "Địa chỉ nhận sao kê");
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
        // System.out.println(dynamicObjects.size());

        ExcelGenerator excelGenerator = new ExcelGenerator();
        // System.out.println("Start generate sheet");
        Sheet sheet = excelGenerator.generateExcel(6, dynamicObjects, false);
        // System.out.println("End generate sheet");

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
        // System.out.println("Start write excel");
        excelGenerator.writeExcel(fileName);

    }

    public static void ISS_012(Branch branch)
            throws FileNotFoundException, IOException, InterruptedException, SQLException, InterruptedException {

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
        // System.out.println(dynamicObjects.get(0).getProperties());

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

    public static void ATM_002()
            throws FileNotFoundException, IOException, InterruptedException, SQLException, InterruptedException {

        Date date = new Date();
        DateFormat dateFNFormat = new SimpleDateFormat("ddMMyyyy");
        String dateFN = dateFNFormat.format(date);
        String fileName = "/FTPData/HSC/ATM_002_" + dateFN + ".xlsx";

        DynamicObject dynamicObject = new DynamicObject();
        Map<String, String> columns = new LinkedHashMap<>();
        columns.put("Terminal_ID", "Terminal ID");
        columns.put("Branch_quan_ly", "Branch quản lý");
        columns.put("ma_chi_nhanh_tq6", "Branch tiếp quỹ");
        columns.put("MA_AM", "Mã AM quản lý máy");
        columns.put("ATM_TYPE", "ATM TYPE");
        columns.put("HANG_ATM", "Hãng ATM");
        columns.put("model", "Model");
        columns.put("ATM_LOCATION", "ATM location");
        columns.put("menh_gia_hop_tien_1", "Mệnh giá hộp tiền 1");
        columns.put("menh_gia_hop_tien_2", "Mệnh giá hộp tiền 2");
        columns.put("menh_gia_hop_tien_3", "Mệnh giá hộp tiền 3");
        columns.put("menh_gia_hop_tien_4", "Mệnh giá hộp tiền 4");
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
        sheet.addMergedRegion(new CellRangeAddress(1, 1, 0, columns.size()));

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

    public static void ACQ_009(Branch branch)
            throws FileNotFoundException, IOException, InterruptedException, SQLException, InterruptedException {
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
        columns.put("OWN_MERCH", "Merchant main ID");
        columns.put("MECHANT_LIENKET", "Merchant liên kết");
        columns.put("MERCHANT", "Merchant ID");
        columns.put("TERM_NBR", "Terminal ID");
        columns.put("MER_TYPE", "MCC");
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
        // System.out.println(dynamicObjects.get(0).getProperties());
        // map dynamicObjects add properties value JD = "" if JD is null
        // for (DynamicObject dynamicObject2 : dynamicObjects) {
        // if (dynamicObject2.getProperties().get("JD") == null) {
        // dynamicObject2.getProperties().put("JD", "");
        // }
        // dynamicObject2.getProperties().put("1", "");
        // }
        ExcelGenerator excelGenerator = new ExcelGenerator();
        Sheet sheet = excelGenerator.generateExcel(8, dynamicObjects, true);

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
                .setCellValue("Chi nhánh: " + branch.getBranchCode());
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
            // sheet.autoSizeColumn(i);
            sheet.setColumnWidth(i, 20 * 256);
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
            sheet.setColumnWidth(i, 20 * 256);
            // sheet.autoSizeColumn(i);
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

    public static void GL_007() throws FileNotFoundException, IOException, InterruptedException {

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

    public static void GL_005_ISS_CT() throws FileNotFoundException, IOException, InterruptedException, SQLException {

        Date date = new Date();
        DateFormat dateFNFormat = new SimpleDateFormat("ddMMyyyy");
        String dateFN = dateFNFormat.format(date);
        String fileName = "/FTPData/HSC/GL_005_ISS_" + dateFN + ".xlsx";
        DynamicObject dynamicObject = new DynamicObject();
        Map<String, String> columns = new LinkedHashMap<>();
        columns.put("ma_cn", "Mã CN CAD");
        columns.put("ma_cn_6_so", "MÃ CN 6 số");
        columns.put("loai_the", "Loại thẻ (phân loại theo TCT)");
        columns.put("so_tk_the", "Số tài khoản thẻ");
        columns.put("cif", "CIF");
        columns.put("ma_khach_hang", "Mã KH (ZCOMCDE)");
        columns.put("ten_khach_hang", "Tên khách hàng");
        columns.put("phan_loai_no", "Phân loại nợ");
        columns.put("du_no_goc_trong_han", "Dư nợ gốc trong hạn");
        columns.put("du_no_goc_qua_han", "Dư nợ quá hạn");
        columns.put("lai_du_no_nhom_1", "Lãi dư nợ nhóm 1");
        columns.put("lai_du_no_nhom_2_5", "Lãi dư nợ nhóm 2-5");
        columns.put("phi_du_no_nhom_1", "Phí dư nợ nhóm 1");
        columns.put("phi_du_no_nhom_2_5", "Phí dư nợ nhóm 2-5");
        columns.put("tong", "Cộng");

        dynamicObject.setColumns(columns);

        // not fill properties for now
        dynamicObject.setProperties(new LinkedHashMap<>());

        Map<Integer, Object> inputParams = new HashMap<>();
        List<Integer> outParams = new ArrayList<>();
        outParams.add(1);

        List<DynamicObject> dynamicObjects = databaseService.callProcedure("REPORT_MIGRATE", "GL_005_ISS_CT", columns,
                inputParams, outParams);
        // List<DynamicObject> dynamicObjects = new ArrayList<>();
        if (dynamicObjects.size() == 0) {
            DynamicObject dynamicObject1 = new DynamicObject();
            dynamicObject1.setColumns(columns);
            dynamicObject1.setProperties(new LinkedHashMap<>());
            dynamicObjects.add(dynamicObject1);
        }
        // dynamicObjects.add(dynamicObject);
        ExcelGenerator excelGenerator = new ExcelGenerator();
        Sheet sheet = excelGenerator.getSheet("Report");
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
        sheet.addMergedRegion(new CellRangeAddress(6, 7, 0, 0));

        for (int i = 1; i < 9; i++) {
            headerRow.createCell(i).setCellValue((String) columns.values().toArray()[i]);
            headerRow.getCell(i).setCellStyle(cellStyle);
            headerRow2.createCell(i).setCellStyle(cellStyle);
            // sheet.autoSizeColumn(i);
            sheet.addMergedRegion(new CellRangeAddress(6, 7, i, i));
        }
        for (int i = 9; i < columns.size(); i++) {
            headerRow2.createCell(i).setCellValue((String) columns.values().toArray()[i]);
            headerRow2.getCell(i).setCellStyle(cellStyle);
            headerRow.createCell(i).setCellStyle(cellStyle);
        }
        headerRow.getCell(9).setCellValue("Dữ liệu chuyển đổi");
        headerRow.createCell(columns.size()).setCellStyle(cellStyle);

        sheet.addMergedRegion(new CellRangeAddress(6, 6, 9, columns.size()));

        excelGenerator.generateExcel(8, dynamicObjects, false, Integer.parseInt(MAX_NUM_ROWS));
        excelGenerator.writeExcel(fileName);
    }

    public static void ISS_009(Branch branch)
            throws FileNotFoundException, IOException, InterruptedException, SQLException, InterruptedException {

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

    public static void ISS_010(Branch branch)
            throws FileNotFoundException, IOException, InterruptedException, SQLException, InterruptedException {

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

    public static void ATM_001() throws FileNotFoundException, IOException, InterruptedException, SQLException {

        Date date = new Date();
        DateFormat dateFNFormat = new SimpleDateFormat("ddMMyyyy");
        String dateFN = dateFNFormat.format(date);
        String fileName = "/FTPData/HSC/ATM_001_" + dateFN + ".xlsx";
        DynamicObject dynamicObject = new DynamicObject();
        Map<String, String> columns = new LinkedHashMap<>();
        columns.put("CNQUANLY", "Chi nhánh");
        columns.put("SL_TCD_ATM", "Số lượng ATM");
        columns.put("SL_TCD_CRM", "Số lượng CRM");
        columns.put("CNQUANLY_SCD", "Branch quản lý");
        columns.put("SL_SCD_ATM", "Số lượng CRM");
        columns.put("SL_SCD_CRM", "Số lượng CRM");

        dynamicObject.setColumns(columns);

        // not fill properties for now
        dynamicObject.setProperties(new LinkedHashMap<>());

        Map<Integer, Object> inputParams = new HashMap<>();
        List<Integer> outParams = new ArrayList<>();
        outParams.add(1);

        List<DynamicObject> dynamicObjects = databaseService.callProcedure("REPORT_MIGRATE", "ATM_001", columns,
                inputParams, outParams);

        // dynamicObjects.add(dynamicObject);
        if (dynamicObjects.size() == 0) {
            DynamicObject dynamicObject1 = new DynamicObject();
            dynamicObject1.setColumns(columns);
            dynamicObject1.setProperties(new LinkedHashMap<>());
            dynamicObjects.add(dynamicObject1);
        }

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

    public static void ISS_001_0() throws FileNotFoundException, IOException, InterruptedException, SQLException {

        Date date = new Date();
        DateFormat dateFNFormat = new SimpleDateFormat("ddMMyyyy");
        String dateFN = dateFNFormat.format(date);
        String fileName = "/FTPData/HSC/ISS_001_0_" + dateFN + ".xlsx";
        DynamicObject dynamicObject = new DynamicObject();
        Map<String, String> columns = new LinkedHashMap<>();
        columns.put("BDS", "BDS");
        columns.put("TONG_CIF_CAD_KHCN", "Tổng số CIF KHCN Cadencie");
        columns.put("TONG_CIF_CAD_KHCN_KHONG_CHUYEN_DOI", "CIF không chuyển đổi");
        columns.put("TONG_CIF_CAD_KHCN_CHUYEN_DOI", "Chuyển đổi (Migration)");
        columns.put("TONG_CIF_WAY4_KHCN", "Cập nhật lên hệ thống Way4");
        columns.put("CHENH_LECH_KHCN", "Chênh lệch");
        columns.put("TONG_CIF_CAD_KHDN", "Tổng số CIF KHDN Cadencie");
        columns.put("TONG_CIF_CAD_KHDN_KHONG_CHUYEN_DOI", "CIF không chuyển đổi");
        columns.put("TONG_CIF_CAD_KHDN_CHUYEN_DOI", "Chuyển đổi (Migration)");
        columns.put("TONG_CIF_WAY4_KHDN", "Cập nhật lên hệ thống Way4");
        columns.put("CHENH_LECH_KHDN", "Chênh lệch");

        dynamicObject.setColumns(columns);

        // not fill properties for now
        dynamicObject.setProperties(new LinkedHashMap<>());

        Map<Integer, Object> inputParams = new HashMap<>();
        List<Integer> outParams = new ArrayList<>();
        outParams.add(1);

        List<DynamicObject> dynamicObjects = databaseService.callProcedure("REPORT_MIGRATE", "ISS_001_0", columns,
                (Map<Integer, Object>) inputParams, outParams);

        // dynamicObjects.add(dynamicObject);

        if (dynamicObjects.size() == 0) {
            DynamicObject dynamicObject1 = new DynamicObject();
            dynamicObject1.setColumns(columns);
            dynamicObject1.setProperties(new LinkedHashMap<>());
            dynamicObjects.add(dynamicObject1);
        }

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
            // sheet.autoSizeColumn(i);
            sheet.setColumnWidth(i, 20 * 256);
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

    public static void ISS_001_1()
            throws FileNotFoundException, IOException, InterruptedException, SQLException, InterruptedException {

        Date date = new Date();
        DateFormat dateFNFormat = new SimpleDateFormat("ddMMyyyy");
        String dateFN = dateFNFormat.format(date);
        String fileName = "/FTPData/HSC/ISS_001_1_" + dateFN + ".xlsx";
        DynamicObject dynamicObject = new DynamicObject();
        Map<String, String> columns = new LinkedHashMap<>();
        columns.put("branch", "BDS");
        columns.put("product", "Product code");
        columns.put("account_all", "Số lượng trên Cadencie");
        columns.put("account_khong_chuyen", "Không chuyển đổi");
        columns.put("account_chuyen", "Chuyển đổi (Migration)");
        columns.put("way4_account", "Cập nhật lên hệ thống Way4");
        columns.put("chenh_lech_account", "Chênh lệch");
        columns.put("card_all", "Số lượng thẻ trên Cadencie");
        columns.put("card_khong_chuyen", "Số lượng thẻ không CĐ");
        columns.put("card_chuyen", "Số lượng thẻ CĐ");
        columns.put("way4_card", "Cập nhật lên W4");
        columns.put("chenh_lech_card", "Chênh lệch");

        dynamicObject.setColumns(columns);

        // not fill properties for now
        dynamicObject.setProperties(new LinkedHashMap<>());

        Map<Integer, Object> inputParams = new HashMap<>();

        List<Integer> outParams = new ArrayList<>();
        outParams.add(1);

        List<DynamicObject> dynamicObjects = databaseService.callProcedure("REPORT_MIGRATE", "ISS_001_1", columns,
                inputParams, outParams);

        if (dynamicObjects.size() == 0) {
            dynamicObjects.add(dynamicObject);
        }

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

        for (int i = 3; i < columns.size() + 1; i++) {
            headerRow.createCell(i).setCellStyle(cellStyle);
            headerRow2.createCell(i).setCellValue((String) columns.values().toArray()[i - 1]);
            // sheet.autoSizeColumn(i);
            sheet.setColumnWidth(i, 20 * 256);
            headerRow2.getCell(i).setCellStyle(cellStyle);
        }
        headerRow.getCell(3).setCellValue("Chuyển đổi tài khoản thẻ");
        headerRow.getCell(8).setCellValue("Chuyển đổi thẻ");

        // row sum after data
        Row sumRow = sheet.createRow(7 + dynamicObjects.size());
        BigDecimal sum3 = new BigDecimal(0);
        BigDecimal sum6 = new BigDecimal(0);
        BigDecimal sum7 = new BigDecimal(0);
        BigDecimal sum8 = new BigDecimal(0);
        BigDecimal sum9 = new BigDecimal(0);

        // sum if value is number
        for (DynamicObject dynamicObject2 : dynamicObjects) {
            if (dynamicObject2.getProperties().get("account_all") != null
                    && dynamicObject2.getProperties().get("account_all") instanceof BigDecimal) {
                sum3 = sum3.add((BigDecimal) dynamicObject2.getProperties().get("account_all"));
            }
            if (dynamicObject2.getProperties().get("way4_account") != null
                    && dynamicObject2.getProperties().get("way4_account") instanceof BigDecimal) {
                sum6 = sum6.add((BigDecimal) dynamicObject2.getProperties().get("way4_account"));
            }
            if (dynamicObject2.getProperties().get("chenh_lech_account") != null
                    && dynamicObject2.getProperties().get("chenh_lech_account") instanceof BigDecimal) {
                sum7 = sum7.add((BigDecimal) dynamicObject2.getProperties().get("chenh_lech_account"));
            }
            if (dynamicObject2.getProperties().get("card_all") != null
                    && dynamicObject2.getProperties().get("card_all") instanceof BigDecimal) {
                sum8 = sum8.add((BigDecimal) dynamicObject2.getProperties().get("card_all"));
            }
            if (dynamicObject2.getProperties().get("card_khong_chuyen") != null
                    && dynamicObject2.getProperties().get("card_khong_chuyen") instanceof BigDecimal) {
                sum9 = sum9.add((BigDecimal) dynamicObject2.getProperties().get("card_khong_chuyen"));

            }
        }

        sumRow.createCell(0).setCellValue("SUM");

        sumRow.createCell(3).setCellValue(sum3.doubleValue());
        sumRow.createCell(6).setCellValue(sum6.doubleValue());
        sumRow.createCell(7).setCellValue(sum7.doubleValue());
        sumRow.createCell(8).setCellValue(sum8.doubleValue());
        sumRow.createCell(9).setCellValue(sum9.doubleValue());

        // merge cell for header
        sheet.addMergedRegion(new CellRangeAddress(6, 6, 3, 7));
        sheet.addMergedRegion(new CellRangeAddress(6, 6, 8, columns.size()));
        sheet.addMergedRegion(new CellRangeAddress(6, 7, 0, 0));
        sheet.addMergedRegion(new CellRangeAddress(6, 7, 1, 1));
        sheet.addMergedRegion(new CellRangeAddress(6, 7, 2, 2));

        Font fontBold = sheet.getWorkbook().createFont();
        fontBold.setBold(true);
        CellStyle styleBold = sheet.getWorkbook().createCellStyle();
        styleBold.setFont(fontBold);

        int endRow = 9 + dynamicObjects.size() + 2;
        Row eRow = sheet.createRow(
                endRow);
        eRow.createCell(0).setCellValue("LẬP BẢNG");
        eRow.createCell(4).setCellValue("NGƯỜI KIỂM SOÁT");

        eRow.getCell(0).setCellStyle(styleBold);
        eRow.getCell(4).setCellStyle(styleBold);

        excelGenerator.writeExcel(fileName);
    }

    public static void ISS_001_2() throws FileNotFoundException, IOException, InterruptedException, SQLException {
        Date date = new Date();
        DateFormat dateFNFormat = new SimpleDateFormat("ddMMyyyy");
        String dateFN = dateFNFormat.format(date);
        String fileName = "/FTPData/HSC/ISS_001_2_" + dateFN + ".xlsx";
        DynamicObject dynamicObject = new DynamicObject();
        Map<String, String> columns = new LinkedHashMap<>();
        columns.put("branch", "BDS");
        columns.put("product", "Product code");
        columns.put("tong_du_no_cad", "Tổng dư nợ trên Cadencie");
        columns.put("w4_tong_du_no", "Tổng dư nợ trên Way4");
        columns.put("chenh_lech_tong_du_no", "Chênh lệch");
        columns.put("du_no_goc_cad", "Cadencie(1)");
        columns.put("w4_open_cash", "W4 Open cash(2)");
        columns.put("w4_open_sale", "W4 Open Sale(3)");
        columns.put("w4_grace_cash", "W4 Grace cash (4)");
        columns.put("w4_open_sale", "W4 Grace sale (5)");
        columns.put("w4_close_cash", "W4 Close Cash (6)");
        columns.put("w4_close_sale", "W4 Close Sale(7)");
        columns.put("w4_open_prin_ins", "W4 Open principle Instalment(7)");
        columns.put("chenh_lech_du_no_goc", "Chênh lệch (8)=(1)-(2)-(3)-(4)- (5)- (6)- (7)- (7')");
        columns.put("du_lai", "Cadencie (9) Lãi");
        columns.put("w4_grace_interest", "W4 Grace Interest (10)");
        columns.put("w4_close_interest", "W4 Close Interest (11)");
        columns.put("chenh_lech_du_lai", "Chênh lệch (12)=(9)-(10)-(11)");
        columns.put("du_phi", "Cadencie (13) Phí");
        columns.put("w4_open_fee", "W4 Open Fee (14)");
        columns.put("w4_grace_fee", "W4 Grace Fee(15)");
        columns.put("w4_open_fee_instalment", "W4 Close Fee (16)");
        columns.put("chenh_lech_du_phi", "W4 Open Fee Instalment (17)");
        columns.put("chenh_lech_du_phi", "Chênh lệch (18)=(13)-(14)-(15)-(16)- (17)");
        columns.put("du_no_tra_gop_chua_phan_bo", "Cadencie (19)");
        columns.put("w4_waiting_principle", "W4 Waiting Principle (20)");
        columns.put("so_sanh_du_no_tra_gop_chua_phan_bo", "So sánh (Pass/ Fail)");
        columns.put("tong_so_gd_tra_gop_chuyen_doi", "Cadencie");
        columns.put("w4_tong_so_luong_gd_tra_gop", "Way4");
        columns.put("so_sanh_tong_gd_tra_gop", "So sánh (Pass/ Fail)");

        dynamicObject.setColumns(columns);

        dynamicObject.setProperties(new LinkedHashMap<>());
        Map<Integer, Object> inputParams = new HashMap<>();

        List<Integer> outParams = new ArrayList<>();
        outParams.add(1);

        List<DynamicObject> dynamicObjects = databaseService.callProcedure("REPORT_MIGRATE", "ISS_001_2", columns,
                inputParams, outParams);

        // dynamicObjects.add(dynamicObject);

        ExcelGenerator excelGenerator = new ExcelGenerator();

        Sheet sheet = excelGenerator.generateExcel(7, dynamicObjects, true);

        // title row 0
        Row titleRow = sheet.createRow(0);
        titleRow.createCell(0)
                .setCellValue(
                        "BÁO CÁO TỔNG HỢP SỐ LIỆU DƯ NỢ THẺ QUỐC TẾ CHUYỂN ĐỔI TOÀN HỆ THỐNG THEO CHI NHÁNH VÀ MÃ SẢN PHẨM THẺ");

        Row row1 = sheet.createRow(1);
        row1.createCell(0).setCellValue("Mã báo cáo: ISS_001.2");

        // date
        DateFormat dateFormat = new SimpleDateFormat("dd/MM/yyyy");
        String dateStr = dateFormat.format(date);
        Row row4 = sheet.createRow(4);
        row4.createCell(0).setCellValue("Ngày báo cáo: " + dateStr);
        // font bold, size 16 for title
        Font font = sheet.getWorkbook().createFont();
        font.setBold(true);
        font.setFontHeightInPoints((short) 16);
        CellStyle style = sheet.getWorkbook().createCellStyle();
        style.setFont(font);
        style.setAlignment(HorizontalAlignment.CENTER);
        titleRow.getCell(0).setCellStyle(style);
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, columns.size()));

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

        headerRow.createCell(0).setCellValue("STT");
        headerRow2.createCell(0).setCellStyle(cellStyle);
        headerRow.getCell(0).setCellStyle(cellStyle);

        String[] header = new String[] { "BDS", "Product code", "Tổng dư nợ trên Cadencie", "Tổng dư nợ trên Way4",
                "Chênh lệch", "Dư nợ gốc", "Dư lãi", "Dư phí", "Dư nợ trả góp chưa phân bổ",
                "Tổng số lượng GD trả góp chuyển đổi" };

        for (int i = 1; i < 6; i++) {
            headerRow.createCell(i).setCellStyle(cellStyle);
            headerRow.getCell(i).setCellValue(header[i - 1]);
            // sheet.autoSizeColumn(i);
            sheet.setColumnWidth(i, 20 * 256);
            headerRow2.createCell(i).setCellStyle(cellStyle);
        }

        for (int i = 6; i < columns.size() + 1; i++) {
            headerRow.createCell(i).setCellStyle(cellStyle);
            headerRow2.createCell(i).setCellValue((String) columns.values().toArray()[i - 1]);
            // sheet.autoSizeColumn(i);
            sheet.setColumnWidth(i, 20 * 256);
            headerRow2.getCell(i).setCellStyle(cellStyle);
        }

        headerRow.getCell(6).setCellValue("Dư nợ gốc");
        headerRow.getCell(15).setCellValue("Dư lãi");
        headerRow.getCell(19).setCellValue("Dư phí");
        headerRow.getCell(25).setCellValue("Dư nợ trả góp chưa phân bổ");
        headerRow.getCell(28).setCellValue("Tổng số lượng GD trả góp chuyển đổi");

        // merge cell for header
        sheet.addMergedRegion(new CellRangeAddress(6, 6, 6, 14));
        sheet.addMergedRegion(new CellRangeAddress(6, 6, 15, 18));
        sheet.addMergedRegion(new CellRangeAddress(6, 6, 19, 24));
        sheet.addMergedRegion(new CellRangeAddress(6, 6, 25, columns.size()));

        sheet.addMergedRegion(new CellRangeAddress(6, 7, 0, 0));
        sheet.addMergedRegion(new CellRangeAddress(6, 7, 1, 1));
        sheet.addMergedRegion(new CellRangeAddress(6, 7, 2, 2));
        sheet.addMergedRegion(new CellRangeAddress(6, 7, 3, 3));
        sheet.addMergedRegion(new CellRangeAddress(6, 7, 4, 4));
        sheet.addMergedRegion(new CellRangeAddress(6, 7, 5, 5));

        // row sum after data
        Row sumRow = sheet.createRow(7 + dynamicObjects.size());
        sumRow.createCell(0).setCellValue("SUM");

        // end
        int endRow = 9 + dynamicObjects.size() + 2;
        Row eRow = sheet.createRow(endRow);
        eRow.createCell(0).setCellValue("LẬP BẢNG");
        eRow.createCell(4).setCellValue("NGƯỜI KIỂM SOÁT");
        eRow.createCell(9).setCellValue("ĐẠI DIỆN CHI NHÁNH");
        eRow.getCell(0).setCellStyle(style);
        eRow.getCell(4).setCellStyle(style);
        eRow.getCell(9).setCellStyle(style);

        excelGenerator.writeExcel(fileName);

    }

    public static void ISS_002() throws FileNotFoundException, IOException, InterruptedException {

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
            // sheet.autoSizeColumn(i);
            sheet.setColumnWidth(i, 20 * 256);
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

    public static void ISS_005() throws SQLException, FileNotFoundException, IOException, InterruptedException {

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

        // List<DynamicObject> dynamicObjects =
        // databaseService.callProcedure("REPORT_MIGRATE", "ISS_005", columns,
        // inputParams, outParams);

        List<DynamicObject> dynamicObjects = new ArrayList<>();

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
            // sheet.autoSizeColumn(i);
            sheet.setColumnWidth(i, 20 * 256);
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

    public static void ISS_003() throws SQLException, FileNotFoundException, IOException, InterruptedException {

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
            // sheet.autoSizeColumn(i);
            sheet.setColumnWidth(i, 20 * 256);
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

    public static void ISS_013(Branch branch)
            throws SQLException, FileNotFoundException, IOException, InterruptedException {

        Date date = new Date();
        DateFormat dateFNFormat = new SimpleDateFormat("ddMMyyyy");
        String dateFN = dateFNFormat.format(date);
        String fileName = branch.getFolderPath() + "ISS_013_" + branch.getBranchCode() + "_" + branch.getBranchName()
                + "_" + dateFN
                + ".xlsx";
        DynamicObject dynamicObject = new DynamicObject();
        Map<String, String> columns = new LinkedHashMap<>();
        columns.put("LOAI_THE", "Loại thẻ");
        columns.put("branch", "CN quản lý");
        columns.put("CLIENT_NUMBER", "Số CIF");
        columns.put("STK_THE_CAP_ISSUE", "Số tài khoản thẻ");
        columns.put("TEN_TK_CAP_ISSUE", "Tên tài khoản");
        columns.put("TRANG_THAI_TK_CAP_ISSUE", "Trạng thái tài khoản");
        columns.put("card_nbr", "Số thẻ");
        columns.put("TRANG_THAI_THE_CAP_CARD", "Trạng thái thẻ");
        columns.put("product", "Mã sản phẩm thẻ");
        columns.put("AM", "AM");
        columns.put("HMTD", "HMTD");
        columns.put("ClassCode", "ClassCode");
        columns.put("InterestCode", "Interest Code");
        columns.put("fee_code", "Feecode");
        columns.put("ID_TSBD", "ID TSBĐ (nếu có)");
        columns.put("SO_NGAY_QUA_HAN", "Số ngày quá hạn");
        columns.put("NGAY_SAO_KE", "Ngày sao kê");
        columns.put("DCNSK", "Địa chỉ nhận sao kê");
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
        columns.put("TY_LE_TRICH_NO_TU_DONG", "Tỷ lệ trích nợ tự động");
        columns.put("TK_TRICH_NO_TU_DONG", "Tài khoản trích nợ tự động");
        columns.put("DU_NO_THE", "Dư nợ thẻ");
        // columns.put("query_amt", "Số tiền khiếu nại đang treo");
        columns.put("TONG_SO_DU_GIAO_DICH_TRA_GOP", "Tổng số dư giao dịch trả góp");
        columns.put("TONG_SO_GD_TRA_GOP_DANG_HOAT_DONG", "Tổng số giao dịch trả góp đang hoạt động");
        columns.put("TONG_GIA_TRI_CAC_GD_CHUA_LEN_DU_NO", "Tổng giá trị các giao dịch chưa lên dư nợ");
        columns.put("DOANH_SO_CHI_TIEU_LUY_KE", "Tổng Doanh số thanh toán");
        dynamicObject.setColumns(columns);

        Map<Integer, Object> inputParams = new HashMap<>();
        inputParams.put(1, branch.getBranchCode());
        List<Integer> outParams = new ArrayList<>();
        outParams.add(2);
        List<DynamicObject> dynamicObjects = databaseService.callProcedure("REPORT_MIGRATE", "ISS_013", columns,
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

    public static void ISS_006() throws FileNotFoundException, IOException, InterruptedException {

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
            // sheet.autoSizeColumn(i);
            sheet.setColumnWidth(i, 20 * 256);
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

    public static void ISS_007() throws FileNotFoundException, IOException, InterruptedException {

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
            // sheet.autoSizeColumn(i);
            sheet.setColumnWidth(i, 20 * 256);
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

    public static void ATM_003()
            throws FileNotFoundException, IOException, InterruptedException, SQLException, InterruptedException {

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

    public static void ISS_008() throws FileNotFoundException, IOException, InterruptedException {

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
        columns.put("8", "W4 Open cash(2)");
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
            // sheet.autoSizeColumn(i);
            sheet.setColumnWidth(i, 20 * 256);
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

    public static void ISS_008_1() throws FileNotFoundException, IOException, InterruptedException {

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
            // sheet.autoSizeColumn(i);
            sheet.setColumnWidth(i, 20 * 256);
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

    public static void ACQ_007() throws FileNotFoundException, IOException, InterruptedException, SQLException {

        Date date = new Date();
        DateFormat dateFNFormat = new SimpleDateFormat("ddMMyyyy");
        String dateFN = dateFNFormat.format(date);
        String fileName = "/FTPData/HSC/ACQ_007_" + dateFN + ".xlsx";
        DynamicObject dynamicObject = new DynamicObject();
        Map<String, String> columns = new LinkedHashMap<>();
        columns.put("STT", "STT");
        columns.put("CN_QUANLY", "CN quản lý");
        columns.put("CIF_KH", "Số Cif Khách hàng");
        columns.put("CIF_NAME", "Tên Khách hàng");
        columns.put("P_MID", "Parent Merchant ID");
        columns.put("MID_CAD", "Merchant ID");
        columns.put("MNAME_CAD", "Tên Merchant");
        columns.put("TID_CAD", "Terminal ID");
        columns.put("MCC", "MCC");
        columns.put("AM", "AM");
        columns.put("RBS_NUM", "RBS number");
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
        List<Integer> outParams = new ArrayList<>();
        outParams.add(1);
        List<DynamicObject> dynamicObjects = databaseService.callProcedure("REPORT_MIGRATE", "ACQ_007", columns,
                inputParams, outParams);

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
            // sheet.autoSizeColumn(i);
            sheet.setColumnWidth(i, 20 * 256);
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
            // sheet.autoSizeColumn(i);
            sheet.setColumnWidth(i, 20 * 256);
            headerRow2.getCell(i).setCellStyle(cellStyle);
        }

        excelGenerator.writeExcel(fileName);
    }

    public static void ACQ_008() throws FileNotFoundException, IOException, InterruptedException, SQLException {

        Date date = new Date();
        DateFormat dateFNFormat = new SimpleDateFormat("ddMMyyyy");
        String dateFN = dateFNFormat.format(date);
        String fileName = "/FTPData/HSC/AQC_008_" + dateFN + ".xlsx";
        DynamicObject dynamicObject = new DynamicObject();
        Map<String, String> columns = new LinkedHashMap<>();
        columns.put("cn_quan_ly_way4", "CN quản lý");
        columns.put("cif_way4", "Số Cif Khách hàng");
        columns.put("ten_kh_way4", "Tên Khách hàng");
        columns.put("merchant_main_way4", "Merchant main ID");
        columns.put("merchant_lien_ket_way4", "Merchant liên kết (Add_info_03)");
        columns.put("merchant_id_way4", "Merchant ID");
        columns.put("merchant_name_way4", "Tên Merchant");
        columns.put("terminal_id_way4", "Terminal ID");
        columns.put("mcc_way4", "MCC");
        columns.put("ma_am_way4", "AM");
        columns.put("RBS_NUMBER", "RBS number");

        dynamicObject.setColumns(columns);

        Map<Integer, Object> inputParams = new HashMap<>();
        List<Integer> outParams = new ArrayList<>();
        outParams.add(1);
        List<DynamicObject> dynamicObjects = databaseService.callProcedure("REPORT_MIGRATE", "ACQ_008", columns,
                inputParams, outParams);

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

    public static void ACQ_006() throws FileNotFoundException, IOException, InterruptedException {

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
            // sheet.autoSizeColumn(i);
            sheet.setColumnWidth(i, 20 * 256);
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

    public static void ACQ_010(Branch branch) throws FileNotFoundException, IOException, InterruptedException {

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
            // sheet.autoSizeColumn(i);
            sheet.setColumnWidth(i, 20 * 256);
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

    public static void ACQ_011() throws FileNotFoundException, IOException, InterruptedException, SQLException {

        Date date = new Date();
        DateFormat dateFNFormat = new SimpleDateFormat("ddMMyyyy");
        String dateFN = dateFNFormat.format(date);
        String fileName = "/FTPData/HSC/ACQ_011_" + dateFN + ".xlsx";
        DynamicObject dynamicObject = new DynamicObject();
        Map<String, String> columns = new LinkedHashMap<>();
        columns.put("STT", "IST");
        columns.put("MSGTYPE", "IST");
        columns.put("PCODE", "IST");
        columns.put("TXNDEST", "IST");
        columns.put("RESPCODE", "IST");
        columns.put("SHCERROR", "IST");
        columns.put("ist_tran_date", "IST");
        columns.put("w4_tran_date", "WAY4");
        columns.put("so_sanh_tran_date", "Check (True/ False)");
        columns.put("ist_pan", "IST");
        columns.put("w4_pan", "WAY4");
        columns.put("so_sanh_pan", "Check (True/ False)");
        columns.put("ist_rrn", "IST");
        columns.put("w4_rrn", "WAY4");
        columns.put("so_sanh_rrn", "Check (True/ False)");
        columns.put("ist_auth_code", "IST");
        columns.put("w4_auth_code", "WAY4");
        columns.put("so_sanh_auth_code", "Check (True/ False)");
        columns.put("ist_merchant", "IST");
        columns.put("w4_merchant", "WAY4");
        columns.put("so_sanh_merchant", "Check (True/ False)");
        columns.put("ist_TERMID", "IST");
        columns.put("w4_TERMID", "WAY4");
        columns.put("so_sanh_TERMID", "Check (True/ False)");
        columns.put("ist_tran_amount", "IST");
        columns.put("w4_tran_amount", "WAY4");
        columns.put("so_sanh_tran_amount", "Check (True/ False)");
        columns.put("27", "SRN (Source Registration Number)");
        columns.put("28", "Target channel");
        columns.put("29", "Source channel");

        dynamicObject.setColumns(columns);

        Map<Integer, Object> inputParams = new HashMap<>();
        List<Integer> outParams = new ArrayList<>();
        outParams.add(1);
        List<DynamicObject> dynamicObjects = databaseService.callProcedure("REPORT_MIGRATE", "ACQ_011", columns,
                inputParams, outParams);

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
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, columns.size() - 1));

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

        for (int i = 1; i < columns.size() - 3; i++) {
            headerRow.createCell(i).setCellStyle(cellStyle);
            headerRow2.createCell(i).setCellValue((String) columns.values().toArray()[i]);
            // resize column width
            // sheet.autoSizeColumn(i);
            sheet.setColumnWidth(i, 20 * 256);
            headerRow2.getCell(i).setCellStyle(cellStyle);
        }

        for (int i = columns.size() - 3; i < columns.size(); i++) {
            headerRow.createCell(i).setCellValue((String) columns.values().toArray()[i]);
            // sheet.autoSizeColumn(i);
            sheet.setColumnWidth(i, 20 * 256);
            headerRow.getCell(i).setCellStyle(cellStyle);
            headerRow2.createCell(i).setCellStyle(cellStyle);
            sheet.addMergedRegion(new CellRangeAddress(5, 6, i, i));
        }
        String[] header = { "MSGTYPE", "PCODE", "TXNDEST", "RESPCODE", "SHCERROR", "Transaction DATE",
                "Target Number (PAN)", "RRN (Retrieval Reference Number)", "Auth Code", "Merchant ID", "Terminal ID",
                "Transaction amount" };
        for (int i = 1; i < 6; i++) {
            // sheet.autoSizeColumn(i);
            sheet.setColumnWidth(i, 20 * 256);
            headerRow.getCell(i).setCellValue(header[i - 1]);
        }
        for (int i = 6; i < 26; i += 3) {
            headerRow.getCell(i).setCellValue(header[i / 3 + 3]);
            // sheet.autoSizeColumn(i);
            sheet.setColumnWidth(i, 20 * 256);
            sheet.addMergedRegion(new CellRangeAddress(5, 5, i, i + 2));
        }

        excelGenerator.writeExcel(fileName);
    }

    public static void ACQ_001() throws FileNotFoundException, IOException, InterruptedException, SQLException {

        Date date = new Date();
        DateFormat dateFNFormat = new SimpleDateFormat("ddMMyyyy");
        String dateFN = dateFNFormat.format(date);
        String fileName = "/FTPData/HSC/ACQ_001_" + dateFN + ".xlsx";
        DynamicObject dynamicObject = new DynamicObject();
        Map<String, String> columns = new LinkedHashMap<>();
        columns.put("STT", "STT");
        columns.put("BDS", "BDS");
        columns.put("tong_cif_truoc_chuyen_doi", "Tổng trước chuyển đổi");
        columns.put("tong_cif_khong_chuyen_doi", "Tổng không chuyển đổi");
        columns.put("tong_cif_duoc_chuyen_doi", "Tổng được chuyển đổi");
        columns.put("tong_cif_chuyen_doi_thanh_cong", "Tổng chuyển đổi thành công");
        columns.put("tong_cif_chuyen_doi_khong_thanh_cong", "Tổng chuyển đổi không thành công");
        columns.put("tong_merchant_truoc_chuyen_doi", "Tổng trước chuyển đổi");
        columns.put("tong_merchant_khong_chuyen_doi", "Tổng không chuyển đổi");
        columns.put("tong_merchant_duoc_chuyen_doi", "Tổng được chuyển đổi");
        columns.put("tong_merchant_chuyen_doi_thanh_cong", "Tổng chuyển đổi thành công");
        columns.put("tong_merchant_chuyen_doi_khong_thanh_cong", "Tổng chuyển đổi không thành công");
        columns.put("tong_device_truoc_chuyen_doi", "Tổng trước chuyển đổi");
        columns.put("tong_device_khong_chuyen_doi", "Tổng không chuyển đổi");
        columns.put("tong_device_duoc_chuyen_doi", "Tổng được chuyển đổi");
        columns.put("tong_device_chuyen_doi_thanh_cong", "Tổng chuyển đổi thành công");
        columns.put("tong_device_chuyen_doi_khong_thanh_cong", "Tổng chuyển đổi không thành công");

        dynamicObject.setColumns(columns);

        Map<Integer, Object> inputParams = new HashMap<>();
        List<Integer> outParams = new ArrayList<>();
        outParams.add(1);
        List<DynamicObject> dynamicObjects = databaseService.callProcedure("REPORT_MIGRATE", "ACQ_001", columns,
                inputParams, outParams);

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
                        "Báo cáo tổng hợp số lượng Khách hàng, MID và TID");
        // font bold, size 16 for title
        Font font = sheet.getWorkbook().createFont();
        font.setBold(true);
        font.setFontHeightInPoints((short) 16);
        CellStyle style = sheet.getWorkbook().createCellStyle();
        style.setFont(font);
        style.setAlignment(HorizontalAlignment.CENTER);

        titleRow.getCell(0).setCellStyle(style);

        // date now format dd/MM/yyyy
        DateFormat dateFormat = new SimpleDateFormat("dd/MM/yyyy");

        String dateStr = dateFormat.format(date);

        // header row 1
        Row row1 = sheet.createRow(1);
        row1.createCell(columns.size() - 2).setCellValue("Mã báo cáo: ACQ_001");
        // header row 2
        Row row2 = sheet.createRow(2);
        row2.createCell(columns.size() - 2).setCellValue("Ngày báo cáo: " + dateStr);

        Font fontBold = sheet.getWorkbook().createFont();
        fontBold.setBold(true);
        CellStyle styleBold = sheet.getWorkbook().createCellStyle();
        styleBold.setFont(fontBold);

        row1.getCell(columns.size() - 2).setCellStyle(styleBold);
        row2.getCell(columns.size() - 2).setCellStyle(styleBold);

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

        headerRow.createCell(1).setCellValue("BDS");
        headerRow.getCell(1).setCellStyle(cellStyle);
        headerRow2.createCell(1).setCellStyle(cellStyle);

        for (int i = 2; i < columns.size(); i++) {
            headerRow2.createCell(i).setCellValue((String) columns.values().toArray()[i]);
            // resize column width
            // sheet.autoSizeColumn(i);
            sheet.setColumnWidth(i, 20 * 256);
            headerRow2.getCell(i).setCellStyle(cellStyle);
            headerRow.createCell(i).setCellStyle(cellStyle);
        }

        headerRow.getCell(2).setCellValue("Số lượng CIF");
        headerRow.getCell(7).setCellValue("Số lượng ĐVCNT (Merchant)");
        headerRow.getCell(12).setCellValue("Số lượng Thiết bị (Device)");

        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, columns.size() - 1));

        sheet.addMergedRegion(new CellRangeAddress(3, 4, 0, 0));
        sheet.addMergedRegion(new CellRangeAddress(3, 4, 1, 1));

        sheet.addMergedRegion(new CellRangeAddress(3, 3, 2, 6));
        sheet.addMergedRegion(new CellRangeAddress(3, 3, 7, 11));
        sheet.addMergedRegion(new CellRangeAddress(3, 3, 12, columns.size() - 1));

        excelGenerator.writeExcel(fileName);
    }

    public static void ISS_004_1() throws FileNotFoundException, IOException, InterruptedException, SQLException {
        Date date = new Date();
        DateFormat dateFNFormat = new SimpleDateFormat("ddMMyyyy");
        String dateFN = dateFNFormat.format(date);
        String fileName = "/FTPData/HSC/ISS_004_1_" + dateFN + ".xlsx";
        DynamicObject dynamicObject = new DynamicObject();
        Map<String, String> columns = new LinkedHashMap<>();
        columns.put("STT", "STT");
        columns.put("ACCOUNT_CONTRACT", "Số tài khoản thẻ");
        columns.put("cif_cad", "Cadencie");
        columns.put("cif_way4", "Way4");
        columns.put("SS_cif", "So sánh (Pass/ Fail)");
        columns.put("cn_cad", "Cadencie");
        columns.put("cn_way4", "Way4");
        columns.put("SS_cn", "So sánh (Pass/ Fail)");
        columns.put("ma_sp_cad", "Cadencie");
        columns.put("ma_sp_way4", "Way4");
        columns.put("SS_ma_sp", "So sánh (Pass/ Fail)");
        columns.put("ngay_mo_cad", "Cadencie");
        columns.put("ngay_mo_way4", "Way4");
        columns.put("SS_ngay_mo", "So sánh (Pass/ Fail)");
        columns.put("ma_gioi_thieu_cad", "Cadencie");
        columns.put("ma_gioi_thieu_way4", "Way4");
        columns.put("SS_ma_gioi_thieu", "So sánh (Pass/ Fail)");
        columns.put("ma_am_cad", "Cadencie");
        columns.put("ma_am_way4", "Way4");
        columns.put("SS_ma_am", "So sánh (Pass/ Fail)");
        columns.put("trang_thai_tk_cad", "Cadencie");
        columns.put("trang_thai_tk_way4", "Way4");
        columns.put("SS_trang_thai_tk", "So sánh (Pass/ Fail)");
        columns.put("feecode_cad", "Cadencie");
        columns.put("feecode_way4", "Way4");
        columns.put("SS_feecode", "So sánh (Pass/ Fail)");
        columns.put("feemonth_cad", "Cadencie");
        columns.put("feemonth_way4", "Way4");
        columns.put("SS_feemonth", "So sánh (Pass/ Fail)");
        columns.put("1", "Cadencie");
        columns.put("2", "Way4");
        columns.put("3", "So sánh (Pass/ Fail)");
        columns.put("fee_delay_cad", "Cadencie");
        columns.put("fee_delay_way4", "Way4");
        columns.put("SS_feedelay", "So sánh (Pass/ Fail)");
        columns.put("delay_active_cad", "Cadencie");
        columns.put("delay_active_way4", "Way4");
        columns.put("SS_delay_active", "So sánh (Pass/ Fail)");
        // columns.put("classcode_cad", "Cadencie");
        // columns.put("classcode_way4", "Way4");
        // columns.put("SS_classcode", "So sánh (Pass/ Fail)");
        columns.put("tk_lien_ket_1_cad", "Cadencie");
        columns.put("tk_lien_ket_1_way", "Way4");
        columns.put("SS_tk_lien_ket_1", "So sánh (Pass/ Fail)");
        columns.put("tk_lien_ket_2_cad", "Cadencie");
        columns.put("tk_lien_ket_2_way", "Way4");
        columns.put("SS_tk_lien_ket_2", "So sánh (Pass/ Fail)");
        columns.put("tk_lien_ket_3_cad", "Cadencie");
        columns.put("tk_lien_ket_3_way", "Way4");
        columns.put("SS_tk_lien_ket_3", "So sánh (Pass/ Fail)");
        columns.put("tk_lien_ket_4_cad", "Cadencie");
        columns.put("tk_lien_ket_4_way", "Way4");
        columns.put("SS_tk_lien_ket_4", "So sánh (Pass/ Fail)");
        columns.put("tk_lien_ket_5_cad", "Cadencie");
        columns.put("tk_lien_ket_5_way", "Way4");
        columns.put("SS_tk_lien_ket_5", "So sánh (Pass/ Fail)");
        columns.put("tk_lien_ket_6_cad", "Cadencie");
        columns.put("tk_lien_ket_6_way", "Way4");
        columns.put("SS_tk_lien_ket_6", "So sánh (Pass/ Fail)");
        columns.put("tk_lien_ket_7_cad", "Cadencie");
        columns.put("tk_lien_ket_7_way", "Way4");
        columns.put("SS_tk_lien_ket_7", "So sánh (Pass/ Fail)");
        columns.put("tk_lien_ket_8_cad", "Cadencie");
        columns.put("tk_lien_ket_8_way", "Way4");
        columns.put("SS_tk_lien_ket_8", "So sánh (Pass/ Fail)");
        columns.put("tk_lien_ket_9_cad", "Cadencie");
        columns.put("tk_lien_ket_9_way", "Way4");
        columns.put("SS_tk_lien_ket_9", "So sánh (Pass/ Fail)");
        columns.put("tk_lien_ket_10_cad", "Cadencie");
        columns.put("tk_lien_ket_10_way", "Way4");
        columns.put("SS_tk_lien_ket_10", "So sánh (Pass/ Fail)");

        dynamicObject.setColumns(columns);

        Map<Integer, Object> inputParams = new HashMap<>();
        List<Integer> outParams = new ArrayList<>();
        outParams.add(1);
        List<DynamicObject> dynamicObjects = databaseService.callProcedure("REPORT_MIGRATE", "ISS_004_1", columns,
                inputParams, outParams);

        if (dynamicObjects.size() == 0) {
            DynamicObject dynamicObject1 = new DynamicObject();
            dynamicObject1.setColumns(columns);
            dynamicObject1.setProperties(new LinkedHashMap<>());
            dynamicObjects.add(dynamicObject1);
        }

        ExcelGenerator excelGenerator = new ExcelGenerator();

        Sheet sheet = excelGenerator.getSheet("Report");

        // title row
        Row titleRow = sheet.createRow(0);
        titleRow.createCell(0)
                .setCellValue(
                        "BÁO CÁO ĐỐI CHIẾU DỮ LIỆU THÔNG TIN TÀI KHOẢN THẺ GHI NỢ TẠI CADENCIE - WAY4");
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
        row1.createCell(0).setCellValue("Mã báo cáo: ISS_004_1");
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
        headerRow2.createCell(0).setCellStyle(cellStyle);
        headerRow.createCell(1).setCellValue("Số tài khoản thẻ");
        headerRow.getCell(1).setCellStyle(cellStyle);
        headerRow2.createCell(1).setCellStyle(cellStyle);

        String[] header = { "Số CIF", "CN quản lý", "Mã sản phẩm", "Ngày mở", "Mã CB giới thiệu", "Mã AM",
                "Trạng thái TK- Reason", "Feemonth", "Mức phí TN", "Fee delay", "Delay until active date", "ClassCode",
                "Tài khoản liên kết 1", "TK liên kết 2", "TK liên kết 3", "TK liên kết 4", "TK liên kết 5",
                "TK liên kết 6", "TK liên kết 7", "TK liên kết 8", "TK liên kết 9", "TK liên kết 10" };

        for (int i = 2; i < columns.size(); i++) {
            headerRow2.createCell(i).setCellValue((String) columns.values().toArray()[i]);
            // resize column width
            // sheet.autoSizeColumn(i);
            sheet.setColumnWidth(i, 20 * 256);
            headerRow2.getCell(i).setCellStyle(cellStyle);
            headerRow.createCell(i).setCellStyle(cellStyle);
        }

        for (int i = 2; i < columns.size(); i += 3) {
            headerRow.getCell(i).setCellValue(header[i / 3]);
            // sheet.autoSizeColumn(i);
            sheet.setColumnWidth(i, 20 * 256);
            sheet.addMergedRegion(new CellRangeAddress(5, 5, i, i + 2));
        }

        sheet.addMergedRegion(new CellRangeAddress(5, 6, 0, 0));
        sheet.addMergedRegion(new CellRangeAddress(5, 6, 1, 1));
        excelGenerator.generateExcel(7, dynamicObjects, true, Integer.parseInt(MAX_NUM_ROWS));

        excelGenerator.writeExcel(fileName);
    }

    public static void ISS_004() throws FileNotFoundException, IOException, InterruptedException {
        Date date = new Date();
        DateFormat dateFNFormat = new SimpleDateFormat("ddMMyyyy");
        String dateFN = dateFNFormat.format(date);
        String fileName = "/FTPData/HSC/ISS_004_" + dateFN + ".xlsx";
        DynamicObject dynamicObject = new DynamicObject();
        Map<String, String> columns = new LinkedHashMap<>();
        columns.put("STT", "STT");
        columns.put("1", "Số tài khoản thẻ");
        columns.put("2", "Cadencie");
        columns.put("3", "Way4");
        columns.put("4", "So sánh (Pass/ Fail)");
        columns.put("5", "Cadencie");
        columns.put("6", "Way4");
        columns.put("7", "So sánh (Pass/ Fail)");
        columns.put("CUSTR_REF", "Cadencie");
        columns.put("9", "Way4");
        columns.put("10", "So sánh (Pass/ Fail)");
        columns.put("BUSI_NAME", "Cadencie");
        columns.put("SHORT_NAME", "Way4");
        columns.put("SS_TKH", "So sánh (Pass/ Fail)");
        columns.put("ZBOC", "Cadencie");
        columns.put("BRANCH", "Way4");
        columns.put("SS_CNPHT", "So sánh (Pass/ Fail)");
        columns.put("17", "Cadencie");
        columns.put("18", "Way4");
        columns.put("19", "So sánh (Pass/ Fail)");
        columns.put("20", "Cadencie");
        columns.put("21", "Way4");
        columns.put("22", "So sánh (Pass/ Fail)");
        columns.put("23", "Cadencie");
        columns.put("24", "Way4");
        columns.put("25", "So sánh (Pass/ Fail)");
        columns.put("26", "Cadencie");
        columns.put("27", "Way4");
        columns.put("28", "So sánh (Pass/ Fail)");
        columns.put("29", "Cadencie");
        columns.put("30", "Way4");
        columns.put("31", "So sánh (Pass/ Fail)");
        columns.put("32", "Cadencie");
        columns.put("33", "Way4");
        columns.put("34", "So sánh (Pass/ Fail)");
        columns.put("35", "Cadencie");
        columns.put("36", "Way4");
        columns.put("37", "So sánh (Pass/ Fail)");
        columns.put("38", "Cadencie");
        columns.put("39", "Way4");
        columns.put("40", "So sánh (Pass/ Fail)");
        columns.put("41", "Cadencie");
        columns.put("42", "Way4");
        columns.put("43", "So sánh (Pass/ Fail)");
        columns.put("44", "Cadencie");
        columns.put("45", "Way4");
        columns.put("46", "So sánh (Pass/ Fail)");
        columns.put("47", "Cadencie");
        columns.put("48", "Way4");
        columns.put("49", "So sánh (Pass/ Fail)");
        columns.put("50", "Cadencie");
        columns.put("51", "Way4");
        columns.put("52", "So sánh (Pass/ Fail)");
        columns.put("53", "Cadencie");
        columns.put("54", "Way4");
        columns.put("55", "So sánh (Pass/ Fail)");
        columns.put("56", "Cadencie");
        columns.put("57", "Way4");
        columns.put("58", "So sánh (Pass/ Fail)");
        columns.put("59", "Cadencie");
        columns.put("60", "Way4");
        columns.put("61", "So sánh (Pass/ Fail)");
        columns.put("62", "Cadencie");
        columns.put("63", "Way4");
        columns.put("64", "So sánh (Pass/ Fail)");
        columns.put("65", "Cadencie");
        columns.put("66", "Way4");
        columns.put("67", "So sánh (Pass/ Fail)");
        columns.put("68", "Cadencie");
        columns.put("69", "Way4");
        columns.put("70", "So sánh (Pass/ Fail)");
        columns.put("71", "Cadencie");
        columns.put("72", "Way4");
        columns.put("73", "So sánh (Pass/ Fail)");
        columns.put("74", "Cadencie");
        columns.put("75", "Way4");
        columns.put("76", "So sánh (Pass/ Fail)");

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
                        "BÁO CÁO ĐỐI CHIẾU DỮ LIỆU THÔNG TIN TÀI KHOẢN THẺ TÍN DỤNG TẠI CADENCIE - WAY4");
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
        row1.createCell(9).setCellValue("Mã báo cáo: ISS_004");
        // header row 2
        Row row2 = sheet.createRow(2);
        row2.createCell(9).setCellValue("Ngày báo cáo: " + dateStr);

        Font fontBold = sheet.getWorkbook().createFont();
        fontBold.setBold(true);
        CellStyle styleBold = sheet.getWorkbook().createCellStyle();
        styleBold.setFont(fontBold);

        row1.getCell(9).setCellStyle(styleBold);
        row2.getCell(9).setCellStyle(styleBold);

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

        String[] header = { "Số tài khoản thẻ", "Tên tài khoản thẻ", "Số CIF", "Tên KH", "CN quản lý", "Mã sản phẩm",
                "Ngày mở", "Mã CB giới thiệu", "Mã AM", "Trạng thái TK - Reason", "Feemonth", "Mức phí TN", "Fee delay",
                "Delay until active date", "Class Code", "Interest code", "Mã chính sách cấp tín dụng", "HMTD",
                "Ngày sao kê", "TK trích nợ tự động", "Tỷ lệ trích nợ tự động", "ID TSĐB", "HMTD tạm thời",
                "Ngày hết hạn HMTD tạm thời", "Số lượng GD trả góp còn hoạt động (trạng thái Active,  N- New)" };

        for (int i = 1; i < columns.size(); i++) {
            headerRow2.createCell(i).setCellValue((String) columns.values().toArray()[i]);
            // resize column width
            // sheet.autoSizeColumn(i);
            sheet.setColumnWidth(i, 20 * 256);
            headerRow2.getCell(i).setCellStyle(cellStyle);
            headerRow.createCell(i).setCellStyle(cellStyle);
        }

        for (int i = 1; i < columns.size(); i += 3) {
            headerRow.getCell(i).setCellValue(header[i / 3]);
            // sheet.autoSizeColumn(i);
            sheet.setColumnWidth(i, 20 * 256);
            sheet.addMergedRegion(new CellRangeAddress(5, 5, i, i + 2));
        }

        int endRow = 7 + dynamicObjects.size() + 2;
        Row eRow = sheet.createRow(endRow);
        eRow.createCell(3).setCellValue("LẬP BẢNG");
        eRow.createCell(9).setCellValue("NGƯỜI KIỂM SOÁT");
        eRow.createCell(14).setCellValue("ĐẠI DIỆN CHI NHÁNH");
        eRow.getCell(3).setCellStyle(styleBold);
        eRow.getCell(9).setCellStyle(styleBold);
        eRow.getCell(14).setCellStyle(styleBold);

        excelGenerator.writeExcel(fileName);

    }

    public static void ACQ_004() throws SQLException, FileNotFoundException, IOException, InterruptedException {
        Date date = new Date();
        DateFormat dateFNFormat = new SimpleDateFormat("ddMMyyyy");
        String dateFN = dateFNFormat.format(date);
        String fileName = "/FTPData/HSC/ACQ_004_" + dateFN + ".xlsx";
        DynamicObject dynamicObject = new DynamicObject();
        Map<String, String> columns = new LinkedHashMap<>();
        columns.put("STT", "STT");
        columns.put("cn_quan_ly_cad", "CAD");
        columns.put("cn_quan_ly_way4", "Way4");
        columns.put("SS_cn_quan_ly", "Check(True/ False)");
        columns.put("merchant_id_cad", "CAD");
        columns.put("merchant_id_way4", "Way4");
        columns.put("SS_merchant_id", "Check(True/ False)");
        columns.put("contract_name_way4", "Way4");
        columns.put("contract_number_way4", "Way4");
        columns.put("product_way4", "Way4");
        columns.put("merchant_main_cad", "CAD");
        columns.put("merchant_main_way4", "Way4");
        columns.put("SS_merchant_main", "Check(True/ False)");
        columns.put("merchant_lien_ket_cad", "CAD");
        columns.put("merchant_lien_ket_way4", "Way4");
        columns.put("SS_merchant_lien_ket", "Check(True/ False)");
        columns.put("cif_cad", "CAD");
        columns.put("cif_way4", "Way4");
        columns.put("SS_cif", "Check(True/ False)");
        columns.put("mcc_nguon_chuyen_doi", "CAD");
        columns.put("mcc_sau_chuyen_doi", "Way4");
        columns.put("SS_mcc", "Check(True/ False)");
        columns.put("status_nguon_chuyen_doi", "CAD");
        columns.put("status_sau_chuyen_doi", "Way4");
        columns.put("SS_status", "Check(True/ False)");
        columns.put("rbs_1_nguon_chuyen_doi", "CAD");
        columns.put("rbs_1_sau_chuyen_doi", "Way4");
        columns.put("SS_rbs", "Check(True/ False)");
        columns.put("rbs_2_nguon_chuyen_doi", "CAD");
        columns.put("rbs_2_sau_chuyen_doi", "Way4");
        columns.put("SS_rbs_2", "Check(True/ False)");
        columns.put("classifier_MMSTAR", "Way4");
        columns.put("classifier_MMDP", "Way4");
        columns.put("classifier_MDR_fee", "Way4");
        columns.put("classifier_add_info_01", "Way4");

        dynamicObject.setColumns(columns);

        Map<Integer, Object> inputParams = new HashMap<>();
        List<Integer> outParams = new ArrayList<>();
        outParams.add(1);
        List<DynamicObject> dynamicObjects = databaseService.callProcedure("REPORT_MIGRATE", "ACQ_004", columns,
                inputParams, outParams);

        // System.out.println(dynamicObjects.get(0).getProperties());
        ExcelGenerator excelGenerator = new ExcelGenerator();

        Sheet sheet = excelGenerator.generateExcel(7, dynamicObjects, true);

        // title row
        Row titleRow = sheet.createRow(0);
        titleRow.createCell(0)
                .setCellValue(
                        "Báo cáo đối chiếu thông tin merchant sau chuyển đổi theo cif khách hàng");
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
        row1.createCell(0).setCellValue("Mã báo cáo: ACQ_004");
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

        headerRow.createCell(0).setCellValue("TT");
        headerRow.getCell(0).setCellStyle(cellStyle);
        headerRow2.createCell(0).setCellStyle(cellStyle);
        headerRow2.getCell(0).setCellValue("Nguồn dữ liệu");

        for (int i = 1; i < 7; i++) {
            headerRow.createCell(i).setCellStyle(cellStyle);
            // sheet.autoSizeColumn(i);
            sheet.setColumnWidth(i, 20 * 256);
            headerRow2.createCell(i).setCellValue((String) columns.values().toArray()[i]);
            headerRow2.getCell(i).setCellStyle(cellStyle);
        }

        for (int i = 7; i < 10; i++) {
            headerRow.createCell(i).setCellStyle(cellStyle);
            // sheet.autoSizeColumn(i);
            sheet.setColumnWidth(i, 20 * 256);
            headerRow2.createCell(i).setCellValue((String) columns.values().toArray()[i]);
            headerRow2.getCell(i).setCellStyle(cellStyle);
        }
        for (int i = 10; i < columns.size() - 3; i++) {
            headerRow.createCell(i).setCellStyle(cellStyle);
            // sheet.autoSizeColumn(i);
            sheet.setColumnWidth(i, 20 * 256);
            headerRow2.createCell(i).setCellValue((String) columns.values().toArray()[i]);
            headerRow2.getCell(i).setCellStyle(cellStyle);
        }
        for (int i = columns.size() - 4; i < columns.size(); i++) {
            headerRow.createCell(i).setCellStyle(cellStyle);
            // sheet.autoSizeColumn(i);
            sheet.setColumnWidth(i, 20 * 256);
            headerRow2.createCell(i).setCellValue((String) columns.values().toArray()[i]);
            headerRow2.getCell(i).setCellStyle(cellStyle);
        }
        headerRow.getCell(1).setCellValue("CN quản lý");
        headerRow.getCell(4).setCellValue("Merchant ID");
        headerRow.getCell(7).setCellValue("Contract name (tên đơn vị CNT)");
        headerRow.getCell(8).setCellValue("Contract number (cấp mid)");
        headerRow.getCell(9).setCellValue("Product (name)");
        headerRow.getCell(10).setCellValue("Merchant main ID");
        headerRow.getCell(13).setCellValue("Merchant chuỗi liên kết (Add_info_03)");
        headerRow.getCell(16).setCellValue("Số Cif Khách hàng");
        headerRow.getCell(19).setCellValue("MCC");
        headerRow.getCell(22).setCellValue("Status");
        headerRow.getCell(25).setCellValue("RBS number 1");
        headerRow.getCell(28).setCellValue("RBS number 2");
        headerRow.getCell(31).setCellValue("MMSTAR");
        headerRow.getCell(32).setCellValue("MMDP");
        headerRow.getCell(33).setCellValue("MDR fee");
        headerRow.getCell(34).setCellValue("Add_info_01");

        Row _headerRow = sheet.createRow(4);
        _headerRow.createCell(31).setCellValue("Classifier");
        _headerRow.getCell(31).setCellStyle(cellStyle);
        _headerRow.createCell(32).setCellStyle(cellStyle);
        _headerRow.createCell(33).setCellStyle(cellStyle);

        sheet.addMergedRegion(new CellRangeAddress(5, 5, 1, 3));
        sheet.addMergedRegion(new CellRangeAddress(5, 5, 4, 6));
        sheet.addMergedRegion(new CellRangeAddress(5, 5, 10, 12));
        sheet.addMergedRegion(new CellRangeAddress(5, 5, 13, 15));
        sheet.addMergedRegion(new CellRangeAddress(5, 5, 16, 18));
        sheet.addMergedRegion(new CellRangeAddress(5, 5, 19, 21));
        sheet.addMergedRegion(new CellRangeAddress(5, 5, 22, 24));
        sheet.addMergedRegion(new CellRangeAddress(5, 5, 25, 27));
        sheet.addMergedRegion(new CellRangeAddress(5, 5, 28, 30));
        sheet.addMergedRegion(new CellRangeAddress(4, 4, 31, 33));
        excelGenerator.writeExcel(fileName);
    }

    public static void ACQ_002() throws FileNotFoundException, IOException, InterruptedException, SQLException {
        Date date = new Date();
        DateFormat dateFNFormat = new SimpleDateFormat("ddMMyyyy");
        String dateFN = dateFNFormat.format(date);
        String fileName = "/FTPData/HSC/ACQ_002_" + dateFN + ".xlsx";
        DynamicObject dynamicObject = new DynamicObject();
        Map<String, String> columns = new LinkedHashMap<>();
        columns.put("STT", "STT");
        columns.put("mid_cad", "CAD");
        columns.put("ccat", "Way4");
        columns.put("clt", "Way4");
        columns.put("cif_cad", "CAD");
        columns.put("cif_prf", "PRF");
        columns.put("cif_way4", "Way4");
        columns.put("SS_cif", "Check (True/ False)");
        columns.put("id_kh_prf", "PRF");
        columns.put("id_kh_way4", "Way4");
        columns.put("SS_id_kh", "Check (True/ False)");
        columns.put("ten_kh_prf", "PRF");
        columns.put("ten_kh_way4", "Way4");
        columns.put("SS_ten_kh", "Check (True/ False)");
        columns.put("dia_chi_kh_prf", "PRF");
        columns.put("dia_chi_kh_way", "Way4");
        columns.put("SS_dia_chi_kh", "Check (True/ False)");
        columns.put("classifier_prf", "PRF");
        columns.put("classifier_way4", "Way4");
        columns.put("SS_classifier", "Check (True/ False)");
        columns.put("sdt_mobile_prf", "PRF");
        columns.put("sdt_mobile_way4", "Way4");
        columns.put("SS_sdt_mobile", "Check (True/ False)");
        columns.put("email_1_prf", "PRF");
        columns.put("email_1_way4", "Way4");
        columns.put("SS_email_1", "Check (True/ False)");

        dynamicObject.setColumns(columns);

        Map<Integer, Object> inputParams = new HashMap<>();
        List<Integer> outParams = new ArrayList<>();
        outParams.add(1);
        List<DynamicObject> dynamicObjects = databaseService.callProcedure("REPORT_MIGRATE", "ACQ_002", columns,
                inputParams, outParams);

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
                        "Báo cáo đối chiếu chi tiết thông tin khách hàng sau chuyển đổi - KHCN");
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
        row1.createCell(0).setCellValue("Mã báo cáo: ACQ_002");

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
        headerRow2.createCell(0).setCellStyle(cellStyle);
        headerRow2.getCell(0).setCellValue("Nguồn dữ liệu");

        for (int i = 1; i < columns.size(); i++) {
            headerRow2.createCell(i).setCellValue((String) columns.values().toArray()[i]);
            // resize column width
            // sheet.autoSizeColumn(i);
            sheet.setColumnWidth(i, 20 * 256);
            headerRow2.getCell(i).setCellStyle(cellStyle);
            headerRow.createCell(i).setCellStyle(cellStyle);
        }

        headerRow.getCell(1).setCellValue("Mid");
        headerRow.getCell(2).setCellValue("Client Catergory");
        headerRow.getCell(3).setCellValue("Client type");
        headerRow.getCell(4).setCellValue("Số CIF KH tại BIDV");

        String header[] = { "Số ID khách hàng (số CCCD)", "Tên KH",
                "Địa chỉ liên hệ tại cấp client đồng bộ từ PRF sang", "Classifier cho phân loại KH zcomcode",
                "Số điện thoại (tất cả sđt gắn client)", "Email (tất cả email gắn client)" };
        for (int i = 8; i < columns.size(); i += 3) {
            headerRow.getCell(i).setCellValue(header[i / 3 - 2]);
            sheet.setColumnWidth(i, 20 * 256);
            // sheet.autoSizeColumn(i);
            sheet.addMergedRegion(new CellRangeAddress(5, 5, i, i + 2));
        }

        excelGenerator.writeExcel(fileName);
    }

    public static void ACQ_005() throws SQLException, FileNotFoundException, IOException, InterruptedException {
        Date date = new Date();
        DateFormat dateFNFormat = new SimpleDateFormat("ddMMyyyy");
        String dateFN = dateFNFormat.format(date);
        String fileName = "/FTPData/HSC/ACQ_005_" + dateFN + ".xlsx";
        DynamicObject dynamicObject = new DynamicObject();
        Map<String, String> columns = new LinkedHashMap<>();
        columns.put("STT", "STT");
        columns.put("cn_quan_ly_cad", "CAD");
        columns.put("cn_quan_ly_way4", "Way4");
        columns.put("SS_cn_quan_ly", "Check (True/ False)");
        columns.put("terminal_id_cad", "CAD");
        columns.put("terminal_id_way4", "Way4");
        columns.put("SS_terminal_id", "Check (True/ False)");
        columns.put("product_way4", "Way4");
        columns.put("device_name_way4", "Way4");
        columns.put("merchant_id_cad", "CAD");
        columns.put("merchant_id_way4", "Way4");
        columns.put("SS_merchant_id", "Check (True/ False)");
        columns.put("cif_cad", "CAD");
        columns.put("cif_way4", "Way4");
        columns.put("SS_cif", "Check (True/ False)");
        columns.put("mcc_cad", "CAD");
        columns.put("mcc_way4", "Way4");
        columns.put("SS_mcc", "Check (True/ False)");
        columns.put("status_cad", "CAD");
        columns.put("classifier_way4", "Way4");
        columns.put("clt", "Check (True/ False)");
        columns.put("ma_am_cad", "CAD");
        columns.put("ma_am_way4", "Way4");
        columns.put("SS_ma_am", "Check (True/ False)");
        columns.put("mid_cad", "CAD");
        columns.put("ccat", "Way4");
        columns.put("clt2", "Check (True/ False)");
        columns.put("postype_cad", "CAD");
        columns.put("postype_way4", "Way4");
        columns.put("SS_postype", "Check (True/ False)");
        columns.put("pos_location_cad", "CAD");
        columns.put("pos_location_way4", "Way4");
        columns.put("SS_pos_location", "Check (True/ False)");
        columns.put("address_type_cad", "CAD");
        columns.put("address_type_way4", "Way4");
        columns.put("SS_address_type", "Check (True/ False)");
        columns.put("serial_cad", "CAD");
        columns.put("serial_way4", "Way4");
        columns.put("SS_serial", "Check (True/ False)");
        columns.put("default_curr_cad", "CAD");
        columns.put("default_curr_way4", "Way4");
        columns.put("SS_default_curr", "Check (True/ False)");
        columns.put("ma_am_comment_cad", "CAD");
        columns.put("ma_am_comment_way4", "Way4");
        columns.put("SS_ma_am_comment", "Check (True/ False)");

        dynamicObject.setColumns(columns);

        Map<Integer, Object> inputParams = new HashMap<>();
        List<Integer> outParams = new ArrayList<>();
        outParams.add(1);

        List<DynamicObject> dynamicObjects = databaseService.callProcedure("REPORT_MIGRATE", "ACQ_005", columns,
                inputParams, outParams);

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
                        "Báo cáo đối chiếu thông tin pos sau chuyển đổi theo cif khách hàng");
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
        row1.createCell(0).setCellValue("Mã báo cáo: ACQ_005");
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
        headerRow2.createCell(0).setCellStyle(cellStyle);
        headerRow2.getCell(0).setCellValue("Nguồn dữ liệu");

        for (int i = 1; i < columns.size(); i++) {
            headerRow2.createCell(i).setCellValue((String) columns.values().toArray()[i]);
            // resize column width
            // sheet.autoSizeColumn(i);
            sheet.setColumnWidth(i, 20 * 256);
            headerRow2.getCell(i).setCellStyle(cellStyle);
            headerRow.createCell(i).setCellStyle(cellStyle);
        }
        headerRow.getCell(1).setCellValue("CN quản lý");
        headerRow.getCell(4).setCellValue("Terminal ID");
        headerRow.getCell(7).setCellValue("Product name");
        headerRow.getCell(8).setCellValue("Device name");

        String[] header = { "Merchant ID", "Số cif KH", "MCC", "Classifier (Force close pos cycle)", "Mã AM", "Status",
                "Postype", "Pos Location", "Address_type (gán cấp device)", "Serial number", "Default curr",
                "Comment Text (mã AM)",
                "Comment Text (mã AM)" };
        for (int i = 9; i < columns.size(); i += 3) {
            headerRow.getCell(i).setCellValue(header[i / 3 - 3]);
            sheet.setColumnWidth(i, 20 * 256);
            // sheet.autoSizeColumn(i);
            sheet.addMergedRegion(new CellRangeAddress(5, 5, i, i + 2));
        }

        sheet.addMergedRegion(new CellRangeAddress(5, 5, 1, 3));
        sheet.addMergedRegion(new CellRangeAddress(5, 5, 4, 6));

        excelGenerator.writeExcel(fileName);
    }

    public static void ACQ_003() throws SQLException, FileNotFoundException, IOException, InterruptedException {
        Date date = new Date();
        DateFormat dateFNFormat = new SimpleDateFormat("ddMMyyyy");
        String dateFN = dateFNFormat.format(date);
        String fileName = "/FTPData/HSC/ACQ_003_" + dateFN + ".xlsx";
        DynamicObject dynamicObject = new DynamicObject();
        Map<String, String> columns = new LinkedHashMap<>();
        columns.put("STT", "STT");
        columns.put("merchant_id", "CAD");
        columns.put("cif_cad", "CAD");
        columns.put("cif_prf", "PRF");
        columns.put("cif_way4", "Way4");
        columns.put("SS_cif", "Check (True/ False)");
        columns.put("ten_kh_prf", "PRF");
        columns.put("ten_kh_way4", "Way4");
        columns.put("SS_ten_kh", "Check (True/ False)");
        columns.put("so_dkkd_prf", "PRF");
        columns.put("so_dkkd_way4", "Way4");
        columns.put("SS_so_dkkd", "Check (True/ False)");
        columns.put("ma_so_thue_prf", "PRF");
        columns.put("ma_so_thue_way4", "Way4");
        columns.put("SS_ma_so_thue", "Check (True/ False)");
        columns.put("dkkd_noi_cap_prf", "PRF");
        columns.put("dkkd_noi_cap_way4", "Way4");
        columns.put("SS_dkkd_noi_cap", "Check (True/ False)");
        columns.put("dia_chi_kh_dkkd_prf", "PRF");
        columns.put("dia_chi_kh_dkkd_way4", "Way4");
        columns.put("SS_dia_chi_kh_dkkd", "Check (True/ False)");
        columns.put("sdt_mobile_prf", "PRF");
        columns.put("sdt_mobile_way4", "Way4");
        columns.put("SS_sdt_mobile", "Check (True/ False)");
        columns.put("sdt_work_prf", "PRF");
        columns.put("sdt_work_way4", "Way4");
        columns.put("SS_sdt_work", "Check (True/ False)");
        columns.put("sdt_home_prf", "PRF");
        columns.put("sdt_home_way4", "Way4");
        columns.put("SS_sdt_home", "Check (True/ False)");
        columns.put("email_1_prf", "PRF");
        columns.put("email_1_way4", "Way4");
        columns.put("SS_email_1", "Check (True/ False)");
        columns.put("email_2_prf", "PRF");
        columns.put("email_2_way4", "Way4");
        columns.put("SS_email_2", "Check (True/ False)");
        columns.put("email_3_prf", "PRF");
        columns.put("email_3_way4", "Way4");
        columns.put("SS_email_3", "Check (True/ False)");
        columns.put("kh_zcomcode_prf", "PRF");
        columns.put("kh_zcomcode_way4", "Way4");
        columns.put("SS_kh_zcomcode", "Check (True/ False)");
        columns.put("nguoi_daidien_1_prf", "PRF");
        columns.put("nguoi_daidien_1_way4", "Way4");
        columns.put("SS_nguoi_daidien_1", "Check (True/ False)");
        columns.put("chuc_vu_1_prf", "PRF");
        columns.put("chuc_vu_1_way4", "Way4");
        columns.put("SS_chuc_vu_1", "Check (True/ False)");
        columns.put("nguoi_daidien_2_prf", "PRF");
        columns.put("nguoi_daidien_2_way4", "Way4");
        columns.put("SS_nguoi_daidien_2", "Check (True/ False)");
        columns.put("chuc_vu_2_prf", "PRF");
        columns.put("chuc_vu_2_way4", "Way4");
        columns.put("SS_chuc_vu_2", "Check (True/ False)");
        columns.put("nguoi_daidien_3_prf", "PRF");
        columns.put("nguoi_daidien_3_way4", "Way4");
        columns.put("SS_nguoi_daidien_3", "Check (True/ False)");
        columns.put("chuc_vu_3_prf", "PRF");
        columns.put("chuc_vu_3_way4", "Way4");
        columns.put("SS_chuc_vu_3", "Check (True/ False)");

        dynamicObject.setColumns(columns);

        Map<Integer, Object> inputParams = new HashMap<>();
        List<Integer> outParams = new ArrayList<>();
        outParams.add(1);

        List<DynamicObject> dynamicObjects = databaseService.callProcedure("REPORT_MIGRATE", "ACQ_003", columns,
                inputParams, outParams);

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
                        "Báo cáo đối chiếu thông tin khách hàng sau chuyển đổi - KHDN");
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
        row1.createCell(0).setCellValue("Mã báo cáo: ACQ_003");
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
        headerRow2.createCell(0).setCellStyle(cellStyle);
        headerRow2.getCell(0).setCellValue("Nguồn dữ liệu");

        for (int i = 1; i < columns.size(); i++) {
            headerRow2.createCell(i).setCellValue((String) columns.values().toArray()[i]);
            // resize column width
            // sheet.autoSizeColumn(i);
            sheet.setColumnWidth(i, 20 * 256);
            headerRow2.getCell(i).setCellStyle(cellStyle);
            headerRow.createCell(i).setCellStyle(cellStyle);
        }

        headerRow.getCell(1).setCellValue("Merchant ID");
        headerRow.getCell(2).setCellValue("Số CIF KH tại BIDV");
        sheet.addMergedRegion(new CellRangeAddress(5, 5, 2, 5));

        String[] header = {
                "Tên KH", "Số ĐKKD", "Mã số thuế", "Ngày ĐKKD + Cơ quan cấp",
                "Địa chỉ khách hàng trên ĐKKD",
                "Số điện thoại Mobile", "Số điện thoại Work", "Số điện thoại Home",
                "Email người đại diện 1", "Email người đại diện 2", "Email người đại diện 3",
                "Classifier cho phân loại KH zcomcode", "Người đại diện 1 (Gắn address_type của client)",
                "Chức danh người đại diện 1", "Người đại diện 2 (Gắn address_type của client)",
                "Chức danh người đại diện 2", "Người đại diện 3 (Gắn address_type của client)",
                "Chức danh người đại diện 3"
        };

        for (int i = 6; i < columns.size(); i += 3) {
            headerRow.getCell(i).setCellValue(header[i / 3 - 2]);
            sheet.addMergedRegion(new CellRangeAddress(5, 5, i, i + 2));
        }

        excelGenerator.writeExcel(fileName);
    }

}
