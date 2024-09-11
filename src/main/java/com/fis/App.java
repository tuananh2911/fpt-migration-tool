package com.fis;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.sql.SQLException;
import java.util.List;

import com.fis.domain.Branch;
import com.fis.services.ExcelReaderService;
import com.fis.services.ReportService;

/**
 * Hello world!
 */
public final class App {
    private App() {
    }

    /**
     * Says hello to the world.
     *
     * @param args The arguments of the program.
     * @throws IOException
     * @throws FileNotFoundException
     * @throws SQLException
     */
    public static void main(String[] args) throws FileNotFoundException, IOException, SQLException {
        List<Branch> ds = ExcelReaderService.readBranchExcel("ds_branch.xlsx");
        // cut branch folderPath from /FTPData
        // for (Branch branch : ds) {
        //     String folderPath = branch.getFolderPath();
        //     if (folderPath.contains("/FTPData")) {
        //         folderPath = folderPath.substring(folderPath.indexOf("/FTPData"));
        //         branch.setFolderPath(folderPath);
        //     }else{
        //         branch.setFolderPath("/FTPData/" );
        //     }
        //     ReportService.ISS011Report(branch);
        //     ReportService.ISS012Report(branch);
        //     ReportService.ACQ009Report(branch);
        // }
        // // ReportService.CMS018Report();
        // ReportService.ATM002REPORT();
        // ReportService.GL007Report();
        ReportService.GL005ISSReport();
    }
}
