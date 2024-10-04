package com.fis;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.sql.SQLException;
import java.util.List;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;

import org.apache.poi.ss.formula.functions.T;

import com.fis.domain.Branch;
import com.fis.services.ExcelReaderService;
import com.fis.services.ProgressTracker;
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
        System.out.println("Start");
        List<Branch> ds = ExcelReaderService.readBranchExcel("ds_branch.xlsx");
        String[] reportClassName = {
                // "ISS009Report",
                // "ISS010Report",
                // "ISS011Report",
                // "ISS012Report",
                // "ISS013Report",
                "ACQ009Report",
                // "ACQ010Report",
        };
        // // no props
        String[] reportHSCClassName = {
                // "ATM001REPORT",
                "ATM002REPORT",
                "ATM003REPORT",
                // "GL005ISSReport",
                // "GL007Report",
                // "ISS0010Report",
                // "ISS011Report",
                // "ISS002Report",
                "ISS003Report",
                // "ISS005Report",
                // "ISS006Report",
                // "ISS007Report",
                // "ISS0081Report",
                // "ACQ001Report",
                // "ACQ006Report",
                // "ACQ007Report",
                // "ACQ008Report",
                // "ACQ011Report",
                // "ISS0012Report",
                // "ISS0041Report",
        };

        // ExecutorService executor = Executors.newFixedThreadPool(reportClassName.length + reportHSCClassName.length);

        // ProgressTracker progressTracker = new ProgressTracker(ds.size() * reportClassName.length);
        // for (Branch branch : ds) {
        //     String folderPath = branch.getFolderPath();
        //     if (folderPath.contains("/FTPData")) {
        //         folderPath = folderPath.substring(folderPath.indexOf("/FTPData"));
        //         branch.setFolderPath(folderPath);
        //     } else {
        //         branch.setFolderPath("/FTPData/");
        //     }

        //     for (String report : reportClassName) {
        //         executor.execute(() -> {
        //             try {
        //                 ReportService.class.getMethod(report, Branch.class).invoke(null, branch);
        //                 Thread.sleep(50);
        //             } catch (Exception e) {
        //                 e.printStackTrace();
        //             } finally {
        //                 progressTracker.taskCompleted();
        //                 System.out.println(report + " " + progressTracker.getProgressBar() + " "
        //                         + progressTracker.getProgressPercentage() + "%");
        //             }
        //         });
        //     }
        // }

        // for (String report : reportHSCClassName) {
        //     ProgressTracker progressTracker1 = new ProgressTracker(1);
        //     executor.execute(() -> {
        //         try {
        //             ReportService.class.getMethod(report).invoke(null);
        //             Thread.sleep(50);
        //         } catch (Exception e) {
        //             e.printStackTrace();
        //         } finally {
        //             progressTracker1.taskCompleted();
        //             System.out.println(report + " " + progressTracker1.getProgressBar() + " "
        //                     + progressTracker1.getProgressPercentage() + "%");
        //         }
        //     });

        // }

        // executor.shutdown();
        Branch branch1 = new Branch("215", "Chi Nhánh Cầu Giấy", "/FTPData/ChiNhanh/MienBac/CauGiay/NHAN/");
        // ReportService.ACQ009Report(branch1);
        // ReportService.ISS009Report(branch1);
        ReportService.ISS011Report(branch1);
        // ReportService.ISS012Report(branch1);
        System.out.println("End");

        // for (Branch branch : ds) {
        // String folderPath = branch.getFolderPath();
        // if (folderPath.contains("/FTPData")) {
        // folderPath = folderPath.substring(folderPath.indexOf("/FTPData"));
        // branch.setFolderPath(folderPath);
        // } else {
        // branch.setFolderPath("/FTPData/");
        // }
        // ReportService.ISS009Report(branch);
        // ReportService.ISS010Report(branch);
        // ReportService.ISS011Report(branch);
        // ReportService.ISS012Report(branch);
        // ReportService.ISS013Report(branch);
        // ReportService.ACQ009Report(branch);
        // ReportService.ACQ010Report(branch);
        // }
        // ReportService.ATM001REPORT();
        // ReportService.ATM002REPORT();
        // ReportService.ATM003REPORT();
        // ReportService.GL005ISSReport();
        // ReportService.GL007Report();
        // ReportService.ISS0010Report();
        // ReportService.ISS0011Report();
        // ReportService.ISS002Report();
        // ReportService.ISS003Report();
        // ReportService.ISS005Report();
        // ReportService.ISS006Report();
        // ReportService.ISS007Report();
        // ReportService.ISS0081Report();
        // ReportService.ACQ001Report();
        // ReportService.ACQ006Report();
        // ReportService.ACQ007Report();
        // ReportService.ACQ008Report();
        // ReportService.ACQ011Report();
    }
}
