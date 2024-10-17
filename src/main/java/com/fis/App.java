package com.fis;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.sql.SQLException;
import java.util.List;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;

import com.fis.domain.Branch;
import com.fis.services.ExcelReaderService;
import com.fis.services.ProgressTracker;
import com.fis.services.ReportService;

/**
 * Hello world!
 */
public final class App {

    public static final String thread_num = System.getenv("THREAD_NUM");
    // public static final int thread_num = 10;

    private App() {
    }

    /**
     * Says hello to the world.
     *
     * @param args The arguments of the program.
     * @throws IOException
     * @throws FileNotFoundException
     * @throws SQLException
     * @throws InterruptedException
     */
    public static void main(String[] args)
            throws FileNotFoundException, IOException, SQLException, InterruptedException {
        // System.out.println("Start");
        List<Branch> ds = ExcelReaderService.readBranchExcel("ds_branch.xlsx");
        String[] reportClassName = {
                "ISS_009",
                "ISS_010",
                "ISS_011",
                "ISS_012",
                "ISS_013",
                "ACQ_009",
                "ACQ_010",
        };
        // // no props
        String[] reportHSCClassName = {
                "ATM_001",
                "ATM_002",
                "ATM_003",
                // "GL_007",
                "ISS_001_0",
                "ISS_001_1",
                "ISS_001_2",
                "ISS_002",
                "ISS_003",
                // "ISS_004",
                "ISS_004_1",
                "ISS_005",
                "ISS_006",
                // "ISS_007",
                // "ISS_008_1",
                "ACQ_001",
                "ACQ_002",
                "ACQ_003",
                "ACQ_004",
                "ACQ_005",
                "ACQ_006",
                "ACQ_007",
                "ACQ_008",
                "GL_005_ISS_CT",
                "ACQ_011",
        };
        System.out.println("Init data");
        // ReportService.initData();
        System.out.println("Init data done");

        ExecutorService executor = Executors.newFixedThreadPool(Integer.parseInt(thread_num.trim()));

        ProgressTracker progressTracker = new ProgressTracker(ds.size() *
                reportClassName.length + reportHSCClassName.length);
        for (Branch branch : ds) {
            String folderPath = branch.getFolderPath();
            if (folderPath.contains("/FTPData")) {
                folderPath = folderPath.substring(folderPath.indexOf("/FTPData"));
                branch.setFolderPath(folderPath);
            } else {
                branch.setFolderPath("/FTPData/");
            }

            if(branch.getReports() != null && branch.getReports().length != 0){
                for (String report : branch.getReports()) {
                    executor.execute(() -> {
                        long startTime = System.currentTimeMillis();
                        try {
                            System.out.println("Thread " + Thread.currentThread().getId() + " "
                                    + Thread.currentThread().getName() + " is working on " + report + " - "
                                    + branch.getBranchCode() + " - " + branch.getBranchName());
                            ReportService.class.getMethod(report, Branch.class).invoke(null, branch);
                            Thread.sleep(50);
                        } catch (Exception e) {
                            e.printStackTrace();
                        } finally {
                            long endTime = System.currentTimeMillis();
                            progressTracker.taskCompleted();
                            System.out.println("Thread " + Thread.currentThread().getId() + " "
                                    + Thread.currentThread().getName() + " finished working on " + report + " - "
                                    + branch.getBranchCode() + " - " + branch.getBranchName()
                                    + " (Duration: " + (endTime - startTime) + " ms)");
                        }
                    });
                }
                continue;
            }

            for (String report : reportClassName) {
                executor.execute(() -> {
                    long startTime = System.currentTimeMillis();
                    try {
                        System.out.println("Thread " + Thread.currentThread().getId() + " "
                                + Thread.currentThread().getName() + " is working on " + report + " - "
                                + branch.getBranchCode() + " - " + branch.getBranchName());
                        ReportService.class.getMethod(report, Branch.class).invoke(null, branch);
                        Thread.sleep(50);
                    } catch (Exception e) {
                        e.printStackTrace();
                    } finally {
                        long endTime = System.currentTimeMillis();
                        progressTracker.taskCompleted();
                        System.out.println("Thread " + Thread.currentThread().getId() + " "
                                + Thread.currentThread().getName() + " finished working on " + report + " - "
                                + branch.getBranchCode() + " - " + branch.getBranchName()
                                + " (Duration: " + (endTime - startTime) + " ms)");
                    }
                });
            }
        }

        for (String report : reportHSCClassName) {

            executor.execute(() -> {
                long startTime = System.currentTimeMillis();
                try {
                    System.out.println("Thread " + Thread.currentThread().getId() + " "
                            + Thread.currentThread().getName() + " is working on " + report);
                    ReportService.class.getMethod(report).invoke(null);
                    Thread.sleep(50);
                } catch (Exception e) {
                    e.printStackTrace();
                } finally {
                    long endTime = System.currentTimeMillis();
                    progressTracker.taskCompleted();
                    System.out.println("Thread " + Thread.currentThread().getId() + " "
                            + Thread.currentThread().getName() + " finished working on " + report
                            + " (Duration: " + (endTime - startTime) + " ms)");
                }
            });

        }

        while (!progressTracker.isFinished()) {
            try {
                // add time delay to reduce CPU usage
                System.out.print("\0337"); // Save cursor position
                System.out.print("\033[999B"); // Move cursor to the bottom of the terminal
                System.out.print("\033[2K"); // Clear the entire line
                System.out.printf("Progress: %d%% %s", progressTracker.getProgressPercentage(),
                        progressTracker.getProgressBar());
                System.out.print("\0338"); // Restore cursor position

                Thread.sleep(100);
            } catch (Exception e) {
                e.printStackTrace();
            }
        }

        executor.shutdown();
        // Branch branch1 = new Branch("120", "Chi Nhánh Sở Giao dịch 1","/FTPData/ChiNhanh/MienBac/SGD1/NHAN/");
        // ReportService.ACQ_009(branch1);
        // // ReportService.ISS_009(branch1);
        // ReportService.ISS_011(branch1);
        // ReportService.ISS_004_1();
        // ReportService.ACQ_004();
        // ReportService.ACQ_003();
        // ReportService.ACQ_001();
        // ReportService.ACQ_002();
        // ReportService.ACQ_005();
        // ReportService.ACQ_007();
        // ReportService.ACQ_008();
        // ReportService.ISS_012(branch1);
        // ReportService.ATM_001();
        // ReportService.ATM_002();
        // ReportService.GL_005_ISS_CT();
        // ReportService.ATM_003();
        // ReportService.ACQ_011();
        // ReportService.ISS_001_1();
        // System.out.println("End");
    }
}
