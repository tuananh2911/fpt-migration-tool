package com.fis;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashSet;
import java.util.List;
import java.util.Scanner;
import java.util.Set;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

import com.fis.domain.Branch;
import com.fis.services.ExcelReaderService;
import com.fis.services.ProgressTracker;
import com.fis.services.ReportService;

/**
 * Hello world!n
 */
public final class App {

    public static final String thread_num = System.getenv("THREAD_NUM");
    // public static final int thread_num = 10;
    private static final Logger logger = LogManager.getLogger(App.class);
    private static final String[] reportClassName = {
            "ISS_009",
            "ISS_010",
            "ISS_011",
            "ISS_012",
            "ISS_013",
            "ACQ_009",
            "ACQ_010",
    };
    // // no props
    private static final String[] reportHSCClassName = {
            "ATM_001",
            "ATM_002",
            "ATM_003",
            "ISS_001_0",
            "ISS_001_1",
            "ISS_001_2",
            "ISS_002",
            "ISS_003",
            "ISS_004",
            "ISS_004_1",
            "ISS_005",
            "ISS_006",
            "ISS_007",
            "ISS_008",
            "ISS_008_1",
            "ACQ_001",
            "ACQ_002",
            "ACQ_003",
            "ACQ_004",
            "ACQ_005",
            "ACQ_006",
            "ACQ_007",
            "ACQ_008",
            "ACQ_011",
            "GL_001_ISS",
            "GL_002_ISS_KH",
            "GL_004_ISS_KH",
            "GL_005_ISS_CT",
            "GL_005_ISS_TH",
    };

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
        List<Branch> ds = ExcelReaderService.readBranchExcel("ds_branch.xlsx");

        for (Branch branch : ds) {
            String folderPath = branch.getFolderPath();
            if (folderPath.contains("/FTPData")) {
                folderPath = folderPath.substring(folderPath.indexOf("FTPData"));
                branch.setFolderPath(folderPath);
            } else {
                branch.setFolderPath("FTPData/");
            }
        }
        // print title --- migrate
        System.out.println("------------------------- MIGRATE -------------------------");
        String[] allReport = new String[reportClassName.length + reportHSCClassName.length];
        System.arraycopy(reportClassName, 0, allReport, 0, reportClassName.length);
        System.arraycopy(reportHSCClassName, 0, allReport, reportClassName.length, reportHSCClassName.length);
        int columns = 3;
        int rows = (int) Math.ceil(allReport.length / (double) columns);
        // --- DANH SACH BAO CAO ---
        System.out.println("-------------------- DANH SACH BAO CAO --------------------");
        for (int i = 0; i < rows; i++) {
            for (int j = 0; j < columns; j++) {
                int index = i + j * rows;
                if (index < allReport.length) {
                    System.out.printf("%-20s", "\"" + allReport[index] + "\",");
                }
            }
            System.out.println();
        }
        System.out.println("-----------------------------------------------------------");

        System.out.print("Tao bao cao (all - tat ca | exit - thoat): ");
        try (Scanner scanner = new Scanner(System.in)) {
            ReportService reportService = new ReportService();
            String reportName = scanner.nextLine().trim();

            do {
                long startTime = System.currentTimeMillis();

                boolean isReportNameInReportClassName = false;
                boolean isReportNameInReportHSCClassName = false;
                boolean isArrayName = false;
                for (String report : reportClassName) {
                    if (report.equalsIgnoreCase(reportName)) {
                        isReportNameInReportClassName = true;
                        break;
                    }
                }
                for (String report : reportHSCClassName) {
                    if (report.equalsIgnoreCase(reportName)) {
                        isReportNameInReportHSCClassName = true;
                        break;
                    }
                }
                String[] arrayName = reportName.split(",");
                for (String report : allReport) {
                    for (String name : arrayName) {
                        if (report.equalsIgnoreCase(name.trim()) && arrayName.length > 1) {
                            isArrayName = true;
                            break;
                        }
                    }
                }
                if (!isReportNameInReportClassName && !isReportNameInReportHSCClassName
                        && !isArrayName
                        && !reportName.equalsIgnoreCase("exit") && !reportName.equalsIgnoreCase("all")) {
                    System.out.println("Bao cao khong ton tai");
                    System.out.print("Tao bao cao: ");
                    reportName = scanner.nextLine();
                    continue;
                } else if (isReportNameInReportClassName) {
                    try {
                        logger.info("Start working on " + reportName);
                        // nhap branch code
                        System.out.print("Nhap branch code: ");
                        String branchCode = scanner.nextLine().trim();
                        if (branchCode.equalsIgnoreCase("all")) {
                            // Thread for each branch

                            ExecutorService executor = Executors
                                    .newFixedThreadPool(Integer.parseInt(thread_num.trim()));
                            final String finalReportName = reportName;
                            for (Branch branch : ds) {
                                executor.execute(() -> {
                                    Long _startTime = System.currentTimeMillis();
                                    try {
                                        logger.info("Thread " + Thread.currentThread().getId() + " "
                                                + Thread.currentThread().getName() + " is working on " + finalReportName
                                                + " - "
                                                + branch.getBranchCode() + " - " + branch.getBranchName());
                                        reportService.getClass().getMethod(finalReportName, Branch.class)
                                                .invoke(reportService, branch);
                                        Thread.sleep(50);
                                    } catch (Exception e) {
                                        e.printStackTrace();
                                    } finally {
                                        long endTime = System.currentTimeMillis();
                                        logger.info("Thread " + Thread.currentThread().getId() + " "
                                                + Thread.currentThread().getName() + " finished working on "
                                                + finalReportName + " - "
                                                + branch.getBranchCode() + " - " + branch.getBranchName()
                                                + " (Duration: " + (endTime - _startTime) + " ms)");
                                    }
                                });
                            }

                            Thread.sleep(50);
                            continue;
                        }
                        Branch branch = null;
                        for (Branch b : ds) {
                            if (b.getBranchCode().equalsIgnoreCase(branchCode)) {
                                branch = b;
                                break;
                            }
                        }
                        if (branch == null) {
                            System.out.println("Branch khong ton tai");
                            System.out.print("Tao bao cao: ");
                            reportName = scanner.nextLine();
                            continue;
                        }
                        reportService.getClass().getMethod(reportName, Branch.class).invoke(reportService, branch);
                        Thread.sleep(50);
                    } catch (Exception e) {
                        e.printStackTrace();
                    } finally {
                        long endTime = System.currentTimeMillis();
                        logger.info("Finished working on " + reportName + " (Duration: " + (endTime - startTime)
                                + " ms)");
                    }
                } else if (isReportNameInReportHSCClassName) {
                    try {
                        logger.info("Start working on " + reportName);
                        reportService.getClass().getMethod(reportName.toUpperCase()).invoke(reportService);
                        Thread.sleep(50);
                    } catch (Exception e) {
                        e.printStackTrace();
                    } finally {
                        long endTime = System.currentTimeMillis();
                        logger.info(
                                "Finished working on " + reportName + " (Duration: " + (endTime - startTime) + " ms)");
                    }
                } else if (isArrayName) {
                    // export arrayName
                    exportListReport(ds, reportService, arrayName);
                } else if (reportName.equalsIgnoreCase("all")) {
                    exportAll(ds, reportService);
                    break;
                } else if (reportName.equalsIgnoreCase("exit")) {
                    break;
                }
                System.out.print("Tao bao cao (all - tat ca | exit - thoat): ");
                reportName = scanner.nextLine();
            } while (!reportName.equalsIgnoreCase("exit"));
        }

    }

    public static void exportAll(List<Branch> ds, ReportService reportService) {
        ExecutorService executor = Executors.newFixedThreadPool(Integer.parseInt(thread_num.trim()));

        ProgressTracker progressTracker = new ProgressTracker(ds.size() *
                reportClassName.length + reportHSCClassName.length);

        for (String report : reportHSCClassName) {

            executor.execute(() -> {
                long startTime = System.currentTimeMillis();
                try {
                    logger.info("Thread " + Thread.currentThread().getId() + " "
                            + Thread.currentThread().getName() + " is working on " + report);
                    reportService.getClass().getMethod(report).invoke(reportService);
                    Thread.sleep(50);
                } catch (Exception e) {
                    e.printStackTrace();
                } finally {
                    long endTime = System.currentTimeMillis();
                    progressTracker.taskCompleted();
                    logger.info("Thread " + Thread.currentThread().getId() + " "
                            + Thread.currentThread().getName() + " finished working on " + report
                            + " (Duration: " + (endTime - startTime) + " ms)");
                }
            });
        }
        for (Branch branch : ds) {
            if (branch.getReports() != null && branch.getReports().length != 0) {
                for (String report : branch.getReports()) {
                    executor.execute(() -> {
                        long startTime = System.currentTimeMillis();
                        try {
                            logger.info("Thread " + Thread.currentThread().getId() + " "
                                    + Thread.currentThread().getName() + " is working on " + report + " - "
                                    + branch.getBranchCode() + " - " + branch.getBranchName());
                            reportService.getClass().getMethod(report, Branch.class).invoke(reportService, branch);
                            Thread.sleep(50);
                        } catch (Exception e) {
                            e.printStackTrace();
                        } finally {
                            long endTime = System.currentTimeMillis();
                            progressTracker.taskCompleted();
                            logger.info("Thread " + Thread.currentThread().getId() + " "
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
                        reportService.getClass().getMethod(report, Branch.class).invoke(reportService, branch);
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
                        logger.info("Thread " + Thread.currentThread().getId() + " "
                                + Thread.currentThread().getName() + " finished working on " + report + " - "
                                + branch.getBranchCode() + " - " + branch.getBranchName()
                                + " (Duration: " + (endTime - startTime) + " ms)");
                    }
                });
            }
        }

        while (!progressTracker.isFinished()) {
            try {
                // add time delay to reduce CPU usage
                System.out.print("\0337"); // Save cursor position
                System.out.print("\033[999B"); // Move cursor to the bottom of the terminal
                System.out.print("\033[2K"); // Clear the entire line
                System.out.printf("Progress: %d%% %s",
                        progressTracker.getProgressPercentage(),
                        progressTracker.getProgressBar());
                System.out.print("\0338"); // Restore cursor position

                Thread.sleep(100);
            } catch (Exception e) {
                e.printStackTrace();
            }
        }

        executor.shutdown();
    }

    public static void exportListReport(List<Branch> ds, ReportService reportService, String[] reportName) {
        ExecutorService executor = Executors.newFixedThreadPool(Integer.parseInt(thread_num.trim()));

        Set<String> reportSet1 = new HashSet<>(Arrays.asList(reportClassName));
        Set<String> reportSet2 = new HashSet<>(Arrays.asList(reportHSCClassName));
        List<String> reportList = new ArrayList<>();
        List<String> reportListHSC = new ArrayList<>();
        for (String report : reportName) {
            if (reportSet1.contains(report.trim())) {
                reportList.add(report.trim());
            } else if (reportSet2.contains(report.trim())) {
                reportListHSC.add(report.trim());
            }
        }
        ProgressTracker progressTracker = new ProgressTracker(ds.size() * reportList.size() + reportListHSC.size());

        for (String report : reportListHSC) {
            executor.execute(() -> {
                long startTime = System.currentTimeMillis();
                try {
                    logger.info("Thread " + Thread.currentThread().getId() + " "
                            + Thread.currentThread().getName() + " is working on " + report);
                    reportService.getClass().getMethod(report).invoke(reportService);
                    Thread.sleep(50);
                } catch (Exception e) {
                    e.printStackTrace();
                } finally {
                    long endTime = System.currentTimeMillis();
                    progressTracker.taskCompleted();
                    logger.info("Thread " + Thread.currentThread().getId() + " "
                            + Thread.currentThread().getName() + " finished working on " + report
                            + " (Duration: " + (endTime - startTime) + " ms)");
                }
            });
        }
        for (Branch branch : ds) {
            for (String report : reportList) {
                executor.execute(() -> {
                    long startTime = System.currentTimeMillis();
                    try {
                        System.out.println("Thread " + Thread.currentThread().getId() + " "
                                + Thread.currentThread().getName() + " is working on " + report + " - "
                                + branch.getBranchCode() + " - " + branch.getBranchName());
                        reportService.getClass().getMethod(report, Branch.class).invoke(reportService, branch);
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
                        logger.info("Thread " + Thread.currentThread().getId() + " "
                                + Thread.currentThread().getName() + " finished working on " + report + " - "
                                + branch.getBranchCode() + " - " + branch.getBranchName()
                                + " (Duration: " + (endTime - startTime) + " ms)");
                    }
                });
            }
        }

        while (!progressTracker.isFinished()) {
            try {
                // add time delay to reduce CPU usage
                System.out.print("\0337"); // Save cursor position
                System.out.print("\033[999B"); // Move cursor to the bottom of the terminal
                System.out.print("\033[2K"); // Clear the entire line
                System.out.printf("Progress: %d%% %s",
                        progressTracker.getProgressPercentage(),
                        progressTracker.getProgressBar());
                System.out.print("\0338"); // Restore cursor position

                Thread.sleep(100);
            } catch (Exception e) {
                e.printStackTrace();
            }
        }

        executor.shutdown();
    }
}
