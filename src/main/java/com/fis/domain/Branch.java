package com.fis.domain;

public class Branch {
    private String branchCode;
    private String branchName;
    private String folderPath;
    private String[] reports;

    public Branch() {
    }

    public Branch(String branchCode, String branchName, String folderPath, String[] reports) {
        this.branchCode = branchCode;
        this.branchName = branchName;
        this.folderPath = folderPath;
        this.reports = reports;
    }

    public String getBranchCode() {
        return branchCode;
    }

    public void setBranchCode(String branchCode) {
        this.branchCode = branchCode;
    }

    public String getBranchName() {
        return branchName;
    }

    public void setBranchName(String branchName) {
        this.branchName = branchName;
    }

    public String getFolderPath() {
        return folderPath;
    }

    public void setFolderPath(String folderPath) {
        this.folderPath = folderPath;
    }

    public String[] getReports() {
        return reports;
    }

    public void setReports(String[] reports) {
        this.reports = reports;
    }

    @Override
    public String toString() {
        return "Branch [branchCode=" + branchCode + ", branchName=" + branchName + ", folderPath=" + folderPath
                + ", reports=" + reports + "]";
    }

}
