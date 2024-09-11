package com.fis.domain;

public class Branch {
    private String branchCode;
    private String branchName;
    private String folderPath;

    public Branch() {
    }

    public Branch(String branchCode, String branchName, String folderPath) {
        this.branchCode = branchCode;
        this.branchName = branchName;
        this.folderPath = folderPath;
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

    @Override
    public String toString() {
        return "Branch{" +
                "branchCode='" + branchCode + '\'' +
                ", branchName='" + branchName + '\'' +
                ", folderPath='" + folderPath + '\'' +
                '}';
    }

}
