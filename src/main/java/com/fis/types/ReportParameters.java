package com.fis.types;

import java.sql.ResultSet;
import java.util.Date;
import java.util.Map;

public class ReportParameters {
    private ResultSet resultSet;
    private String baseFileName;
    private Map<String, String> columns;
    private String[] groupHeaders;

    private String reportName;
    private String branchCode;
    private String branchName;
    private String reportCode;
    private Date reportDate;

    // Constructor
    public ReportParameters(ResultSet resultSet, String baseFileName, Map<String, String> columns, String[] groupHeaders,
                            String reportName, String branchCode, String branchName, String reportCode, Date reportDate) {
        this.resultSet = resultSet;
        this.baseFileName = baseFileName;
        this.columns = columns;
        this.groupHeaders = groupHeaders;
        this.reportName = reportName;
        this.branchCode = branchCode;
        this.branchName = branchName;
        this.reportCode = reportCode;
        this.reportDate = reportDate;
    }

    // Getters và Setters
    public ResultSet getResultSet() {
        return resultSet;
    }

    public void setResultSet(ResultSet resultSet) {
        this.resultSet = resultSet;
    }

    public String getBaseFileName() {
        return baseFileName;
    }

    public void setBaseFileName(String baseFileName) {
        this.baseFileName = baseFileName;
    }

    public Map<String, String> getColumns() {
        return columns;
    }

    public void setColumns(Map<String, String> columns) {
        this.columns = columns;
    }

    public String[] getGroupHeaders() {
        return groupHeaders;
    }

    public void setGroupHeaders(String[] groupHeaders) {
        this.groupHeaders = groupHeaders;
    }

    public String getReportName() {
        return reportName;
    }

    public void setReportName(String reportName) {
        this.reportName = reportName;
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

    public String getReportCode() {
        return reportCode;
    }

    public void setReportCode(String reportCode) {
        this.reportCode = reportCode;
    }

    public Date getReportDate() {
        return reportDate;
    }

    public void setReportDate(Date reportDate) {
        this.reportDate = reportDate;
    }

    @Override
    public String toString() {
        return "ReportParameters{" +
            "resultSet=" + resultSet +
            ", baseFileName='" + baseFileName + '\'' +
            ", columns=" + columns +
            ", groupHeaders=" + String.join(", ", groupHeaders) +
            ", reportName='" + reportName + '\'' +
            ", branchCode='" + branchCode + '\'' +
            ", branchName='" + branchName + '\'' +
            ", reportCode='" + reportCode + '\'' +
            ", reportDate=" + reportDate +
            '}';
    }
}
