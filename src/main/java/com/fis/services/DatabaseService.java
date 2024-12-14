package com.fis.services;

import java.io.BufferedWriter;
import java.io.FileWriter;
import java.io.IOException;
import java.sql.*;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Date;
import javax.sql.DataSource;

import com.fis.types.JDBCResult;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

import com.fis.config.DataSourceConfig;

public class DatabaseService {

    private DataSource dataSource;

    public DatabaseService() {
        this.dataSource = DataSourceConfig.getDataSource();
    }

    private final Logger logger = LogManager.getLogger(DatabaseService.class);
    // Oracle JDBC connection details
    private final String url = "jdbc:oracle:thin:@//10.53.115.61:1521/way4migrate";
    private final String username = "dmreport";
    private final String password = "Bidv@123456";
    // private final String url = System.getenv("DB_URL");
    // private final String username = System.getenv("DB_USERNAME");
    // private final String password = System.getenv("DB_PASSWORD");

    public void initData(String packageName, String procedureName) {
        // call procedure
        String call = "{call " + packageName + "." + procedureName + "}";
        Connection connection = null;
        try {
            connection = dataSource.getConnection();
            CallableStatement callableStatement = connection.prepareCall(call);
            callableStatement.execute();
        } catch (SQLException e) {
            e.printStackTrace();
        } finally {
            try {
                connection.close();
            } catch (SQLException e) {
                e.printStackTrace();
            }
        }
    }

    // Method to dynamically call a stored procedure from an Oracle package and map
    // result to DynamicObject
    public JDBCResult callProcedureV2(String packageName, String procedureName, Map<String, String> columns,
            Map<Integer, Object> inputParams, List<Integer> outParams) throws SQLException {
        long startTime = System.currentTimeMillis();
        // Build the SQL call string dynamically based on the number of parameters
        String call = buildProcedureCall(packageName, procedureName, inputParams.size(), outParams.size());
        Connection connection = dataSource.getConnection();
        try (
            CallableStatement statement = connection.prepareCall(call)) {
//            statement.registerOutParameter(1, Types.REF_CURSOR);
//            statement.setFetchSize(1000);
//            for (Map.Entry<Integer, Object> entry : inputParams.entrySet()) {
//                statement.setObject(entry.getKey(), entry.getValue());
//            }
//            for (Integer outParamIndex : outParams) {
//                statement.registerOutParameter(outParamIndex, Types.REF_CURSOR);
//            }
            System.out.println(" inputParams = " + inputParams + ", outParams = " + outParams);
//            statement.execute();

            long endGet = System.currentTimeMillis();
            System.out.println("TOTAL TIME GET DATA: " + (endGet - startTime) + "ms");
            return new JDBCResult(connection, null, call);

        } catch (Exception e) {
            e.printStackTrace();
        }
        long endTime = System.currentTimeMillis();
        logger.trace("Thread " + Thread.currentThread().getId() + " " + Thread.currentThread().getName()
                + " query done on " + procedureName + " (Duration: " + (endTime - startTime) + " ms)");
        return null;
    }

    public List<DynamicObject> callProcedure(String packageName, String procedureName, Map<String, String> columns,
                                             Map<Integer, Object> inputParams, List<Integer> outParams) throws SQLException {
        List<DynamicObject> resultList = new ArrayList<>();
        long startTime = System.currentTimeMillis();
        // Build the SQL call string dynamically based on the number of parameters
        String call = buildProcedureCall(packageName, procedureName, inputParams.size(), outParams.size());

        Connection connection = null;
        try {
            connection = dataSource.getConnection();
            CallableStatement statement = connection.prepareCall(call);
            statement.registerOutParameter(1, Types.REF_CURSOR);
            statement.setFetchSize(1000);
            statement.execute();
            ResultSet resultSet = (ResultSet) statement.getObject(1);
            // Set input parameters
            for (Map.Entry<Integer, Object> entry : inputParams.entrySet()) {
                statement.setObject(entry.getKey(), entry.getValue());
            }

//            // Register output parameters
//            for (Integer outParamIndex : outParams) {
//                callableStatement.registerOutParameter(outParamIndex, Types.REF_CURSOR);
//            }
//            callableStatement.setFetchSize(1000);
            // Execute the procedure
//            callableStatement.execute();

            // Map the result set to a list of DynamicObjects
//            for (Integer outParamIndex : outParams) {
//                ResultSet resultSet = (ResultSet) callableStatement.getObject(outParamIndex);
//                // check resultSet is empty
//
//            }


            if (resultSet.isBeforeFirst()) {
                resultList.addAll(mapResultSetToDynamicObject(resultSet, columns));
            }



            // Map the output parameters to a DynamicObject
            // if (!outParams.isEmpty()) {
            // resultList.add(mapOutParamsToDynamicObject(callableStatement, outParams,
            // columns));
            // }

        } catch (SQLException e) {
            e.printStackTrace();
        } finally {
            connection.close();
        }
        long endTime = System.currentTimeMillis();
        logger.trace("Thread " + Thread.currentThread().getId() + " " + Thread.currentThread().getName()
            + " query done on " + procedureName + " (Duration: " + (endTime - startTime) + " ms)");
        return resultList;
    }
    private static void writeDataToFile(ResultSet resultSet, BufferedWriter writer, long startTime) throws Exception {
        // Ghi tiêu đề
        writer.write("ID,Name,Value");
        writer.newLine();

        // Ghi dữ liệu
        int rowCount = 0;
        System.out.println("Size of resultSet: " + resultSet.getFetchSize());
        while (resultSet.next()) {

            ResultSetMetaData metaData = resultSet.getMetaData();
            int columnCount = metaData.getColumnCount();
            StringBuilder row = new StringBuilder();

            for (int i = 1; i <= columnCount; i++) {
                if (i > 1) {
                    row.append(",");
                }
                // Xử lý giá trị null
                String value = resultSet.getString(i);
                row.append(value != null ? value : "");
            }
            writer.write(row.toString());
            writer.newLine();
            rowCount++;

            if (rowCount % 100000 == 0) {
                System.out.println("Đã xử lý: " + rowCount + " bản ghi");
                System.out.println("START TIME: " + new Date());
            }
        }

        System.out.println("\nTổng số bản ghi đã xử lý: " + rowCount);
        System.out.println("END TIME: " + new Date());
        long endTime = System.currentTimeMillis();
        System.out.println("TOTAL TIME: " + (endTime - startTime)/1000 + " seconds");
    }

    // 30-9-2024 vietnq
    // Helper method to dynamically build the procedure call
    private String buildProcedureCall(String packageName, String procedureName, int inputParamCount,
            int outputParamCount) {
        StringBuilder sb = new StringBuilder("{call ");
        sb.append(packageName).append(".").append(procedureName).append("(");

        // Add placeholders for input and output parameters
        int totalParams = inputParamCount + outputParamCount;
        for (int i = 0; i < totalParams; i++) {
            if (i > 0)
                sb.append(", ");
            sb.append("?");
        }
        sb.append(")}");

        return sb.toString();
    }

    // 30-9-2024 vietnq
    // Helper method to map the result set to a list of DynamicObjects
    public List<DynamicObject> mapResultSetToDynamicObject(ResultSet resultSet, Map<String, String> columns) throws SQLException {
        List<DynamicObject> resultList = new ArrayList<>();
        ResultSetMetaData metaData = resultSet.getMetaData();
        int columnCount = metaData.getColumnCount();

        // Pre-compute column name mapping to avoid repeated lookups
        Map<Integer, String> columnIndexToKey = new HashMap<>();
        for (int i = 1; i <= columnCount; i++) {
            String columnName = metaData.getColumnName(i);
            for (String key : columns.keySet()) {
                if (key.equalsIgnoreCase(columnName)) {
                    columnIndexToKey.put(i, key);
                    break;
                }
            }
        }

        // Create date formatter once
        DateFormat dateFormatter = new SimpleDateFormat("dd/MM/yyyy");

        // Process in batches
        int batchSize = 1000;
        int rowCount = 0;

        while (resultSet.next()) {
            Map<String, Object> properties = new LinkedHashMap<>(columns.size());

            // Initialize properties with empty values only for relevant columns
            for (String key : columns.keySet()) {
                if (!key.equalsIgnoreCase("STT") && !key.equalsIgnoreCase("TT")) {
                    properties.put(key, "");
                }
            }

            // Direct column mapping using pre-computed index
            for (int i = 1; i <= columnCount; i++) {
                String key = columnIndexToKey.get(i);
                if (key != null) {
                    Object value = resultSet.getObject(i);
                    if (value != null) {
                        if ("DATE".equals(columns.get(key))) {
                            value = dateFormatter.format(resultSet.getDate(i));
                        }
                        properties.put(key, value);
                    }
                }
            }

            resultList.add(new DynamicObject(properties, columns));

            // Log only every 1000 records
            if (++rowCount % batchSize == 0) {
                System.out.println("Processed " + rowCount + " records");
            }
        }

        return resultList;
    }

    private List<DynamicObject> mapResultSetToDynamicObjectWithKey(ResultSet resultSet, Map<String, String> columns)
            throws SQLException {
        List<DynamicObject> resultList = new ArrayList<>();
        ResultSetMetaData metaData = resultSet.getMetaData();
        int columnCount = metaData.getColumnCount();

        while (resultSet.next()) {
            Map<String, Object> properties = new LinkedHashMap<>();
            for (int j = 0; j < columns.size(); j++) {
                for (int i = 1; i <= columnCount; i++) {
                    String columnName = metaData.getColumnName(i);
                    String columnType = columns.get(columnName);
                    Object columnValue = resultSet.getObject(i);

                    // Perform type-specific operations if needed
                    if ("DATE".equals(columnType)) {
                        columnValue = resultSet.getDate(i);
                    }

                    // check null return ""
                    if (columnValue == null) {
                        columnValue = "";
                    }

                    // check column[j] key is equal to column name
                    if (columns.keySet().toArray()[j].toString().equalsIgnoreCase(columnName)) {

                        properties.put(columnName, columnValue);
                    }
                }
            }
            resultList.add(new DynamicObject(properties, columns));
        }
        return resultList;
    }

    // Helper method to map the output parameters to a DynamicObject
    private DynamicObject mapOutParamsToDynamicObject(CallableStatement callableStatement, List<Integer> outParams,
            Map<String, String> columns) throws SQLException {
        Map<String, Object> properties = new HashMap<>();
        for (Integer outParamIndex : outParams) {
            ResultSet resultSet = (ResultSet) callableStatement.getObject(outParamIndex);
            ResultSetMetaData metaData = resultSet.getMetaData();
            int columnCount = metaData.getColumnCount();

            while (resultSet.next()) {
                for (int i = 1; i <= columnCount; i++) {
                    String columnName = metaData.getColumnName(i);
                    String columnType = columns.get(columnName);
                    Object columnValue = resultSet.getObject(i);

                    // Perform type-specific operations if needed
                    if ("DATE".equals(columnType)) {
                        columnValue = resultSet.getDate(i);
                    }
                    properties.put(columnName, columnValue);
                }
            }
        }
        return new DynamicObject(properties, columns);
    }
}
