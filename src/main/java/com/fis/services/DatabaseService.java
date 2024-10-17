package com.fis.services;

import java.sql.*;
import java.util.*;

public class DatabaseService {

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
            connection = DriverManager.getConnection(url, username, password);
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
    public List<DynamicObject> callProcedure(String packageName, String procedureName, Map<String, String> columns,
            Map<Integer, Object> inputParams, List<Integer> outParams) throws SQLException {
        List<DynamicObject> resultList = new ArrayList<>();

        // Build the SQL call string dynamically based on the number of parameters
        String call = buildProcedureCall(packageName, procedureName, inputParams.size(), outParams.size());

        Connection connection = null;
        try {
            connection = DriverManager.getConnection(url, username, password);
            CallableStatement callableStatement = connection.prepareCall(call);

            // Set input parameters
            for (Map.Entry<Integer, Object> entry : inputParams.entrySet()) {
                callableStatement.setObject(entry.getKey(), entry.getValue());
            }

            // Register output parameters
            for (Integer outParamIndex : outParams) {
                callableStatement.registerOutParameter(outParamIndex, Types.REF_CURSOR);
            }

            // Execute the procedure
            callableStatement.execute();

            // Map the result set to a list of DynamicObjects
            for (Integer outParamIndex : outParams) {
                ResultSet resultSet = (ResultSet) callableStatement.getObject(outParamIndex);
                // check resultSet is empty
                if (resultSet.isBeforeFirst()) {
                    resultList.addAll(mapResultSetToDynamicObject(resultSet, columns));
                }
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

        return resultList;
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
    private List<DynamicObject> mapResultSetToDynamicObject(ResultSet resultSet, Map<String, String> columns)
            throws SQLException {
        List<DynamicObject> resultList = new ArrayList<>();
        ResultSetMetaData metaData = resultSet.getMetaData();
        int columnCount = metaData.getColumnCount();

        while (resultSet.next()) {
            Map<String, Object> properties = new LinkedHashMap<>();
            // for each key in columns init properties with key and value = ""
            for (String key : columns.keySet()) {
                if (key.equalsIgnoreCase("STT") || key.equalsIgnoreCase("TT")) {
                    continue;
                }
                properties.put(key, "");
            }

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

                        properties.put(columns.keySet().toArray()[j].toString(), columnValue);
                    }
                }
            }
            resultList.add(new DynamicObject(properties, columns));
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
