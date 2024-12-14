package com.fis.types;

import java.sql.CallableStatement;
import java.sql.Connection;
import java.sql.ResultSet;

public class JDBCResult {
    private  Connection connection;
    private  CallableStatement statement;
    private  String  call;

    public JDBCResult(Connection connection, CallableStatement statement, String call) {
        this.connection = connection;
        this.statement = statement;
        this.call = call;
    }

    public JDBCResult(Connection connection, CallableStatement statement) {
        this.connection = connection;
        this.statement = statement;
    }

    public String getCall() {
        return call;
    }
    public Connection getConnection() {
        return connection;
    }
    public CallableStatement getStatement() {
        return statement;
    }


    public void close() throws Exception {
        if (statement != null) statement.close();
        if (connection != null) connection.close();
    }
}
