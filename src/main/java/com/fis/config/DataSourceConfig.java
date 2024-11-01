package com.fis.config;

import com.zaxxer.hikari.HikariConfig;
import com.zaxxer.hikari.HikariDataSource;

import javax.sql.DataSource;

public class DataSourceConfig {
    private final static String url = "jdbc:oracle:thin:@//10.53.115.61:1521/way4migrate";
    private final static String username = "dmreport";
    private final static String password = "Bidv@123456";

    // private final String url = System.getenv("DB_URL");
    // private final String username = System.getenv("DB_USERNAME");
    // private final String password = System.getenv("DB_PASSWORD");

    public static DataSource getDataSource() {
        HikariConfig config = new HikariConfig();
        config.setJdbcUrl(url);
        config.setUsername(username);
        config.setPassword(password);
        config.setMaximumPoolSize(10); // Số kết nối tối đa
        config.setMinimumIdle(2); // Số kết nối tối thiểu
        config.setConnectionTimeout(30000); // Thời gian timeout kết nối
        // config.setIdleTimeout(600000); // Thời gian tối đa cho một kết nối không sử dụng

        return new HikariDataSource(config);
    }
}
