package com.example.util;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;

public class DatabaseConnection {
    private static final String URL = "jdbc:mysql://localhost:3306/second_pj?useSSL=false&serverTimezone=UTC";
    private static final String USER = "root"; // Replace with  MySQL username
    private static final String PASSWORD = "root"; // Replace with  MySQL password

    public static Connection getConnection() throws SQLException {
        try {
            Class.forName("com.mysql.cj.jdbc.Driver");
            return DriverManager.getConnection(URL, USER, PASSWORD);
        } catch (ClassNotFoundException e) {
            throw new SQLException("MySQL driver not found: " + e.getMessage(), e);
        }
    }
}	