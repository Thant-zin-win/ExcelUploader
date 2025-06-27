package com.example.servlet;

import javax.servlet.ServletException;
import javax.servlet.annotation.WebServlet;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import org.mindrot.jbcrypt.BCrypt;

import com.example.util.DatabaseConnection;

import java.io.IOException;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;

@WebServlet("/register")
public class RegisterServlet extends HttpServlet {
    protected void doPost(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
        String username = request.getParameter("username");
        String email = request.getParameter("email");
        String password = request.getParameter("password");
        String confirmPassword = request.getParameter("confirmPassword");

        if (!password.equals(confirmPassword)) {
            response.sendRedirect("register.html?error=passwords_do_not_match");
            return;
        }

        if (!email.matches("^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\\.[a-zA-Z]{2,}$")) {
            response.sendRedirect("register.html?error=invalid_email");
            return;
        }

        try (Connection conn = DatabaseConnection.getConnection()) {
            String checkSql = "SELECT COUNT(*) FROM Users WHERE Username = ? OR Email = ?";
            try (PreparedStatement checkStmt = conn.prepareStatement(checkSql)) {
                checkStmt.setString(1, username);
                checkStmt.setString(2, email);
                try (ResultSet rs = checkStmt.executeQuery()) {
                    if (rs.next() && rs.getInt(1) > 0) {
                        response.sendRedirect("register.html?error=user_exists");
                        return;
                    }
                }
            }

            String passwordHash = BCrypt.hashpw(password, BCrypt.gensalt());
            String insertSql = "INSERT INTO Users (Username, Email, PasswordHash) VALUES (?, ?, ?)";
            try (PreparedStatement insertStmt = conn.prepareStatement(insertSql)) {
                insertStmt.setString(1, username);
                insertStmt.setString(2, email);
                insertStmt.setString(3, passwordHash);
                insertStmt.executeUpdate();
            }

            response.sendRedirect("login.html?registered=1");
        } catch (SQLException e) {
            response.sendRedirect("register.html?error=database_error");
        }
    }
}