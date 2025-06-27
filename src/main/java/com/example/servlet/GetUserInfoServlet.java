package com.example.servlet;

import com.example.util.DatabaseConnection;
import org.json.JSONObject;

import javax.servlet.ServletException;
import javax.servlet.annotation.WebServlet;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import javax.servlet.http.HttpSession;
import java.io.IOException;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;

@WebServlet("/getUserInfo")
public class GetUserInfoServlet extends HttpServlet {
    protected void doGet(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
        response.setContentType("application/json; charset=UTF-8");
        response.setCharacterEncoding("UTF-8");
        JSONObject jsonResponse = new JSONObject();
        HttpSession session = request.getSession(false);

        if (session != null && session.getAttribute("userId") != null) {
            int userId = (int) session.getAttribute("userId");
            try (Connection conn = DatabaseConnection.getConnection()) {
                String sql = "SELECT Username FROM Users WHERE UserID = ?";
                try (PreparedStatement stmt = conn.prepareStatement(sql)) {
                    stmt.setInt(1, userId);
                    try (ResultSet rs = stmt.executeQuery()) {
                        if (rs.next()) {
                            jsonResponse.put("loggedIn", true);
                            jsonResponse.put("username", rs.getString("Username"));
                        } else {
                            jsonResponse.put("loggedIn", false);
                            jsonResponse.put("message", "User not found.");
                        }
                    }
                }
            } catch (SQLException e) {
                jsonResponse.put("loggedIn", false);
                jsonResponse.put("message", "Database error: " + e.getMessage());
                response.setStatus(HttpServletResponse.SC_INTERNAL_SERVER_ERROR);
            }
        } else {
            jsonResponse.put("loggedIn", false);
            jsonResponse.put("message", "Not logged in.");
        }
        response.getWriter().write(jsonResponse.toString());
    }
}