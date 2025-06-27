package com.example.servlet;

import com.example.util.DatabaseConnection;
import javax.servlet.*;
import javax.servlet.annotation.WebServlet;
import javax.servlet.http.*;
import java.io.*;
import java.sql.*;
import org.json.JSONArray;
import org.json.JSONObject;
import org.slf4j.Logger; // NEW
import org.slf4j.LoggerFactory; // NEW

@WebServlet("/templates")
public class TemplatesServlet extends HttpServlet {
    private static final Logger logger = LoggerFactory.getLogger(TemplatesServlet.class); // NEW

    protected void doGet(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {

        HttpSession session = request.getSession(false);
        if (session == null || session.getAttribute("userId") == null) {
            response.sendError(HttpServletResponse.SC_UNAUTHORIZED, "Please log in");
            logger.warn("Unauthorized access to /templates. Session invalid or userId missing."); // NEW
            return;
        }

        response.setCharacterEncoding("UTF-8");
        response.setContentType("application/json; charset=UTF-8");

        JSONArray jsonArray = new JSONArray();

        try (Connection conn = DatabaseConnection.getConnection();
             Statement stmt = conn.createStatement();
             ResultSet rs = stmt.executeQuery(
                 "SELECT T.TemplateCategory, MIN(T.TemplateID) AS SampleTemplateID " +
                 "FROM Templates T " +
                 "JOIN Responses R ON T.TemplateID = R.TemplateID " +
                 "GROUP BY T.TemplateCategory " +
                 "ORDER BY T.TemplateCategory"
             )) {

            while (rs.next()) {
                JSONObject jsonObject = new JSONObject();
                jsonObject.put("TemplateCategory", rs.getString("TemplateCategory"));
                jsonObject.put("SampleTemplateID", rs.getInt("SampleTemplateID"));
                jsonArray.put(jsonObject);
            }
            response.getWriter().write(jsonArray.toString());
            logger.info("Successfully fetched {} template categories.", jsonArray.length()); // NEW
        } catch (SQLException e) {
            logger.error("Database error fetching templates: {}", e.getMessage(), e); // NEW
            response.sendError(HttpServletResponse.SC_INTERNAL_SERVER_ERROR, "Error fetching templates: " + e.getMessage());
        } catch (Exception e) {
            logger.error("Unexpected error fetching templates: {}", e.getMessage(), e); // NEW
            response.sendError(HttpServletResponse.SC_INTERNAL_SERVER_ERROR, "An unexpected error occurred: " + e.getMessage()); // Provide more specific error
        }
    }
}