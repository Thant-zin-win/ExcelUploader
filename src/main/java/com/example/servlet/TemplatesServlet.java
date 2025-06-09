package com.example.servlet;

import com.example.util.DatabaseConnection;
import javax.servlet.*;
import javax.servlet.annotation.WebServlet;
import javax.servlet.http.*;
import java.io.*;
import java.sql.*;
import org.json.JSONArray;
import org.json.JSONObject;

@WebServlet("/templates")
public class TemplatesServlet extends HttpServlet {
    protected void doGet(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
        response.setContentType("application/json");
        JSONArray jsonArray = new JSONArray();

        try (Connection conn = DatabaseConnection.getConnection();
             Statement stmt = conn.createStatement();
             ResultSet rs = stmt.executeQuery("SELECT TemplateID, TemplateName, UploadDate FROM Templates")) {

            while (rs.next()) {
                JSONObject jsonObject = new JSONObject();
                jsonObject.put("TemplateID", rs.getInt("TemplateID"));
                jsonObject.put("TemplateName", rs.getString("TemplateName"));
                jsonObject.put("UploadDate", rs.getString("UploadDate"));
                jsonArray.put(jsonObject);
            }
            response.getWriter().write(jsonArray.toString());
        } catch (SQLException e) {
            response.sendError(HttpServletResponse.SC_INTERNAL_SERVER_ERROR, "Error fetching templates: " + e.getMessage());
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}