package com.example.servlet;

import com.example.util.DatabaseConnection;
import javax.servlet.*;
import javax.servlet.annotation.WebServlet;
import javax.servlet.http.*;
import java.io.*;
import java.sql.*;
import org.json.JSONObject;

@WebServlet("/delete")
public class DeleteTemplateServlet extends HttpServlet {
    protected void doDelete(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
        response.setContentType("application/json");
        JSONObject jsonResponse = new JSONObject();

        String templateIdStr = request.getParameter("templateId");
        if (templateIdStr == null || templateIdStr.isEmpty()) {
            jsonResponse.put("status", "error");
            jsonResponse.put("message", "Template ID is required");
            response.setStatus(HttpServletResponse.SC_BAD_REQUEST);
            response.getWriter().write(jsonResponse.toString());
            return;
        }

        int templateId;
        try {
            templateId = Integer.parseInt(templateIdStr);
        } catch (NumberFormatException e) {
            jsonResponse.put("status", "error");
            jsonResponse.put("message", "Invalid Template ID");
            response.setStatus(HttpServletResponse.SC_BAD_REQUEST);
            response.getWriter().write(jsonResponse.toString());
            return;
        }

        try (Connection conn = DatabaseConnection.getConnection()) {
            conn.setAutoCommit(false);
            try {
                // Delete from EvaluationData
                String deleteEvalSql = "DELETE ed FROM EvaluationData ed JOIN Responses r ON ed.ResponseID = r.ResponseID WHERE r.TemplateID = ?";
                try (PreparedStatement pstmt = conn.prepareStatement(deleteEvalSql)) {
                    pstmt.setInt(1, templateId);
                    pstmt.executeUpdate();
                }

                // Delete from ResponseMetadata
                String deleteMetaSql = "DELETE rm FROM ResponseMetadata rm JOIN Responses r ON rm.ResponseID = r.ResponseID WHERE r.TemplateID = ?";
                try (PreparedStatement pstmt = conn.prepareStatement(deleteMetaSql)) {
                    pstmt.setInt(1, templateId);
                    pstmt.executeUpdate();
                }

                // Delete from Responses
                String deleteResponsesSql = "DELETE FROM Responses WHERE TemplateID = ?";
                try (PreparedStatement pstmt = conn.prepareStatement(deleteResponsesSql)) {
                    pstmt.setInt(1, templateId);
                    pstmt.executeUpdate();
                }

                // Delete from Templates
                String deleteTemplateSql = "DELETE FROM Templates WHERE TemplateID = ?";
                try (PreparedStatement pstmt = conn.prepareStatement(deleteTemplateSql)) {
                    pstmt.setInt(1, templateId);
                    int rowsAffected = pstmt.executeUpdate();
                    if (rowsAffected > 0) {
                        jsonResponse.put("status", "success");
                        jsonResponse.put("message", "Template deleted successfully");
                    } else {
                        jsonResponse.put("status", "error");
                        jsonResponse.put("message", "Template not found");
                        response.setStatus(HttpServletResponse.SC_NOT_FOUND);
                    }
                }

                conn.commit();
            } catch (SQLException e) {
                conn.rollback();
                jsonResponse.put("status", "error");
                jsonResponse.put("message", "Failed to delete template: " + e.getMessage());
                response.setStatus(HttpServletResponse.SC_INTERNAL_SERVER_ERROR);
            } finally {
                conn.setAutoCommit(true);
            }
        } catch (SQLException e) {
            jsonResponse.put("status", "error");
            jsonResponse.put("message", "Database connection error: " + e.getMessage());
            response.setStatus(HttpServletResponse.SC_INTERNAL_SERVER_ERROR);
        }

        response.getWriter().write(jsonResponse.toString());
    }
}