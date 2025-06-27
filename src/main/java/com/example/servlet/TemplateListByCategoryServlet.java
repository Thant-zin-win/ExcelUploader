package com.example.servlet;

import com.example.util.DatabaseConnection;
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
import org.json.JSONArray;
import org.json.JSONObject;

@WebServlet("/templatesByCategory")
public class TemplateListByCategoryServlet extends HttpServlet {
    private static final long serialVersionUID = 1L;

    protected void doGet(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
        HttpSession session = request.getSession(false);
        if (session == null || session.getAttribute("userId") == null) {
            response.sendError(HttpServletResponse.SC_UNAUTHORIZED, "Please log in");
            return;
        }

        response.setCharacterEncoding("UTF-8");
        response.setContentType("application/json; charset=UTF-8");
        request.setCharacterEncoding("UTF-8");

        String templateCategoryDisplayName = request.getParameter("templateCategory"); // This is the user-friendly name, e.g., "Template Type 1"
        JSONArray jsonArray = new JSONArray();

        if (templateCategoryDisplayName == null || templateCategoryDisplayName.isEmpty()) {
            response.setStatus(HttpServletResponse.SC_BAD_REQUEST);
            JSONObject error = new JSONObject();
            error.put("status", "error");
            error.put("message", "Template category is required.");
            response.getWriter().write(error.toString());
            return;
        }

        try (Connection conn = DatabaseConnection.getConnection()) {
            // Step 1: Find the TemplateID for the given TemplateCategory (display name)
            // Since TemplateCategory is unique due to generateUniqueTemplateDisplayName, this will yield one TemplateID
            String getTemplateIdSql = "SELECT TemplateID FROM Templates WHERE TemplateCategory = ?";
            int templateId = -1;
            try (PreparedStatement pstmt = conn.prepareStatement(getTemplateIdSql)) {
                pstmt.setString(1, templateCategoryDisplayName);
                try (ResultSet rs = pstmt.executeQuery()) {
                    if (rs.next()) {
                        templateId = rs.getInt("TemplateID");
                    }
                }
            }

            if (templateId == -1) {
                // Should not happen if loadTemplateCategories works correctly, but as a safeguard
                response.setStatus(HttpServletResponse.SC_NOT_FOUND);
                JSONObject error = new JSONObject();
                error.put("status", "error");
                error.put("message", "Template category '" + templateCategoryDisplayName + "' not found.");
                response.getWriter().write(error.toString());
                return;
            }

            // Step 2: Now, select all individual responses (uploaded files/sheets) that belong to this template ID
            // We need ResponseID and OriginalFileName for display and deletion.
            String sql = "SELECT ResponseID, OriginalFileName, SheetName FROM Responses WHERE TemplateID = ? ORDER BY OriginalFileName, SheetName";
            try (PreparedStatement pstmt = conn.prepareStatement(sql)) {
                pstmt.setInt(1, templateId); // Filter by the found TemplateID
                try (ResultSet rs = pstmt.executeQuery()) {
                    while (rs.next()) {
                        JSONObject jsonObject = new JSONObject();
                        jsonObject.put("ResponseID", rs.getInt("ResponseID")); // This is what we will delete
                        jsonObject.put("OriginalFileName", rs.getString("OriginalFileName"));
                        jsonObject.put("SheetName", rs.getString("SheetName"));
                        jsonArray.put(jsonObject);
                    }
                }
            }
            response.getWriter().write(jsonArray.toString());
        } catch (SQLException e) {
            response.setStatus(HttpServletResponse.SC_INTERNAL_SERVER_ERROR);
            JSONObject error = new JSONObject();
            error.put("status", "error");
            error.put("message", "Database error fetching templates by category: " + e.getMessage());
            response.getWriter().write(error.toString());
        } catch (Exception e) {
            response.setStatus(HttpServletResponse.SC_INTERNAL_SERVER_ERROR);
            JSONObject error = new JSONObject();
            error.put("status", "error");
            error.put("message", "Server error: " + e.getMessage());
            response.getWriter().write(error.toString());
        }
    }
}