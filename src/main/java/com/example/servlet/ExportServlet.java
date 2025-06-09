package com.example.servlet;

import com.example.processor.DatabaseExcelExporter;
import com.example.util.DatabaseConnection;
import javax.servlet.*;
import javax.servlet.annotation.WebServlet;
import javax.servlet.http.*;

import org.json.JSONObject;

import java.io.*;
import java.sql.*;

@WebServlet("/export")
public class ExportServlet extends HttpServlet {
	protected void doGet(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
	    String templateIdStr = request.getParameter("templateId");
	    if (templateIdStr == null || templateIdStr.isEmpty()) {
	        sendJsonError(response, HttpServletResponse.SC_BAD_REQUEST, "Template ID is required");
	        return;
	    }

	    int templateId;
	    try {
	        templateId = Integer.parseInt(templateIdStr);
	    } catch (NumberFormatException e) {
	        sendJsonError(response, HttpServletResponse.SC_BAD_REQUEST, "Invalid Template ID");
	        return;
	    }

	    // Check if data exists for the template
	    try (Connection conn = DatabaseConnection.getConnection();
	         PreparedStatement pstmt = conn.prepareStatement("SELECT COUNT(*) FROM Responses WHERE TemplateID = ?")) {
	        pstmt.setInt(1, templateId);
	        try (ResultSet rs = pstmt.executeQuery()) {
	            if (rs.next() && rs.getInt(1) == 0) {
	                sendJsonError(response, HttpServletResponse.SC_NOT_FOUND, "No data found for Template ID: " + templateId);
	                return;
	            }
	        }
	    } catch (SQLException e) {
	        sendJsonError(response, HttpServletResponse.SC_INTERNAL_SERVER_ERROR, "Database error: " + e.getMessage());
	        return;
	    }

	    // Generate the Excel file first without committing the response
	    try (Connection conn = DatabaseConnection.getConnection();
	         ByteArrayOutputStream baos = new ByteArrayOutputStream()) {
	        // Export to ByteArrayOutputStream
	        DatabaseExcelExporter.exportDatabaseToExcel(conn, baos, templateId);

	        // If successful, set headers and write to output stream
	        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
	        response.setHeader("Content-Disposition", "attachment; filename=\"template_export_" + templateId + ".xlsx\"");
	        try (OutputStream out = response.getOutputStream()) {
	            baos.writeTo(out);
	            out.flush();
	        }
	    } catch (Exception e) {
	        // Only send JSON error if the response isn’t committed
	        if (!response.isCommitted()) {
	            sendJsonError(response, HttpServletResponse.SC_INTERNAL_SERVER_ERROR, "Error exporting data: " + e.getMessage());
	        } else {
	            // Log the error if the response is already committed
	            System.err.println("Error exporting data after response committed: " + e.getMessage());
	        }
	    }
	}

    private void sendJsonError(HttpServletResponse response, int status, String message) throws IOException {
        response.setContentType("application/json");
        response.setStatus(status);
        JSONObject json = new JSONObject();
        json.put("status", "error");
        json.put("message", message);
        response.getWriter().write(json.toString());
    }
}