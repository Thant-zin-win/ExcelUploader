// File: ExportServlet.java
package com.example.servlet;

import com.example.processor.DatabaseExcelExporter;
import com.example.util.DatabaseConnection;
import javax.servlet.*;
import javax.servlet.annotation.WebServlet;
import javax.servlet.http.*;

import org.json.JSONObject;

import java.io.*;
import java.sql.*;
import java.net.URLEncoder; // Added for URL encoding
import java.nio.charset.StandardCharsets; // Added for UTF-8 charset

@WebServlet("/export")
public class ExportServlet extends HttpServlet {
	protected void doGet(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
		
		HttpSession session = request.getSession(false);
		if (session == null || session.getAttribute("userId") == null) {
		    response.sendError(HttpServletResponse.SC_UNAUTHORIZED, "Please log in");
		    return;
		}
		
		request.setCharacterEncoding("UTF-8");
	    String templateCategory = request.getParameter("templateCategory");
	    if (templateCategory == null || templateCategory.isEmpty()) {
	        sendJsonError(response, HttpServletResponse.SC_BAD_REQUEST, "Template Category is required");
	        return;
	    }

	    // Check if data exists for the template category
	    try (Connection conn = DatabaseConnection.getConnection();
	         PreparedStatement pstmt = conn.prepareStatement("SELECT COUNT(*) FROM Templates WHERE TemplateCategory = ?")) {
	        pstmt.setString(1, templateCategory);
	        try (ResultSet rs = pstmt.executeQuery()) {
	            if (rs.next() && rs.getInt(1) == 0) {
	                sendJsonError(response, HttpServletResponse.SC_NOT_FOUND, "No data found for Template Category: " + templateCategory);
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
	        
	        DatabaseExcelExporter.exportDatabaseToExcel(conn, baos, templateCategory);

	        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
	        
	        // FIXED: URL-encode the filename for Content-Disposition header
	        String fileName = templateCategory.replaceAll(" ", "_") + "_export.xlsx";
	        String encodedFileName = URLEncoder.encode(fileName, StandardCharsets.UTF_8.toString()).replaceAll("\\+", "%20"); // Handle spaces
	        
	        response.setHeader("Content-Disposition", "attachment; filename=\"" + encodedFileName + "\"");
	        // For better compatibility with some browsers, might also use filename* as per RFC 5987,
	        // but this simple filename encoding often suffices to prevent the IllegalArgumentException.
	        // response.setHeader("Content-Disposition", "attachment; filename*=UTF-8''" + encodedFileName);


	        try (OutputStream out = response.getOutputStream()) {
	            baos.writeTo(out);
	            out.flush();
	        }
	    } catch (Exception e) {
	        System.err.println("Error during Excel export for category: " + templateCategory);
	        e.printStackTrace(System.err); 

	        if (!response.isCommitted()) {
	            sendJsonError(response, HttpServletResponse.SC_INTERNAL_SERVER_ERROR, "Error exporting data: " + e.getMessage());
	        } else {
	            System.err.println("Response already committed. Cannot send JSON error response.");
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