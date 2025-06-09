package com.example.servlet;

import com.example.processor.ExcelProcessor;
import com.example.util.DatabaseConnection;
import javax.servlet.*;
import javax.servlet.annotation.MultipartConfig;
import javax.servlet.annotation.WebServlet;
import javax.servlet.http.*;
import java.io.*;
import java.sql.Connection;

import org.json.JSONObject;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

@WebServlet("/upload")
@MultipartConfig
public class UploadServlet extends HttpServlet {
    private static final Logger logger = LoggerFactory.getLogger(UploadServlet.class);

    protected void doPost(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
        response.setContentType("application/json; charset=UTF-8");
        JSONObject jsonResponse = new JSONObject();

        try {
            Part filePart = request.getPart("file");
            String fileName = filePart.getSubmittedFileName();
            if (!fileName.endsWith(".xlsx")) {
                jsonResponse.put("status", "error");
                jsonResponse.put("message", "Please upload an .xlsx file");
                response.setStatus(HttpServletResponse.SC_BAD_REQUEST);
                response.getWriter().write(jsonResponse.toString());
                return;
            }

            try (Connection conn = DatabaseConnection.getConnection();
                 InputStream fileContent = filePart.getInputStream()) {
                ExcelProcessor.processExcelFile(conn, fileContent, fileName);
                jsonResponse.put("status", "success");
                jsonResponse.put("fileName", fileName);
                response.setStatus(HttpServletResponse.SC_OK);
                response.getWriter().write(jsonResponse.toString());
            } catch (Exception e) {
                logger.error("Failed to process file: " + fileName, e);
                jsonResponse.put("status", "error");
                jsonResponse.put("message", "Processing failed: " + e.getMessage());
                response.setStatus(HttpServletResponse.SC_INTERNAL_SERVER_ERROR);
                response.getWriter().write(jsonResponse.toString());
            }
        } catch (Exception e) {
            logger.error("Failed to handle upload", e);
            jsonResponse.put("status", "error");
            jsonResponse.put("message", "Upload failed: " + e.getMessage());
            response.setStatus(HttpServletResponse.SC_INTERNAL_SERVER_ERROR);
            response.getWriter().write(jsonResponse.toString());
        }
    }
}