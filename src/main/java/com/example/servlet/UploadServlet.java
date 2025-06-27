package com.example.servlet;

import com.example.processor.ExcelProcessor;
import com.example.util.DatabaseConnection;
import javax.servlet.*;
import javax.servlet.annotation.MultipartConfig;
import javax.servlet.annotation.WebServlet;
import javax.servlet.http.*;
import java.io.*;
import java.sql.Connection;
import java.util.ArrayList;
import java.util.Collection;
import java.util.List;
import org.json.JSONArray;
import org.json.JSONObject;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

@WebServlet("/upload")
@MultipartConfig
public class UploadServlet extends HttpServlet {
    private static final Logger logger = LoggerFactory.getLogger(UploadServlet.class);

    protected void doPost(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
    	
    	 HttpSession session = request.getSession(false);
         if (session == null || session.getAttribute("userId") == null) {
             response.sendError(HttpServletResponse.SC_UNAUTHORIZED, "Please log in");
             return;
         }
    	
    	request.setCharacterEncoding("UTF-8");
        response.setContentType("application/json; charset=UTF-8");
        JSONObject jsonResponse = new JSONObject();
        
        // --- MODIFIED: To store structured info for the modal ---
        List<JSONObject> uploadedFilesInfo = new ArrayList<>();
        int uploadedFilesCount = 0;


        try {
            Collection<Part> parts = request.getParts();
            boolean anyFileProcessed = false;

            for (Part filePart : parts) {
                if ("files".equals(filePart.getName())) { // Ensure it's the file input part
                    String fileName = filePart.getSubmittedFileName();
                    if (fileName != null && !fileName.isEmpty() && filePart.getContentType() != null) {

                        if (!fileName.toLowerCase().endsWith(".xlsx")) {
                            logger.warn("Skipping non-.xlsx file: {}", fileName);
                            continue;
                        }

                        try (Connection conn = DatabaseConnection.getConnection();
                             InputStream fileContent = filePart.getInputStream()) {
                            
                            ExcelProcessor.TemplateCreationResult result = ExcelProcessor.processExcelFile(conn, fileContent, fileName);
                            
                            // --- MODIFIED: Store filename and category for the modal ---
                            JSONObject fileInfo = new JSONObject();
                            fileInfo.put("fileName", fileName);
                            fileInfo.put("templateCategory", result.templateCategoryDisplayName);
                            uploadedFilesInfo.add(fileInfo);

                            uploadedFilesCount++;
                            anyFileProcessed = true;

                        } catch (Exception e) {
                            logger.error("Failed to process file: " + fileName, e);
                        }
                    }
                }
            }

            if (anyFileProcessed) {
                jsonResponse.put("status", "success");
                jsonResponse.put("uploadedFilesCount", uploadedFilesCount);
                // --- MODIFIED: Send the structured info ---
                jsonResponse.put("uploadedFilesInfo", new JSONArray(uploadedFilesInfo)); 
                response.setStatus(HttpServletResponse.SC_OK);
            } else {
                jsonResponse.put("status", "error");
                jsonResponse.put("message", "No valid .xlsx files were uploaded or found in the request.");
                response.setStatus(HttpServletResponse.SC_BAD_REQUEST);
            }

        } catch (Exception e) {
            logger.error("Failed to handle upload", e);
            jsonResponse.put("status", "error");
            jsonResponse.put("message", "Upload failed: " + e.getMessage());
            response.setStatus(HttpServletResponse.SC_INTERNAL_SERVER_ERROR);
        }

        response.getWriter().write(jsonResponse.toString());
    }
}