// File: DeleteTemplateServlet.java
package com.example.servlet;

import com.example.util.DatabaseConnection;
import javax.servlet.*;
import javax.servlet.annotation.WebServlet;
import javax.servlet.http.*;
import java.io.*;
import java.sql.*;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;
import java.util.HashSet; // NEW
import java.util.Set; // NEW
import org.json.JSONObject;
import org.json.JSONArray; // NEW

// Import for Logger (assuming SLF4J is set up)
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

@WebServlet("/delete")
public class DeleteTemplateServlet extends HttpServlet {
    private static final Logger logger = LoggerFactory.getLogger(DeleteTemplateServlet.class);

    protected void doDelete(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
    	
    	
    	HttpSession session = request.getSession(false);
    	if (session == null || session.getAttribute("userId") == null) {
    	    response.sendError(HttpServletResponse.SC_UNAUTHORIZED, "Please log in");
    	    return;
    	}
    	
    	request.setCharacterEncoding("UTF-8");
        response.setContentType("application/json");
        JSONObject jsonResponse = new JSONObject();

        // CHANGED: Read responseIds from JSON request body
        List<Integer> responseIdsToDelete = new ArrayList<>();
        try (BufferedReader reader = request.getReader()) {
            StringBuilder sb = new StringBuilder();
            String line;
            while ((line = reader.readLine()) != null) {
                sb.append(line);
            }
            JSONObject requestBody = new JSONObject(sb.toString());
            JSONArray jsonResponseIds = requestBody.optJSONArray("responseIds");

            if (jsonResponseIds == null || jsonResponseIds.isEmpty()) {
                jsonResponse.put("status", "error");
                jsonResponse.put("message", "Response IDs are required in the request body.");
                response.setStatus(HttpServletResponse.SC_BAD_REQUEST);
                response.getWriter().write(jsonResponse.toString());
                logger.warn("No responseIds provided in batch delete request.");
                return;
            }

            for (int i = 0; i < jsonResponseIds.length(); i++) {
                responseIdsToDelete.add(jsonResponseIds.getInt(i));
            }
        } catch (Exception e) {
            jsonResponse.put("status", "error");
            jsonResponse.put("message", "Invalid request body format: " + e.getMessage());
            response.setStatus(HttpServletResponse.SC_BAD_REQUEST);
            response.getWriter().write(jsonResponse.toString());
            logger.error("Failed to parse JSON body for batch delete: {}", e.getMessage(), e);
            return;
        }

        logger.info("Received batch delete request for ResponseIDs: {}", responseIdsToDelete); // NEW LOG

        int deletedCount = 0;
        Set<Integer> affectedTemplateIds = new HashSet<>(); // NEW: To track unique TemplateIDs affected by this batch deletion

        try (Connection conn = DatabaseConnection.getConnection()) {
            conn.setAutoCommit(false); // Start transaction for the entire batch
            try {
                // Prepare statements outside the loop for efficiency
                String getTemplateIdSql = "SELECT TemplateID FROM Responses WHERE ResponseID = ?";
                String deleteEvalSql = "DELETE FROM EvaluationData WHERE ResponseID = ?";
                String deleteMetaSql = "DELETE FROM ResponseMetadata WHERE ResponseID = ?";
                String deleteResponseSql = "DELETE FROM Responses WHERE ResponseID = ?";

                try (PreparedStatement pstmtGetTemplateId = conn.prepareStatement(getTemplateIdSql);
                     PreparedStatement pstmtDeleteEval = conn.prepareStatement(deleteEvalSql);
                     PreparedStatement pstmtDeleteMeta = conn.prepareStatement(deleteMetaSql);
                     PreparedStatement pstmtDeleteResponse = conn.prepareStatement(deleteResponseSql)) {

                    for (Integer responseId : responseIdsToDelete) {
                        int currentTemplateId = -1;
                        // Get the TemplateID before deleting the response itself
                        pstmtGetTemplateId.setInt(1, responseId);
                        try (ResultSet rs = pstmtGetTemplateId.executeQuery()) {
                            if (rs.next()) {
                                currentTemplateId = rs.getInt("TemplateID");
                                affectedTemplateIds.add(currentTemplateId); // Track affected template IDs
                                logger.debug("Found TemplateID {} for ResponseID {}", currentTemplateId, responseId); // NEW LOG
                            }
                        }

                        if (currentTemplateId == -1) {
                            logger.warn("ResponseID {} not found or already deleted. Skipping.", responseId); // NEW LOG
                            continue; // Skip to next responseId
                        }

                        // Delete from EvaluationData
                        pstmtDeleteEval.setInt(1, responseId);
                        int deletedEvalRows = pstmtDeleteEval.executeUpdate();
                        logger.debug("Deleted {} rows from EvaluationData for ResponseID {}", deletedEvalRows, responseId);

                        // Delete from ResponseMetadata
                        pstmtDeleteMeta.setInt(1, responseId);
                        int deletedMetaRows = pstmtDeleteMeta.executeUpdate();
                        logger.debug("Deleted {} rows from ResponseMetadata for ResponseID {}", deletedMetaRows, responseId);

                        // Delete the Response record itself
                        pstmtDeleteResponse.setInt(1, responseId);
                        int rowsAffected = pstmtDeleteResponse.executeUpdate();
                        if (rowsAffected > 0) {
                            deletedCount++;
                            logger.info("Deleted ResponseID {} successfully. Total deleted in batch: {}", responseId, deletedCount);
                        } else {
                            logger.warn("ResponseID {} not found during deletion. No rows affected.", responseId);
                        }
                    }
                }

                // --- After ALL responses in the batch are deleted, check affected TemplateIDs ---
                for (Integer templateId : affectedTemplateIds) {
                    String countResponsesSql = "SELECT COUNT(*) FROM Responses WHERE TemplateID = ?";
                    int remainingResponses = 0;
                    try (PreparedStatement pstmt = conn.prepareStatement(countResponsesSql)) {
                        pstmt.setInt(1, templateId);
                        try (ResultSet rs = pstmt.executeQuery()) {
                            if (rs.next()) {
                                remainingResponses = rs.getInt(1);
                            }
                        }
                    }
                    logger.info("TemplateID {} has {} remaining responses after batch deletions.", templateId, remainingResponses);

                    if (remainingResponses == 0) {
                        // No more responses for this template, delete the template entry
                        String deleteTemplateSql = "DELETE FROM Templates WHERE TemplateID = ?";
                        try (PreparedStatement pstmt = conn.prepareStatement(deleteTemplateSql)) {
                            pstmt.setInt(1, templateId);
                            int deletedTemplateRows = pstmt.executeUpdate();
                            if (deletedTemplateRows > 0) {
                                logger.info("Template (ID: {}) deleted as it has no more associated responses.", templateId);
                            } else {
                                logger.warn("Template (ID: {}) expected to be deleted but no rows affected. Already gone?", templateId);
                            }
                        }
                    }
                }
                
                // --- Finally, call renumberTemplateCategories ONCE after all batch processing ---
                renumberTemplateCategories(conn);

                conn.commit(); // Commit the entire transaction
                jsonResponse.put("status", "success");
                jsonResponse.put("message", deletedCount + " file(s)/sheet(s) deleted successfully.");
                jsonResponse.put("deletedCount", deletedCount); // NEW: Return count of successfully deleted items
                response.setStatus(HttpServletResponse.SC_OK);
                logger.info("Batch delete transaction committed successfully. Total deleted: {}", deletedCount);
            } catch (SQLException e) {
                conn.rollback(); // Rollback the entire transaction on any error
                jsonResponse.put("status", "error");
                jsonResponse.put("message", "Failed to delete files/sheets: " + e.getMessage());
                response.setStatus(HttpServletResponse.SC_INTERNAL_SERVER_ERROR);
                logger.error("Batch delete transaction rolled back: {}", e.getMessage(), e);
            } finally {
                conn.setAutoCommit(true);
            }
        } catch (SQLException e) {
            jsonResponse.put("status", "error");
            jsonResponse.put("message", "Database connection error: " + e.getMessage());
            response.setStatus(HttpServletResponse.SC_INTERNAL_SERVER_ERROR);
            logger.error("Database connection error during batch delete: {}", e.getMessage(), e);
        }

        response.getWriter().write(jsonResponse.toString());
    }

    /**
     * Renumbers TemplateCategory names (e.g., "Template Type 1", "Template Type 2")
     * to ensure they are sequential without gaps after a deletion.
     * This method also now attempts to delete Template entries that are "Template Type X"
     * but have no associated responses, cleaning up orphaned entries.
     * @param conn The database connection.
     * @throws SQLException If a database access error occurs.
     */
    private void renumberTemplateCategories(Connection conn) throws SQLException {
        logger.info("Starting renumberTemplateCategories...");
        List<TemplateCategoryInfo> categoriesToRenumber = new ArrayList<>();
        List<Integer> templateIdsToDelete = new ArrayList<>();

        // Fetch all "Template Type X" categories and their IDs
        String selectSql = "SELECT TemplateID, TemplateCategory FROM Templates WHERE TemplateCategory LIKE 'Template Type %' ORDER BY LENGTH(TemplateCategory), TemplateCategory";
        try (PreparedStatement pstmt = conn.prepareStatement(selectSql)) {
            try (ResultSet rs = pstmt.executeQuery()) {
                while (rs.next()) {
                    int templateId = rs.getInt("TemplateID");
                    String categoryName = rs.getString("TemplateCategory");
                    logger.debug("Found template: ID={}, Category='{}'", templateId, categoryName);
                    
                    // Check if this template category has any responses linked to it
                    if (hasResponses(conn, templateId)) {
                        categoriesToRenumber.add(new TemplateCategoryInfo(templateId, categoryName));
                        logger.debug("Template ID {} ('{}') has responses, keeping for renumbering.", templateId, categoryName);
                    } else {
                        // If it's a "Template Type X" and has no responses, mark for deletion
                        templateIdsToDelete.add(templateId);
                        logger.info("Marking TemplateID {} ('{}') for deletion as it has no associated responses.", templateId, categoryName);
                    }
                }
            }
        }
        logger.info("Found {} templates to renumber, {} templates to delete (orphaned).", categoriesToRenumber.size(), templateIdsToDelete.size());

        // NEW: Delete orphaned "Template Type X" entries first
        if (!templateIdsToDelete.isEmpty()) {
            String deleteOrphanSql = "DELETE FROM Templates WHERE TemplateID IN (" +
                                     String.join(",", java.util.Collections.nCopies(templateIdsToDelete.size(), "?")) + ")";
            try (PreparedStatement pstmt = conn.prepareStatement(deleteOrphanSql)) {
                int i = 1;
                for (Integer id : templateIdsToDelete) {
                    pstmt.setInt(i++, id);
                }
                int deletedRows = pstmt.executeUpdate();
                logger.info("Deleted {} orphaned template categories from DB.", deletedRows);
            }
        }

        // Sort the remaining active "Template Type X" categories numerically
        categoriesToRenumber.sort(Comparator.comparingInt(t -> {
            try {
                String numStr = t.categoryName.replace("Template Type ", "");
                return Integer.parseInt(numStr);
            } catch (NumberFormatException e) {
                logger.warn("Non-numeric category name found during sort: {}", t.categoryName);
                return Integer.MAX_VALUE; // Put malformed names at the end
            }
        }));
        logger.debug("Sorted categories to renumber: {}", categoriesToRenumber.stream().map(t -> t.categoryName).collect(java.util.stream.Collectors.joining(", ")));

        // Update the categories to be sequential
        String updateSql = "UPDATE Templates SET TemplateCategory = ? WHERE TemplateID = ?";
        try (PreparedStatement pstmt = conn.prepareStatement(updateSql)) {
            int newNumber = 1;
            for (TemplateCategoryInfo info : categoriesToRenumber) {
                String expectedNewName = "Template Type " + newNumber;
                if (!info.categoryName.equals(expectedNewName)) {
                    pstmt.setString(1, expectedNewName);
                    pstmt.setInt(2, info.templateId);
                    pstmt.addBatch();
                    logger.info("Adding batch update: Renaming TemplateID {} from '{}' to '{}'", info.templateId, info.categoryName, expectedNewName);
                } else {
                    logger.debug("TemplateID {} ('{}') already has expected name, no update needed.", info.templateId, info.categoryName);
                }
                newNumber++;
            }
            int[] batchResults = pstmt.executeBatch();
            logger.info("Executed batch updates for renumbering. {} items updated.", batchResults.length);
            logger.info("Template categories renumbered successfully.");
        }
    }

    // Helper method: Checks if a given TemplateID has any associated responses
    private boolean hasResponses(Connection conn, int templateId) throws SQLException {
        logger.debug("Checking hasResponses for TemplateID {}", templateId);
        String countResponsesSql = "SELECT COUNT(*) FROM Responses WHERE TemplateID = ?";
        try (PreparedStatement pstmt = conn.prepareStatement(countResponsesSql)) {
            pstmt.setInt(1, templateId);
            try (ResultSet rs = pstmt.executeQuery()) {
                if (rs.next()) {
                    int count = rs.getInt(1);
                    logger.debug("TemplateID {} has {} responses.", templateId, count);
                    return count > 0;
                }
            }
        }
        logger.debug("TemplateID {} has no responses (or query failed).", templateId);
        return false;
    }

    // Helper class to hold template category information for sorting and renumbering
    private static class TemplateCategoryInfo {
        int templateId;
        String categoryName;

        TemplateCategoryInfo(int templateId, String categoryName) {
            this.templateId = templateId;
            this.categoryName = categoryName;
        }
    }
}