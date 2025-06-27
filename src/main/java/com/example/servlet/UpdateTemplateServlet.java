// File: UpdateTemplateServlet.java
package com.example.servlet;

import com.example.util.DatabaseConnection;
import org.json.JSONObject;

import javax.servlet.ServletException;
import javax.servlet.annotation.WebServlet;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import javax.servlet.http.HttpSession;
import java.io.BufferedReader;
import java.io.IOException;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.ArrayList; // NEW
import java.util.Comparator; // NEW
import java.util.List; // NEW

// NEW: Import for Logger
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

@WebServlet("/updateTemplateCategory")
public class UpdateTemplateServlet extends HttpServlet {
    // NEW: Logger instance
    private static final Logger logger = LoggerFactory.getLogger(UpdateTemplateServlet.class);

    protected void doPut(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
        HttpSession session = request.getSession(false);
        if (session == null || session.getAttribute("userId") == null) {
            response.sendError(HttpServletResponse.SC_UNAUTHORIZED, "Please log in");
            return;
        }

        response.setContentType("application/json; charset=UTF-8");
        response.setCharacterEncoding("UTF-8");
        JSONObject jsonResponse = new JSONObject();

        StringBuilder sb = new StringBuilder();
        try (BufferedReader reader = request.getReader()) {
            String line;
            while ((line = reader.readLine()) != null) {
                sb.append(line);
            }
        }

        JSONObject requestBody = new JSONObject(sb.toString());
        int templateIdToUpdate = requestBody.optInt("templateId", -1);
        String newCategoryName = requestBody.optString("newCategoryName", "").trim();

        logger.info("Received update request: TemplateID={}, NewCategoryName='{}'", templateIdToUpdate, newCategoryName); // NEW LOG

        if (templateIdToUpdate == -1 || newCategoryName.isEmpty()) {
            jsonResponse.put("status", "error");
            jsonResponse.put("message", "Invalid template ID or new category name provided.");
            response.setStatus(HttpServletResponse.SC_BAD_REQUEST);
            response.getWriter().write(jsonResponse.toString());
            logger.warn("Invalid request: templateIdToUpdate={}, newCategoryName='{}'", templateIdToUpdate, newCategoryName); // NEW LOG
            return;
        }

        try (Connection conn = DatabaseConnection.getConnection()) {
            conn.setAutoCommit(false); // Start transaction

            try {
                // 1. Check for conflict: Does the newCategoryName already exist for a different template?
                // Select the TemplateID of any conflicting template category
                String checkSql = "SELECT TemplateID FROM Templates WHERE TemplateCategory = ?";
                int conflictingTemplateId = -1;
                try (PreparedStatement checkStmt = conn.prepareStatement(checkSql)) {
                    checkStmt.setString(1, newCategoryName);
                    try (ResultSet rs = checkStmt.executeQuery()) {
                        if (rs.next()) {
                            conflictingTemplateId = rs.getInt("TemplateID");
                            logger.info("Found conflicting TemplateID {} for new category name '{}'.", conflictingTemplateId, newCategoryName); // NEW LOG
                        }
                    }
                }

                if (conflictingTemplateId != -1 && conflictingTemplateId != templateIdToUpdate) {
                    // There's a conflict with another template category.
                    // Now, check if this conflicting category has any associated responses.
                    if (hasResponses(conn, conflictingTemplateId)) { // NEW LOGGING INSIDE hasResponses
                        // The conflicting category is active (has responses), so block the rename.
                        jsonResponse.put("status", "error");
                        jsonResponse.put("message", "Template category '" + newCategoryName + "' is already in use by another active template.");
                        response.setStatus(HttpServletResponse.SC_CONFLICT); // 409 Conflict
                        response.getWriter().write(jsonResponse.toString());
                        conn.rollback();
                        logger.warn("Rename blocked: New category name '{}' is in use by active TemplateID {}", newCategoryName, conflictingTemplateId); // NEW LOG
                        return;
                    } else {
                        // The conflicting category exists but has no responses.
                        // We can "reclaim" this name by deleting the old, empty template entry.
                        String deleteConflictingTemplateSql = "DELETE FROM Templates WHERE TemplateID = ?";
                        try (PreparedStatement pstmt = conn.prepareStatement(deleteConflictingTemplateSql)) {
                            pstmt.setInt(1, conflictingTemplateId);
                            int deletedRows = pstmt.executeUpdate(); // NEW
                            logger.info("Deleted conflicting (empty) TemplateID {}: {} rows affected, freeing up name '{}'.", conflictingTemplateId, deletedRows, newCategoryName); // NEW LOG
                        }
                        // After deleting, renumber the auto-generated categories if necessary
                        renumberTemplateCategories(conn); // NEW LOGGING INSIDE renumberTemplateCategories
                    }
                }

                // 2. Perform the update on the original template
                String updateSql = "UPDATE Templates SET TemplateCategory = ? WHERE TemplateID = ?";
                try (PreparedStatement pstmt = conn.prepareStatement(updateSql)) {
                    pstmt.setString(1, newCategoryName);
                    pstmt.setInt(2, templateIdToUpdate);
                    int rowsAffected = pstmt.executeUpdate();

                    if (rowsAffected > 0) {
                        jsonResponse.put("status", "success");
                        jsonResponse.put("message", "Template category updated successfully to '" + newCategoryName + "'");
                        conn.commit(); // Commit transaction
                        logger.info("Successfully updated TemplateID {} to '{}'. Transaction committed.", templateIdToUpdate, newCategoryName); // NEW LOG
                    } else {
                        jsonResponse.put("status", "error");
                        jsonResponse.put("message", "Template with ID " + templateIdToUpdate + " not found or no changes made.");
                        response.setStatus(HttpServletResponse.SC_NOT_FOUND);
                        conn.rollback();
                        logger.warn("Update failed: TemplateID {} not found or no changes. Transaction rolled back.", templateIdToUpdate); // NEW LOG
                    }
                }
            } catch (SQLException e) {
                conn.rollback(); // Rollback on error
                jsonResponse.put("status", "error");
                jsonResponse.put("message", "Database error: " + e.getMessage());
                response.setStatus(HttpServletResponse.SC_INTERNAL_SERVER_ERROR);
                logger.error("Transaction rolled back for TemplateID {}: {}", templateIdToUpdate, e.getMessage(), e); // NEW LOG
            } finally {
                conn.setAutoCommit(true); // Restore auto-commit
            }
        } catch (SQLException e) {
            jsonResponse.put("status", "error");
            jsonResponse.put("message", "Database connection error: " + e.getMessage());
            response.setStatus(HttpServletResponse.SC_INTERNAL_SERVER_ERROR);
            logger.error("Database connection error: {}", e.getMessage(), e); // NEW LOG
        }

        response.getWriter().write(jsonResponse.toString());
    }

    // Helper method: Checks if a given TemplateID has any associated responses
    private boolean hasResponses(Connection conn, int templateId) throws SQLException {
        logger.debug("Checking hasResponses for TemplateID {}", templateId); // NEW LOG
        String countResponsesSql = "SELECT COUNT(*) FROM Responses WHERE TemplateID = ?";
        try (PreparedStatement pstmt = conn.prepareStatement(countResponsesSql)) {
            pstmt.setInt(1, templateId);
            try (ResultSet rs = pstmt.executeQuery()) {
                if (rs.next()) {
                    int count = rs.getInt(1);
                    logger.debug("TemplateID {} has {} responses.", templateId, count); // NEW LOG
                    return count > 0;
                }
            }
        }
        logger.debug("TemplateID {} has no responses (or query failed).", templateId); // NEW LOG
        return false;
    }

    /**
     * Renumbers TemplateCategory names (e.g., "Template Type 1", "Template Type 2")
     * to ensure they are sequential without gaps after a deletion.
     * This method also now attempts to delete Template entries that are "Template Type X"
     * but have no associated responses, cleaning up orphaned entries.
     * This is duplicated from DeleteTemplateServlet for now to ensure this servlet
     * has access to the same cleanup logic.
     * @param conn The database connection.
     * @throws SQLException If a database access error occurs.
     */
    private void renumberTemplateCategories(Connection conn) throws SQLException {
        logger.info("Starting renumberTemplateCategories (from UpdateTemplateServlet)..."); // NEW LOG
        List<TemplateCategoryInfo> categoriesToRenumber = new ArrayList<>();
        List<Integer> templateIdsToDelete = new ArrayList<>();

        String selectSql = "SELECT TemplateID, TemplateCategory FROM Templates WHERE TemplateCategory LIKE 'Template Type %' ORDER BY LENGTH(TemplateCategory), TemplateCategory";
        try (PreparedStatement pstmt = conn.prepareStatement(selectSql)) {
            try (ResultSet rs = pstmt.executeQuery()) {
                while (rs.next()) {
                    int templateId = rs.getInt("TemplateID");
                    String categoryName = rs.getString("TemplateCategory");
                    logger.debug("Found template: ID={}, Category='{}'", templateId, categoryName); // NEW LOG
                    
                    if (hasResponses(conn, templateId)) { // NEW LOGGING INSIDE hasResponses
                        categoriesToRenumber.add(new TemplateCategoryInfo(templateId, categoryName));
                        logger.debug("Template ID {} ('{}') has responses, keeping for renumbering.", templateId, categoryName); // NEW LOG
                    } else {
                        templateIdsToDelete.add(templateId);
                        logger.info("Marking TemplateID {} ('{}') for deletion as it has no associated responses.", templateId, categoryName); // NEW LOG
                    }
                }
            }
        }
        logger.info("Found {} templates to renumber, {} templates to delete (orphaned).", categoriesToRenumber.size(), templateIdsToDelete.size()); // NEW LOG

        if (!templateIdsToDelete.isEmpty()) {
            String deleteOrphanSql = "DELETE FROM Templates WHERE TemplateID IN (" +
                                     String.join(",", java.util.Collections.nCopies(templateIdsToDelete.size(), "?")) + ")";
            try (PreparedStatement pstmt = conn.prepareStatement(deleteOrphanSql)) {
                int i = 1;
                for (Integer id : templateIdsToDelete) {
                    pstmt.setInt(i++, id);
                }
                int deletedRows = pstmt.executeUpdate(); // NEW
                logger.info("Deleted {} orphaned template categories from DB.", deletedRows); // NEW LOG
            }
        }

        categoriesToRenumber.sort(Comparator.comparingInt(t -> {
            try {
                String numStr = t.categoryName.replace("Template Type ", "");
                return Integer.parseInt(numStr);
            } catch (NumberFormatException e) {
                logger.warn("Non-numeric category name found during sort: {}", t.categoryName); // NEW LOG
                return Integer.MAX_VALUE;
            }
        }));
        logger.debug("Sorted categories to renumber: {}", categoriesToRenumber.stream().map(t -> t.categoryName).collect(java.util.stream.Collectors.joining(", "))); // NEW LOG

        String updateSql = "UPDATE Templates SET TemplateCategory = ? WHERE TemplateID = ?";
        try (PreparedStatement pstmt = conn.prepareStatement(updateSql)) {
            int newNumber = 1;
            for (TemplateCategoryInfo info : categoriesToRenumber) {
                String expectedNewName = "Template Type " + newNumber;
                if (!info.categoryName.equals(expectedNewName)) {
                    pstmt.setString(1, expectedNewName);
                    pstmt.setInt(2, info.templateId);
                    pstmt.addBatch();
                    logger.info("Adding batch update: Renaming TemplateID {} from '{}' to '{}'", info.templateId, info.categoryName, expectedNewName); // NEW LOG
                } else {
                    logger.debug("TemplateID {} ('{}') already has expected name, no update needed.", info.templateId, info.categoryName); // NEW LOG
                }
                newNumber++;
            }
            int[] batchResults = pstmt.executeBatch(); // NEW
            logger.info("Executed batch updates for renumbering. {} items updated.", batchResults.length); // NEW LOG
            logger.info("Template categories renumbered successfully."); // NEW LOG
        }
    }

    private static class TemplateCategoryInfo {
        int templateId;
        String categoryName;

        TemplateCategoryInfo(int templateId, String categoryName) {
            this.templateId = templateId;
            this.categoryName = categoryName;
        }
    }
}