package com.example.servlet;

import com.example.util.DatabaseConnection;
import org.json.JSONArray;
import org.json.JSONObject;

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
import java.sql.Timestamp;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.TimeZone;

@WebServlet("/recentFiles")
public class RecentFilesServlet extends HttpServlet {
    private static final long serialVersionUID = 1L;

    protected void doGet(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
        HttpSession session = request.getSession(false);
        if (session == null || session.getAttribute("userId") == null) {
            response.sendError(HttpServletResponse.SC_UNAUTHORIZED, "Please log in");
            return;
        }

        response.setCharacterEncoding("UTF-8");
        response.setContentType("application/json; charset=UTF-8");
        JSONObject jsonResponse = new JSONObject();
        JSONArray filesArray = new JSONArray();

        int page = 1;
        int limit = 10;

        try {
            page = Integer.parseInt(request.getParameter("page"));
        } catch (NumberFormatException e) {
            // Default to 1 if not provided or invalid
        }
        try {
            limit = Integer.parseInt(request.getParameter("limit"));
        }
        catch (NumberFormatException e) {
            // Default to 10 if not provided or invalid
        }

        int offset = (page - 1) * limit;

        try (Connection conn = DatabaseConnection.getConnection()) {
            String countSql = "SELECT COUNT(DISTINCT CONCAT(r.OriginalFileName, '_', t.TemplateCategory)) " +
                              "FROM Responses r JOIN Templates t ON r.TemplateID = t.TemplateID";
            int totalRecords = 0;
            try (PreparedStatement countStmt = conn.prepareStatement(countSql);
                 ResultSet rs = countStmt.executeQuery()) {
                if (rs.next()) {
                    totalRecords = rs.getInt(1);
                }
            }

            // *** MODIFIED SQL QUERY: Added MAX(r.ResponseID) DESC as a secondary sort ***
            String sql = """
                SELECT r.OriginalFileName, t.TemplateCategory, MAX(r.LastUpdated) AS LatestFileUpdate, MAX(r.ResponseID) AS LatestResponseId
                FROM Responses r
                JOIN Templates t ON r.TemplateID = t.TemplateID
                GROUP BY r.OriginalFileName, t.TemplateCategory
                ORDER BY LatestFileUpdate DESC, LatestResponseId DESC -- Sort by most recent update, then by latest response ID for stable ordering
                LIMIT ? OFFSET ?
            """;
            try (PreparedStatement pstmt = conn.prepareStatement(sql)) {
                pstmt.setInt(1, limit);
                pstmt.setInt(2, offset);
                try (ResultSet rs = pstmt.executeQuery()) {
                    SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
                    dateFormat.setTimeZone(TimeZone.getTimeZone("UTC"));

                    while (rs.next()) {
                        JSONObject file = new JSONObject();
                        file.put("FileName", rs.getString("OriginalFileName"));
                        file.put("TemplateCategory", rs.getString("TemplateCategory"));

                        Timestamp uploadTimestamp = rs.getTimestamp("LatestFileUpdate");
                        if (uploadTimestamp != null) {
                            file.put("UploadDate", dateFormat.format(new Date(uploadTimestamp.getTime())));
                        } else {
                            file.put("UploadDate", "N/A");
                        }

                        filesArray.put(file);
                    }
                }
            }

            jsonResponse.put("status", "success");
            jsonResponse.put("files", filesArray);
            jsonResponse.put("totalRecords", totalRecords);
            jsonResponse.put("currentPage", page);
            jsonResponse.put("limit", limit);

        } catch (SQLException e) {
            jsonResponse.put("status", "error");
            jsonResponse.put("message", "Database error: " + e.getMessage());
            response.setStatus(HttpServletResponse.SC_INTERNAL_SERVER_ERROR);
        } catch (Exception e) {
            jsonResponse.put("status", "error");
            jsonResponse.put("message", "Server error: " + e.getMessage());
            response.setStatus(HttpServletResponse.SC_INTERNAL_SERVER_ERROR);
        }

        response.getWriter().write(jsonResponse.toString());
    }
}