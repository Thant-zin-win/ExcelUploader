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
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Collectors;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

@WebServlet("/preview")
public class ResultPreviewServlet extends HttpServlet {
    private static final Logger logger = LoggerFactory.getLogger(ResultPreviewServlet.class);

    private static final Pattern MAIN_ITEM_NUMBER_EXTRACTOR = Pattern.compile("^(?:[0-9０-９]+)");

    private static Integer extractLeadingNumber(String s) {
        Matcher matcher = MAIN_ITEM_NUMBER_EXTRACTOR.matcher(s);
        return matcher.find() ? Integer.parseInt(matcher.group(0)) : null;
    }


    protected void doGet(HttpServletRequest request, HttpServletResponse response) throws IOException {

        HttpSession session = request.getSession(false);
        if (session == null || session.getAttribute("userId") == null) {
            response.sendError(HttpServletResponse.SC_UNAUTHORIZED, "Please log in");
            return;
        }

        request.setCharacterEncoding("UTF-8");
        response.setContentType("application/json; charset=UTF-8");
        response.setCharacterEncoding("UTF-8");
        JSONObject jsonResponse = new JSONObject();

        String templateCategory = request.getParameter("templateCategory");
        String sheetNameParam = request.getParameter("sheetName");

        logger.info("Received preview request: templateCategory={}, sheetName={}", templateCategory, sheetNameParam);

        if (templateCategory == null || templateCategory.isEmpty()) {
            jsonResponse.put("status", "error");
            jsonResponse.put("message", "Template category is required.");
            response.setStatus(HttpServletResponse.SC_BAD_REQUEST);
            response.getWriter().append(jsonResponse.toString());
            logger.warn("Invalid or missing templateCategory: {}", templateCategory);
            return;
        }

        try (Connection conn = DatabaseConnection.getConnection()) {
            String getTemplateIdSql = "SELECT TemplateID FROM Templates WHERE TemplateCategory = ?";
            int templateId = -1;
            try (PreparedStatement pstmt = conn.prepareStatement(getTemplateIdSql)) {
                pstmt.setString(1, templateCategory);
                try (ResultSet rs = pstmt.executeQuery()) {
                    if (rs.next()) {
                        templateId = rs.getInt("TemplateID");
                    }
                }
            }

            if (templateId == -1) {
                jsonResponse.put("status", "error");
                jsonResponse.put("message", "No templates found for category: " + templateCategory);
                response.setStatus(HttpServletResponse.SC_NOT_FOUND);
                logger.warn("No templates found for category: {}", templateCategory);
                response.getWriter().append(jsonResponse.toString());
                return;
            }

            JSONArray sheetNames = new JSONArray();
            String sheetSql = "SELECT SheetName FROM Responses WHERE TemplateID = ? GROUP BY SheetName ORDER BY MIN(ResponseID)";
            try (PreparedStatement pstmt = conn.prepareStatement(sheetSql)) {
                pstmt.setInt(1, templateId);
                try (ResultSet rs = pstmt.executeQuery()) {
                    while (rs.next()) {
                        String name = rs.getString("SheetName");
                        if (name != null) {
                            name = name.trim();
                            if (!name.isEmpty()) {
                                sheetNames.put(name);
                            }
                        }
                    }
                }
            }

            if (sheetNames.length() == 0) {
                jsonResponse.put("status", "error");
                jsonResponse.put("message", "No sheets found for Template Category: " + templateCategory);
                response.setStatus(HttpServletResponse.SC_NOT_FOUND);
                logger.warn("No sheets found for Template Category: {}", templateCategory);
                response.getWriter().append(jsonResponse.toString());
                return;
            }

            String selectedSheet = sheetNameParam != null ? sheetNameParam.trim() : null;
            if (selectedSheet == null || selectedSheet.isEmpty() || !containsSheet(sheetNames, selectedSheet)) {
                selectedSheet = sheetNames.getString(0);
            }
            logger.info("Selected sheet after validation: {}", selectedSheet);

            List<String> sortedMetadataHeaders = new ArrayList<>();
            String headerSql = """
                SELECT rm.HeaderKey
                FROM ResponseMetadata rm
                JOIN Responses r ON rm.ResponseID = r.ResponseID
                WHERE r.TemplateID = ? AND r.SheetName = ? AND rm.HeaderValue IS NOT NULL AND rm.HeaderValue != ''
                GROUP BY rm.HeaderKey
                ORDER BY MIN(rm.MetadataID)
            """;
            try (PreparedStatement pstmt = conn.prepareStatement(headerSql)) {
                pstmt.setInt(1, templateId);
                pstmt.setString(2, selectedSheet);
                try (ResultSet rs = pstmt.executeQuery()) {
                    while (rs.next()) {
                        String key = rs.getString("HeaderKey");
                        if (key != null && !key.trim().isEmpty()) {
                             sortedMetadataHeaders.add(key);
                        }
                    }
                }
            }

            logger.info("Found {} active HeaderKeys for sheet {}: {}", sortedMetadataHeaders.size(), selectedSheet, sortedMetadataHeaders);


            Map<String, Map<String, String>> responsesMetadata = new LinkedHashMap<>();
            Map<String, List<String>> mainItemsMap = new LinkedHashMap<>();
            Map<String, Map<String, List<String>>> subItemsMap = new LinkedHashMap<>();
            Map<String, Map<String, Map<String, Map<String, String>>>> evalDataMap = new LinkedHashMap<>();

            // *** MODIFIED SQL QUERY: PRIMARY ORDERING BY IsReuploaded ASC, then ResponseID ASC ***
            String dataSql = """
                SELECT r.ResponseID, r.OriginalFileName, r.SheetName, rm.HeaderKey, rm.HeaderValue,
                       e.MainItem, e.SubItem, e.Evaluation, e.Comment, r.IsReuploaded -- Select IsReuploaded
                FROM Responses r
                LEFT JOIN ResponseMetadata rm ON r.ResponseID = rm.ResponseID
                LEFT JOIN EvaluationData e ON r.ResponseID = e.ResponseID
                WHERE r.TemplateID = ? AND r.SheetName = ?
                ORDER BY r.IsReuploaded ASC, r.ResponseID ASC, e.DataID -- PRIMARY CHANGE HERE
            """;
            try (PreparedStatement pstmt = conn.prepareStatement(dataSql)) {
                pstmt.setInt(1, templateId);
                pstmt.setString(2, selectedSheet);
                try (ResultSet rs = pstmt.executeQuery()) {
                    while (rs.next()) {
                        String responseId = String.valueOf(rs.getInt("ResponseID"));
                        String originalFileName = rs.getString("OriginalFileName");
                        // Use IsReuploaded to augment the unique key if desired for debug, not strictly needed for display
                        String uniqueKey = originalFileName + " (ResponseID: " + responseId + ")";

                        responsesMetadata.computeIfAbsent(uniqueKey, k -> new HashMap<>());
                        mainItemsMap.computeIfAbsent(uniqueKey, k -> new ArrayList<>());
                        subItemsMap.computeIfAbsent(uniqueKey, k -> new LinkedHashMap<>());
                        evalDataMap.computeIfAbsent(uniqueKey, k -> new LinkedHashMap<>());

                        String headerKey = rs.getString("HeaderKey");
                        String headerValue = rs.getString("HeaderValue");
                        if (headerKey != null && headerValue != null && !headerValue.trim().isEmpty()) {
                            responsesMetadata.get(uniqueKey).put(headerKey, headerValue);
                        }

                        String mainItem = rs.getString("MainItem");
                        String subItem = rs.getString("SubItem") != null ? rs.getString("SubItem") : "";
                        String eval = rs.getString("Evaluation") != null ? rs.getString("Evaluation") : "";
                        String comment = rs.getString("Comment") != null ? rs.getString("Comment") : "";

                        if (mainItem != null) {
                            List<String> mainItems = mainItemsMap.get(uniqueKey);
                            if (!mainItems.contains(mainItem)) {
                                mainItems.add(mainItem);
                            }

                            Map<String, List<String>> currentSubItems = subItemsMap.get(uniqueKey);
                            List<String> subItemsList = currentSubItems.computeIfAbsent(mainItem, k -> new ArrayList<>());
                            if (!subItemsList.contains(subItem)) {
                                subItemsList.add(subItem);
                            }

                            evalDataMap.get(uniqueKey)
                                             .computeIfAbsent(mainItem, k -> new LinkedHashMap<>())
                                             .put(subItem, Map.of("Evaluation", eval, "Comment", comment));
                        }
                    }
                }
            }

            JSONArray headers = new JSONArray();
            JSONArray row0 = new JSONArray();
            JSONArray row1 = new JSONArray();
            JSONArray row2 = new JSONArray();

            for (String headerKey : sortedMetadataHeaders) {
                row0.put(new JSONObject().put("label", headerKey).put("rowspan", 3).put("colspan", 1));
            }

            Set<String> allMainItemsInSheet = new LinkedHashSet<>();
            Map<String, Set<String>> allSubItemsInSheetByMainItem = new LinkedHashMap<>();

            for(Map.Entry<String, Map<String, Map<String, Map<String, String>>>> entryByUniqueKey : evalDataMap.entrySet()) {
                Map<String, Map<String, Map<String, String>>> evalDataByMainItem = entryByUniqueKey.getValue();
                for(Map.Entry<String, Map<String, Map<String, String>>> entryByMainItem : evalDataByMainItem.entrySet()) {
                    String mainItem = entryByMainItem.getKey();

                    allMainItemsInSheet.add(mainItem);

                    Map<String, Map<String, String>> evalDataBySubItem = entryByMainItem.getValue();
                    for (String subItem : evalDataBySubItem.keySet()) {
                        Map<String, String> result = evalDataBySubItem.get(subItem);
                        if (result != null && (!result.getOrDefault("Evaluation", "").isEmpty() || !result.getOrDefault("Comment", "").isEmpty())) {
                             allSubItemsInSheetByMainItem.computeIfAbsent(mainItem, k -> new LinkedHashSet<>()).add(subItem);
                        }
                    }
                    if (mainItem.contains("ご要望等")) {
                        Map<String, Map<String, String>> mainEvalForRequests = evalDataByMainItem.getOrDefault(mainItem, new LinkedHashMap<String, Map<String, String>>());
                        Map<String, String> result = mainEvalForRequests.getOrDefault("", Collections.emptyMap());
                        if (result != null && !result.getOrDefault("Comment", "").isEmpty()) {
                            allMainItemsInSheet.add(mainItem);
                            allSubItemsInSheetByMainItem.computeIfAbsent(mainItem, k -> new LinkedHashSet<>()).add("");
                        }
                    }
                }
            }


            List<String> sortedAllMainItemsInSheet = new ArrayList<>(allMainItemsInSheet);
            sortedAllMainItemsInSheet.sort(new Comparator<String>() {
                @Override
                public int compare(String item1, String item2) {
                    boolean isPriority1 = item1.contains("より満足いただくために");
                    boolean isRequests1 = item1.contains("ご要望等");
                    boolean isPriority2 = item2.contains("より満足いただくために");
                    boolean isRequests2 = item2.contains("ご要望等");

                    if ((isPriority1 || isRequests1) && !(isPriority2 || isRequests2)) {
                        return 1;
                    }
                    if (!(isPriority1 || isRequests1) && (isPriority2 || isRequests2)) {
                        return -1;
                    }
                    Integer num1 = extractLeadingNumber(item1);
                    Integer num2 = extractLeadingNumber(item2);

                    if (num1 != null && num2 != null) {
                        int numCompare = num1.compareTo(num2);
                        if (numCompare != 0) {
                            return numCompare;
                        }
                    }
                    return item1.compareTo(item2);
                }
            });


            for (String mainItem : sortedAllMainItemsInSheet) {
                List<String> allSubItems = new ArrayList<>(allSubItemsInSheetByMainItem.getOrDefault(mainItem, Collections.emptySet()));

                List<String> validSubItems = allSubItems.stream()
                                                       .filter(s -> s != null && !s.trim().isEmpty())
                                                       .collect(Collectors.toList());

                boolean isPrioritySection = mainItem.contains("より満足いただくために");
                boolean isRequestsSection = mainItem.contains("ご要望等");

                String displayMainItem = mainItem;
                String leadingNumberStr = extractLeadingNumber(mainItem) != null ? extractLeadingNumber(mainItem) + "." : "";
                if (mainItem.contains("より満足いただくために、弊社が真っ先に解決/取組むべき項目/事柄はどのようなものだとお考えですか。")) {
                    displayMainItem = leadingNumberStr + "より満足いただくため";
                } else if (mainItem.contains("ご要望等がございましたらご記入ください。")) {
                    displayMainItem = leadingNumberStr + "ご要望";
                }

                if (validSubItems.isEmpty() || isRequestsSection) {
                    row0.put(new JSONObject().put("label", displayMainItem).put("rowspan", 3).put("colspan", 1));
                } else {
                    int colspan = 2 * validSubItems.size();
                    row0.put(new JSONObject().put("label", displayMainItem).put("rowspan", 1).put("colspan", colspan));

                    for (String subItem : validSubItems) {
                        String subItemHeader = isPrioritySection ? "<" + (validSubItems.indexOf(subItem) + 1) + ">" : subItem;
                        row1.put(new JSONObject().put("label", subItemHeader).put("colspan", 2));
                    }
                    for (String subItem : validSubItems) {
                        row2.put(new JSONObject().put("label", "Evaluation").put("colspan", 1));
                        row2.put(new JSONObject().put("label", "Comment").put("colspan", 1));
                    }
                }
            }

            headers.put(row0);
            headers.put(row1);
            headers.put(row2);

            JSONArray dataRows = new JSONArray();
            for (Map.Entry<String, Map<String, String>> responseEntry : responsesMetadata.entrySet()) {
                String uniqueKey = responseEntry.getKey();
                JSONArray row = new JSONArray();

                Map<String, String> metadata = responseEntry.getValue();
                for (String headerKey : sortedMetadataHeaders) {
                    row.put(metadata.getOrDefault(headerKey, ""));
                }

                Map<String, Map<String, Map<String, String>>> currentResponseEvalDataMap = evalDataMap.getOrDefault(uniqueKey, Collections.emptyMap());

                for (String mainItem : sortedAllMainItemsInSheet) {
                    List<String> allSubItems = new ArrayList<>(allSubItemsInSheetByMainItem.getOrDefault(mainItem, Collections.emptySet()));
                    List<String> validSubItems = allSubItems.stream()
                                                           .filter(s -> s != null && !s.trim().isEmpty())
                                                           .collect(Collectors.toList());

                    boolean isPrioritySection = mainItem.contains("より満足いただくために");
                    if (isPrioritySection) {
                        validSubItems.sort(Comparator.comparingInt(s -> {
                            Matcher matcher = Pattern.compile("^[＜<]?(\\d+)[＞>]?[\\.．①②③④⑤⑥⑦].*").matcher(s);
                            return matcher.find() ? Integer.parseInt(matcher.group(1)) : Integer.MAX_VALUE;
                        }));
                    } else {
                        Collections.sort(validSubItems);
                    }

                    boolean isRequestsSection = mainItem.contains("ご要望等");

                    if (validSubItems.isEmpty() || isRequestsSection) {
                        Map<String, Map<String, String>> mainEval = currentResponseEvalDataMap.getOrDefault(mainItem, Collections.emptyMap());
                        Map<String, String> result = mainEval.getOrDefault("", Collections.emptyMap());
                        row.put(result.getOrDefault("Comment", ""));
                    } else {
                        for (String subItem : validSubItems) {
                            Map<String, Map<String, String>> mainEval = currentResponseEvalDataMap.getOrDefault(mainItem, Collections.emptyMap());
                            Map<String, String> result = mainEval.getOrDefault(subItem, Collections.emptyMap());
                            row.put(result.getOrDefault("Evaluation", ""));
                            row.put(result.getOrDefault("Comment", ""));
                        }
                    }
                }
                dataRows.put(row);
            }

            JSONObject sheetData = new JSONObject();
            sheetData.put("SheetName", selectedSheet);
            sheetData.put("Headers", headers);
            sheetData.put("Rows", dataRows);

            jsonResponse.put("status", "success");
            jsonResponse.put("sheetNames", sheetNames);
            jsonResponse.put("data", sheetData);

            logger.info("Sending response for category {}: {}", templateCategory, jsonResponse.toString());
        } catch (SQLException e) {
            logger.error("Database error for templateCategory: {}", templateCategory, e);
            jsonResponse.put("status", "error");
            jsonResponse.put("message", "Database error: " + e.getMessage());
            response.setStatus(HttpServletResponse.SC_INTERNAL_SERVER_ERROR);
            response.getWriter().append(jsonResponse.toString());
        } catch (Exception e) {
            logger.error("Unexpected error for templateCategory: {}", templateCategory, e);
            jsonResponse.put("status", "error");
            jsonResponse.put("message", "Unexpected error: " + e.getMessage());
            response.setStatus(HttpServletResponse.SC_INTERNAL_SERVER_ERROR);
            response.getWriter().append(jsonResponse.toString());
        }

        response.getWriter().append(jsonResponse.toString());
    }

    private boolean containsSheet(JSONArray sheetNames, String sheetName) {
        String trimmedSheetName = sheetName != null ? sheetName.trim() : "";
        for (Object name : sheetNames) {
            String trimmedName = name != null ? name.toString().trim() : "";
            if (trimmedName.equals(trimmedSheetName)) {
                return true;
            }
        }
        return false;
    }
}