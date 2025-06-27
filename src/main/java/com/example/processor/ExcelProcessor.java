package com.example.processor;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.*;
import java.sql.*;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.Date;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Collectors;

public class ExcelProcessor {

    private static final Pattern MAIN_CRITERION_PATTERN = Pattern.compile("^(?:[0-9０-９]+\\.|(?:I{1,3}|IV|V|VI|VII|VIII)\\.)[^①②③④⑤⑥⑦].*");
    private static final Pattern PRIORITY_PATTERN = Pattern.compile("^(?:[１２３４１-４]|[1-4])\\.[①②③④⑤⑥⑦]|(?:<|＜)(?:[1-4]|[１-４])(?:>|＞)");
    private static final Pattern PRIORITY_NUMBER_EXTRACTOR = Pattern.compile("^[＜<]?(\\d+)[＞>]?[\\.．①②③④⑤⑥⑦].*");
    private static final Pattern EVALUATION_PATTERN = Pattern.compile("^\\d+:(?:Not Related|[A-Za-z]+)$");
    private static final Pattern NOTE_PATTERN = Pattern.compile("^(※.*|.*(?:お願い|注意|注|備考|説明|選択ボックス).*)", Pattern.UNICODE_CASE);
    private static final String EVAL_TABLE_HEADER = "評価項目";

    public static class TemplateCreationResult {
        public int templateId;
        public String templateCategoryDisplayName;
        public String internalTemplateCategory;

        public TemplateCreationResult(int templateId, String templateCategoryDisplayName, String internalTemplateCategory) {
            this.templateId = templateId;
            this.templateCategoryDisplayName = templateCategoryDisplayName;
            this.internalTemplateCategory = internalTemplateCategory;
        }
    }

    public static TemplateCreationResult processExcelFile(Connection conn, InputStream inputStream, String originalFileName) throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook(inputStream)) {
            String internalTemplateCategory = generateInternalTemplateCategory(workbook);

            TemplateCreationResult creationResult = getOrCreateTemplate(conn, originalFileName, internalTemplateCategory);
            int templateId = creationResult.templateId;

            for (Sheet sheet : workbook) {
                String sheetName = sheet.getSheetName();
                System.out.println("Processing sheet: " + sheetName);

                if (sheetName.startsWith("評価結果リスト_") || sheetName.equals("表紙") || sheetName.isEmpty()) {
                    System.out.println("Skipping summary/cover/empty sheet: " + sheetName);
                    continue;
                }

                int responseIdToUse;
                boolean isReupload = false; // Flag to track if this is a re-upload of an existing ResponseID
                int existingResponseId = getExistingResponseIdForFileNameAndSheet(conn, templateId, originalFileName, sheetName);

                conn.setAutoCommit(false); // Start transaction for atomicity

                try {
                    if (existingResponseId != -1) {
                        System.out.println("Existing response found for OriginalFileName '" + originalFileName + "' and SheetName '" + sheetName + "'. ResponseID: " + existingResponseId + ". Deleting old data for update.");
                        deleteResponseData(conn, existingResponseId);
                        responseIdToUse = existingResponseId;
                        isReupload = true; // Mark as re-upload
                    } else {
                        System.out.println("No existing response found for OriginalFileName '" + originalFileName + "' and SheetName '" + sheetName + "'. Inserting new Response record.");
                        responseIdToUse = insertNewResponse(conn, templateId, sheetName, originalFileName);
                        // isReupload remains false for new inserts
                    }

                    List<Map<String, String>> metadataList = extractMetadata(sheet);
                    List<Map<String, String>> evalData = extractEvaluationData(sheet);

                    if (!metadataList.isEmpty()) {
                        storeResponseMetadata(conn, responseIdToUse, metadataList.get(0));
                    } else {
                        System.out.println("No metadata found for sheet: " + sheetName + ". Skipping metadata insert for ResponseID " + responseIdToUse + ".");
                    }

                    if (!evalData.isEmpty()) {
                        storeEvaluationData(conn, responseIdToUse, evalData);
                        System.out.println("Stored evaluation data for sheet: " + sheetName + " with ResponseID=" + responseIdToUse);
                    } else {
                        System.out.println("No evaluation data to store for sheet: " + sheetName + " for ResponseID " + responseIdToUse + ".");
                    }

                    // *** MODIFIED: Update LastUpdated timestamp and IsReuploaded flag for the Response ***
                    String updateResponseSql = "UPDATE Responses SET LastUpdated = NOW(), IsReuploaded = ? WHERE ResponseID = ?";
                    try (PreparedStatement pstmt = conn.prepareStatement(updateResponseSql)) {
                        pstmt.setBoolean(1, isReupload);
                        pstmt.setInt(2, responseIdToUse);
                        pstmt.executeUpdate();
                        System.out.println("Updated LastUpdated and IsReuploaded (" + isReupload + ") for ResponseID: " + responseIdToUse);
                    }

                    conn.commit(); // Commit transaction
                    System.out.println("Transaction committed for ResponseID: " + responseIdToUse);

                } catch (SQLException e) {
                    conn.rollback(); // Rollback on error
                    System.err.println("Transaction rolled back for sheet " + sheetName + ". Error: " + e.getMessage());
                    throw e; // Re-throw to be caught by outer try-catch
                } finally {
                    conn.setAutoCommit(true); // Restore auto-commit
                }
            }
            return creationResult;
        }
    }

    private static String generateInternalTemplateCategory(XSSFWorkbook workbook) {
        List<String> relevantSheetNames = new ArrayList<>();
        for (Sheet sheet : workbook) {
            String sheetName = sheet.getSheetName().trim();
            if (!sheetName.startsWith("評価結果リスト_") && !sheetName.equals("表紙") && !sheetName.isEmpty()) {
                relevantSheetNames.add(sheetName);
            }
        }
        Collections.sort(relevantSheetNames);
        if (relevantSheetNames.isEmpty()) {
            return "EmptyTemplate";
        }
        return String.join("_", relevantSheetNames);
    }

    private static TemplateCreationResult getOrCreateTemplate(Connection conn, String originalFileName, String internalTemplateCategory) throws SQLException {
        String selectSql = "SELECT TemplateID, TemplateCategory FROM Templates WHERE InternalTemplateCategory = ?";
        try (PreparedStatement pstmt = conn.prepareStatement(selectSql)) {
            pstmt.setString(1, internalTemplateCategory);
            try (ResultSet rs = pstmt.executeQuery()) {
                if (rs.next()) {
                    System.out.println("Existing template structure found. ID: " + rs.getInt("TemplateID") + ", Display Name: " + rs.getString("TemplateCategory"));
                    return new TemplateCreationResult(rs.getInt("TemplateID"), rs.getString("TemplateCategory"), internalTemplateCategory);
                }
            }
        }

        String newDisplayName = generateUniqueTemplateDisplayName(conn);
        String insertSql = "INSERT INTO Templates (TemplateName, TemplateCategory, InternalTemplateCategory, UploadDate) VALUES (?, ?, ?, NOW())";
        int newTemplateId;
        try (PreparedStatement pstmt = conn.prepareStatement(insertSql, Statement.RETURN_GENERATED_KEYS)) {
            pstmt.setString(1, originalFileName);
            pstmt.setString(2, newDisplayName);
            pstmt.setString(3, internalTemplateCategory);
            pstmt.executeUpdate();
            try (ResultSet rs = pstmt.getGeneratedKeys()) {
                if (rs.next()) {
                    newTemplateId = rs.getInt(1);
                    System.out.println("New template created. ID: " + newTemplateId + ", Display Name: " + newDisplayName + ", Internal Category: " + internalTemplateCategory);
                } else {
                    throw new SQLException("Failed to retrieve generated TemplateID.");
                }
            }
        }
        return new TemplateCreationResult(newTemplateId, newDisplayName, internalTemplateCategory);
    }

    private static String generateUniqueTemplateDisplayName(Connection conn) throws SQLException {
        int maxNum = 0;
        String getMaxNumSql = "SELECT MAX(CAST(SUBSTRING_INDEX(TemplateCategory, 'Template Type ', -1) AS UNSIGNED)) AS MaxNum FROM Templates WHERE TemplateCategory LIKE 'Template Type %'";
        try (PreparedStatement pstmt = conn.prepareStatement(getMaxNumSql);
             ResultSet rs = pstmt.executeQuery()) {
            if (rs.next()) {
                maxNum = rs.getInt("MaxNum");
            }
        }
        return "Template Type " + (maxNum + 1);
    }

    private static int getExistingResponseIdForFileNameAndSheet(Connection conn, int templateId, String originalFileName, String sheetName) throws SQLException {
        String sql = "SELECT ResponseID FROM Responses WHERE TemplateID = ? AND OriginalFileName = ? AND SheetName = ?";
        try (PreparedStatement pstmt = conn.prepareStatement(sql)) {
            pstmt.setInt(1, templateId);
            pstmt.setString(2, originalFileName);
            pstmt.setString(3, sheetName);
            try (ResultSet rs = pstmt.executeQuery()) {
                if (rs.next()) {
                    return rs.getInt("ResponseID");
                }
            }
        }
        return -1;
    }

    private static int insertNewResponse(Connection conn, int templateId, String sheetName, String originalFileName) throws SQLException {
        // *** MODIFIED: Set IsReuploaded to 0 (false) for new inserts ***
        String insertResponseSql = "INSERT INTO Responses (TemplateID, SheetName, OriginalFileName, LastUpdated, IsReuploaded) VALUES (?, ?, ?, NOW(), 0)";
        int responseId;
        try (PreparedStatement pstmt = conn.prepareStatement(insertResponseSql, Statement.RETURN_GENERATED_KEYS)) {
            pstmt.setInt(1, templateId);
            pstmt.setString(2, sheetName);
            pstmt.setString(3, originalFileName);
            pstmt.executeUpdate();
            try (ResultSet rs = pstmt.getGeneratedKeys()) {
                if (rs.next()) {
                    responseId = rs.getInt(1);
                } else {
                    throw new SQLException("Failed to retrieve generated ResponseID.");
                }
            }
        }
        return responseId;
    }

    private static void deleteResponseData(Connection conn, int responseId) throws SQLException {
        String deleteEvalSql = "DELETE FROM EvaluationData WHERE ResponseID = ?";
        try (PreparedStatement pstmt = conn.prepareStatement(deleteEvalSql)) {
            pstmt.setInt(1, responseId);
            pstmt.executeUpdate();
        }
        String deleteMetaSql = "DELETE FROM ResponseMetadata WHERE ResponseID = ?";
        try (PreparedStatement pstmt = conn.prepareStatement(deleteMetaSql)) {
            pstmt.setInt(1, responseId);
            pstmt.executeUpdate();
        }
        System.out.println("Successfully deleted existing metadata and evaluation data for ResponseID: " + responseId);
    }

    private static void storeResponseMetadata(Connection conn, int responseId, Map<String, String> metadata) throws SQLException {
        String insertMetadataSql = "INSERT INTO ResponseMetadata (ResponseID, HeaderKey, HeaderValue) VALUES (?, ?, ?)";
        try (PreparedStatement pstmt = conn.prepareStatement(insertMetadataSql)) {
            for (Map.Entry<String, String> entry : metadata.entrySet()) {
                pstmt.setInt(1, responseId);
                pstmt.setString(2, entry.getKey());
                pstmt.setString(3, entry.getValue());
                pstmt.addBatch();
            }
            pstmt.executeBatch();
            System.out.println("Stored new metadata for ResponseID=" + responseId + ": " + metadata);
        }
    }

    private static void storeEvaluationData(Connection conn, int responseId, List<Map<String, String>> evalData) throws SQLException {
        String insertSql = "INSERT INTO EvaluationData (ResponseID, MainItem, SubItem, Evaluation, Comment) VALUES (?, ?, ?, ?, ?)";

        try (PreparedStatement pstmtInsert = conn.prepareStatement(insertSql)) {
            for (Map<String, String> entry : evalData) {
                String mainItem = entry.get("MainItem");
                String subItem = entry.get("SubItem");
                String evaluation = entry.get("Evaluation");
                String comment = entry.get("Comment");

                pstmtInsert.setInt(1, responseId);
                pstmtInsert.setString(2, mainItem);
                pstmtInsert.setString(3, subItem);
                pstmtInsert.setString(4, evaluation);
                pstmtInsert.setString(5, comment);
                pstmtInsert.addBatch();
            }
            pstmtInsert.executeBatch();
            System.out.println("Inserted evaluation data batch for ResponseID=" + responseId + ". Total items: " + evalData.size());
        }
    }

    private static List<Map<String, String>> extractMetadata(Sheet sheet) {
        List<Map<String, String>> metadataList = new ArrayList<>();
        int evalStartRow = findEvaluationStartRow(sheet);
        int maxRowsToSearch = (evalStartRow == -1) ? sheet.getLastRowNum() + 1 : evalStartRow - 1;

        System.out.println("Scanning for metadata in sheet: " + sheet.getSheetName() + " up to row: " + maxRowsToSearch);
        Map<String, String> metadata = new LinkedHashMap<>();
        readMetadataTable(sheet, 0, metadata, maxRowsToSearch);

        if (!metadata.isEmpty()) {
            metadataList.add(metadata);
            System.out.println("Metadata found: " + metadata);
        } else {
            System.out.println("No metadata found in sheet: " + sheet.getSheetName());
        }

        return metadataList;
    }

    private static void readMetadataTable(Sheet sheet, int startRow, Map<String, String> metadata, int maxRows) {
        for (int i = startRow; i < maxRows; i++) {
            Row row = sheet.getRow(i);
            if (row == null) continue;

            boolean isHeaderRow = false;
            for (Cell cell : row) {
                if ("評価項目".equals(getCellValue(cell))) {
                    isHeaderRow = true;
                    break;
                }
            }
            if (isHeaderRow) {
                System.out.println("Skipping header row at: " + i);
                continue;
            }

            int maxCols = row.getLastCellNum();
            int j = 0;
            while (j < maxCols) {
                while (j < maxCols && getCellValue(row.getCell(j)).isEmpty()) j++;
                if (j >= maxCols) break;
                String header = getCellValue(row.getCell(j));
                if (isNote(header)) {
                    j++;
                    continue;
                }
                int m = j + 1;
                while (m < maxCols && getCellValue(row.getCell(m)).isEmpty()) m++;
                String value = (m < maxCols) ? getCellValue(row.getCell(m)) : "";
                if (!value.isEmpty() && !isNote(value)) {
                    metadata.put(header, value);
                    System.out.println("Stored: " + header + " = " + value + " at row " + i + ", header col " + j + ", value col " + m);
                    j = m + 1;
                } else {
                    j++;
                }
            }
        }
    }

    private static boolean isNote(String value) {
        if (value == null || value.isEmpty()) return false;
        return NOTE_PATTERN.matcher(value).matches();
    }

    private static int[] findEvaluationAndCommentColumns(Sheet sheet, int evalStartRow) {
        Row headerRow = sheet.getRow(evalStartRow - 1);
        if (headerRow == null) {
            System.out.println("Header row " + (evalStartRow - 1) + " is null in sheet: " + sheet.getSheetName());
            return new int[]{-1, -1};
        }

        int evalColumn = -1;
        int commentColumn = -1;
        int maxColumnsToCheck = headerRow.getLastCellNum() > 0 ? headerRow.getLastCellNum() : 52;

        for (int j = 0; j < maxColumnsToCheck; j++) {
            Cell cell = headerRow.getCell(j, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
            String value = getCellValue(cell).trim();
            if ("評価".equals(value) && evalColumn == -1) {
                evalColumn = j;
            } else if ("コメント".equals(value) && commentColumn == -1) {
                commentColumn = j;
            }
        }

        if (evalColumn == -1) {
            evalColumn = findDataPatternColumn(sheet, evalStartRow, "evaluation");
        }
        if (commentColumn == -1 || commentColumn == evalColumn) {
            commentColumn = findDataPatternColumn(sheet, evalStartRow, "comment");
        }

        if (commentColumn == -1 || commentColumn == evalColumn) {
            for (int j = evalColumn + 1; j < maxColumnsToCheck; j++) {
                if (hasCommentData(sheet, evalStartRow, j)) {
                    commentColumn = j;
                    break;
                }
            }
        }

        if (evalColumn == -1) {
            System.out.println("Could not dynamically detect eval column in sheet: " + sheet.getSheetName());
            return new int[]{-1, -1};
        }
        if (commentColumn == -1) {
            commentColumn = evalColumn + 1;
            System.out.println("Comment column not found, defaulting to evalColumn + 1: " + commentColumn);
        }

        System.out.println("Detected columns in sheet " + sheet.getSheetName() + ": evalColumn=" + evalColumn + ", commentColumn=" + commentColumn);
        return new int[]{evalColumn, commentColumn};
    }

    private static boolean hasCommentData(Sheet sheet, int startRow, int column) {
        for (int i = startRow; i < startRow + 10 && i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            if (row != null) {
                String value = getCellValue(row.getCell(column, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK));
                if (!value.isEmpty() && !EVALUATION_PATTERN.matcher(value).matches() && !value.matches("^[＜＞]")) {
                    System.out.println("Comment data found at col " + column + ", row " + i + ": " + value);
                    return true;
                }
            }
        }
        return false;
    }

    private static int findDataPatternColumn(Sheet sheet, int startRow, String type) {
        for (int i = startRow; i < startRow + 10 && i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            if (row != null) {
                for (int j = 0; j < row.getLastCellNum(); j++) {
                    String value = getCellValue(row.getCell(j, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK));
                    if ("evaluation".equals(type) && EVALUATION_PATTERN.matcher(value).matches()) {
                        System.out.println(type + " data found at col " + j + ", row " + i + ": " + value);
                        return j;
                    } else if ("comment".equals(type) && !value.isEmpty() && !EVALUATION_PATTERN.matcher(value).matches() && !value.matches("^[＜＞]")) {
                        System.out.println(type + " data found at col " + j + ", row " + i + ": " + value);
                        return j;
                    }
                }
            }
        }
        return -1;
    }

    private static List<Map<String, String>> extractEvaluationData(Sheet sheet) {
        List<Map<String, String>> data = new ArrayList<>();
        int evalStartRow = findEvaluationStartRow(sheet);
        if (evalStartRow == -1) {
            System.out.println("Evaluation data not found in sheet: " + sheet.getSheetName());
            return data;
        }

        int[] columns = findEvaluationAndCommentColumns(sheet, evalStartRow);
        int evalColumn = columns[0];
        int commentColumn = columns[1];
        if (evalColumn == -1) {
            return data;
        }

        String currentMainItem = null;
        Map<String, Map<String, String>> mainSubItemMap = new HashMap<>();

        for (int i = evalStartRow; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            if (row == null) {
                System.out.println("Skipping null row " + i);
                continue;
            }

            String mainItem = getMainItemText(row);
            if (!mainItem.isEmpty()) {
                currentMainItem = mainItem;
                System.out.println("Main item detected at row " + i + ": " + currentMainItem);
                for (int j = i + 1; j <= sheet.getLastRowNum(); j++) {
                    Row subRow = sheet.getRow(j);
                    if (subRow == null) break;
                    String subNumber = "";
                    String subItemText = "";
                    for (int k = 0; k < subRow.getLastCellNum(); k++) {
                        String cellValue = getCellValue(subRow.getCell(k));
                        if (cellValue.matches("^[①②③④⑤⑥⑦１-７1-7]") && subNumber.isEmpty()) {
                            subNumber = cellValue;
                        } else if (!cellValue.isEmpty() && !cellValue.matches("^[＜＞]") && !EVALUATION_PATTERN.matcher(cellValue).matches()) {
                            subItemText = cellValue;
                            break;
                        }
                    }
                    if (!subNumber.isEmpty() && !subItemText.isEmpty()) {
                        mainSubItemMap.computeIfAbsent(currentMainItem, k -> new HashMap<>()).put(subNumber, subItemText);
                    }
                    if (!getMainItemText(subRow).isEmpty()) break;
                }
                continue;
            }

            if (currentMainItem != null) {
                if (currentMainItem.contains("より満足いただくために")) {
                    System.out.println("Priority items section started at row " + i + ": " + currentMainItem);
                    int nextMainItemRow = findNextMainItemRow(sheet, i);

                    int priorityHeadersRowIdx = -1;
                    for (int rowNum = i + 1; rowNum < nextMainItemRow && rowNum <= sheet.getLastRowNum(); rowNum++) {
                        Row potentialHeaderRow = sheet.getRow(rowNum);
                        if (potentialHeaderRow == null) continue;

                        boolean foundAnyPriorityHeader = false;
                        for (int col = 0; col < potentialHeaderRow.getLastCellNum(); col++) {
                            String cellValue = getCellValue(potentialHeaderRow.getCell(col));
                            if (PRIORITY_PATTERN.matcher(cellValue).matches()) {
                                foundAnyPriorityHeader = true;
                                break;
                            }
                        }
                        if (foundAnyPriorityHeader) {
                            priorityHeadersRowIdx = rowNum;
                            break;
                        }
                    }

                    if (priorityHeadersRowIdx == -1) {
                        System.out.println("Priority headers row not found for " + currentMainItem + ". Skipping section.");
                        i = nextMainItemRow - 1;
                        continue;
                    }

                    Row priorityHeadersRow = sheet.getRow(priorityHeadersRowIdx);
                    Row dataRowForPriorities = sheet.getRow(priorityHeadersRowIdx + 1);

                    if (dataRowForPriorities == null) {
                        System.out.println("No data row found for priority items below headers. Skipping section.");
                        i = nextMainItemRow - 1;
                        continue;
                    }

                    Map<String, Integer> priorityHeaderCols = new LinkedHashMap<>();
                    for (int col = 0; col < priorityHeadersRow.getLastCellNum(); col++) {
                        String headerText = getCellValue(priorityHeadersRow.getCell(col));
                        if (PRIORITY_PATTERN.matcher(headerText).matches()) {
                            priorityHeaderCols.put(headerText, col);
                        }
                    }

                    if (priorityHeaderCols.isEmpty()) {
                        System.out.println("No priority sub-headers (<1>, <2> etc.) found for " + currentMainItem + ". Skipping section.");
                        i = nextMainItemRow - 1;
                        continue;
                    }

                    List<String> orderedPriorityHeaders = new ArrayList<>(priorityHeaderCols.keySet());

                    for (int headerBlockIdx = 0; headerBlockIdx < orderedPriorityHeaders.size(); headerBlockIdx++) {
                        String priorityHeader = orderedPriorityHeaders.get(headerBlockIdx);
                        int currentBlockStartCol = priorityHeaderCols.get(priorityHeader);

                        String evaluationToStore = "";
                        String description = "";

                        String potentialEvaluation = getCellValue(dataRowForPriorities.getCell(currentBlockStartCol, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK));
                        if (!potentialEvaluation.isEmpty()) {
                            evaluationToStore = potentialEvaluation;
                        }

                        int nextBlockStartCol = dataRowForPriorities.getLastCellNum() + 1;
                        if (headerBlockIdx + 1 < orderedPriorityHeaders.size()) {
                            nextBlockStartCol = priorityHeaderCols.get(orderedPriorityHeaders.get(headerBlockIdx + 1));
                        }

                        StringBuilder descriptionBuilder = new StringBuilder();
                        int descStartCol = currentBlockStartCol + 1;

                        for (int m = descStartCol; m < dataRowForPriorities.getLastCellNum(); m++) {
                            if (m >= nextBlockStartCol) {
                                break;
                            }
                            String cellValue = getCellValue(dataRowForPriorities.getCell(m, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK));
                            if (!cellValue.isEmpty() && !isNote(cellValue) && !EVALUATION_PATTERN.matcher(cellValue).matches()) {
                                if (descriptionBuilder.length() > 0) {
                                    descriptionBuilder.append(" ");
                                }
                                descriptionBuilder.append(cellValue);
                            }
                        }
                        description = descriptionBuilder.toString().trim();

                        if (!evaluationToStore.isEmpty() || !description.isEmpty()) {
                            Map<String, String> entry = new HashMap<>();

                            entry.put("MainItem", currentMainItem);
                            entry.put("SubItem", description);
                            entry.put("Evaluation", evaluationToStore);
                            entry.put("Comment", description);
                            data.add(entry);
                            System.out.println("Priority item extracted: Header='" + priorityHeader + "', Eval='" + evaluationToStore + "', Desc='" + description + "'");
                        } else {
                            System.out.println("Skipping empty priority item block for header: " + priorityHeader);
                        }
                    }
                    i = nextMainItemRow - 1;
                    continue;
                }

                if (currentMainItem.contains("ご要望等")) {
                    System.out.println("Requests section started at row " + i + ": " + currentMainItem);
                    int nextMainItemRow = findNextMainItemRow(sheet, i);
                    for (int j = i; j <= sheet.getLastRowNum() && j < nextMainItemRow; j++) {
                        Row requestRow = sheet.getRow(j);
                        if (requestRow == null) {
                            System.out.println("Skipping null request row " + j);
                            continue;
                        }
                        StringBuilder requestTextBuilder = new StringBuilder();
                        boolean hasContent = false;
                        System.out.println("Inspecting request row " + j + " for content:");
                        for (int k = 0; k < requestRow.getLastCellNum(); k++) {
                            String cellValue = getCellValue(requestRow.getCell(k));
                            System.out.println("  Cell at col " + k + ": '" + cellValue + "'");
                            if (!cellValue.isEmpty() && !EVALUATION_PATTERN.matcher(cellValue).matches()) {
                                if (requestTextBuilder.length() > 0) {
                                    requestTextBuilder.append(" | ");
                                }
                                requestTextBuilder.append(cellValue);
                                hasContent = true;
                            }
                        }
                        if (hasContent) {
                            Map<String, String> entry = new HashMap<>();
                            entry.put("MainItem", currentMainItem);
                            entry.put("SubItem", "");
                            entry.put("Evaluation", "");
                            entry.put("Comment", requestTextBuilder.toString());
                            data.add(entry);
                            System.out.println("Request at row " + j + ": " + requestTextBuilder.toString());
                        } else {
                            System.out.println("No valid content found in request row " + j);
                        }
                    }
                    i = nextMainItemRow - 1;
                    continue;
                }

                String subItem = "";
                String evaluation = "";
                String comment = "";
                for (int j = 0; j < row.getLastCellNum(); j++) {
                    String value = getCellValue(row.getCell(j));
                    if (value.matches("^[①②③④⑤⑥⑦１-７1-7]") && subItem.isEmpty()) {
                        subItem = value;
                    } else if (j == evalColumn && EVALUATION_PATTERN.matcher(value).matches()) {
                        evaluation = value;
                        System.out.println("Evaluation found at row " + i + ", col " + j + ": " + evaluation);
                    } else if (j == commentColumn && !EVALUATION_PATTERN.matcher(value).matches() && !value.isEmpty()) {
                        comment = value;
                    } else if (j > evalColumn && j != commentColumn && !EVALUATION_PATTERN.matcher(value).matches() && !value.isEmpty()) {
                        comment = comment.isEmpty() ? value : comment + " " + value;
                    }
                }

                if (!subItem.isEmpty() || !evaluation.isEmpty() || !comment.isEmpty()) {
                    String effectiveSubItemText = "";
                    for (int j = 0; j < row.getLastCellNum(); j++) {
                        String value = getCellValue(row.getCell(j));
                        if (!value.isEmpty() && !value.matches("^[＜＞]") && !EVALUATION_PATTERN.matcher(value).matches() && !value.matches("^[①②③④⑤⑥⑦１-７1-7]")) {
                            effectiveSubItemText = value;
                            break;
                        }
                    }
                    if (!effectiveSubItemText.isEmpty() || !subItem.isEmpty()) {
                        Map<String, String> entry = new HashMap<>();
                        entry.put("MainItem", currentMainItem);
                        entry.put("SubItem", subItem.isEmpty() ? effectiveSubItemText : subItem + " " + effectiveSubItemText);
                        entry.put("Evaluation", evaluation);
                        entry.put("Comment", comment.trim());
                        data.add(entry);
                        System.out.println("Evaluation item at row " + i + ": Main=" + currentMainItem +
                                ", Sub=" + entry.get("SubItem") + ", Eval=" + evaluation +
                                ", Comment=" + comment.trim());
                    }
                }
            }
        }
        return data;
    }

    private static String getMainItemText(Row row) {
        for (int j = 0; j < row.getLastCellNum(); j++) {
            String value = getCellValue(row.getCell(j));
            if (MAIN_CRITERION_PATTERN.matcher(value).find()) {
                return value;
            }
        }
        return "";
    }

    private static int findNextMainItemRow(Sheet sheet, int startRow) {
        for (int i = startRow + 1; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            if (row != null) {
                for (int j = 0; j < row.getLastCellNum(); j++) {
                    String number = getCellValue(row.getCell(j));
                    if (MAIN_CRITERION_PATTERN.matcher(number).find()) {
                        return i;
                    }
                }
            }
        }
        return sheet.getLastRowNum() + 1;
    }

    private static String getMainItemByNumber(String mainItemNo, Map<String, Map<String, String>> mainSubItemMap) {
        return mainSubItemMap.keySet().stream()
                .filter(key -> key.startsWith(mainItemNo + "."))
                .findFirst()
                .orElse(null);
    }

    private static String getSubItemByNumber(String mainItem, String subItemNo, Map<String, Map<String, String>> mainSubItemMap) {
        if (mainItem == null) return null;
        Map<String, String> subItems = mainSubItemMap.get(mainItem);
        if (subItems == null) return null;
        return subItems.get(subItemNo);
    }

    private static int findEvaluationStartRow(Sheet sheet) {
        for (int i = 0; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            if (row != null) {
                for (int j = 0; j < row.getLastCellNum(); j++) {
                    Cell cell = row.getCell(j);
                    if (cell != null && EVAL_TABLE_HEADER.equals(getCellValue(cell))) {
                        System.out.println("Evaluation table start found at row " + i + ", col " + j);
                        return i + 1;
                    }
                }
            }
        }
        System.out.println("Evaluation data not found in sheet: " + sheet.getSheetName());
        return -1;
    }

    private static String getCellValue(Cell cell) {
        if (cell == null) {
            System.out.println("Cell is null");
            return "";
        }

        System.out.println("Cell at row=" + cell.getRowIndex() + ", col=" + cell.getColumnIndex() +
                ", type=" + cell.getCellType());

        DataFormatter formatter = new DataFormatter();
        String value;

        try {
            switch (cell.getCellType()) {
                case NUMERIC:
                    double numericValue = cell.getNumericCellValue();
                    String formatString = cell.getCellStyle().getDataFormatString();
                    System.out.println("Numeric value: " + numericValue +
                            ", Format string: " + formatString +
                            ", Is date formatted: " + DateUtil.isCellDateFormatted(cell));

                    if (DateUtil.isCellDateFormatted(cell)) {
                        value = formatter.formatCellValue(cell);
                    } else {
                        if (numericValue > 25569 && numericValue < 73050) {
                            Date date = DateUtil.getJavaDate(numericValue, false);
                            SimpleDateFormat sdf = new SimpleDateFormat("MM/dd/yyyy", Locale.US);
                            value = sdf.format(date);
                        } else {
                            value = formatter.formatCellValue(cell);
                        }
                    }
                    break;

                case STRING:
                    value = cell.getStringCellValue();
                    break;

                case FORMULA:
                    FormulaEvaluator evaluator = cell.getSheet().getWorkbook().getCreationHelper().createFormulaEvaluator();
                    CellValue formulaResult = evaluator.evaluate(cell);
                    System.out.println("Formula: " + cell.getCellFormula() +
                            ", Result type: " + formulaResult.getCellType() +
                            ", Result value: " + formulaResult.formatAsString());
                    value = formatter.formatCellValue(cell, evaluator);
                    break;

                case BLANK:
                    value = "";
                    break;

                default:
                    value = formatter.formatCellValue(cell);
                    break;
            }
        } catch (Exception e) {
            System.out.println("Error processing cell: " + e.getMessage());
            value = formatter.formatCellValue(cell);
        }

        value = value.trim()
                .replaceAll("\\s+", " ")
                .replaceAll("　", " ")
                .replaceAll("．", ".")
                .replaceAll("✕", "XX");

        System.out.println("Final formatted value: '" + value + "'");
        return value;
    }
}