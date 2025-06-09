package com.example.processor;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.*;
import java.sql.*;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.Date;
import java.util.regex.Pattern;

public class ExcelProcessor {

    private static final Pattern MAIN_CRITERION_PATTERN = Pattern.compile("^(?:[0-9０-９]+\\.|(?:I{1,3}|IV|V|VI{0,3})\\.)[^①②③④⑤⑥⑦].*");
    private static final Pattern PRIORITY_PATTERN = Pattern.compile("^(?:[１２３４１-４]|[1-4])\\.[①②③④⑤⑥⑦]|(?:<|＜)(?:[1-4]|[１-４])(?:>|＞)");
    private static final Pattern EVALUATION_PATTERN = Pattern.compile("^\\d+:(?:Not Related|[A-Za-z]+)$");
    private static final Pattern NOTE_PATTERN = Pattern.compile("^(※.*|.*(?:お願い|注意|注|備考|説明|選択ボックス).*)", Pattern.UNICODE_CASE);
    private static final String EVAL_TABLE_HEADER = "評価項目";

    public static int processExcelFile(Connection conn, InputStream inputStream, String templateName) throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook(inputStream)) {
            int templateId = getOrCreateTemplate(conn, templateName);

            for (Sheet sheet : workbook) {
                String sheetName = sheet.getSheetName();
                System.out.println("Processing sheet: " + sheetName);

                if (sheetName.startsWith("評価結果リスト_")) {
                    System.out.println("Skipping summary sheet: " + sheetName);
                    continue;
                }

                if (sheetName.equals("表紙")) {
                    System.out.println("Skipping cover/empty sheet: " + sheetName);
                    continue;
                }

                // Check if a response for this TemplateID and SheetName already exists
                int existingResponseId = getExistingResponseId(conn, templateId, sheetName);
                if (existingResponseId != -1) {
                    System.out.println("Response already exists for TemplateID=" + templateId + ", SheetName=" + sheetName + ", ResponseID=" + existingResponseId + ". Skipping.");
                    continue;
                }

                List<Map<String, String>> metadataList = extractMetadata(sheet);
                List<Map<String, String>> evalData = extractEvaluationData(sheet);

                int firstResponseId = -1;
                if (metadataList.isEmpty()) {
                    System.out.println("No metadata found for sheet: " + sheetName + ". Creating default response.");
                    Map<String, String> defaultMetadata = new HashMap<>();
                    firstResponseId = storeResponse(conn, templateId, sheetName, defaultMetadata);
                } else {
                    for (Map<String, String> metadata : metadataList) {
                        int responseId = storeResponse(conn, templateId, sheetName, metadata);
                        if (firstResponseId == -1) firstResponseId = responseId;
                    }
                }

                if (firstResponseId != -1 && !evalData.isEmpty()) {
                    storeEvaluationData(conn, firstResponseId, evalData);
                    System.out.println("Stored evaluation data for sheet: " + sheetName + " with ResponseID=" + firstResponseId);
                } else {
                    System.out.println("No evaluation data stored for sheet: " + sheetName + " (ResponseID=" + firstResponseId + ", EvalDataSize=" + evalData.size() + ")");
                }
            }
            return templateId;
        }
    }

    private static int getExistingResponseId(Connection conn, int templateId, String sheetName) throws SQLException {
        String sql = "SELECT ResponseID FROM Responses WHERE TemplateID = ? AND SheetName = ?";
        try (PreparedStatement pstmt = conn.prepareStatement(sql)) {
            pstmt.setInt(1, templateId);
            pstmt.setString(2, sheetName);
            try (ResultSet rs = pstmt.executeQuery()) {
                if (rs.next()) {
                    return rs.getInt("ResponseID");
                }
            }
        }
        return -1;
    }

    private static List<Map<String, String>> extractMetadata(Sheet sheet) {
        List<Map<String, String>> metadataList = new ArrayList<>();
        int evalStartRow = findEvaluationStartRow(sheet);
        int maxRowsToSearch = (evalStartRow == -1) ? sheet.getLastRowNum() + 1 : evalStartRow - 1;

        System.out.println("Scanning for metadata in sheet: " + sheet.getSheetName() + " up to row: " + maxRowsToSearch);
        Map<String, String> metadata = new HashMap<>();
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

            // Skip the row if it contains EVAL_TABLE_HEADER
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
                // Find the next non-empty cell as header
                while (j < maxCols && getCellValue(row.getCell(j)).isEmpty()) j++;
                if (j >= maxCols) break;
                String header = getCellValue(row.getCell(j));
                if (isNote(header)) {
                    j++;
                    continue;
                }
                // Find the next non-empty cell as value
                int m = j + 1;
                while (m < maxCols && getCellValue(row.getCell(m)).isEmpty()) m++;
                String value = (m < maxCols) ? getCellValue(row.getCell(m)) : "";
                if (!value.isEmpty() && !isNote(value)) {
                    metadata.put(header, value);
                    System.out.println("Stored: " + header + " = " + value + " at row " + i + ", header col " + j + ", value col " + m);
                    j = m + 1; // Move past the value
                } else {
                    j++; // Move to next potential header
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
                    for (int j = i + 1; j < nextMainItemRow && j <= sheet.getLastRowNum(); j++) {
                        Row priorityRow = sheet.getRow(j);
                        if (priorityRow == null) continue;
                        for (int k = 0; k < priorityRow.getLastCellNum(); k++) {
                            String priorityItem = getCellValue(priorityRow.getCell(k));
                            if (PRIORITY_PATTERN.matcher(priorityItem).matches()) {
                                String description = "";
                                for (int m = k + 1; m < priorityRow.getLastCellNum(); m++) {
                                    String cellValue = getCellValue(priorityRow.getCell(m));
                                    if (!cellValue.isEmpty() && !PRIORITY_PATTERN.matcher(cellValue).matches()) {
                                        description = cellValue;
                                        break;
                                    }
                                }
                                if (!description.isEmpty()) {
                                    Map<String, String> entry = new HashMap<>();
                                    String mainItemNo = priorityItem.replaceAll("[^0-9１-４]", "");
                                    String subItemNo = priorityItem.replaceAll(".*[\\.]([①②③④⑤⑥⑦])", "$1");
                                    String fullMainItem = getMainItemByNumber(mainItemNo, mainSubItemMap);
                                    String fullSubItem = getSubItemByNumber(fullMainItem, subItemNo, mainSubItemMap);
                                    entry.put("MainItem", currentMainItem);
                                    entry.put("SubItem", (fullMainItem != null && fullSubItem != null) ?
                                            fullMainItem + " - " + fullSubItem : priorityItem);
                                    entry.put("Evaluation", "");
                                    entry.put("Comment", description);
                                    data.add(entry);
                                    System.out.println("Priority item at row " + j + ", col " + k + ": Sub=" + entry.get("SubItem") + ", Comment=" + description);
                                } else {
                                    System.out.println("Skipping priority item at row " + j + ", col " + k + ": " + priorityItem + " due to empty description");
                                }
                            }
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
                    String effectiveSubItem = subItem;
                    for (int j = 0; j < row.getLastCellNum(); j++) {
                        String value = getCellValue(row.getCell(j));
                        if (!value.isEmpty() && !value.matches("^[＜＞]") && !EVALUATION_PATTERN.matcher(value).matches() && !value.matches("^[①②③④⑤⑥⑦１-７1-7]")) {
                            effectiveSubItem = value;
                            break;
                        }
                    }
                    if (!effectiveSubItem.isEmpty()) {
                        Map<String, String> entry = new HashMap<>();
                        entry.put("MainItem", currentMainItem);
                        entry.put("SubItem", subItem.isEmpty() ? effectiveSubItem : subItem + " " + effectiveSubItem);
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

    private static int getOrCreateTemplate(Connection conn, String templateName) throws SQLException {
        String sql = "INSERT IGNORE INTO Templates (TemplateName, UploadDate) VALUES (?, NOW())";
        try (PreparedStatement pstmt = conn.prepareStatement(sql, Statement.RETURN_GENERATED_KEYS)) {
            pstmt.setString(1, templateName);
            pstmt.executeUpdate();
            try (ResultSet rs = pstmt.getGeneratedKeys()) {
                if (rs.next()) return rs.getInt(1);
            }
        }
        try (PreparedStatement pstmt = conn.prepareStatement("SELECT TemplateID FROM Templates WHERE TemplateName = ?")) {
            pstmt.setString(1, templateName);
            try (ResultSet rs = pstmt.executeQuery()) {
                return rs.next() ? rs.getInt("TemplateID") : -1;
            }
        }
    }

    private static int storeResponse(Connection conn, int templateId, String sheetName, Map<String, String> metadata) throws SQLException {
        // Check if response already exists
        int existingResponseId = getExistingResponseId(conn, templateId, sheetName);
        if (existingResponseId != -1) {
            System.out.println("Response already exists for TemplateID=" + templateId + ", SheetName=" + sheetName + ", ResponseID=" + existingResponseId + ". Skipping metadata insert.");
            return existingResponseId;
        }

        String insertResponseSql = "INSERT INTO Responses (TemplateID, SheetName) VALUES (?, ?)";
        int responseId;
        try (PreparedStatement pstmt = conn.prepareStatement(insertResponseSql, Statement.RETURN_GENERATED_KEYS)) {
            pstmt.setInt(1, templateId);
            pstmt.setString(2, sheetName);
            pstmt.executeUpdate();
            try (ResultSet rs = pstmt.getGeneratedKeys()) {
                if (rs.next()) {
                    responseId = rs.getInt(1);
                } else {
                    throw new SQLException("Failed to retrieve ResponseID");
                }
            }
        }

        String insertMetadataSql = "INSERT INTO ResponseMetadata (ResponseID, HeaderKey, HeaderValue) VALUES (?, ?, ?)";
        try (PreparedStatement pstmt = conn.prepareStatement(insertMetadataSql)) {
            for (Map.Entry<String, String> entry : metadata.entrySet()) {
                pstmt.setInt(1, responseId);
                pstmt.setString(2, entry.getKey());
                pstmt.setString(3, entry.getValue());
                pstmt.addBatch();
            }
            pstmt.executeBatch();
            System.out.println("Stored metadata for ResponseID=" + responseId + ": " + metadata);
        }

        return responseId;
    }

    private static void storeEvaluationData(Connection conn, int responseId, List<Map<String, String>> evalData) throws SQLException {
        String checkSql = "SELECT DataID FROM EvaluationData WHERE ResponseID = ? AND MainItem = ? AND SubItem = ? AND Evaluation = ? AND Comment = ?";
        String insertSql = "INSERT INTO EvaluationData (ResponseID, MainItem, SubItem, Evaluation, Comment) VALUES (?, ?, ?, ?, ?)";

        try (PreparedStatement pstmtCheck = conn.prepareStatement(checkSql);
             PreparedStatement pstmtInsert = conn.prepareStatement(insertSql)) {
            for (Map<String, String> entry : evalData) {
                String mainItem = entry.get("MainItem");
                String subItem = entry.get("SubItem");
                String evaluation = entry.get("Evaluation");
                String comment = entry.get("Comment");

                pstmtCheck.setInt(1, responseId);
                pstmtCheck.setString(2, mainItem);
                pstmtCheck.setString(3, subItem);
                pstmtCheck.setString(4, evaluation);
                pstmtCheck.setString(5, comment);
                try (ResultSet rs = pstmtCheck.executeQuery()) {
                    if (!rs.next()) {
                        pstmtInsert.setInt(1, responseId);
                        pstmtInsert.setString(2, mainItem);
                        pstmtInsert.setString(3, subItem);
                        pstmtInsert.setString(4, evaluation);
                        pstmtInsert.setString(5, comment);
                        pstmtInsert.addBatch();
                        System.out.println("Inserted new evaluation data: ResponseID=" + responseId + ", MainItem=" + mainItem + ", SubItem=" + subItem);
                    } else {
                        System.out.println("Skipping duplicate evaluation data: ResponseID=" + responseId + ", MainItem=" + mainItem + ", SubItem=" + subItem);
                    }
                }
            }
            pstmtInsert.executeBatch();
        }
    }

    private static String getCellValue(Cell cell) {
        // Handle null cells
        if (cell == null) {
            System.out.println("Cell is null");
            return "";
        }

        // Log cell details for debugging
        System.out.println("Cell at row=" + cell.getRowIndex() + ", col=" + cell.getColumnIndex() + 
                          ", type=" + cell.getCellType());

        DataFormatter formatter = new DataFormatter();
        String value;

        try {
            switch (cell.getCellType()) {
                case NUMERIC:
                    // Log numeric details
                    double numericValue = cell.getNumericCellValue();
                    String formatString = cell.getCellStyle().getDataFormatString();
                    System.out.println("Numeric value: " + numericValue + 
                                     ", Format string: " + formatString + 
                                     ", Is date formatted: " + DateUtil.isCellDateFormatted(cell));

                    // Check if explicitly date-formatted
                    if (DateUtil.isCellDateFormatted(cell)) {
                        value = formatter.formatCellValue(cell);
                    } else {
                        // Heuristic: If numeric value looks like a date (post-1970, ~25569)
                        if (numericValue > 25569 && numericValue < 73050) { // Up to ~2100
                            Date date = DateUtil.getJavaDate(numericValue, false);
                            SimpleDateFormat sdf = new SimpleDateFormat("MM/dd/yyyy", Locale.US);
                            value = sdf.format(date);
                        } else {
                            value = formatter.formatCellValue(cell); // Plain number
                        }
                    }
                    break;

                case STRING:
                    value = cell.getStringCellValue();
                    break;

                case FORMULA:
                    // Evaluate formula and format result
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
            value = formatter.formatCellValue(cell); // Fallback to DataFormatter
        }

        // Clean up the value
        value = value.trim()
                     .replaceAll("\\s+", " ")
                     .replaceAll("\u3000", " ")
                     .replaceAll("．", ".")
                     .replaceAll("\u2715", "XX");

        System.out.println("Final formatted value: '" + value + "'");
        return value;
    }
}