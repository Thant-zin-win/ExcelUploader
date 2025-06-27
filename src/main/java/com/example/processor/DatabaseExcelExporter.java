package com.example.processor;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.sql.*;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Collectors;

public class DatabaseExcelExporter {

    // Regex to extract number from priority items like "<1>" or "1."
    private static final Pattern PRIORITY_NUMBER_EXTRACTOR = Pattern.compile("^[＜<]?(\\d+)[＞>]?[\\.．①②③④⑤⑥⑦].*");
    // Regex to extract main item number for general sorting (e.g., "1" from "1. Development")
    private static final Pattern MAIN_ITEM_NUMBER_EXTRACTOR = Pattern.compile("^(?:[0-9０-９]+)");

    // Helper to extract leading number and dot from main item for shortening
    private static String extractAndKeepLeadingNumber(String mainItem) {
        Matcher matcher = MAIN_ITEM_NUMBER_EXTRACTOR.matcher(mainItem);
        if (matcher.find()) {
            return matcher.group(0) + "."; // e.g., "5." or "6."
        }
        return ""; // No leading number found
    }

    // Helper method to get metadata headers in correct order
    private static List<String> getOrderedMetadataHeadersForSheet(Connection conn, int templateId, String sheetName) throws SQLException {
        List<String> headerKeys = new ArrayList<>();
        String sql = "SELECT rm.HeaderKey " +
                "FROM ResponseMetadata rm " +
                "JOIN Responses r ON rm.ResponseID = r.ResponseID " +
                "WHERE r.TemplateID = ? AND r.SheetName = ? AND rm.HeaderValue IS NOT NULL AND rm.HeaderValue != '' " +
                "GROUP BY rm.HeaderKey " +
                "ORDER BY MIN(rm.MetadataID)";
        try (PreparedStatement pstmt = conn.prepareStatement(sql)) {
            pstmt.setInt(1, templateId);
            pstmt.setString(2, sheetName);
            try (ResultSet rs = pstmt.executeQuery()) {
                while (rs.next()) {
                    headerKeys.add(rs.getString("HeaderKey"));
                }
            }
        }
        return headerKeys;
    }

    public static void exportDatabaseToExcel(Connection conn, OutputStream outputStream, String templateCategory) throws Exception {

        // Step 1: Find all TemplateIDs for the given templateCategory
        Set<Integer> templateIdsInCategory = new LinkedHashSet<>();
        String templateIdsSql = "SELECT TemplateID FROM Templates WHERE TemplateCategory = ?";
        try (PreparedStatement ps = conn.prepareStatement(templateIdsSql)) {
            ps.setString(1, templateCategory);
            try (ResultSet rs = ps.executeQuery()) {
                while (rs.next()) {
                    templateIdsInCategory.add(rs.getInt("TemplateID"));
                }
            }
        }

        if (templateIdsInCategory.isEmpty()) {
            throw new IllegalArgumentException("No templates found for the given category: " + templateCategory);
        }

        String placeholders = String.join(",", Collections.nCopies(templateIdsInCategory.size(), "?"));
        String dataSql = """
                SELECT r.ResponseID, r.TemplateID, t.TemplateName, t.UploadDate,
                       r.SheetName, rm.HeaderKey, rm.HeaderValue,
                       e.DataID, e.MainItem, e.SubItem, e.Evaluation, e.Comment, r.IsReuploaded
                FROM Templates t
                JOIN Responses r ON t.TemplateID = r.TemplateID
                LEFT JOIN ResponseMetadata rm ON r.ResponseID = rm.ResponseID
                LEFT JOIN EvaluationData e ON r.ResponseID = e.ResponseID
                WHERE r.TemplateID IN (""" + placeholders + ") " +
                "ORDER BY r.IsReuploaded ASC, r.ResponseID ASC, r.SheetName, e.DataID"; // PRIMARY CHANGE HERE

        Map<String, Map<String, Map<String, String>>> uniqueResponsesBySheet = new LinkedHashMap<>(); // SheetName -> UniqueKey -> Metadata Map (HeaderKey -> HeaderValue)
        Map<String, Map<String, List<String>>> mainItemsByResponse = new LinkedHashMap<>(); // SheetName -> UniqueKey -> List<MainItem>
        Map<String, Map<String, Map<String, List<String>>>> subItemsByResponse = new LinkedHashMap<>(); // SheetName -> UniqueKey -> MainItem -> List<SubItem>
        Map<String, Map<String, Map<String, Map<String, Map<String, String>>>> > evalDataByResponse = new LinkedHashMap<>(); // SheetName -> UniqueKey -> MainItem -> SubItem -> {Evaluation, Comment}

        Set<String> allUniqueSheetNames = new LinkedHashSet<>();

        Map<String, Integer> sheetToTemplateId = new HashMap<>(); // keep track of TemplateID by sheet

        try (PreparedStatement ps = conn.prepareStatement(dataSql)) {
            int paramIndex = 1;
            for (Integer templateId : templateIdsInCategory) {
                ps.setInt(paramIndex++, templateId);
            }

            try (ResultSet rs = ps.executeQuery()) {
                while (rs.next()) {
                    String sheetName = rs.getString("SheetName");
                    allUniqueSheetNames.add(sheetName);

                    int thisTemplateId = rs.getInt("TemplateID");
                    sheetToTemplateId.putIfAbsent(sheetName, thisTemplateId);

                    String responseId = String.valueOf(rs.getInt("ResponseID"));
                    boolean isReuploaded = rs.getBoolean("IsReuploaded");
                    String uniqueKey = rs.getString("TemplateName") + " (ID: " + responseId + (isReuploaded ? " - Reupload)" : ")");

                    // --- FIX: Correctly initialize all nested maps before accessing them ---
                    Map<String, Map<String, String>> sheetResponses = uniqueResponsesBySheet.computeIfAbsent(sheetName, k -> new LinkedHashMap<>());
                    Map<String, List<String>> sheetMainItems = mainItemsByResponse.computeIfAbsent(sheetName, k -> new LinkedHashMap<>());
                    Map<String, Map<String, List<String>>> sheetSubItems = subItemsByResponse.computeIfAbsent(sheetName, k -> new LinkedHashMap<>());
                    Map<String, Map<String, Map<String, Map<String, String>>>> sheetEvalData = evalDataByResponse.computeIfAbsent(sheetName, k -> new LinkedHashMap<>());

                    // Now, get/create the uniqueKey-level maps within those sheet maps
                    Map<String, String> uniqueKeyMetadata = sheetResponses.computeIfAbsent(uniqueKey, k -> new HashMap<>());
                    List<String> uniqueKeyMainItems = sheetMainItems.computeIfAbsent(uniqueKey, k -> new ArrayList<>());
                    Map<String, List<String>> uniqueKeySubItems = sheetSubItems.computeIfAbsent(uniqueKey, k -> new LinkedHashMap<>());
                    Map<String, Map<String, Map<String, String>>> uniqueKeyEvalData = sheetEvalData.computeIfAbsent(uniqueKey, k -> new LinkedHashMap<>());
                    // --- END FIX ---


                    String headerKey = rs.getString("HeaderKey");
                    String headerValue = rs.getString("HeaderValue");
                    if (headerKey != null && headerValue != null && !headerValue.trim().isEmpty()) {
                        uniqueKeyMetadata.put(headerKey, headerValue); // Use the correctly obtained map
                    }

                    String mainItem = rs.getString("MainItem");
                    String subItem = rs.getString("SubItem") != null ? rs.getString("SubItem") : "";
                    String eval = rs.getString("Evaluation") != null ? rs.getString("Evaluation") : "";
                    String comment = rs.getString("Comment") != null ? rs.getString("Comment") : "";

                    if (mainItem != null) {
                        List<String> mainItems = uniqueKeyMainItems; // Use the correctly obtained map
                        if (!mainItems.contains(mainItem)) {
                            mainItems.add(mainItem);
                        }

                        Map<String, List<String>> subItemsMapForResponse = uniqueKeySubItems; // Use the correctly obtained map
                        List<String> subItemsList = subItemsMapForResponse.computeIfAbsent(mainItem, _ -> new ArrayList<>());
                        if (!subItemsList.contains(subItem)) {
                            subItemsList.add(subItem);
                        }

                        uniqueKeyEvalData // Use the correctly obtained map
                                .computeIfAbsent(mainItem, _ -> new LinkedHashMap<>())
                                .put(subItem, Map.of("Evaluation", eval, "Comment", comment));
                    }
                }
            }
        }

        Workbook wb = new XSSFWorkbook();

        if (allUniqueSheetNames.isEmpty()) {
            allUniqueSheetNames.add("Default Sheet");
        }

        for (String sheetNameToProcess : allUniqueSheetNames) {
            Sheet sheet = wb.createSheet(sheetNameToProcess);
            int rowNum = 0;

            Font headerFont = wb.createFont();
            headerFont.setBold(true);
            headerFont.setFontName("Times New Roman");
            headerFont.setFontHeightInPoints((short) 12);

            Font dataFont = wb.createFont();
            dataFont.setFontName("Times New Roman");
            dataFont.setFontHeightInPoints((short) 12);

            CellStyle headerStyle = wb.createCellStyle();
            headerStyle.setFont(headerFont);
            headerStyle.setAlignment(HorizontalAlignment.CENTER);
            headerStyle.setVerticalAlignment(VerticalAlignment.CENTER);
            headerStyle.setWrapText(true);
            headerStyle.setBorderTop(BorderStyle.THIN);
            headerStyle.setBorderBottom(BorderStyle.THIN);
            headerStyle.setBorderLeft(BorderStyle.THIN);
            headerStyle.setBorderRight(BorderStyle.THIN);

            CellStyle dataStyle = wb.createCellStyle();
            dataStyle.setFont(dataFont);
            dataStyle.setWrapText(true);
            dataStyle.setVerticalAlignment(VerticalAlignment.TOP);
            dataStyle.setBorderTop(BorderStyle.THIN);
            dataStyle.setBorderBottom(BorderStyle.THIN);
            dataStyle.setBorderLeft(BorderStyle.THIN);
            dataStyle.setBorderRight(BorderStyle.THIN);

            Font resultTitleFont = wb.createFont();
            resultTitleFont.setBold(true);
            resultTitleFont.setFontName("Times New Roman");
            resultTitleFont.setFontHeightInPoints((short) 16);
            CellStyle resultTitleStyle = wb.createCellStyle();
            resultTitleStyle.setFont(resultTitleFont);
            resultTitleStyle.setAlignment(HorizontalAlignment.LEFT);
            resultTitleStyle.setVerticalAlignment(VerticalAlignment.CENTER);

            Row resultTitleRow = sheet.createRow(rowNum++);
            resultTitleRow.setHeightInPoints(25);
            Cell resultTitleCell = resultTitleRow.createCell(0);
            resultTitleCell.setCellValue("Result");
            resultTitleCell.setCellStyle(resultTitleStyle);

            rowNum++;

            Map<String, Map<String, String>> responsesForCurrentSheet = uniqueResponsesBySheet.getOrDefault(sheetNameToProcess, Collections.emptyMap());

            int targetTemplateId = -1;
            if (sheetToTemplateId.containsKey(sheetNameToProcess)) {
                targetTemplateId = sheetToTemplateId.get(sheetNameToProcess);
            } else if (!templateIdsInCategory.isEmpty()) {
                targetTemplateId = templateIdsInCategory.iterator().next();
            }
            List<String> sortedMetadataHeaders = getOrderedMetadataHeadersForSheet(conn, targetTemplateId, sheetNameToProcess);

            Set<String> sheetMainItems = new LinkedHashSet<>();
            Map<String, Set<String>> sheetSubItemsByMainItem = new LinkedHashMap<>();

            Map<String, Map<String, Map<String, Map<String, String>>>> evalDataForCurrentSheet =
                evalDataByResponse.getOrDefault(sheetNameToProcess, new LinkedHashMap<String, Map<String, Map<String, Map<String, String>>>>());

            for (Map.Entry<String, Map<String, Map<String, Map<String, String>>>> entryByUniqueKey : evalDataForCurrentSheet.entrySet()) {
                Map<String, Map<String, Map<String, String>>> evalDataByMainItem = entryByUniqueKey.getValue();
                for (Map.Entry<String, Map<String, Map<String, String>>> entryByMainItem : evalDataByMainItem.entrySet()) {
                    String mainItem = entryByMainItem.getKey();
                    sheetMainItems.add(mainItem);

                    Map<String, Map<String, String>> evalDataBySubItem = entryByMainItem.getValue();
                    for (String subItem : evalDataBySubItem.keySet()) {
                        Map<String, String> result = evalDataBySubItem.get(subItem);
                        if (result != null && (!result.getOrDefault("Evaluation", "").isEmpty() || !result.getOrDefault("Comment", "").isEmpty())) {
                             sheetSubItemsByMainItem.computeIfAbsent(mainItem, _ -> new LinkedHashSet<>()).add(subItem);
                        }
                    }
                    if (mainItem.contains("ご要望等")) {
                        Map<String, String> result = evalDataByMainItem.getOrDefault(mainItem, Collections.emptyMap()).getOrDefault("", Collections.emptyMap());
                        if (result != null && !result.getOrDefault("Comment", "").isEmpty()) {
                            sheetMainItems.add(mainItem);
                            sheetSubItemsByMainItem.computeIfAbsent(mainItem, _ -> new LinkedHashSet<>()).add("");
                        }
                    }
                }
            }

            List<String> sortedSheetMainItems = new ArrayList<>(sheetMainItems);
            sortedSheetMainItems.sort(new Comparator<String>() {
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

                private Integer extractLeadingNumber(String s) {
                    Matcher matcher = MAIN_ITEM_NUMBER_EXTRACTOR.matcher(s);
                    return matcher.find() ? Integer.parseInt(matcher.group(0)) : null;
                }
            });

            Row row0 = sheet.createRow(rowNum++);
            Row row1 = sheet.createRow(rowNum++);
            Row row2 = sheet.createRow(rowNum++);

            int colIndex = 0;
            for (String col : sortedMetadataHeaders) {
                Cell cell0 = row0.createCell(colIndex);
                cell0.setCellValue(col);
                cell0.setCellStyle(headerStyle);
                row1.createCell(colIndex).setCellStyle(headerStyle);
                row2.createCell(colIndex).setCellStyle(headerStyle);

                sheet.addMergedRegion(new CellRangeAddress(row0.getRowNum(), row2.getRowNum(), colIndex, colIndex));
                colIndex++;
            }

            // This map stores the column indices for each main item and its sub-items (Evaluation/Comment columns)
            Map<String, Map<String, Integer>> evalColIndexMap = new LinkedHashMap<>();

            for (String mainItem : sortedSheetMainItems) {
                int mainStart = colIndex;
                Map<String, Integer> subMap = new LinkedHashMap<>(); // This subMap holds column indices for current mainItem
                Set<String> subItemsSet = sheetSubItemsByMainItem.getOrDefault(mainItem, new LinkedHashSet<>());

                List<String> validSubItems = subItemsSet.stream()
                        .filter(subItem -> subItem != null && !subItem.trim().isEmpty())
                        .collect(Collectors.toCollection(ArrayList::new));

                boolean isPrioritySection = mainItem.contains("より満足いただくために");
                if (isPrioritySection) {
                    validSubItems.sort(Comparator.comparingInt(s -> {
                        Matcher matcher = PRIORITY_NUMBER_EXTRACTOR.matcher(s);
                        return matcher.find() ? Integer.parseInt(matcher.group(1)) : Integer.MAX_VALUE;
                    }));
                } else {
                    Collections.sort(validSubItems);
                }

                boolean isRequestsSection = mainItem.contains("ご要望等");

                String displayMainItem = mainItem;
                String leadingNumber = extractAndKeepLeadingNumber(mainItem);
                if (mainItem.contains("より満足いただくために、弊社が真っ先に解決/取組むべき項目/事柄はどのようなものだとおthink?")) {
                    displayMainItem = leadingNumber + "より満足いただくため";
                } else if (mainItem.contains("ご要望等がございましたらご記入ください。")) {
                    displayMainItem = leadingNumber + "ご要望";
                }

                if (validSubItems.isEmpty() || isRequestsSection) {
                    Cell cell0 = row0.createCell(colIndex);
                    cell0.setCellValue(displayMainItem);
                    cell0.setCellStyle(headerStyle);
                    row1.createCell(colIndex).setCellStyle(headerStyle);
                    row2.createCell(colIndex).setCellStyle(headerStyle);

                    sheet.addMergedRegion(new CellRangeAddress(row0.getRowNum(), row2.getRowNum(), colIndex, colIndex));

                    subMap.put("", colIndex);
                    colIndex++;
                } else {
                    if (isPrioritySection) {
                        int priorityIndex = 1;
                        for (String subItem : validSubItems) {
                            String subItemHeader = "<" + priorityIndex + ">";
                            int subStart = colIndex;

                            Cell cell1 = row1.createCell(colIndex);
                            cell1.setCellValue(subItemHeader);
                            cell1.setCellStyle(headerStyle);
                            row1.createCell(colIndex + 1).setCellStyle(headerStyle);

                            Cell evalCell = row2.createCell(colIndex);
                            evalCell.setCellValue("Evaluation");
                            evalCell.setCellStyle(headerStyle);
                            colIndex++;

                            Cell commentCell = row2.createCell(colIndex);
                            commentCell.setCellValue("Comment");
                            commentCell.setCellStyle(headerStyle);
                            colIndex++;

                            sheet.addMergedRegion(new CellRangeAddress(row1.getRowNum(), row1.getRowNum(), subStart, colIndex - 1));

                            subMap.put(subItem, subStart);
                            priorityIndex++;
                        }
                    } else {
                        for (String subItem : validSubItems) {
                            int subStart = colIndex;
                            Cell cell1 = row1.createCell(colIndex);
                            cell1.setCellValue(subItem);
                            cell1.setCellStyle(headerStyle);
                            row1.createCell(colIndex + 1).setCellStyle(headerStyle);

                            Cell evalCell = row2.createCell(colIndex);
                            evalCell.setCellValue("Evaluation");
                            evalCell.setCellStyle(headerStyle);
                            colIndex++;

                            Cell commentCell = row2.createCell(colIndex);
                            commentCell.setCellValue("Comment");
                            commentCell.setCellStyle(headerStyle);
                            colIndex++;

                            sheet.addMergedRegion(new CellRangeAddress(row1.getRowNum(), row1.getRowNum(), subStart, colIndex - 1));

                            subMap.put(subItem, subStart);
                        }
                    }
                }

                if ((colIndex - mainStart) > 1) {
                    for (int k = mainStart; k < colIndex; k++) {
                        Cell cell = row0.getCell(k);
                        if (cell == null) {
                            cell = row0.createCell(k);
                        }
                        cell.setCellStyle(headerStyle);
                    }
                    Cell mainCell = row0.getCell(mainStart);
                    mainCell.setCellValue(displayMainItem);
                    sheet.addMergedRegion(new CellRangeAddress(row0.getRowNum(), row0.getRowNum(), mainStart, colIndex - 1));
                } else {
                    Cell mainCell = row0.createCell(mainStart);
                    mainCell.setCellValue(displayMainItem);
                    mainCell.setCellStyle(headerStyle);
                }

                evalColIndexMap.put(mainItem, subMap); // Store the subMap for this mainItem
            }

            if (colIndex > 1) {
                sheet.addMergedRegion(new CellRangeAddress(resultTitleRow.getRowNum(), resultTitleRow.getRowNum(), 0, colIndex - 1));
            }

            Map<String, Map<String, String>> responsesForCurrentSheetOrdered = responsesForCurrentSheet; // This LinkedHashMap inherently preserves order from SQL
            Map<String, Map<String, Map<String, Map<String, String>>>> currentSheetEvalData =
                evalDataByResponse.getOrDefault(sheetNameToProcess, new LinkedHashMap<String, Map<String, Map<String, Map<String, String>>>>());

            for (Map.Entry<String, Map<String, String>> responseEntry : responsesForCurrentSheetOrdered.entrySet()) {
                String uniqueKey = responseEntry.getKey();
                Row dataRow = sheet.createRow(rowNum++);

                for (int i = 0; i < colIndex; i++) {
                    Cell cell = dataRow.createCell(i);
                    cell.setCellStyle(dataStyle);
                }

                int currentMetadataCol = 0;
                for (String col : sortedMetadataHeaders) {
                    String val = responseEntry.getValue().get(col);
                    Cell cell = dataRow.getCell(currentMetadataCol++);
                    cell.setCellValue(val != null ? val : "");
                }

                Map<String, Map<String, Map<String, String>>> responseEvalDataMap =
                    currentSheetEvalData.getOrDefault(uniqueKey, new LinkedHashMap<>());

                for (String mainItem : sortedSheetMainItems) {
                    // FIX: Retrieve the correct sub-item column map for the current mainItem
                    Map<String, Integer> currentMainItemColMap = evalColIndexMap.get(mainItem);
                    if (currentMainItemColMap == null) {
                        System.err.println("ERROR: Column index map not found for MainItem: '" + mainItem + "'. This indicates a header construction issue or data inconsistency.");
                        continue;
                    }

                    List<String> allSubItems = new ArrayList<>(sheetSubItemsByMainItem.getOrDefault(mainItem, Collections.emptySet()));
                    List<String> validSubItems = allSubItems.stream()
                                                       .filter(s -> s != null && !s.trim().isEmpty())
                                                       .collect(Collectors.toCollection(ArrayList::new));

                    boolean isPrioritySection = mainItem.contains("より満足いただくために");
                    if (isPrioritySection) {
                        validSubItems.sort(Comparator.comparingInt(s -> {
                            Matcher matcher = PRIORITY_NUMBER_EXTRACTOR.matcher(s);
                            return matcher.find() ? Integer.parseInt(matcher.group(1)) : Integer.MAX_VALUE;
                        }));
                    } else {
                        Collections.sort(validSubItems);
                    }

                    boolean isRequestsSection = mainItem.contains("ご要望等");

                    if (validSubItems.isEmpty() || isRequestsSection) {
                        Map<String, Map<String, String>> mainEval = responseEvalDataMap.getOrDefault(mainItem, new LinkedHashMap<>());
                        Map<String, String> result = mainEval.getOrDefault("", Collections.emptyMap());

                        Integer columnIndex = currentMainItemColMap.get(""); // Use the retrieved map
                        if (columnIndex != null) {
                            Cell cell = dataRow.getCell(columnIndex);
                            cell.setCellValue(result.getOrDefault("Comment", ""));
                        } else {
                            System.err.println("ERROR: Column index not found for empty subItem in MainItem: '" + mainItem + "' (Requests/No Sub-items section).");
                        }
                    } else {
                        for (String subItem : validSubItems) {
                            Map<String, Map<String, String>> mainEval = responseEvalDataMap.getOrDefault(mainItem, Collections.emptyMap());
                            Map<String, String> result = mainEval.getOrDefault(subItem, Collections.emptyMap());

                            Integer subStart = currentMainItemColMap.get(subItem); // Use the retrieved map
                            if (subStart != null) {
                                Cell evalCell = dataRow.getCell(subStart);
                                evalCell.setCellValue(result.getOrDefault("Evaluation", ""));

                                Cell commentCell = dataRow.getCell(subStart + 1);
                                commentCell.setCellValue(result.getOrDefault("Comment", ""));
                            } else {
                                System.err.println("ERROR: Column index not found for SubItem: '" + subItem + "' in MainItem: '" + mainItem + "'.");
                            }
                        }
                    }
                }
            }

            for (int i = 0; i < colIndex; i++) {
                sheet.autoSizeColumn(i);
            }
        }

        wb.write(outputStream);
        wb.close();
    }
}