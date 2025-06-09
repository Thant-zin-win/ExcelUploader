package com.example.processor;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.sql.*;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.Date;
import java.util.stream.Collectors;

public class DatabaseExcelExporter {

    public static void exportDatabaseToExcel(Connection conn, OutputStream outputStream, int templateId) throws Exception {

        // Step 1: Find all SheetNames and their ResponseIDs for the given templateId
        Map<String, Set<Integer>> sheetNameToResponseIds = new LinkedHashMap<>();
        String sheetSql = "SELECT SheetName, ResponseID FROM Responses WHERE TemplateID = ?";
        try (PreparedStatement ps = conn.prepareStatement(sheetSql)) {
            ps.setInt(1, templateId);
            try (ResultSet rs = ps.executeQuery()) {
                while (rs.next()) {
                    String sheetName = rs.getString("SheetName");
                    int responseId = rs.getInt("ResponseID");
                    sheetNameToResponseIds.computeIfAbsent(sheetName, k -> new LinkedHashSet<>()).add(responseId);
                }
            }
        }

        if (sheetNameToResponseIds.isEmpty()) {
            throw new IllegalArgumentException("No sheets found for the given template ID: " + templateId);
        }

        // Step 2: Fetch data for records that match the templateId and sheet names
        Set<String> targetSheetNames = sheetNameToResponseIds.keySet();
        String placeholders = String.join(",", Collections.nCopies(targetSheetNames.size(), "?"));
        String dataSql = """
        	    SELECT r.ResponseID, r.TemplateID, t.TemplateName, t.UploadDate,
        	           r.SheetName, rm.HeaderKey, rm.HeaderValue,
        	           e.DataID, e.MainItem, e.SubItem, e.Evaluation, e.Comment
        	    FROM Templates t
        	    JOIN Responses r ON t.TemplateID = r.TemplateID
        	    LEFT JOIN ResponseMetadata rm ON r.ResponseID = rm.ResponseID
        	    LEFT JOIN EvaluationData e ON r.ResponseID = e.ResponseID
        	    WHERE r.TemplateID = ? AND r.SheetName IN (""" + placeholders + ") " +
        	    "ORDER BY r.ResponseID, r.SheetName, e.DataID";

        // Data structures
        Map<String, Map<String, Map<String, String>>> uniqueResponsesBySheet = new LinkedHashMap<>();
        Map<String, Map<String, List<String>>> mainItemsByResponse = new LinkedHashMap<>();
        Map<String, Map<String, Map<String, List<String>>>> subItemsByResponse = new LinkedHashMap<>();
        Map<String, Map<String, Map<String, Map<String, Map<String, String>>>>> evalDataByResponse = new LinkedHashMap<>();
        Map<String, Set<String>> nonEmptyColsBySheet = new LinkedHashMap<>();

        try (PreparedStatement ps = conn.prepareStatement(dataSql)) {
            ps.setInt(1, templateId);
            int i = 2;
            for (String sheetName : targetSheetNames) {
                ps.setString(i++, sheetName);
                nonEmptyColsBySheet.put(sheetName, new LinkedHashSet<>());
            }

            try (ResultSet rs = ps.executeQuery()) {
                while (rs.next()) {
                    String sheetName = rs.getString("SheetName");
                    String responseId = String.valueOf(rs.getInt("ResponseID"));
                    String uniqueKey = responseId;

                    uniqueResponsesBySheet.computeIfAbsent(sheetName, k -> new LinkedHashMap<>())
                                          .putIfAbsent(uniqueKey, new HashMap<>());
                    mainItemsByResponse.computeIfAbsent(sheetName, k -> new LinkedHashMap<>())
                                       .putIfAbsent(uniqueKey, new ArrayList<>());
                    subItemsByResponse.computeIfAbsent(sheetName, k -> new LinkedHashMap<>())
                                      .putIfAbsent(uniqueKey, new LinkedHashMap<>());
                    evalDataByResponse.computeIfAbsent(sheetName, k -> new LinkedHashMap<>())
                                      .putIfAbsent(uniqueKey, new LinkedHashMap<>());

                    String headerKey = rs.getString("HeaderKey");
                    String headerValue = rs.getString("HeaderValue");
                    if (headerKey != null && headerValue != null && !headerValue.trim().isEmpty()) {
                        uniqueResponsesBySheet.get(sheetName).get(uniqueKey).put(headerKey, headerValue);
                        nonEmptyColsBySheet.get(sheetName).add(headerKey);
                    }

                    String mainItem = rs.getString("MainItem");
                    String subItem = rs.getString("SubItem") != null ? rs.getString("SubItem") : "";
                    String eval = rs.getString("Evaluation") != null ? rs.getString("Evaluation") : "";
                    String comment = rs.getString("Comment") != null ? rs.getString("Comment") : "";

                    if (mainItem != null) {
                        List<String> mainItems = mainItemsByResponse.get(sheetName).get(uniqueKey);
                        if (!mainItems.contains(mainItem)) {
                            mainItems.add(mainItem);
                        }

                        Map<String, List<String>> subItemsMap = subItemsByResponse.get(sheetName).get(uniqueKey);
                        List<String> subItems = subItemsMap.computeIfAbsent(mainItem, k -> new ArrayList<>());
                        if (!subItems.contains(subItem)) {
                            subItems.add(subItem);
                        }

                        evalDataByResponse.get(sheetName).get(uniqueKey)
                                .computeIfAbsent(mainItem, k -> new LinkedHashMap<>())
                                .put(subItem, Map.of("Evaluation", eval, "Comment", comment));
                    }
                }
            }
        }

        Workbook wb = new XSSFWorkbook();

        // === Output Sheet Creation ===
        for (String sheetName : uniqueResponsesBySheet.keySet()) {
            Sheet sheet = wb.createSheet(sheetName);
            int rowNum = 0;

            // === Styles ===
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

            // Step 3: Fetch headers for this specific sheet
            List<String> sheetHeaders = new ArrayList<>();
            String headerSql = "SELECT rm.HeaderKey FROM ResponseMetadata rm WHERE rm.ResponseID IN (SELECT ResponseID FROM Responses WHERE TemplateID = ? AND SheetName = ?) ORDER BY rm.MetadataID";
            try (PreparedStatement ps = conn.prepareStatement(headerSql)) {
                ps.setInt(1, templateId);
                ps.setString(2, sheetName);
                try (ResultSet rs = ps.executeQuery()) {
                    while (rs.next()) {
                        String headerKey = rs.getString("HeaderKey");
                        if (!sheetHeaders.contains(headerKey)) {
                            sheetHeaders.add(headerKey);
                        }
                    }
                }
            }

            // Collect all unique columns and evaluation data
            Set<String> allActiveCols = new LinkedHashSet<>();
            Set<String> allMainItems = new LinkedHashSet<>();
            Map<String, Set<String>> allSubItemsByMainItem = new LinkedHashMap<>();
            Map<String, Map<String, String>> responses = uniqueResponsesBySheet.get(sheetName);
            Map<String, Map<String, List<String>>> subItems = subItemsByResponse.get(sheetName);

            if (subItems == null) {
                subItems = new LinkedHashMap<>();
            }

            // Add SheetName as the first column
            allActiveCols.add("SheetName");

            // Add metadata headers specific to this sheet
            for (String header : sheetHeaders) {
                allActiveCols.add(header);
            }

            for (Map.Entry<String, Map<String, List<String>>> entry : subItems.entrySet()) {
                String uniqueKey = entry.getKey();
                Map<String, List<String>> subItemsMap = entry.getValue();
                if (subItemsMap != null) {
                    allMainItems.addAll(subItemsMap.keySet());
                    for (Map.Entry<String, List<String>> subEntry : subItemsMap.entrySet()) {
                        allSubItemsByMainItem.computeIfAbsent(subEntry.getKey(), k -> new LinkedHashSet<>()).addAll(subEntry.getValue());
                    }
                }
            }

            // Track merged regions
            Set<String> mergedRegions = new HashSet<>();

            // Create header rows
            Row row0 = sheet.createRow(rowNum++);
            Row row1 = sheet.createRow(rowNum++);
            Row row2 = sheet.createRow(rowNum++);

            int colIndex = 0;
            for (String col : allActiveCols) {
                Cell cell0 = row0.createCell(colIndex);
                cell0.setCellValue(col);
                cell0.setCellStyle(headerStyle);
                row1.createCell(colIndex).setCellStyle(headerStyle);
                row2.createCell(colIndex).setCellStyle(headerStyle);
                String regionKey = "0:2:" + colIndex + ":" + colIndex;
                if (!mergedRegions.contains(regionKey)) {
                    sheet.addMergedRegion(new CellRangeAddress(0, 2, colIndex, colIndex));
                    mergedRegions.add(regionKey);
                }
                colIndex++;
            }

            Map<String, Map<String, Integer>> evalColIndexMap = new LinkedHashMap<>();
            for (String mainItem : allMainItems) {
                int mainStart = colIndex;
                Map<String, Integer> subMap = new LinkedHashMap<>();
                Set<String> subItemsSet = allSubItemsByMainItem.getOrDefault(mainItem, new LinkedHashSet<>());

                List<String> validSubItems = subItemsSet.stream()
                        .filter(subItem -> subItem != null && !subItem.trim().isEmpty())
                        .collect(Collectors.toList());

                if (!subItemsSet.isEmpty()) {
                    boolean isPrioritySection = mainItem.contains("より満足いただくために");

                    if (isPrioritySection) {
                        int priorityIndex = 1;
                        System.out.println("Processing priority section with " + subItemsSet.size() + " sub-items");
                        System.out.println("Valid sub-items: " + validSubItems);

                        for (String subItem : validSubItems) {
                            String subItemHeader = "<" + priorityIndex + ">";
                            int subStart = colIndex;
                            System.out.println("Priority item " + subItemHeader + " for '" + subItem + "' starts at column " + subStart);

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

                            if (subStart < colIndex - 1) {
                                String subRegionKey = "1:1:" + subStart + ":" + (colIndex - 1);
                                System.out.println("Adding merge region: " + subRegionKey);
                                try {
                                    sheet.addMergedRegion(new CellRangeAddress(1, 1, subStart, colIndex - 1));
                                    mergedRegions.add(subRegionKey);
                                } catch (IllegalArgumentException e) {
                                    System.err.println("Failed to add merge region " + subRegionKey + ": " + e.getMessage());
                                }
                            }
                            subMap.put(subItem, subStart);
                            priorityIndex++;
                        }
                    } else {
                        for (String subItem : subItemsSet) {
                            if (mainItem.contains("ご要望等がございましたらご記入ください。") && subItem.isEmpty()) {
                                Cell cell0 = row0.createCell(colIndex);
                                cell0.setCellValue(mainItem);
                                cell0.setCellStyle(headerStyle);
                                row1.createCell(colIndex).setCellStyle(headerStyle);
                                row2.createCell(colIndex).setCellStyle(headerStyle);
                                String regionKey = "0:2:" + colIndex + ":" + colIndex;
                                if (!mergedRegions.contains(regionKey)) {
                                    sheet.addMergedRegion(new CellRangeAddress(0, 2, colIndex, colIndex));
                                    mergedRegions.add(regionKey);
                                }
                                subMap.put(subItem, colIndex);
                                colIndex++;
                                continue;
                            }

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

                            if (subStart < colIndex - 1) {
                                String subRegionKey = "1:1:" + subStart + ":" + (colIndex - 1);
                                if (!mergedRegions.contains(subRegionKey)) {
                                    sheet.addMergedRegion(new CellRangeAddress(1, 1, subStart, colIndex - 1));
                                    mergedRegions.add(subRegionKey);
                                }
                            }
                            subMap.put(subItem, subStart);
                        }
                    }
                }

                if (mainStart < colIndex && !validSubItems.isEmpty()) {
                    for (int i = mainStart; i < colIndex; i++) {
                        Cell cell0 = row0.getCell(i);
                        if (cell0 == null) {
                            cell0 = row0.createCell(i);
                            cell0.setCellStyle(headerStyle);
                        }
                    }

                    String mainRegionKey = "0:0:" + mainStart + ":" + (colIndex - 1);
                    if (!mergedRegions.contains(mainRegionKey)) {
                        sheet.addMergedRegion(new CellRangeAddress(0, 0, mainStart, colIndex - 1));
                        mergedRegions.add(mainRegionKey);
                    }

                    Cell mainCell = row0.getCell(mainStart);
                    if (mainCell == null) {
                        mainCell = row0.createCell(mainStart);
                    }
                    mainCell.setCellValue(mainItem);
                    mainCell.setCellStyle(headerStyle);
                }

                evalColIndexMap.put(mainItem, subMap);
                System.out.println("evalColIndexMap for " + mainItem + ": " + subMap);
            }

            // Add data rows
            for (Map.Entry<String, Map<String, String>> responseEntry : responses.entrySet()) {
                String uniqueKey = responseEntry.getKey();
                Row dataRow = sheet.createRow(rowNum++);
                dataRow.setHeightInPoints(40);
                int c = 0;
                for (String col : allActiveCols) {
                    String val = col.equals("SheetName") ? sheetName : responseEntry.getValue().get(col);
                    Cell cell = dataRow.createCell(c++);
                    cell.setCellValue(val != null ? val : "");
                    cell.setCellStyle(dataStyle);
                }
                
                Map<String, List<String>> subItemsMap = subItemsByResponse.get(sheetName).get(uniqueKey);
                Map<String, Map<String, Map<String, String>>> evalDataMap = evalDataByResponse.get(sheetName).get(uniqueKey);

                for (String mainItem : allMainItems) {
                    List<String> subItemList = subItemsMap != null ? subItemsMap.getOrDefault(mainItem, new ArrayList<>()) : new ArrayList<>();
                    if (subItemList != null) {
                        boolean isPrioritySection = mainItem.contains("より満足いただくために");
                        if (isPrioritySection) {
                            for (String subItem : subItemList) {
                                if (subItem == null || subItem.trim().isEmpty()) continue;
                                Map<String, String> result = evalDataMap != null ? evalDataMap.getOrDefault(mainItem, new LinkedHashMap<>())
                                        .getOrDefault(subItem, Map.of()) : Map.of();
                                String evaluation = result.getOrDefault("Evaluation", "");
                                String comment = result.getOrDefault("Comment", "");
                                Integer columnIndex = evalColIndexMap.get(mainItem).get(subItem);
                                if (columnIndex != null) {
                                    Cell evalCell = dataRow.createCell(columnIndex);
                                    evalCell.setCellValue(evaluation);
                                    evalCell.setCellStyle(dataStyle);
                                    Cell commentCell = dataRow.createCell(columnIndex + 1);
                                    commentCell.setCellValue(comment);
                                    commentCell.setCellStyle(dataStyle);
                                    System.out.println("Writing priority data for " + subItem + " at column " + columnIndex);
                                }
                            }
                        } else {
                            for (String subItem : subItemList) {
                                Map<String, String> result = evalDataMap != null ? evalDataMap.getOrDefault(mainItem, new LinkedHashMap<>())
                                        .getOrDefault(subItem, Map.of()) : Map.of();
                                String evaluation = result.getOrDefault("Evaluation", "");
                                String comment = result.getOrDefault("Comment", "");

                                if (mainItem.contains("ご要望等がございましたらご記入ください。") && subItem.isEmpty()) {
                                    Integer commentStart = evalColIndexMap.get(mainItem).get(subItem);
                                    if (commentStart != null) {
                                        Cell cell = dataRow.createCell(commentStart);
                                        cell.setCellValue(comment);
                                        cell.setCellStyle(dataStyle);
                                    }
                                    continue;
                                }

                                Integer subStart = evalColIndexMap.get(mainItem).get(subItem);
                                if (subStart != null) {
                                    Cell evalCell = dataRow.createCell(subStart);
                                    evalCell.setCellValue(evaluation);
                                    evalCell.setCellStyle(dataStyle);

                                    Cell commentCell = dataRow.createCell(subStart + 1);
                                    commentCell.setCellValue(comment);
                                    commentCell.setCellStyle(dataStyle);
                                }
                            }
                        }
                    }
                }
            }

            for (int i = 0; i < 100; i++) {
                sheet.autoSizeColumn(i);
            }
        }

        wb.write(outputStream);
        wb.close();
    }
}