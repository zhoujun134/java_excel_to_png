package com.zj.excel.utils;

import com.zj.excel.FileTypeEnum;
import com.zj.excel.ai.AiInvokeUtils;
import com.zj.excel.domian.dto.RowIndexInfoDTO;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.io.InputStream;
import java.net.HttpURLConnection;
import java.net.URL;
import java.util.*;
import java.util.Map.Entry;
import java.util.concurrent.atomic.AtomicInteger;

@Slf4j
public class ExcelUtils {
    /**
     * 识别二维表基础信息
     *
     * @param sheetData 二维表格数据
     * @return 返回表头信息
     */
    public static RowIndexInfoDTO extractRowDesc(List<List<Object>> sheetData) {
        if (CollectionUtils.isEmpty(sheetData)) {
            return null;
        }
        addRowNumberIndex(sheetData);
        String prompt = "下面的数据是一个二维的表格数据，该数据的第一列为每行数据的行号: \n" +
                "\n" +
                "二维表数据如下: \n" +
                "$rowDataList\n" +
                "\n" +
                "请你识别下面的数据，前几行数据属于表头信息，按照如下的格式返回，{\"headerRowIndexList\":[],\"dataRowStartIndex\":0, otherRowIndexList:[]}\n" +
                "其中: \n" +
                "\theaderRowIndexList: 为表头的行号列表。\n" +
                "\tdataRowStartIndex:  为除开表头，实际的数据开始的行号。\n" +
                "\totherRowIndexList:  为除开表头和数据行的其他行的行号，如果不存在则返回空列表 \n" +
                "要求，请你按照上面的格式返回识别出来的 json 结果。";
        String promptJson = prompt.replace("$rowDataList", GsonUtils.toJSONString(sheetData));
        RowIndexInfoDTO result = AiInvokeUtils.invokeWithJson(promptJson, RowIndexInfoDTO.class);
        log.info("ai 识别表头信息结果为: result:{}", GsonUtils.toJSONString(result));
        if (Objects.nonNull(result) && CollectionUtils.isNotEmpty(result.getHeaderRowIndexList())) {
            List<Integer> headerRowIndexList = result.getHeaderRowIndexList();
            List<Integer> newHeaderRowIndexList = new ArrayList<>();
            headerRowIndexList.forEach(rowNumber -> {
                if (rowNumber > 0) {
                    newHeaderRowIndexList.add(rowNumber - 1);
                } else {
                    newHeaderRowIndexList.add(rowNumber);
                }
            });
            result.setHeaderRowIndexList(newHeaderRowIndexList);
        }
        return result;
    }

    private static void addRowNumberIndex(List<List<Object>> sheetData) {
        if (CollectionUtils.isEmpty(sheetData)) {
            return;
        }
        AtomicInteger indexNumber = new AtomicInteger(1);
        sheetData.forEach(oneRowData -> {
            if (CollectionUtils.isEmpty(oneRowData)) {
                return;
            }
            oneRowData.add(0, indexNumber.getAndIncrement());
        });
    }

    public static InputStream downloadFileByUrl(String url) {
        long s1 = System.currentTimeMillis();
        try {
            log.info("importFile::开始下载文件：{}", url);
            URL httpUrl = new URL(url);
            HttpURLConnection connection = (HttpURLConnection) httpUrl.openConnection();
            connection.setRequestMethod("GET");
            connection.setConnectTimeout(5000);
            connection.setReadTimeout(5000);
            InputStream inputStream = connection.getInputStream();
            log.info("importFile::下载文件完成，耗时：{} url={}", System.currentTimeMillis() - s1, url);
            return inputStream;
        } catch (Exception exception) {
            log.error("downloadFileByUrl::error, url={}", url, exception);
        }
        return null;
    }

    public static Map<String, List<List<Object>>> readExcelToList(String url, FileTypeEnum fileType) {
        InputStream inputStream = null;
        HttpURLConnection connection = null;
        long s1 = System.currentTimeMillis();
        try {
            log.info("importFile::开始下载文件：{}", url);
            URL httpUrl = new URL(url);
            connection = (HttpURLConnection) httpUrl.openConnection();
            connection.setRequestMethod("GET");
            connection.setConnectTimeout(5000);
            connection.setReadTimeout(5000);
            inputStream = connection.getInputStream();
            log.info("importFile::下载文件完成，耗时：{} url={}", System.currentTimeMillis() - s1, url);
            return ExcelUtils.readExcelToList(inputStream, fileType);
        } catch (Exception e) {
            log.error("importFile::error url=" + url, e);
        } finally {
            if (inputStream != null) {
                try {
                    inputStream.close();
                } catch (IOException e) {
                    log.error("文件输入流关闭失败！");
                    throw new RuntimeException(e);
                }
            }
            if (connection != null) {
                connection.disconnect();
            }
        }
        return null;
    }

    /* 返回当前 (r,c) 所在的合并区域，找不到返回 null */
    private static CellRangeAddress getMergedRegion(Sheet sheet, int r, int c) {
        for (CellRangeAddress range : sheet.getMergedRegions()) {
            if (range.isInRange(r, c)) {
                return range;
            }
        }
        return null;
    }

    /* 当前 (r,c) 是否是合并区域的左上角 */
    private static boolean isMergedTopLeft(Sheet sheet, int r, int c) {
        CellRangeAddress range = getMergedRegion(sheet, r, c);
        return range != null && range.getFirstRow() == r && range.getFirstColumn() == c;
    }

    /**
     * 将 Excel 文件转换为 HTML 表格
     *
     * @param fis excel 文件流
     * @return 返回 html 表格字符串，key 为 sheetIndex_sheetName，value 为 html 表格字符串
     * @throws IOException 可能会存在的 io 异常
     */
    public static Map<String, String> excelToHtml(InputStream fis) throws IOException {
        Map<String, String> result = new HashMap<>();
        try (Workbook wb = new XSSFWorkbook(fis)) {
            int numberOfSheets = wb.getNumberOfSheets();
            for (int i = 0; i < numberOfSheets; i++) {
                Sheet sheet = wb.getSheetAt(i);
                String sheetName = sheet.getSheetName();
                log.info("excelToHtml::解析表格Sheet-{} {}", i + 1, sheetName);
                String oneTableString = convertOneSheetToOneHtmlTable(sheet, wb);
                result.put(String.format("%s_%s", i, sheetName), oneTableString);
            }
        }
        return result;
    }

    private static String convertOneSheetToOneHtmlTable(Sheet sheet, Workbook wb) {
        StringBuilder html = new StringBuilder();
        html.append("<table border='1' cellspacing='0' cellpadding='4'>\n");
        int lastRow = sheet.getLastRowNum();
        for (int r = 0; r <= lastRow; r++) {
            Row row = sheet.getRow(r);
            if (row == null) {          // 空行直接补<tr></tr>
                html.append("<tr></tr>\n");
                continue;
            }
            html.append("<tr>");
            int lastCol = row.getLastCellNum();
            for (int c = 0; c < lastCol; c++) {
                /* 1. 如果当前格被合并但不是左上角，跳过 */
                if (getMergedRegion(sheet, r, c) != null &&
                        !isMergedTopLeft(sheet, r, c)) {
                    continue;
                }

                /* 2. 构造 <td> 属性 */
                StringBuilder tdAttr = new StringBuilder();
                CellRangeAddress merged = getMergedRegion(sheet, r, c);
                if (merged != null) {
                    int rs = merged.getLastRow() - merged.getFirstRow() + 1;
                    int cs = merged.getLastColumn() - merged.getFirstColumn() + 1;
                    if (rs > 1) {
                        tdAttr.append(" rowspan='").append(rs).append("'");
                    }
                    if (cs > 1) {
                        tdAttr.append(" colspan='").append(cs).append("'");
                    }
                }

                /* 3. 单元格内容 */
                Cell cell = row.getCell(c);
                String content = getCellText(cell, wb);

                html.append("<td").append(tdAttr).append(">")
                        .append(escapeHtml(content))
                        .append("</td>");
            }
            html.append("</tr>\n");
        }
        html.append("</table>");
        return html.toString();
    }

    /* 统一提取文本，日期、数字、公式都能转字符串 */
    private static String getCellText(Cell cell, Workbook workbook) {
        FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
        Object data = formatCellValue(cell, evaluator);
        return data == null ? "" : String.valueOf(data);
    }

    /* 最简单转义，防止 < > & 破坏 html */
    private static String escapeHtml(String s) {
        if (s == null) {
            return "";
        }
        return s.replace("&", "&amp;")
                .replace("<", "&lt;")
                .replace(">", "&gt;");
    }

    public static Map<String, List<List<Object>>> readExcelToList(InputStream inputStream,
                                                                  FileTypeEnum fileTypeEnum) {
        Map<String, List<List<Object>>> result = new HashMap<>();
        try {
            // 解析 xls
            Workbook workbook;
            if (fileTypeEnum == FileTypeEnum.XLS) {
                workbook = new HSSFWorkbook(inputStream);
            } else {
                // 解析xlsx
                workbook = new XSSFWorkbook(inputStream);
            }
            log.info("importExcelToList::开始解析表格文件 ");
            FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                if (workbook.isSheetHidden(i) || workbook.isSheetVeryHidden(i)) {
                    continue;
                }
                Sheet sheet = workbook.getSheetAt(i);
                log.info("importExcelToList::解析表格Sheet-{} {}", i + 1, sheet.getSheetName());
                buildSheetData(sheet, result, evaluator);
            }
            log.info("importExcelToList::表格文件解析完成 ");
        } catch (Exception e) {
            log.error("importExcelToList::error", e);
        }
        return result;
    }

    public static List<List<Object>> getSheetData(Sheet sheet, FormulaEvaluator evaluator) {
        List<List<Object>> sheetData = new ArrayList<>();
        boolean sheetAllCellsEmpty = true;
        // 获取合并单元格信息
        Map<String, Object> mergedCellsMap = getMergedCells(sheet, evaluator);
        log.info("获取当前合并单元格的信息: {}", GsonUtils.toJSONString(mergedCellsMap));

        for (int rowNum = 0; rowNum <= sheet.getLastRowNum(); rowNum++) {
            try {
                Row row = sheet.getRow(rowNum);
                if (row == null) {
                    continue;
                }
                // 构建行
                List<Object> rowData = new ArrayList<>();
                boolean rowAllCellsEmpty = true;
                for (int cellNum = 0; cellNum < row.getLastCellNum(); cellNum++) {
                    try {
                        Cell cell = row.getCell(cellNum);
                        Object cellData;
                        String cellKey = rowNum + "_" + cellNum;
                        if (cell == null) {
                            cellData = "";
                        } else {
                            cellData = formatCellValue(cell, evaluator);
                        }
                        if (cellData != null && !cellData.toString().isEmpty()) {
                            rowAllCellsEmpty = false;
                            sheetAllCellsEmpty = false;
                        }
                        rowData.add(cellData);
                    } catch (Exception e) {
                        rowData.add("");
                        log.error("buildSheetData::单元格处理异常 error:{}", e.getMessage());
                    }
                }
                if (!rowAllCellsEmpty) {
                    sheetData.add(rowData);
                }
            } catch (Exception e) {
                log.error("buildSheetData::sheet 页处理异常 error:{}", e.getMessage());
            }
        }
        if (!sheetAllCellsEmpty) {
            sheetData = parseSheetData(sheetData);
        }
        return sheetData;
    }

    public static void buildSheetData(Sheet sheet, Map<String, List<List<Object>>> result,
                                      FormulaEvaluator evaluator) {
        List<List<Object>> sheetData = getSheetData(sheet, evaluator);
        result.put(sheet.getSheetName(), sheetData);
    }

    public static List<List<Object>> parseSheetData(List<List<Object>> sheetData) {
        if (sheetData == null || sheetData.isEmpty()) {
            return sheetData;
        }
        // 找到最长的子列表长度
        Map<Integer, Integer> lengthMap = new HashMap<>();
        for (List<Object> row : sheetData) {
            int rowLength = row.size();
            lengthMap.put(rowLength, lengthMap.getOrDefault(rowLength, 0) + 1);
        }
        // 使用 Stream API 找到 value 的最大值
        int maxLength = lengthMap.entrySet().stream().max(Entry.comparingByValue()).map(Entry::getKey).orElse(100);
        if (maxLength >= 100) {
            log.warn("parseSheetData::sheet 页数据异常，最大长度为：{}", maxLength);
            maxLength = 100;
        }

        // 填充空字符串
        List<List<Object>> result = new ArrayList<>();
        for (List<Object> row : sheetData) {
            List<Object> newRow = new ArrayList<>(row);
            while (newRow.size() < maxLength) {
                newRow.add("");
            }
            if (newRow.size() > maxLength) {
                // 截断
                newRow = newRow.subList(0, maxLength);
            }
            result.add(newRow);
        }
        return result;
    }

    private static Map<String, Object> getMergedCells(Sheet sheet, FormulaEvaluator evaluator) {
        int mergedRegionsCount = sheet.getNumMergedRegions();
        Map<String, Object> mergedCellsMap = new HashMap<>();
        for (int regionIndex = 0; regionIndex < mergedRegionsCount; regionIndex++) {
            CellRangeAddress region = sheet.getMergedRegion(regionIndex);
            Cell cell = sheet.getRow(region.getFirstRow()).getCell(region.getFirstColumn());
            Object cellValue = formatCellValue(cell, evaluator);
            for (int row = region.getFirstRow(); row <= region.getLastRow(); row++) {
                for (int col = region.getFirstColumn(); col <= region.getLastColumn(); col++) {
                    mergedCellsMap.put(row + "_" + col, cellValue);
                }
            }
        }
        return mergedCellsMap;
    }

    private static Object formatCellValue(Cell cell, FormulaEvaluator evaluator) {
        DataFormatter formatter = new DataFormatter();
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return formatter.formatCellValue(cell, evaluator);
                } else {
                    return cell.getNumericCellValue();
                }
            case BOOLEAN:
                return cell.getBooleanCellValue();
            case BLANK:
                return "";
            case FORMULA:
                Object defaultCellValue = null;
                try {
                    defaultCellValue = cell.getNumericCellValue();
                } catch (Exception exception) {
                    log.error("获取默认值: formatCellValue::error, {}", exception.getMessage());
                }
                if (Objects.isNull(defaultCellValue) && cell instanceof XSSFCell) {
                    String rawValue = ((XSSFCell) cell).getRawValue();
                    if (StringUtils.isNotBlank(rawValue)) {
                        defaultCellValue = rawValue;
                    }
                }
                try {
                    CellValue cellValue = evaluator.evaluate(cell);
                    switch (cellValue.getCellType()) {
                        case STRING:
                            return cellValue.getStringValue();
                        case NUMERIC:
                            if (DateUtil.isCellDateFormatted(cell)) {
                                return formatter.formatCellValue(cell, evaluator);
                            } else {
                                return cell.getNumericCellValue();
                            }
                        case BOOLEAN:
                            return cellValue.getBooleanValue();
                        default:
                            return "";
                    }
                } catch (Exception e) {
                    log.error("计算公式出现异常: formatCellValue::error{}", e.getMessage());
                    if (defaultCellValue != null) {
                        return defaultCellValue;
                    }
                    throw e;
                }
            default:
                return cell.toString();
        }
    }
}
