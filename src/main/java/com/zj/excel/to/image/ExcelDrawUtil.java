package com.zj.excel.to.image;

import cn.hutool.json.JSONObject;
import cn.hutool.json.JSONUtil;
import com.zj.excel.FileTypeEnum;
import com.zj.excel.domian.dto.RowIndexInfoDTO;
import com.zj.excel.graph.JDrawTableUtil;
import com.zj.excel.graph.domain.JExtendedCell;
import com.zj.excel.graph.domain.JTable;
import com.zj.excel.graph.domain.JTableMergeConfig;
import com.zj.excel.to.image.dto.ExcelDrawImageRequest;
import com.zj.excel.utils.ExcelUtils;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.awt.Color;
import java.awt.Font;
import java.awt.image.BufferedImage;
import java.io.IOException;
import java.io.InputStream;
import java.util.List;
import java.util.*;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.stream.IntStream;

/**
 * @author zhoujun134
 * Created on 2025-09-30
 */
@Slf4j
public class ExcelDrawUtil {

    private final static String COLOR_JSON =
            "{\"0\":\"#000000\",\"1\":\"#FFFFFF\",\"2\":\"#FF0000\",\"3\":\"#00FF00\",\"4\":\"#0000FF\","
                    + "\"5\":\"#FFFF00\",\"6\":\"#FF00FF\",\"7\":\"#00FFFF\",\"8\":\"#000000\",\"9\":\"#FFFFFF\","
                    + "\"10\":\"#FF0000\",\"11\":\"#00FF00\",\"12\":\"#0000FF\",\"13\":\"#FFFF00\","
                    + "\"14\":\"#FF00FF\",\"15\":\"#00FFFF\",\"16\":\"#800000\",\"17\":\"#008000\","
                    + "\"18\":\"#000080\",\"19\":\"#808000\",\"20\":\"#800080\",\"21\":\"#008080\","
                    + "\"22\":\"#C0C0C0\",\"23\":\"#808080\",\"24\":\"#9999FF\",\"25\":\"#993366\","
                    + "\"26\":\"#FFFFCC\",\"27\":\"#CCFFFF\",\"28\":\"#660066\",\"29\":\"#FF8080\","
                    + "\"30\":\"#0066CC\",\"31\":\"#CCCCFF\",\"40\":\"#00CCFF\",\"41\":\"#CCFFFF\","
                    + "\"42\":\"#CCFFCC\",\"43\":\"#FFFF99\",\"44\":\"#99CCFF\",\"45\":\"#FF99CC\","
                    + "\"46\":\"#CC99FF\",\"47\":\"#FFCC99\",\"48\":\"#3366FF\",\"49\":\"#33CCCC\","
                    + "\"50\":\"#99CC00\",\"51\":\"#FFCC00\",\"52\":\"#FF9900\",\"53\":\"#FF6600\","
                    + "\"54\":\"#666699\",\"55\":\"#969696\",\"56\":\"#003366\",\"57\":\"#339966\","
                    + "\"58\":\"#003300\",\"59\":\"#333300\",\"60\":\"#993300\",\"61\":\"#993366\","
                    + "\"62\":\"#333399\",\"63\":\"#333333\",\"64\":\"#FFFFFF\"}";

    private final static JSONObject COLOR_INDEX_RGB_MAP = JSONUtil.parseObj(COLOR_JSON);

    /**
     * 将 Excel 文件转换为 HTML 表格
     *
     * @param request 画图请求体
     * @return 返回 html 表格字符串，key 为 sheetIndex_sheetName，value 为 html 表格字符串
     * @throws IOException 可能会存在的 io 异常
     */
    public static Map<String, BufferedImage> excelToPngWithColor(ExcelDrawImageRequest request)
            throws IOException {
        InputStream fis = request.getExcelStream();
        FileTypeEnum fileTypeEnum = request.getFileTypeEnum();
        int defaultRowLength = request.getDefaultRowLength();
        int defaultColumnLength = request.getDefaultColumnLength();
        boolean needHeader = request.isNeedHeader();
        List<Integer> headerRowIndexList = new ArrayList<>();
        Map<String, BufferedImage> result = new HashMap<>();
        try (Workbook wb = fileTypeEnum == FileTypeEnum.XLS ? new HSSFWorkbook(fis) : new XSSFWorkbook(fis)) {
            int numberOfSheets = wb.getNumberOfSheets();
            FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();
            for (int i = 0; i < numberOfSheets; i++) {
                Sheet sheet = wb.getSheetAt(i);
                String sheetName = sheet.getSheetName();
                log.info("excelToPng::解析表格Sheet-{} {}", i + 1, sheetName);
                if (needHeader) {
                    List<List<Object>> sheetData = ExcelUtils.getSheetData(sheet, evaluator);
                    RowIndexInfoDTO rowIndexDesc = ExcelUtils.extractRowDesc(sheetData);
                    if (Objects.nonNull(rowIndexDesc) && CollectionUtils.isNotEmpty(rowIndexDesc.getHeaderRowIndexList())) {
                        headerRowIndexList = rowIndexDesc.getHeaderRowIndexList();
                    }
                }
                List<BufferedImage> images =
                        convertOneSheetToOnePngTable(sheet, wb, defaultRowLength,
                                defaultColumnLength, headerRowIndexList, needHeader);
                IntStream.range(0, images.size()).forEach(index -> {
                    String imageName =
                            String.format("%s_%s_%s.png", System.currentTimeMillis(), sheet.getSheetName(), index);
                    BufferedImage image = images.get(index);
                    result.put(imageName, image);
                });
            }
        }
        return result;
    }


    /**
     * 将 Excel 文件转换为 HTML 表格
     *
     * @param fis excel 文件流
     * @return 返回 html 表格字符串，key 为 sheetIndex_sheetName，value 为 html 表格字符串
     * @throws IOException 可能会存在的 io 异常
     */
    public static Map<String, BufferedImage> excelToPngWithColor(InputStream fis, FileTypeEnum fileTypeEnum)
            throws IOException {
        Map<String, BufferedImage> result = new HashMap<>();
        boolean needHeader = false;
        try (Workbook wb = fileTypeEnum == FileTypeEnum.XLS ? new HSSFWorkbook(fis) : new XSSFWorkbook(fis)) {
            int numberOfSheets = wb.getNumberOfSheets();
            for (int i = 0; i < numberOfSheets; i++) {
                Sheet sheet = wb.getSheetAt(i);
                String sheetName = sheet.getSheetName();
                log.info("excelToPng::解析表格Sheet-{} {}", i + 1, sheetName);
                List<BufferedImage> images = convertOneSheetToOnePngTable(sheet, wb, 100,
                        10, new ArrayList<>(), needHeader);
                IntStream.range(0, images.size()).forEach(index -> {
                    String imageName =
                            String.format("%s_%s_%s.png", System.currentTimeMillis(), sheet.getSheetName(), index);
                    BufferedImage image = images.get(index);
                    result.put(imageName, image);
                });
            }
        }
        return result;
    }

    private static List<BufferedImage> convertOneSheetToOnePngTable(Sheet sheet, Workbook wb,
                                                                    int defaultRowLength,
                                                                    int defaultColumnLength,
                                                                    List<Integer> headerRowIndexList,
                                                                    boolean needHeader) {
        List<BufferedImage> tableImages = new ArrayList<>();
        List<List<JExtendedCell>> tableRowContents = new ArrayList<>();
        List<JTableMergeConfig> mergeConfigs = new ArrayList<>();
        int lastRow = sheet.getLastRowNum();
        AtomicInteger rowIndex = new AtomicInteger(0);
        List<List<JExtendedCell>> headerRowContents = new ArrayList<>();
        AtomicInteger pageNumber = new AtomicInteger(0);
        for (int r = 0; r <= lastRow; r++) {
            List<JExtendedCell> oneRowContent = new ArrayList<>();
            Row row = sheet.getRow(r);
            if (row == null) {
                // 空行直接补
                tableRowContents.add(oneRowContent);
                rowIndex.incrementAndGet();
                if (headerRowIndexList.contains(r)) {
                    headerRowContents.add(oneRowContent);
                }
                continue;
            }
            int lastCol = row.getLastCellNum();
            for (int c = 0; c < lastCol; c++) {
                Cell cell = row.getCell(c);
                // 获取单元格颜色
                Color backgroundColor = getCellBackgroundColor(cell);
                Color textColor = getCellTextColor(cell, wb);
                /* 1. 如果当前格被合并但不是左上角，跳过 */
                if (getMergedRegion(sheet, r, c) != null &&
                        !isMergedTopLeft(sheet, r, c)) {
                    oneRowContent.add(new JExtendedCell("\n")
                            .setBackgroundColor(backgroundColor)
                            .setTextColor(textColor));
                    continue;
                }
                /* 2. 构造 <td> 属性 */
                CellRangeAddress merged = getMergedRegion(sheet, r, c);
                if (merged != null) {
                    int rs = merged.getLastRow() - merged.getFirstRow() + 1;
                    int cs = merged.getLastColumn() - merged.getFirstColumn() + 1;
                    int mgFirstRow = merged.getFirstRow() % defaultRowLength;
                    int mgLastRow = merged.getLastRow() % defaultRowLength;
                    int mgFirstColumn = merged.getFirstColumn();
                    int mgLastColumn = merged.getLastColumn();
                    if (rs > 1) {
                        JTableMergeConfig mergeConfig =
                                new JTableMergeConfig(mgFirstRow + 1, mgLastRow + 1,
                                        mgFirstColumn + 1, mgLastColumn + 1, false);
                        mergeConfigs.add(mergeConfig);
                    }
                    if (cs > 1) {
                        JTableMergeConfig mergeConfig =
                                new JTableMergeConfig(mgFirstRow + 1, mgLastRow + 1,
                                        mgFirstColumn + 1, mgLastColumn + 1, true);
                        mergeConfigs.add(mergeConfig);
                    }
                }
                /* 3. 单元格内容 */
                String content = getCellText(cell, wb);
                oneRowContent.add(new JExtendedCell(content)
                        .setBackgroundColor(backgroundColor)
                        .setTextColor(textColor));
            }
            tableRowContents.add(oneRowContent);
            if (headerRowIndexList.contains(r)) {
                headerRowContents.add(oneRowContent);
            }
            int curRowIndex = rowIndex.incrementAndGet();
            if (curRowIndex % defaultRowLength == 0) {
                drawImageForCurPage(needHeader, tableImages, tableRowContents,
                        mergeConfigs, headerRowContents, pageNumber);
                tableRowContents = new ArrayList<>();
                mergeConfigs = new ArrayList<>();
                pageNumber.incrementAndGet();
            }
        }
        if (CollectionUtils.isNotEmpty(tableRowContents)) {
            drawImageForCurPage(needHeader, tableImages, tableRowContents,
                    mergeConfigs, headerRowContents, pageNumber);
            pageNumber.incrementAndGet();
        }
        return tableImages;
    }

    private static void drawImageForCurPage(boolean needHeader, List<BufferedImage> tableImages, List<List<JExtendedCell>> tableRowContents, List<JTableMergeConfig> mergeConfigs, List<List<JExtendedCell>> headerRowContents, AtomicInteger pageNumber) {
        JTable tableGraph = new JTable()
                .setCellFont(new Font("宋体", Font.PLAIN, 24))
                .setHeaderFont(new Font("宋体", Font.BOLD, 24))
                .setHeaderBackGroundColor(Color.gray)
                .setMergeConfigs(mergeConfigs)
                .setRowHeight(50);            // 计算表头信息
        if (needHeader && pageNumber.get() != 0 && CollectionUtils.size(headerRowContents) > 0) {
            // 将表头添加到 tableRowContents 中，并重新计算索引
            tableRowContents.addAll(0, headerRowContents);
        }
        BufferedImage curTableImage = JDrawTableUtil.drawTableWithColor(tableGraph, tableRowContents);
        tableImages.add(curTableImage);
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
     * 获取单元格背景色
     */
    private static Color getCellBackgroundColor(Cell cell) {
        if (cell == null) {
            return null;
        }

        CellStyle style = cell.getCellStyle();
        if (style == null) {
            return null;
        }
        if (style instanceof XSSFCellStyle) {
            XSSFCellStyle newCellStyle = (XSSFCellStyle) style;
            XSSFColor backGroundColor = newCellStyle.getFillForegroundColorColor(); // 关键方法
            if (backGroundColor != null) {
                byte[] rgb = backGroundColor.getRGB();
                return new Color(rgb[0] & 0xFF, rgb[1] & 0xFF, rgb[2] & 0xFF);
            }
        }
        if (style instanceof HSSFCellStyle) {
            HSSFCellStyle newCellStyle = (HSSFCellStyle) style;
            HSSFColor backGroundColor = newCellStyle.getFillForegroundColorColor(); // 关键方法
            if (backGroundColor != null) {
                short index = backGroundColor.getIndex();
                short[] triplet = backGroundColor.getTriplet();
                Color color = new Color(triplet[0], triplet[1], triplet[2]);
                if (index == IndexedColors.AUTOMATIC.index) {
                    color = Color.white;
                }
                return color;
            }
        }

        short fillForegroundColor = style.getFillForegroundColor();
        if (fillForegroundColor == IndexedColors.AUTOMATIC.getIndex()) {
            return null; // 无背景色
        }

        // 获取填充颜色
        return getColorFromIndexedColor(style.getFillForegroundColor());
    }

    /**
     * 获取单元格字体颜色
     */
    private static Color getCellTextColor(Cell cell, Workbook wb) {
        if (cell == null) {
            return Color.BLACK;
        }
        CellStyle style = cell.getCellStyle();
        if (style == null) {
            return Color.BLACK;
        }
        if (style instanceof XSSFCellStyle) {
            XSSFCellStyle newCellStyle = (XSSFCellStyle) style;
            XSSFColor fontColor = newCellStyle.getFont().getXSSFColor(); // 关键方法
            if (fontColor != null) {
                byte[] rgb = fontColor.getRGB();
                return new Color(rgb[0] & 0xFF, rgb[1] & 0xFF, rgb[2] & 0xFF);
            }
        }
        if (style instanceof HSSFCellStyle) {
            HSSFWorkbook newWb = (HSSFWorkbook) wb;
            int fontIndex = style.getFontIndex();
            HSSFFont fontAt = newWb.getFontAt(fontIndex);
            short fontColorIndex = fontAt.getColor();
            return getColorFromIndexedColor(fontColorIndex);
        }
        org.apache.poi.ss.usermodel.Font fontAt = cell.getSheet().getWorkbook().getFontAt(style.getFontIndex());
        short colorIndex = fontAt.getColor();
        return getColorFromIndexedColor(colorIndex);
    }

    /**
     * 从索引颜色获取 Color 对象
     */
    private static Color getColorFromIndexedColor(short colorIndex) {
        // 尝试从标准索引颜色获取
        IndexedColors indexedColor = IndexedColors.fromInt(colorIndex);
        if (indexedColor != null) {
            IndexedColors indexedColors = IndexedColors.fromInt(colorIndex);
            if (indexedColors == null) {
                indexedColors = IndexedColors.BLACK;
            }
            String rgbString = convertRgbHex(indexedColors.index);
            if (StringUtils.isBlank(rgbString)) {
                return Color.BLACK;
            }
            return new Color(Integer.parseInt(rgbString.replace("#", ""), 16));
        }
        return Color.BLACK;
    }

    /* 统一提取文本，日期、数字、公式都能转字符串 */
    private static String getCellText(Cell cell, Workbook workbook) {
        FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
        Object data = formatCellValue(cell, evaluator);
        return data == null ? "" : String.valueOf(data);
    }

    private static Object formatCellValue(Cell cell, FormulaEvaluator evaluator) {
        if (cell == null) {
            return "";
        }
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
                } catch (Exception ignored) {

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
                    if (defaultCellValue != null) {
                        return defaultCellValue;
                    }
                    log.error("计算公式出现异常: formatCellValue::error{}", e.getMessage());
                    throw e;
                }
            default:
                return cell.toString();
        }
    }

    public static String convertRgbHex(short idx) {
        String key = String.valueOf(idx & 63);
        if (!COLOR_INDEX_RGB_MAP.containsKey(key)) {
            return "";
        }
        Object string = COLOR_INDEX_RGB_MAP.get(key);
        return String.valueOf(string);
    }
}

