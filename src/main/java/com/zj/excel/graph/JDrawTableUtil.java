package com.zj.excel.graph;

import java.awt.BasicStroke;
import java.awt.Color;
import java.awt.Font;
import java.awt.FontMetrics;
import java.awt.Graphics2D;
import java.awt.RenderingHints;
import java.awt.Stroke;
import java.awt.image.BufferedImage;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Collection;
import java.util.Comparator;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.stream.Collectors;
import java.util.stream.IntStream;
import java.util.stream.Stream;

import com.zj.excel.graph.domain.JCell;
import com.zj.excel.graph.domain.JExtendedCell;
import com.zj.excel.graph.domain.JTable;
import com.zj.excel.graph.domain.JTableMergeConfig;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.tuple.ImmutableTriple;
import org.apache.commons.lang3.tuple.Pair;

import lombok.SneakyThrows;
import lombok.extern.slf4j.Slf4j;

/**
 * JAVA 绘制表格
 *
 * @author sk
 * @date 2022/7/17 22:41
 */
@Slf4j
public class JDrawTableUtil {
    @SneakyThrows
    public static BufferedImage drawTableWithColor(JTable table, List<List<JExtendedCell>> tableRowContents) {
        checkHead(table);

        List<JTableMergeConfig> mergeConfigs = table.getMergeConfigs();

        // 生成表格各单元格内容对象
        List<JCell> contents = getTableContentWithColors(tableRowContents, table);

        // 合并单元格
        List<JCell> cells = mergeCells(contents, mergeConfigs, new ArrayList<>());

        // 计算表格高度
        Map<Integer, Integer> contentRowHeight = getContentRowHeightFromExtended(tableRowContents, table);
        int tableHeight = contentRowHeight.values().stream().reduce(Integer::sum).orElse(0);

        // 绘制表格
        return starDrawTableWithColors(0, cells, table, tableHeight);
    }
    /**
     * 计算带扩展内容的列宽
     */
    private static List<Integer> calculateColumnWidthsWithExtended(List<List<JExtendedCell>> tableRowContents, JTable table) {
        List<Integer> widths = new ArrayList<>();
        if (tableRowContents.isEmpty()) {
            return widths;
        }

        // 计算最大列数
        int maxCols = 0;
        for (List<JExtendedCell> row : tableRowContents) {
            maxCols = Math.max(maxCols, row.size());
        }

        // 为每一列计算宽度
        for (int col = 0; col < maxCols; col++) {
            int maxWidth = 100; // 默认最小宽度
            for (List<JExtendedCell> row : tableRowContents) {
                if (col < row.size()) {
                    String content = row.get(col).getContent();
                    // 计算内容宽度
                    int contentWidth = calculateContentWidth(content, table.getCellFont());
                    maxWidth = Math.max(maxWidth, contentWidth);
                }
            }
            widths.add(Math.max(100, maxWidth + 20)); // 添加一些边距
        }
        return widths;
    }
    /**
     * 生成带颜色的表格内容
     */
    public static List<JCell> getTableContentWithColors(List<List<JExtendedCell>> tableRowContents, JTable table) {
        List<JCell> contents = new ArrayList<>();
        int marginX = table.getMarginX();
        int marginY = table.getMarginY();

        if (tableRowContents.isEmpty()) {
            return contents;
        }

        // 计算每列的宽度（根据内容长度）
        List<Integer> columnWidths = calculateColumnWidthsWithExtended(tableRowContents, table);

        HashMap<Integer, Integer> contentRowHeight = getContentRowHeightFromExtended(tableRowContents, table);

        for (int i = 0; i < tableRowContents.size(); i++) {
            List<JExtendedCell> rowContent = tableRowContents.get(i);
            int currentX = marginX;
            int currentY = marginY + IntStream.rangeClosed(1, i).map(contentRowHeight::get).sum();

            for (int j = 0; j < rowContent.size(); j++) {
                JExtendedCell extendedCell = rowContent.get(j);
                JCell cell = new JCell();
                cell.setRow(i + 1);
                cell.setColumn(j + 1);
                cell.setX(currentX);
                cell.setY(currentY);
                cell.setWidth(columnWidths.get(j));
                cell.setTextAlign(true);
                cell.setHeight(contentRowHeight.get(i + 1));
                cell.setContent(extendedCell.getContent());
                cell.setBackgroundColor(extendedCell.getBackgroundColor());
                cell.setTextColor(extendedCell.getTextColor());
                contents.add(cell);

                currentX += columnWidths.get(j);
            }
        }
        return contents;
    }

    /**
     * 从扩展内容计算行高
     */
    private static HashMap<Integer, Integer> getContentRowHeightFromExtended(List<List<JExtendedCell>> tableRowContents, JTable table) {
        return IntStream.rangeClosed(1, tableRowContents.size())
                .mapToObj(row -> dealRowHeightFromExtended(tableRowContents, table, row))
                .reduce(new HashMap<>(16), (a, b) -> {
                    a.put(b.getLeft(), b.getRight());
                    return a;
                }, (c, d) -> c);
    }

    /**
     * 计算扩展内容的行高
     */
    private static Pair<Integer, Integer> dealRowHeightFromExtended(List<List<JExtendedCell>> tableRowContents, JTable table, int row) {
        Integer cellRows = tableRowContents.get(row - 1).stream()
                .map(cell -> StringUtils.countMatches(cell.getContent(), "\n")).max(Integer::compare).orElse(1);
        int appendHeight = table.getCellFont().getSize() * cellRows;
        return Pair.of(row, table.getRowHeight() + appendHeight);
    }

    /**
     * 绘制带颜色的表格
     */
    public static BufferedImage starDrawTableWithColors(int headRow, List<JCell> tableCells, JTable table, int tableHeight)
            throws IOException {
        if (tableCells.isEmpty()) {
            return new BufferedImage(400, 200, BufferedImage.TYPE_INT_RGB);
        }

        int marginY = table.getMarginY();
        int marginX = table.getMarginX();
        Map<Integer, List<JCell>> allTableRows = tableCells.stream().collect(Collectors.groupingBy(JCell::getRow));
        int allRows = allTableRows.size();

        if (allRows == 0) {
            return new BufferedImage(400, 200, BufferedImage.TYPE_INT_RGB);
        }

        // 画布高度
        int imageHeight = tableHeight + marginY * 2;

        // 计算画布宽度：取所有单元格的最大X+Width
        int tableWidth = tableCells.stream()
                .mapToInt(cell -> cell.getX() + cell.getWidth())
                .max()
                .orElse(400);
        int imageWidth = tableWidth + marginX * 2;

        BufferedImage image = new BufferedImage(imageWidth, imageHeight, BufferedImage.TYPE_INT_RGB);
        Graphics2D graphics = image.createGraphics();
        graphics.setColor(Color.white);
        // 默认白色背景
        graphics.fillRect(0, 0, imageWidth, imageHeight);

        // 绘制所有单元格
        for (int i = 1; i <= allRows; i++) {
            List<JCell> perRowCells = allTableRows.get(i);
            if (perRowCells == null) continue;

            for (int j = 0; j < perRowCells.size(); j++) {
                JCell cell = perRowCells.get(j);
                // 绘制单元格背景色
                drawCellBackground(graphics, cell);
                // 当前单元格是否为本行最后一个
                boolean lastCellInRow = j == perRowCells.size() - 1;
                coreMethodV2(graphics, table.getCellFont(), lastCellInRow, cell);
            }
        }

        graphics.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
        graphics.setRenderingHint(RenderingHints.KEY_STROKE_CONTROL, RenderingHints.VALUE_STROKE_DEFAULT);
        graphics.setRenderingHint(RenderingHints.KEY_TEXT_ANTIALIASING, RenderingHints.VALUE_TEXT_ANTIALIAS_ON);
        // 设置高质量渲染提示
        graphics.setRenderingHint(RenderingHints.KEY_INTERPOLATION, RenderingHints.VALUE_INTERPOLATION_BICUBIC);
        graphics.setRenderingHint(RenderingHints.KEY_RENDERING, RenderingHints.VALUE_RENDER_QUALITY);
        Stroke basicStroke = new BasicStroke(1);
        graphics.setStroke(basicStroke);

        graphics.dispose();
        return image;
    }

    /**
     * 绘制单元格背景色
     */
    private static void drawCellBackground(Graphics2D graphics, JCell cell) {
        Color bgColor = cell.getBackgroundColor();
        if (bgColor != null) {
            graphics.setColor(bgColor);
        } else {
            graphics.setColor(Color.WHITE); // 默认白色背景
        }
        graphics.fillRect(cell.getX(), cell.getY(), cell.getWidth(), cell.getHeight());
    }

    /**
     * 绘制核心方法
     */
    private static void coreMethodV2(Graphics2D graphics, Font font, boolean lastCellInRow, JCell cell) {
        // 绘制单元格边框
        graphics.setColor(Color.black);
        graphics.drawRect(cell.getX(), cell.getY(), cell.getWidth(), cell.getHeight());

        // 设置字体颜色
        Color textColor = cell.getTextColor();
        if (textColor != null) {
            graphics.setColor(textColor);
        } else {
            graphics.setColor(Color.BLACK); // 默认黑色
        }

        graphics.setFont(font);
        String content = StringUtils.defaultIfBlank(cell.getContent(), "-");
        String[] split = StringUtils.splitPreserveAllTokens(content, "\n");
        for (int i = 0; i < split.length; i++) {
            //1.计算单元格内容的横坐标，加1是为了防止文字紧贴在单元格上
            int contentX = cell.getX() + 1;
            String cellRowContent = split[i];
            if (cell.isTextAlign()) {
                // 初始执行比较耗时
                FontMetrics fontMetrics = graphics.getFontMetrics();
                int contentLen = fontMetrics.stringWidth(cellRowContent);
                contentX += (cell.getWidth() - contentLen) / 2;
            }
            // 2. 计算单元格纵坐标,默认居中(!!!注意：内容是从下向上从左向右渲染,所在在单元格的基础上又加了字体font.getSize())
            int startY = cell.getY() + font.getSize();
            // 单元格内第一行文字的纵坐标
            int cellFirstRowPosition = (cell.getHeight() - font.getSize() * (split.length)) / 2;
            // 偏移量（在行一行文字纵坐标的基础上进行累加），加1是为了防止每行文字粘在一起。
            int offset = (font.getSize() + 1) * i;
            int contentY = startY + cellFirstRowPosition + offset;
            // 3. 写入单元格内容
            graphics.drawString(cellRowContent, contentX, contentY);
        }
    }

    /**
     * 绘制表格
     *
     * @param table 表格属性
     * @param tableRowContents 表格内容
     */
    @SneakyThrows
    public static BufferedImage drawTable(JTable table, List<List<String>> tableRowContents) {
        checkHead(table);

        List<JCell> headCells = table.getHeadCells();
        List<JTableMergeConfig> mergeConfigs = table.getMergeConfigs();

        // 如果没有表头，则直接处理数据
        if (CollectionUtils.isEmpty(headCells)) {
            // 生成表格各单元格内容对象（作为表头）
            List<JCell> contents = getTableContentWithoutHeader(tableRowContents, table);
            // 合并单元格
            List<JCell> cells = mergeCells(contents, mergeConfigs, new ArrayList<>());
            // 计算表格高度
            HashMap<Integer, Integer> contentRowHeight = getContentRowHeight(tableRowContents, table);
            int tableHeight = contentRowHeight.values().stream().reduce(Integer::sum).orElse(0);
            // 绘制表格
            return starDrawTable(0, cells, table, tableHeight);
        }

        // 有表头的情况
        Map<Integer, List<JCell>> rows = headCells.stream().collect(Collectors.groupingBy(JCell::getRow));
        //表头的最后一行实际有多少个单元格，有合并单元格的情况下按垂直投影的方式获取列；
        List<JCell> actualLastHeadRowColumnCell = findHeadColumns(headCells);
        HashMap<Integer, Integer> headRealRowHead = getHeadRowHeight(rows, table);
        HashMap<Integer, Integer> contentRowHeight = getContentRowHeight(tableRowContents, table);
        //处理表头
        dealTableHead(rows, table, actualLastHeadRowColumnCell, headCells, headRealRowHead);
        // 生成表格各单元格内容对象
        List<JCell> contents =
                getTableContent(tableRowContents, actualLastHeadRowColumnCell, rows.size(), contentRowHeight);
        // 合并单元格
        List<JCell> cells = mergeCells(contents, mergeConfigs, headCells);
        // 计算表格高度
        Integer tableHeight = Stream.of(headRealRowHead.values(), contentRowHeight.values())
                .flatMap(Collection::stream)
                .reduce(Integer::sum)
                .orElseThrow(() -> new RuntimeException("表格高度错误"));
        // 绘制表格
        return starDrawTable(rows.size(), cells, table, tableHeight);
    }

    /**
     * 没有表头时生成表格内容
     */
    public static List<JCell> getTableContentWithoutHeader(List<List<String>> tableRowContents, JTable table) {
        List<JCell> contents = new ArrayList<>();
        int marginX = table.getMarginX();
        int marginY = table.getMarginY();

        if (tableRowContents.isEmpty()) {
            return contents;
        }

        // 计算每列的宽度（根据内容长度）
        List<Integer> columnWidths = calculateColumnWidths(tableRowContents, table);

        HashMap<Integer, Integer> contentRowHeight = getContentRowHeight(tableRowContents, table);

        for (int i = 0; i < tableRowContents.size(); i++) {
            List<String> rowContent = tableRowContents.get(i);
            int currentX = marginX;
            int currentY = marginY + IntStream.rangeClosed(1, i).map(contentRowHeight::get).sum();

            for (int j = 0; j < rowContent.size(); j++) {
                String cellContent = rowContent.get(j);
                JCell cell = new JCell();
                cell.setRow(i + 1);
                cell.setColumn(j + 1);
                cell.setX(currentX);
                cell.setY(currentY);
                cell.setWidth(columnWidths.get(j));
                cell.setTextAlign(true);
                cell.setHeight(contentRowHeight.get(i + 1));
                cell.setContent(cellContent);
                contents.add(cell);

                currentX += columnWidths.get(j);
            }
        }
        return contents;
    }

    /**
     * 计算列宽
     */
    private static List<Integer> calculateColumnWidths(List<List<String>> tableRowContents, JTable table) {
        List<Integer> widths = new ArrayList<>();
        if (tableRowContents.isEmpty()) {
            return widths;
        }

        // 计算最大列数
        int maxCols = 0;
        for (List<String> row : tableRowContents) {
            maxCols = Math.max(maxCols, row.size());
        }

        // 为每一列计算宽度
        for (int col = 0; col < maxCols; col++) {
            int maxWidth = 100; // 默认最小宽度
            for (List<String> row : tableRowContents) {
                if (col < row.size()) {
                    String content = row.get(col);
                    // 计算内容宽度
                    int contentWidth = calculateContentWidth(content, table.getCellFont());
                    maxWidth = Math.max(maxWidth, contentWidth);
                }
            }
            widths.add(Math.max(100, maxWidth + 20)); // 添加一些边距
        }
        return widths;
    }

    /**
     * 计算内容宽度
     */
    private static int calculateContentWidth(String content, Font font) {
        BufferedImage tempImage = new BufferedImage(1, 1, BufferedImage.TYPE_INT_ARGB);
        Graphics2D g2d = tempImage.createGraphics();
        g2d.setFont(font);
        FontMetrics fm = g2d.getFontMetrics();
        String[] lines = content.split("\n");
        int maxWidth = 0;
        for (String line : lines) {
            int width = fm.stringWidth(line);
            maxWidth = Math.max(maxWidth, width);
        }
        g2d.dispose();
        tempImage.flush();
        return maxWidth;
    }

    /**
     * 表头格式校验
     *
     * @param table 表头参数信息
     */
    public static void checkHead(JTable table) {
        if (Objects.isNull(table.getHeaderFont())) {
            table.setHeaderFont(new Font("楷体", Font.BOLD, 15));
        }
        if (Objects.isNull(table.getCellFont())) {
            table.setCellFont(new Font("宋体", Font.PLAIN, 12));
        }
        if (Objects.isNull(table.getHeaderBackGroundColor())) {
            table.setHeaderBackGroundColor(Color.gray);
        }
        if (table.getRowHeight() <= 30) {
            table.setRowHeight(30);
        }
        if (table.getMarginX() < 0 || table.getMarginY() < 0) {
            throw new RuntimeException("表格边距不可小于0");
        }

        // 如果有表头，则校验表头
        if (!CollectionUtils.isEmpty(table.getHeadCells())) {
            table.getHeadCells().forEach(head -> {
                if (head.getRow() <= 0) {
                    throw new RuntimeException("行不可小于0" + head.getRow() + "," + head.getColumn());
                }
                if (head.getColumn() <= 0) {
                    throw new RuntimeException("列不可小于0" + head.getRow() + "," + head.getColumn());
                }
                if (head.getBelongColumn() < 0) {
                    throw new RuntimeException("从属单元格列不可小于0" + head.getRow() + "," + head.getColumn());
                }
                if (head.getWidth() < 0) {
                    throw new RuntimeException("列宽不可小于0" + head.getRow() + "," + head.getColumn());
                }
            });
            checkBelongs(table.getHeadCells());
            ImmutableTriple<ArrayList<Integer>, ArrayList<Integer>, ArrayList<Integer>> triple =
                    table.getHeadCells().stream()
                            .reduce(ImmutableTriple.of(new ArrayList<>(), new ArrayList<>(), new ArrayList<>()),
                                    (a, b) -> {
                                        a.getLeft().add(b.getRow());
                                        a.getMiddle().add(b.getColumn());
                                        a.getRight().add(b.getBelongColumn());
                                        return a;
                                    }, (a, b) -> a);
            checkConsistency(triple.getLeft());
            checkConsistency(triple.getMiddle());
            table.getHeadCells().stream().map(JCell::getBelongColumn).distinct()
                    .filter(column -> column > 0)
                    .filter(column -> !triple.getMiddle().contains(column))
                    .findAny().ifPresent(a -> {
                        throw new RuntimeException("从属单元格" + a + "列不存在");
                    });
        }
    }

    /**
     * 表格连续性校验
     *
     * @param data 行或列序列
     */
    private static void checkConsistency(ArrayList<Integer> data) {
        List<Integer> collect = data.stream().distinct().sorted().collect(Collectors.toList());
        for (int i = 0; i < collect.size() - 1; i++) {
            Integer integer = collect.get(i);
            if (!collect.get(i + 1).equals(integer + 1)) {
                throw new RuntimeException("行或者列不连续");
            }
        }
    }

    /**
     * 从属单元格校验
     *
     * @param headCells 表头
     */
    public static void checkBelongs(List<JCell> headCells) {
        Map<Integer, List<JCell>> collect = headCells.stream().collect(Collectors.groupingBy(JCell::getRow));
        collect.forEach((key, value) -> {
            List<JCell> rowCells =
                    value.stream().sorted(Comparator.comparing(JCell::getColumn)).collect(Collectors.toList());
            int belongColumn = 0;
            for (JCell cell : rowCells) {
                if (cell.getBelongColumn() >= belongColumn) {
                    belongColumn = cell.getBelongColumn();
                    continue;
                }
                throw new RuntimeException("从属单元格配置错误" + cell.getRow() + "," + cell.getColumn());
            }
        });
    }

    /**
     * 处理表头
     *
     * @param rows 表头行内容
     * @param table 行高
     * @param actualLastHeadRowColumnCell 表头单元格
     * @param headCells 表头单元格
     */
    public static void dealTableHead(Map<Integer, List<JCell>> rows, JTable table, List<JCell> actualLastHeadRowColumnCell,
                                     List<JCell> headCells,
                                     HashMap<Integer, Integer> headRealRowHead) {
        int marginX = table.getMarginX();
        int marginY = table.getMarginY();
        rows.keySet().stream().sorted().forEach(row -> {
            List<JCell> perRowCells =
                    rows.get(row).stream().sorted(Comparator.comparing(JCell::getColumn)).collect(Collectors.toList());
            int startFrom = marginX;
            // 非第一行的情况，有可能起始列不是第一列，这里需要取前几列的宽度作为起始宽度
            if (row > 1 && perRowCells.get(0).getColumn() > 1) {
                int column = perRowCells.get(0).getColumn();
                int offset =
                        IntStream.range(1, column).map(index -> actualLastHeadRowColumnCell.get(index - 1).getWidth())
                                .sum();
                startFrom = marginX + offset;
            }
            for (JCell cell : perRowCells) {
                // 获取从属单元格,单元格的宽度由从属单元格确定
                List<JCell> attached = getCellSonCells(headCells, cell);
                int cellWidth = attached.stream().mapToInt(JCell::getWidth).sum();
                cell.setHeight(headRealRowHead.get(row));
                // 如果该单元格有从属，则高度为行高，反之为单元格所跨行数*行高
                // attached
                //  长度为1：当不是自己时才说明该单元格是他的从属
                //  长度大于1：有从属
                if (attached.size() == 1 && attached.get(0).equals(cell)) {
                    int sum = IntStream.rangeClosed(cell.getRow(), rows.size()).map(headRealRowHead::get).sum();
                    cell.setHeight(sum);
                }
                if (cellWidth <= 0) {
                    throw new RuntimeException("行:" + cell.getRow() + " 列:" + cell.getColumn() + "宽度必须大于0");
                }
                cell.setWidth(cellWidth);
                cell.setX(startFrom);
                startFrom += cellWidth;
                int sum = IntStream.rangeClosed(1, row - 1).map(headRealRowHead::get).sum();
                cell.setY(sum + marginY);
            }
        });
    }

    /**
     * 计算表头行高
     *
     * @param rows 表头行
     * @param table 表属性
     * @return 表头行与行高的对应 key:行 value:行高
     */
    private static HashMap<Integer, Integer> getHeadRowHeight(Map<Integer, List<JCell>> rows, JTable table) {
        return rows.keySet().stream().map(row1 -> {
            Integer cellRow = rows.get(row1).stream().map(cell -> StringUtils.countMatches(cell.getContent(), "\n"))
                    .max(Integer::compare).orElse(1);
            int appendHeight = cellRow * table.getHeaderFont().getSize();
            return Pair.of(row1, table.getRowHeight() + appendHeight);
        }).reduce(new HashMap<>(16), (a, b) -> {
            a.put(b.getLeft(), b.getRight());
            return a;
        }, (c, d) -> c);
    }

    /**
     * 生成表格各单元格内容对象
     *
     * @param tableRowContents 表格内容
     * @param actualLastHeadRowColumnCell 表头单元格
     * @param headRowSize 表头行数
     * @param contentRowHeight 各行的行高关系
     * @return 表格各单元格内容对象
     */
    public static List<JCell> getTableContent(List<List<String>> tableRowContents,
                                             List<JCell> actualLastHeadRowColumnCell,
                                             int headRowSize, HashMap<Integer, Integer> contentRowHeight) {
        List<JCell> contents = new ArrayList<>();
        for (int i = 0; i < tableRowContents.size(); i++) {
            List<String> rowContent = tableRowContents.get(i);
            for (int j = 0; j < rowContent.size(); j++) {
                String cellContent = rowContent.get(j);
                JCell cell = new JCell();
                JCell lastHeadColumnCell = actualLastHeadRowColumnCell.get(j);
                cell.setRow(i + 1 + headRowSize);
                cell.setColumn(lastHeadColumnCell.getColumn());
                cell.setX(lastHeadColumnCell.getX());
                // 单元格纵坐标
                int sum = IntStream.rangeClosed(1, i).map(contentRowHeight::get).sum();
                int y = (lastHeadColumnCell.getY() + lastHeadColumnCell.getHeight()) + sum;
                cell.setY(y);
                cell.setWidth(lastHeadColumnCell.getWidth());
                cell.setTextAlign(true);
                cell.setHeight(contentRowHeight.get(i + 1));
                cell.setContent(cellContent);
                contents.add(cell);
            }
        }
        return contents;
    }

    /**
     * 计算单元格高度
     *
     * @param tableRowContents 表格内容
     * @param table 表格属性
     * @return 行与行高对应关系  key:行 value:行高
     */
    private static HashMap<Integer, Integer> getContentRowHeight(List<List<String>> tableRowContents, JTable table) {
        return IntStream.rangeClosed(1, tableRowContents.size())
                .mapToObj(row -> dealRowHeight(tableRowContents, table, row))
                .reduce(new HashMap<>(16), (a, b) -> {
                    a.put(b.getLeft(), b.getRight());
                    return a;
                }, (c, d) -> c);
    }

    /**
     * 计算单元格高度主方法
     *
     * @param tableRowContents 表格内容
     * @param table 表格属性
     * @param row 行内容
     * @return 行与行高的对应关系  key:行 value:行高
     */
    private static Pair<Integer, Integer> dealRowHeight(List<List<String>> tableRowContents, JTable table, int row) {
        Integer cellRows = tableRowContents.get(row - 1).stream()
                .map(content -> StringUtils.countMatches(content, "\n")).max(Integer::compare).orElse(1);
        int appendHeight = table.getCellFont().getSize() * cellRows;
        return Pair.of(row, table.getRowHeight() + appendHeight);
    }

    /**
     * 绘制表格
     *
     * @param tableCells 表格所有单元格
     * @param table 表格属性
     */
    public static BufferedImage starDrawTable(int headRow, List<JCell> tableCells, JTable table, int tableHeight)
            throws IOException {
        if (tableCells.isEmpty()) {
            return new BufferedImage(400, 200, BufferedImage.TYPE_INT_RGB);
        }

        int marginY = table.getMarginY();
        int marginX = table.getMarginX();
        Map<Integer, List<JCell>> allTableRows = tableCells.stream().collect(Collectors.groupingBy(JCell::getRow));
        int allRows = allTableRows.size();

        if (allRows == 0) {
            return new BufferedImage(400, 200, BufferedImage.TYPE_INT_RGB);
        }

        // 画布高度
        int imageHeight = tableHeight + marginY * 2;

        // 计算画布宽度：取所有单元格的最大X+Width
        int tableWidth = tableCells.stream()
                .mapToInt(cell -> cell.getX() + cell.getWidth())
                .max()
                .orElse(400);
        int imageWidth = tableWidth + marginX * 2;

        BufferedImage image = new BufferedImage(imageWidth, imageHeight, BufferedImage.TYPE_INT_RGB);
        Graphics2D graphics = image.createGraphics();
        graphics.setColor(Color.white);
        // 默认白色背景
        graphics.fillRect(0, 0, imageWidth, imageHeight);

        // 绘制所有单元格
        for (int i = 1; i <= allRows; i++) {
            List<JCell> perRowCells = allTableRows.get(i);
            if (perRowCells == null) {
                continue;
            }

            if (i <= headRow) {
                for (int h = 0; h < perRowCells.size(); h++) {
                    JCell cell = perRowCells.get(h);
                    //表头单元格填充背景色
                    graphics.setColor(table.getHeaderBackGroundColor());
                    graphics.fillRect(cell.getX(), cell.getY(), cell.getWidth(), cell.getHeight());
                    // 当前单元格是否为本行最后一个
                    boolean lastCellInRow = h == perRowCells.size() - 1;
                    coreMethod(graphics, table.getHeaderFont(), lastCellInRow, cell);
                }
            } else {
                for (int j = 0; j < perRowCells.size(); j++) {
                    JCell cell = perRowCells.get(j);
                    // 当前单元格是否为本行最后一个
                    boolean lastCellInRow = j == perRowCells.size() - 1;
                    coreMethod(graphics, table.getCellFont(), lastCellInRow, cell);
                }
            }
        }

        graphics.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
        graphics.setRenderingHint(RenderingHints.KEY_STROKE_CONTROL, RenderingHints.VALUE_STROKE_DEFAULT);
        graphics.setRenderingHint(RenderingHints.KEY_TEXT_ANTIALIASING, RenderingHints.VALUE_TEXT_ANTIALIAS_ON);
        // 设置高质量渲染提示
        graphics.setRenderingHint(RenderingHints.KEY_INTERPOLATION, RenderingHints.VALUE_INTERPOLATION_BICUBIC);
        graphics.setRenderingHint(RenderingHints.KEY_RENDERING, RenderingHints.VALUE_RENDER_QUALITY);
        Stroke basicStroke = new BasicStroke(1);
        graphics.setStroke(basicStroke);

        graphics.dispose();
        return convertImageByDpi(image, table.getImageDpi());
    }

    public static BufferedImage convertImageByDpi(BufferedImage inputImage, int dpi) {
        if (dpi <= 0 || dpi == 72) {
            log.info("dpi <= 0 or dpi = 72, use default dpi 72 ");
            return inputImage;
        }
        int width = inputImage.getWidth();
        int height = inputImage.getHeight();

        // 计算新的宽度和高度
        int newWidth = width * dpi / 72;
        int newHeight = height * dpi / 72;

        // 创建一个新的BufferedImage对象
        BufferedImage outputImage = new BufferedImage(newWidth, newHeight, BufferedImage.TYPE_INT_RGB);

        // 将输入图像绘制到新的BufferedImage对象中
        Graphics2D g2d = outputImage.createGraphics();
        g2d.drawImage(inputImage, 0, 0, newWidth, newHeight, null);
        g2d.dispose();
        return outputImage;
    }

    /**
     * 绘制核心方法
     *
     * @param graphics 绘制类
     * @param font 字体信息
     * @param lastCellInRow 是否为行内最后一个单元格
     * @param cell 要绘制的单元格
     */
    private static void coreMethod(Graphics2D graphics, Font font, boolean lastCellInRow, JCell cell) {
        graphics.setColor(Color.red);
        //绘制单元格边框
        graphics.drawRect(cell.getX(), cell.getY(), cell.getWidth(), cell.getHeight());

        //单元格内容
        // 如果是水平居中,根据"字体"计算出实际起始坐标位置
        graphics.setFont(font);
        String content = StringUtils.defaultIfBlank(cell.getContent(), "-");
        String[] split = StringUtils.splitPreserveAllTokens(content, "\n");
        for (int i = 0; i < split.length; i++) {
            //1.计算单元格内容的横坐标，加1是为了防止文字紧贴在单元格上
            int contentX = cell.getX() + 1;
            String cellRowContent = split[i];
            if (cell.isTextAlign()) {
                // 初始执行比较耗时
                FontMetrics fontMetrics = graphics.getFontMetrics();
                int contentLen = fontMetrics.stringWidth(cellRowContent);
                contentX += (cell.getWidth() - contentLen) / 2;
            }
            // 2. 计算单元格纵坐标,默认居中(!!!注意：内容是从下向上从左向右渲染,所在在单元格的基础上又加了字体font.getSize())
            int startY = cell.getY() + font.getSize();
            // 单元格内第一行文字的纵坐标
            int cellFirstRowPosition = (cell.getHeight() - font.getSize() * (split.length)) / 2;
            // 偏移量（在行一行文字纵坐标的基础上进行累加），加1是为了防止每行文字粘在一起。
            int offset = (font.getSize() + 1) * i;
            int contentY = startY + cellFirstRowPosition + offset;
            // 3. 写入单元格内容
            graphics.drawString(cellRowContent, contentX, contentY);
        }
    }

    /**
     * 按列合并单元格
     *
     * @param contents 单元格内容
     * @param mergeConfigs 要合并的配置
     * @param headCells 表头单元格
     */
    public static List<JCell> mergeCells(List<JCell> contents, List<JTableMergeConfig> mergeConfigs,
                                        List<JCell> headCells) {
        if (CollectionUtils.isEmpty(mergeConfigs)) {
            contents.addAll(headCells);
            return contents;
        }

        List<JCell> allCells = new ArrayList<>(contents);
        allCells.addAll(headCells);

        if (CollectionUtils.isEmpty(contents)) {
            contents.addAll(headCells);
            return contents;
        }

        // 按配置进行合并
        for (JTableMergeConfig config : mergeConfigs) {
            allCells = performMerge(allCells, config);
        }

        return allCells;
    }

    /**
     * 执行具体的合并操作
     */
    private static List<JCell> performMerge(List<JCell> allCells, JTableMergeConfig config) {
        List<JCell> result = new ArrayList<>(allCells);

        if (config.isHorizontal()) {
            // 横向合并
            result = mergeHorizontally(allCells, config);
        } else {
            // 竖向合并
            result = mergeVertically(allCells, config);
        }

        return result;
    }

    /**
     * 横向合并单元格
     */
    private static List<JCell> mergeHorizontally(List<JCell> allCells, JTableMergeConfig config) {
        List<JCell> result = new ArrayList<>(allCells);

        // 找到要合并的单元格
        List<JCell> cellsToMerge = allCells.stream()
                .filter(cell -> cell.getRow() == config.getStartRow())
                .filter(cell -> cell.getColumn() >= config.getStartCol() && cell.getColumn() <= config.getEndCol())
                .sorted(Comparator.comparingInt(JCell::getColumn))
                .collect(Collectors.toList());

        if (cellsToMerge.size() < 2) {
            return result;
        }

        List<String> contentList = cellsToMerge.stream()
                .map(JCell::getContent)
                .filter(StringUtils::isNotBlank)
                .collect(Collectors.toList());
        String content = String.join("\n", contentList);
        // 创建合并后的单元格
        JCell mergedCell = new JCell();
        mergedCell.setRow(config.getStartRow());
        mergedCell.setColumn(config.getStartCol());
        mergedCell.setX(cellsToMerge.get(0).getX());
        mergedCell.setY(cellsToMerge.get(0).getY());
        mergedCell.setHeight(cellsToMerge.get(0).getHeight());
        mergedCell.setWidth(cellsToMerge.stream().mapToInt(JCell::getWidth).sum());
        mergedCell.setContent(content); // 使用第一个单元格的内容
        mergedCell.setTextAlign(cellsToMerge.get(0).isTextAlign());
        mergedCell.setBackgroundColor(cellsToMerge.get(0).getBackgroundColor());
        mergedCell.setTextColor(cellsToMerge.get(0).getTextColor());

        // 移除被合并的单元格
        List<JCell> cellsToRemove = cellsToMerge.stream()
                .filter(cell -> cell.getColumn() >= config.getStartCol() && cell.getColumn() <= config.getEndCol())
                .collect(Collectors.toList());

        result.removeAll(cellsToRemove);
        result.add(mergedCell);

        return result;
    }

    /**
     * 竖向合并单元格
     */
    private static List<JCell> mergeVertically(List<JCell> allCells, JTableMergeConfig config) {
        List<JCell> result = new ArrayList<>(allCells);

        // 找到要合并的单元格
        List<JCell> cellsToMerge = allCells.stream()
                .filter(cell -> cell.getColumn() == config.getStartCol())
                .filter(cell -> cell.getRow() >= config.getStartRow() && cell.getRow() <= config.getEndRow())
                .sorted(Comparator.comparingInt(JCell::getRow))
                .collect(Collectors.toList());

        if (cellsToMerge.size() < 2) {
            return result;
        }
        List<String> contentList = cellsToMerge.stream()
                .map(JCell::getContent)
                .filter(StringUtils::isNotBlank)
                .collect(Collectors.toList());
        String content = String.join("\n", contentList);
        // 创建合并后的单元格
        JCell mergedCell = new JCell();
        mergedCell.setRow(config.getStartRow());
        mergedCell.setColumn(config.getStartCol());
        mergedCell.setX(cellsToMerge.get(0).getX());
        mergedCell.setY(cellsToMerge.get(0).getY());
        mergedCell.setWidth(cellsToMerge.get(0).getWidth());
        mergedCell.setHeight(cellsToMerge.stream().mapToInt(JCell::getHeight).sum());
        mergedCell.setContent(content); // 使用第一个单元格的内容
        mergedCell.setTextAlign(cellsToMerge.get(0).isTextAlign());
        mergedCell.setBackgroundColor(cellsToMerge.get(0).getBackgroundColor());
        mergedCell.setTextColor(cellsToMerge.get(0).getTextColor());

        // 移除被合并的单元格
        List<JCell> cellsToRemove = cellsToMerge.stream()
                .filter(cell -> cell.getRow() >= config.getStartRow() && cell.getRow() <= config.getEndRow())
                .collect(Collectors.toList());

        result.removeAll(cellsToRemove);
        result.add(mergedCell);

        return result;
    }

    /**
     * 获取表头最后一行的所有单元格
     *
     * @param cells 表头所有的单元格
     */
    public static List<JCell> findHeadColumns(List<JCell> cells) {
        Map<Integer, List<JCell>> rows = cells.stream().collect(Collectors.groupingBy(JCell::getRow));
        //取第一行的所有单元格
        List<JCell> firstRow = rows.get(1);
        if (CollectionUtils.isEmpty(firstRow)) {
            throw new RuntimeException("缺少第一行");
        }
        // 按列升序排序
        List<JCell> perRowCells =
                firstRow.stream().sorted(Comparator.comparing(JCell::getColumn)).collect(Collectors.toList());
        List<JCell> headColumns = new ArrayList<>();
        for (JCell cell : perRowCells) {
            // 获取每个单元格下的所有从属单元格，如果没有则返回其本身
            headColumns.addAll(getCellSonCells(cells, cell));
        }
        return headColumns;
    }

    /**
     * 获取所有的从属单元格
     *
     * @param cells 所有表头单元格
     * @param cell 要获取从属的单元格
     * @return 该单元格下的从属单元格
     */
    public static List<JCell> getCellSonCells(List<JCell> cells, JCell cell) {
        // 从属单元格：该单元格所在列中的所有单元格
        List<JCell> collect = cells.stream()
                .filter(cel -> cel.getRow() == cell.getRow() + 1)
                .filter(cel -> cel.getBelongColumn() == cell.getColumn())
                .collect(Collectors.toList());
        // 没有从属单元格则返回本身
        if (CollectionUtils.isEmpty(collect)) {
            ArrayList<JCell> objects = new ArrayList<>();
            objects.add(cell);
            return objects;
        }
        //有从属单元格则遍历每一个单元格来获取每个单元格下的从属单元格
        List<List<JCell>> allSonCells =
                collect.stream().map(s -> getCellSonCells(cells, s)).collect(Collectors.toList());
        return allSonCells.stream().flatMap(Collection::stream).collect(Collectors.toList());
    }
}
