package com.zj.excel.graph.domain;

import com.zj.excel.graph.JDrawTableUtil;
import lombok.Data;
import lombok.SneakyThrows;
import lombok.experimental.Accessors;

import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.image.BufferedImage;
import java.io.File;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.stream.Collectors;
import java.util.stream.Stream;

/**
 * 单元格属性
 */
@Data
@Accessors(chain = true)
public class JCell {
    /**
     * 行
     */
    private int row;
    /**
     * 列
     */
    private int column;
    /**
     * x 坐标
     */
    private int x;
    /**
     * y 坐标
     */
    private int y;
    /**
     * 当前单元格的父级单元格所在的列
     */
    private int belongColumn;
    /**
     * 单元格宽度,对于有从属单元格的，其宽度由从属单元格决定
     */
    private int width;
    /**
     * 合并单元格的情况会用到，默认为行高
     */
    private int height;
    /**
     * 单元格内容
     */
    private String content;
    /**
     * 是否水平居中
     */
    private boolean textAlign;

    private Color backgroundColor;

    private Color textColor;


    public JCell(int row, int column, int width, int belongColumn) {
        this.row = row;
        this.column = column;
        this.width = width;
        this.belongColumn = belongColumn;
    }

    public JCell() {
    }

    @Override
    public boolean equals(Object o) {
        if (this == o) {
            return true;
        }
        if (o == null || getClass() != o.getClass()) {
            return false;
        }
        JCell cell = (JCell) o;
        return row == cell.row && column == cell.column;
    }

    @Override
    public int hashCode() {
        return Objects.hash(row, column);
    }

    public static void main(String[] args) {
        style0(); // 有表头
        style1(); // 有表头，带合并
        style2(); // 无表头
        style3(); // 无表头，带合并
    }

    @SneakyThrows
    private static void style0() {
        List<JCell> headCells = new ArrayList<>();
        headCells.add(new JCell(1, 1, 100, 1).setTextAlign(true).setContent("姓名"));
        headCells.add(new JCell(1, 2, 100, 1).setTextAlign(true).setContent("年龄"));
        headCells.add(new JCell(1, 3, 100, 1).setTextAlign(true).setContent("城市"));
        List<List<String>> tableRowContents = new ArrayList<>();
        List<String> a = Stream.of("张三", "25", "北京").collect(Collectors.toList());
        List<String> b = Stream.of("李四", "30", "上海").collect(Collectors.toList());
        List<String> c = Stream.of("王五", "28", "广州").collect(Collectors.toList());
        tableRowContents.add(a);
        tableRowContents.add(b);
        tableRowContents.add(c);
        String path = "0_with_header.png";
        JTable table = new JTable()
                .setCellFont(new Font("宋体", Font.PLAIN, 18))
                .setHeadCells(headCells)
                .setHeaderFont(new Font("宋体", Font.BOLD, 24))
                .setHeaderBackGroundColor(Color.gray)
                .setRowHeight(50);
        BufferedImage image = JDrawTableUtil.drawTable(table, tableRowContents);
        ImageIO.write(image, "png", new File(path));
    }

    @SneakyThrows
    private static void style1() {
        List<JCell> headCells = new ArrayList<>();
        headCells.add(new JCell(1, 1, 100, 0).setTextAlign(true).setContent("信息"));
        headCells.add(new JCell(2, 1, 100, 1).setTextAlign(true).setContent("姓名"));
        headCells.add(new JCell(2, 2, 100, 1).setTextAlign(true).setContent("年龄"));
        headCells.add(new JCell(2, 3, 100, 1).setTextAlign(true).setContent("城市"));
        List<List<String>> tableRowContents = new ArrayList<>();
        List<String> a = Stream.of("张三", "25", "北京").collect(Collectors.toList());
        List<String> b = Stream.of("李四", "30", "上海").collect(Collectors.toList());
        List<String> c = Stream.of("王五", "28", "广州").collect(Collectors.toList());
        tableRowContents.add(a);
        tableRowContents.add(b);
        tableRowContents.add(c);
        List<JTableMergeConfig> mergeConfigs = new ArrayList<>();
        mergeConfigs.add(new JTableMergeConfig(2, 2, 1, 2, true)); // 横向合并第一行的第1-2列
        String path = "1_with_header_merge.png";
        JTable table = new JTable().setCellFont(new Font("宋体", Font.PLAIN, 10))
                .setHeadCells(headCells).setHeaderFont(new Font("宋体", Font.BOLD, 15))
                .setHeaderBackGroundColor(Color.gray)
                .setMergeConfigs(mergeConfigs);
        BufferedImage image = JDrawTableUtil.drawTable(table, tableRowContents);
        ImageIO.write(image, "png", new File(path));
    }

    @SneakyThrows
    private static void style2() {
        List<List<String>> tableRowContents = new ArrayList<>();
        List<String> a = Stream.of("张三", "25", "北京", "技术部").collect(Collectors.toList());
        List<String> b = Stream.of("李四", "30", "上海", "销售部").collect(Collectors.toList());
        List<String> c = Stream.of("王五", "28", "广州", "财务部").collect(Collectors.toList());
        tableRowContents.add(a);
        tableRowContents.add(b);
        tableRowContents.add(c);
        String path = "2_without_header.png";
        JTable table = new JTable().setCellFont(new Font("宋体", Font.PLAIN, 12))
                .setHeaderFont(new Font("宋体", Font.BOLD, 15))
                .setHeaderBackGroundColor(Color.gray)
                .setRowHeight(40);
        BufferedImage image = JDrawTableUtil.drawTable(table, tableRowContents);
        ImageIO.write(image, "png", new File(path));
    }

    @SneakyThrows
    private static void style3() {
        List<List<String>> tableRowContents = new ArrayList<>();
        List<String> a = Stream.of("张三", "25", "北京", "技术部").collect(Collectors.toList());
        List<String> b = Stream.of("文伟", "26", "北京", "销售部").collect(Collectors.toList()); // 张三合并
        List<String> c = Stream.of("李四", "30", "上海", "财务部").collect(Collectors.toList());
        List<String> d = Stream.of("", "31", "上海", "人事部").collect(Collectors.toList()); // 李四合并

        tableRowContents.add(a);
        tableRowContents.add(b);
        tableRowContents.add(c);
        tableRowContents.add(d);

        List<JTableMergeConfig> mergeConfigs = new ArrayList<>();
        // 竖向合并姓名列的张三和李四
        mergeConfigs.add(new JTableMergeConfig(1, 2, 1, 1, false)); // 张三竖向合并
        mergeConfigs.add(new JTableMergeConfig(3, 4, 1, 1, false)); // 李四竖向合并

        String path = "3_without_header_merge.png";
        JTable table = new JTable().setCellFont(new Font("宋体", Font.PLAIN, 16))
                .setHeaderFont(new Font("宋体", Font.BOLD, 18))
                .setHeaderBackGroundColor(Color.gray)
                .setMergeConfigs(mergeConfigs)
                .setRowHeight(40);
        BufferedImage image = JDrawTableUtil.drawTable(table, tableRowContents);
        ImageIO.write(image, "png", new File(path));
    }
}
