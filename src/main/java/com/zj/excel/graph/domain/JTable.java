package com.zj.excel.graph.domain;

import lombok.Data;
import lombok.experimental.Accessors;

import java.awt.*;
import java.util.ArrayList;
import java.util.List;

/**
 * 表样属性
 */
@Data
@Accessors(chain = true)
public class JTable {
    /**
     * 表头背景色
     */
    private Color headerBackGroundColor;
    /**
     * 表头字体
     */
    private Font headerFont;
    /**
     * 单元格字体
     */
    private Font cellFont;
    /**
     * 表头单元格
     */
    private java.util.List<JCell> headCells;
    /**
     * 合并配置
     */
    private List<JTableMergeConfig> mergeConfigs = new ArrayList<>();
    /**
     * 行高
     */
    private int rowHeight = 30;
    /**
     * 上边距
     */
    private int marginY = 10;
    /***
     * 左边距
     */
    private int marginX = 10;
    /**
     * 默认生成图片的 dpi
     */
    private int imageDpi = 72;
}
