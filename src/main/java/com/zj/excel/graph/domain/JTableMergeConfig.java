package com.zj.excel.graph.domain;

import lombok.Data;
import lombok.experimental.Accessors;

/**
 * 合并配置
 */
@Data
@Accessors(chain = true)
public class JTableMergeConfig {
    private int startRow;
    private int endRow;
    private int startCol;
    private int endCol;
    private boolean horizontal; // true为横向合并，false为竖向合并

    public JTableMergeConfig(int startRow, int endRow, int startCol, int endCol, boolean horizontal) {
        this.startRow = startRow;
        this.endRow = endRow;
        this.startCol = startCol;
        this.endCol = endCol;
        this.horizontal = horizontal;
    }
}
