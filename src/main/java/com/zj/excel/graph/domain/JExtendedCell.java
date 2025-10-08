package com.zj.excel.graph.domain;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import lombok.experimental.Accessors;

import java.awt.*;

/**
 * @ClassName JExtendedCell
 * @Author zj
 * @Description
 * @Date 2025/10/1 15:57
 * @Version v1.0
 **/
@Data
@Accessors(chain = true)
@AllArgsConstructor
@NoArgsConstructor
public class JExtendedCell {

    private String content;

    private Color backgroundColor;

    private Color textColor;

    public JExtendedCell(String content) {
        this.content = content;
    }
}
