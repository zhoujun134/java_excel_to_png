package com.zj.excel.to.image.dto;

import com.zj.excel.FileTypeEnum;
import com.zj.excel.ai.AiInvokeUtils;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import lombok.experimental.Accessors;

import java.io.InputStream;

/**
 * @author zhoujun134
 * Created on 2025-09-30
 */
@Data
@AllArgsConstructor
@NoArgsConstructor
@Accessors(chain = true)
public class ExcelDrawImageRequest {

    /*
     * 表格的默认行数
     */
    private int defaultRowLength = 100;

    /*
     * 表格的默认列数
     */
    private int defaultColumnLength = 10;

    /**
     * excel 文件流
     */
    private InputStream excelStream;

    /**
     * 文件类型
     * {@link FileTypeEnum#XLS} or {@link FileTypeEnum#XLSX}
     */
    private FileTypeEnum fileTypeEnum = FileTypeEnum.XLSX;

    /**
     * 是否需要表头
     * <p> 每张图片带上表头需要 ai 能力的支持。{@link AiInvokeUtils#setAiConfig } aiConfig 中，配置相关的 api url, apiKey 和 调用的 model 即可实现这个功能 </p>
     */
    private boolean needHeader = false;
}
