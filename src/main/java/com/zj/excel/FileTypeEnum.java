package com.zj.excel;

import org.apache.commons.lang3.StringUtils;

import java.util.Arrays;

/**
 * @ClassName FileTypeEnum
 * @Author zj
 * @Description
 * @Date 2025/10/1 15:54
 * @Version v1.0
 **/
public enum FileTypeEnum {
    /**
     * pdf 文件
     */
    PDF,
    /**
     * excel xlsx 文件
     */
    XLSX,
    /**
     * excel xls 文件
     */
    XLS,
    /**
     * word 文件
     */
    WORD,
    /**
     * 其他文件
     */
    OTHER;

    public static FileTypeEnum getFileType(String fileTypeName) {
        if (StringUtils.isBlank(fileTypeName)) {
            return FileTypeEnum.OTHER;
        }
        return Arrays.stream(FileTypeEnum.values())
                .filter(fileTypeEnum -> StringUtils.equals(fileTypeEnum.name(), fileTypeName)).findFirst()
                .orElse(FileTypeEnum.OTHER);
    }

    public static FileTypeEnum getFileTypeByFileName(String fileName) {
        FileTypeEnum fileType = FileTypeEnum.OTHER;
        if (StringUtils.isBlank(fileName)) {
            return fileType;
        }
        if (fileName.endsWith(".xlsx")) {
            fileType = FileTypeEnum.XLSX;
        } else if (fileName.endsWith(".xls")) {
            fileType = FileTypeEnum.XLS;
        } else if (fileName.endsWith(".pdf")) {
            fileType = FileTypeEnum.PDF;
        } else if (fileName.endsWith(".docx")) {
            fileType = FileTypeEnum.WORD;
        }
        return fileType;
    }
}
