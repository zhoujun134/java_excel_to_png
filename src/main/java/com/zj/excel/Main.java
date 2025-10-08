package com.zj.excel;

import cn.hutool.core.io.FileUtil;
import com.zj.excel.to.image.ExcelDrawUtil;
import com.zj.excel.to.image.dto.ExcelDrawImageRequest;
import lombok.extern.slf4j.Slf4j;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.BufferedInputStream;
import java.io.File;
import java.io.IOException;
import java.util.Map;

/**
 * @Author zj
 * @Description
 * @Date 2025/9/29 22:19
 * @Version v1.0
 **///TIP To <b>Run</b> code, press <shortcut actionId="Run"/> or
// click the <icon src="AllIcons.Actions.Execute"/> icon in the gutter.
@Slf4j
public class Main {
    public static void main(String[] args) throws IOException {
        String excelPath = "/Users/zj/Desktop/test_excel/课程信息表4.xlsx";
        BufferedInputStream excelInputStream = FileUtil.getInputStream(excelPath);
        ExcelDrawImageRequest request = new ExcelDrawImageRequest()
                .setExcelStream(excelInputStream)
                .setNeedHeader(true)
                .setDefaultColumnLength(2)
                .setDefaultRowLength(2)
                .setFileTypeEnum(FileTypeEnum.XLSX);
        Map<String, BufferedImage> imageMap = ExcelDrawUtil.excelToPngWithColor(request);
        imageMap.forEach((imageName, image) -> {
            System.out.println("imageName: " + imageName);
            System.out.println("image: " + image);
            try {
                ImageIO.write(image, "png", new File("/Users/zj/Desktop/test_excel/image/" + imageName));
            } catch (IOException e) {
                throw new RuntimeException(e);
            }
        });
    }
}

