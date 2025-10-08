package com.zj.excel.domian.dto;

import lombok.Data;

import java.util.ArrayList;
import java.util.List;

@Data
public class RowIndexInfoDTO {
    private List<Integer> headerRowIndexList = new ArrayList<>();
    private Integer dataRowStartIndex = 0;
    private List<Integer> otherRowIndexList = new ArrayList<>();
}
