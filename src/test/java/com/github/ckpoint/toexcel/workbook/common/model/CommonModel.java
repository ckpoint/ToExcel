package com.github.ckpoint.toexcel.workbook.common.model;

import com.github.ckpoint.toexcel.annotation.ExcelHeader;
import lombok.Data;

@Data
public class CommonModel {

    @ExcelHeader(headerName = "no", priority = 100)
    private String id;
}
