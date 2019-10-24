package com.github.ckpoint.toexcel.core.converter;

import com.github.ckpoint.toexcel.annotation.ExcelHeader;

public class ExcelHeaderDefatulConverter implements ExcelHeaderConverter {

    @Override
    public String headerKeyConverter(ExcelHeader excelHeader) {
        return excelHeader.headerName();
    }
}
